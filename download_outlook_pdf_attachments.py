#!/usr/bin/env python3
"""Download PDF attachments from Outlook via Microsoft Graph.

This works cross-platform because it talks to the mailbox through Graph instead of
using the local Outlook desktop COM API.

Supported mail sources:
- Outlook desktop mailboxes connected to Microsoft 365 or Exchange Online
- Outlook on the web mailboxes exposed through Microsoft Graph

Authentication modes:
- Device code flow for interactive sign-in
- Client credentials for app-only access

Examples:
- Interactive sign-in, Inbox, save into ./attachments:
  python download_outlook_pdf_attachments.py --account user@company.com

- Specific folder path:
  python download_outlook_pdf_attachments.py --account user@company.com --folder "Inbox/Invoices"

- Application permissions using client credentials:
  python download_outlook_pdf_attachments.py --tenant-id ... --client-id ... --client-secret ... --account user@company.com
"""

from __future__ import annotations

import argparse
import base64
import re
import socket
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable, Iterator
from urllib.parse import quote

import requests
import urllib3.util.connection

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["Mail.Read"]
DEFAULT_PAGE_SIZE = 100
PDF_CONTENT_TYPE = "application/pdf"
BODY_URL_KEYWORDS = ["www.ediarchive.eu", "www.szamlazz.hu"]
URL_FILTER_PATTERNS = [
    r'https?://(?:[^\s]*\.)?ediarchive\.eu/[^\s]*/downloadurl[^\s]*',
    r'https?://(?:[^\s]*\.)?szamlazz\.hu/szamla/fiok/[^\s]*',
]


@dataclass(frozen=True)
class AuthConfig:
    tenant_id: str | None
    client_id: str | None
    client_secret: str | None
    use_device_code: bool
    scopes: list[str]


class GraphError(RuntimeError):
    pass


def force_ipv4_only() -> None:
    urllib3.util.connection.allowed_gai_family = lambda: socket.AF_INET


def sanitize_filename(value: str, fallback: str = "untitled") -> str:
    cleaned = re.sub(r'[<>:"/|?*\x00-\x1f]', "_", value).strip(" ._")
    return cleaned or fallback


def unique_path(directory: Path, stem: str, suffix: str) -> Path:
    candidate = directory / f"{stem}{suffix}"
    index = 1
    while candidate.exists():
        candidate = directory / f"{stem}_{index}{suffix}"
        index += 1
    return candidate


def normalize_folder_path(folder_path: str | None) -> list[str]:
    if not folder_path:
        return ["inbox"]
    return [part for part in re.split(r"[\\/]+", folder_path.strip()) if part]


def graph_get(session: requests.Session, url: str, *, params: dict | None = None) -> dict:
    response = session.get(url, params=params, timeout=60)
    if response.status_code >= 400:
        raise GraphError(f"Graph request failed ({response.status_code}): {response.text}")
    return response.json()


def graph_post(session: requests.Session, url: str, *, data: dict | None = None, json_body: dict | None = None) -> dict:
    response = session.post(url, data=data, json=json_body, timeout=60)
    if response.status_code >= 400:
        raise GraphError(f"Graph request failed ({response.status_code}): {response.text}")
    return response.json()


def acquire_token(auth: AuthConfig) -> str:
    if auth.client_id and auth.client_secret:
        if not auth.tenant_id:
            raise SystemExit("--tenant-id is required with --client-id and --client-secret")
        token_url = f"https://login.microsoftonline.com/{auth.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": auth.client_id,
            "client_secret": auth.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }
        response = requests.post(token_url, data=data, timeout=60)
        if response.status_code >= 400:
            raise SystemExit(f"Token request failed ({response.status_code}): {response.text}")
        return response.json()["access_token"]

    if not auth.client_id:
        raise SystemExit("--client-id is required for device code flow")

    if not auth.tenant_id:
        tenant = "common"
    else:
        tenant = auth.tenant_id

    device_code_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
    token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    scope = " ".join(auth.scopes)

    device = requests.post(device_code_url, data={"client_id": auth.client_id, "scope": scope}, timeout=60)
    if device.status_code >= 400:
        raise SystemExit(f"Device code request failed ({device.status_code}): {device.text}")

    payload = device.json()
    print(payload["message"])

    data = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": auth.client_id,
        "device_code": payload["device_code"],
    }

    while True:
        token = requests.post(token_url, data=data, timeout=60)
        if token.status_code == 200:
            return token.json()["access_token"]

        error = token.json().get("error")
        if error in {"authorization_pending", "slow_down"}:
            continue
        raise SystemExit(f"Device sign-in failed: {token.text}")


def build_session(access_token: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    })
    return session


def get_user_root(session: requests.Session, account: str | None) -> str:
    if account:
        return f"{GRAPH_ROOT}/users/{quote(account, safe='')}"
    return f"{GRAPH_ROOT}/me"


def resolve_mail_folder(session: requests.Session, user_root: str, folder_path: str | None) -> str:
    parts = normalize_folder_path(folder_path)
    if not parts:
        return f"{user_root}/mailFolders/inbox"

    first = parts[0].lower()
    folder_id_map = {
        "inbox": "inbox",
        "sent items": "sentitems",
        "deleted items": "deleteditems",
        "drafts": "drafts",
        "archive": "archive",
        "junk email": "junkemail",
    }

    if first in folder_id_map:
        current_url = f"{user_root}/mailFolders/{folder_id_map[first]}"
    else:
        current_url = f"{user_root}/mailFolders"
        folders = graph_get(session, current_url).get("value", [])
        match = next((item for item in folders if item.get("displayName", "").lower() == parts[0].lower()), None)
        if not match:
            raise SystemExit(f"Could not find top-level folder: {parts[0]}")
        current_url = match["@odata.id"] if "@odata.id" in match else f"{user_root}/mailFolders/{match['id']}"

    for part in parts[1:]:
        children = graph_get(session, f"{current_url}/childFolders").get("value", [])
        match = next((item for item in children if item.get("displayName", "").lower() == part.lower()), None)
        if not match:
            raise SystemExit(f"Could not find subfolder: {part}")
        current_url = f"{current_url}/childFolders/{match['id']}"

    return current_url


def iter_message_pages(session: requests.Session, folder_url: str, lookback_days: int) -> Iterator[dict]:
    days_ago = (datetime.now(timezone.utc) - timedelta(days=lookback_days)).isoformat()
    filter_param = f"receivedDateTime ge {days_ago}"
    next_url = f"{folder_url}/messages?$top={DEFAULT_PAGE_SIZE}&$select=id,subject,receivedDateTime,hasAttachments,body&$filter={quote(filter_param, safe='')}"
    while next_url:
        payload = graph_get(session, next_url)
        for message in payload.get("value", []):
            # Filter by keywords in body content or hasAttachments
            body_content = (message.get("body", {}).get("content") or "").lower()
            has_keyword = any(kw.lower() in body_content for kw in BODY_URL_KEYWORDS)
            if message.get("hasAttachments") or has_keyword:
                yield message
        next_url = payload.get("@odata.nextLink")


def iter_folder_tree(session: requests.Session, folder_url: str) -> Iterator[str]:
    yield folder_url
    children = graph_get(session, f"{folder_url}/childFolders?$select=id,displayName").get("value", [])
    for child in children:
        child_url = f"{folder_url}/childFolders/{child['id']}"
        yield from iter_folder_tree(session, child_url)


def list_attachments(session: requests.Session, message_url: str) -> list[dict]:
    payload = graph_get(session, f"{message_url}/attachments?$select=id,name,contentType,size,isInline")
    return payload.get("value", [])


def parse_received_time(value: str | None) -> str:
    if not value:
        return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return dt.astimezone(timezone.utc).strftime("%Y%m%d_%H%M%S")


def decode_file_attachment(attachment: dict) -> bytes:
    content_bytes = attachment.get("contentBytes")
    if content_bytes:
        return base64.b64decode(content_bytes)
    raise GraphError("Attachment payload did not include contentBytes")


def extract_matching_links(body_content: str) -> list[str]:
    """Extract URLs from body content that contain keywords or match specific patterns."""
    url_pattern = r'https?://[^\s<>"{}|\\^`\[\]]*'
    urls = re.findall(url_pattern, body_content)
    

    matching_urls = []
    
    for url in urls:
        # Check if URL matches any regex pattern
        if any(re.match(pattern, url) for pattern in URL_FILTER_PATTERNS):
            matching_urls.append(url)
    
    return list(set(matching_urls))  # Remove duplicates


def download_pdf_attachments(
    session: requests.Session,
    folder_url: str,
    output_dir: Path,
    recursive: bool,
    lookback_days: int,
) -> int:
    output_dir.mkdir(parents=True, exist_ok=True)
    saved = 0

    folder_urls: Iterable[str] = iter_folder_tree(session, folder_url) if recursive else [folder_url]
    for current_folder_url in folder_urls:
        for message in iter_message_pages(session, current_folder_url, lookback_days):
            if not message.get("hasAttachments"):
                body_content = message.get("body", {}).get("content") or ""
                matching_links = extract_matching_links(body_content)
                if matching_links:
                    subject = sanitize_filename(message.get("subject") or "message")
                    timestamp = parse_received_time(message.get("receivedDateTime"))
                    
                    for link in matching_links:
                        # Check if link is ediarchive download URL
                        if any(re.match(pattern, link) for pattern in URL_FILTER_PATTERNS[:1]):  # First pattern is ediarchive
                            try:
                                response = session.get(link, timeout=60)
                                if response.status_code == 200:
                                    file_name = sanitize_filename("invoice")
                                    stem = f"{timestamp}_{subject}_{file_name}"
                                    destination = unique_path(output_dir, stem, ".pdf")
                                    destination.write_bytes(response.content)
                                    saved += 1
                                    print(f"Saved: {destination}")
                            except Exception as e:
                                print(f"Failed to download from {link}: {e}")
                        
                continue
                
            subject = sanitize_filename(message.get("subject") or "message")
            timestamp = parse_received_time(message.get("receivedDateTime"))
            message_url = f"{current_folder_url}/messages/{message['id']}"
            attachments = list_attachments(session, message_url)

            for index, attachment in enumerate(attachments, start=1):
                name = attachment.get("name") or "attachment"
                content_type = (attachment.get("contentType") or "").lower()
                is_pdf_name = name.lower().endswith(".pdf")
                is_pdf_type = content_type == PDF_CONTENT_TYPE
                if not (is_pdf_name or is_pdf_type):
                    continue

                file_name = sanitize_filename(Path(name).stem)
                stem = f"{timestamp}_{subject}_{index}_{file_name}"
                destination = unique_path(output_dir, stem, ".pdf")

                if attachment.get("contentBytes"):
                    destination.write_bytes(decode_file_attachment(attachment))
                else:
                    attachment_payload = graph_get(session, f"{message_url}/attachments/{attachment['id']}")
                    if attachment_payload.get("@odata.type") == "#microsoft.graph.fileAttachment":
                        destination.write_bytes(decode_file_attachment(attachment_payload))
                    else:
                        continue

                saved += 1
                print(f"Saved: {destination}")

    return saved


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Download PDF attachments from Outlook mailboxes using Microsoft Graph."
    )
    parser.add_argument("--account", help="Mailbox address to read, e.g. user@company.com. Defaults to /me with device sign-in.")
    parser.add_argument("--folder", default="Inbox", help="Mail folder path, e.g. Inbox or Inbox/Invoices")
    parser.add_argument("--output", default="invoices", help="Directory where PDFs will be saved")
    parser.add_argument("--recursive", action="store_true", help="Scan subfolders recursively")
    parser.add_argument("--tenant-id", help="Microsoft Entra tenant id for authentication")
    parser.add_argument("--client-id", help="App registration client id")
    parser.add_argument("--client-secret", help="App registration client secret for client credentials")
    parser.add_argument("--scope", action="append", dest="scopes", help="OAuth scope for device flow. Can be repeated. Defaults to Mail.Read.")
    parser.add_argument("--days", type=int, default=7, help="Number of days to look back for emails")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    # Ensure days is at least 1
    if args.days < 1:
        print("Number of days must be at least 1")
        return 1
    output_dir = Path(args.output).expanduser().resolve()

    force_ipv4_only()

    auth = AuthConfig(
        tenant_id=args.tenant_id,
        client_id=args.client_id,
        client_secret=args.client_secret,
        use_device_code=not bool(args.client_secret),
        scopes=args.scopes or DEFAULT_SCOPES,
    )

    access_token = acquire_token(auth)
    session = build_session(access_token)
    user_root = get_user_root(session, args.account)
    folder_url = resolve_mail_folder(session, user_root, args.folder)

    saved = download_pdf_attachments(session, folder_url, output_dir, args.recursive, args.days)
    print(f"Done. Saved {saved} PDF attachment(s) to {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
