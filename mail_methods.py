import requests
import re
import base64
import html
from pathlib import Path
from typing import Iterable, Iterator
from urllib.parse import parse_qs, quote, urljoin, urlparse
from datetime import datetime, timedelta, timezone

DEFAULT_PAGE_SIZE = 100
PDF_CONTENT_TYPE = "application/pdf"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SZAMLAZZ_ROOT = "https://www.szamlazz.hu"
BODY_URL_KEYWORDS = ["www.ediarchive.eu", "www.szamlazz.hu"]
URL_FILTER_PATTERNS = [
    r'https?://(?:[^\s]*\.)?ediarchive\.eu/[^\s]*/downloadurl[^\s]*',
    r'https?://(?:[^\s]*\.)?szamlazz\.hu/szamla(?:/fiok/[^\s]*|/\?action=szamlapdf[^\s]*)',
]

class GraphError(RuntimeError):
    pass


class MailMethods:
    def __init__(self, session: requests.Session, account: str | None, folder: str | None = "Inbox", lookback_days: int = 7, recursive: bool = False, output_dir: str | None = "attachments"):
        self.session = session
        self.account = account
        self.folder = folder
        self.lookback_days = lookback_days
        self.recursive = recursive
        self.output_dir = output_dir
        self.user_root = self.get_user_root()
        self.folder_url = self.resolve_mail_folder()
        self.saved_attachments = self.download_pdf_attachments()
    
    def get_user_root(self) -> str:
        if self.account:
            return f"{GRAPH_ROOT}/users/{quote(self.account, safe='')}"
        return f"{GRAPH_ROOT}/me"
    
    
    def resolve_mail_folder(self) -> str:
        parts = self.normalize_folder_path(self.folder)
        if not parts:
            return f"{self.user_root}/mailFolders/inbox"

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
            current_url = f"{self.user_root}/mailFolders/{folder_id_map[first]}"
        else:
            current_url = f"{self.user_root}/mailFolders"
            folders = self.graph_get(current_url).get("value", [])
            match = next((item for item in folders if item.get("displayName", "").lower() == parts[0].lower()), None)
            if not match:
                raise SystemExit(f"Could not find top-level folder: {parts[0]}")
            current_url = match["@odata.id"] if "@odata.id" in match else f"{self.user_root}/mailFolders/{match['id']}"

        for part in parts[1:]:
            children = self.graph_get(f"{current_url}/childFolders").get("value", [])
            match = next((item for item in children if item.get("displayName", "").lower() == part.lower()), None)
            if not match:
                raise SystemExit(f"Could not find subfolder: {part}")
            current_url = f"{current_url}/childFolders/{match['id']}"

        return current_url
    
    
    def normalize_folder_path(self, folder_path: str | None) -> list[str]:
        if not folder_path:
            return ["inbox"]
        return [part for part in re.split(r"[\\/]+", folder_path.strip()) if part]


    def graph_get(self, url: str, *, params: dict | None = None) -> dict:
        response = self.session.get(url, params=params, timeout=60)
        if response.status_code >= 400:
            raise GraphError(f"Graph request failed ({response.status_code}): {response.text}")
        return response.json()


    def graph_post(self, url: str, *, data: dict | None = None, json_body: dict | None = None) -> dict:
        response = self.session.post(url, data=data, json=json_body, timeout=60)
        if response.status_code >= 400:
            raise GraphError(f"Graph request failed ({response.status_code}): {response.text}")
        return response.json()
    
    
    def download_pdf_attachments(self) -> int:
        output_dir = Path(self.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        saved = 0

        folder_urls: Iterable[str] = self.iter_folder_tree(self.folder_url) if self.recursive else [self.folder_url]
        for current_folder_url in folder_urls:
            for message in self.iter_message_pages(current_folder_url):
                if not message.get("hasAttachments"):
                    body_content = message.get("body", {}).get("content") or ""
                    matching_links = self.extract_matching_links(body_content)
                    if matching_links:
                        subject = self.sanitize_filename(message.get("subject") or "message")
                        timestamp = self.parse_received_time(message.get("receivedDateTime"))
                        
                        for link in matching_links:
                            # Check if link matches first pattern (ediarchive - direct download)
                            if re.match(URL_FILTER_PATTERNS[0], link):
                                try:
                                    response = self.session.get(link, timeout=60)
                                    if response.status_code == 200:
                                        file_name = self.sanitize_filename("invoice")
                                        stem = f"{timestamp}_{subject}_{file_name}"
                                        destination = self.unique_path(output_dir, stem, ".pdf")
                                        destination.write_bytes(response.content)
                                        saved += 1
                                        print(f"Saved: {destination}")
                                except Exception as e:
                                    print(f"Failed to download from {link}: {e}")
                            
                            # Check if link matches second pattern (szamlazz - redirect to page with link)
                            elif re.match(URL_FILTER_PATTERNS[1], link):
                                try:
                                    # Follow redirect to get the page containing the actual PDF link
                                    response = self.session.get(link, timeout=60, allow_redirects=True)
                                    if response.status_code == 200:
                                        # Extract links from the redirected page content
                                        redirected_links = self.extract_matching_links(response.text)
                                        for pdf_link in redirected_links:
                                            if pdf_link != link and self.is_attachment_download_url(pdf_link):
                                                try:
                                                    pdf_response = self.session.get(pdf_link, timeout=60)
                                                    if pdf_response.status_code == 200:
                                                        file_name = self.sanitize_filename("invoice")
                                                        stem = f"{timestamp}_{subject}_{file_name}"
                                                        destination = self.unique_path(output_dir, stem, ".pdf")
                                                        destination.write_bytes(pdf_response.content)
                                                        saved += 1
                                                        print(f"Saved: {destination}")
                                                        break
                                                except Exception as e:
                                                    print(f"Failed to download from redirected link {pdf_link}: {e}")
                                except Exception as e:
                                    print(f"Failed to follow redirect from {link}: {e}")
                            
                    continue
                    
                subject = self.sanitize_filename(message.get("subject") or "message")
                timestamp = self.parse_received_time(message.get("receivedDateTime"))
                message_url = f"{current_folder_url}/messages/{message['id']}"
                attachments = self.list_attachments(message_url)

                for index, attachment in enumerate(attachments, start=1):
                    name = attachment.get("name") or "attachment"
                    content_type = (attachment.get("contentType") or "").lower()
                    is_pdf_name = name.lower().endswith(".pdf")
                    is_pdf_type = content_type == PDF_CONTENT_TYPE
                    if not (is_pdf_name or is_pdf_type):
                        continue

                    file_name = self.sanitize_filename(Path(name).stem)
                    stem = f"{timestamp}_{subject}_{index}_{file_name}"
                    destination = self.unique_path(output_dir, stem, ".pdf")

                    if attachment.get("contentBytes"):
                        destination.write_bytes(self.decode_file_attachment(attachment))
                    else:
                        attachment_payload = self.graph_get(f"{message_url}/attachments/{attachment['id']}")
                        if attachment_payload.get("@odata.type") == "#microsoft.graph.fileAttachment":
                            destination.write_bytes(self.decode_file_attachment(attachment_payload))
                        else:
                            continue

                    saved += 1
                    print(f"Saved: {destination}")

        return saved
    
    
    def iter_folder_tree(self, folder_url: str) -> Iterator[str]:
        yield folder_url
        children = self.graph_get(f"{folder_url}/childFolders?$select=id,displayName").get("value", [])
        for child in children:
            child_url = f"{folder_url}/childFolders/{child['id']}"
            yield from self.iter_folder_tree(child_url)
            
            
    def iter_message_pages(self, folder_url: str) -> Iterator[dict]:
        days_ago = (datetime.now(timezone.utc) - timedelta(days=self.lookback_days)).isoformat()
        filter_param = f"receivedDateTime ge {days_ago}"
        next_url = f"{folder_url}/messages?$top={DEFAULT_PAGE_SIZE}&$select=id,subject,receivedDateTime,hasAttachments,body&$filter={quote(filter_param, safe='')}"
        while next_url:
            payload = self.graph_get(next_url)
            for message in payload.get("value", []):
                # Filter by keywords in body content or hasAttachments
                body_content = message.get("body", {}).get("content") or ""
                lowered_body_content = body_content.lower()
                has_keyword = any(kw.lower() in lowered_body_content for kw in BODY_URL_KEYWORDS)
                has_matching_link = bool(self.extract_matching_links(body_content))
                if message.get("hasAttachments") or has_keyword or has_matching_link:
                    yield message
            next_url = payload.get("@odata.nextLink")
            
            
    def list_attachments(self, message_url: str) -> list[dict]:
        payload = self.graph_get(f"{message_url}/attachments?$select=id,name,contentType,size,isInline")
        return payload.get("value", [])


    def parse_received_time(self, value: str | None) -> str:
        if not value:
            return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.astimezone(timezone.utc).strftime("%Y%m%d_%H%M%S")


    def decode_file_attachment(self, attachment: dict) -> bytes:
        content_bytes = attachment.get("contentBytes")
        if content_bytes:
            return base64.b64decode(content_bytes)
        raise GraphError("Attachment payload did not include contentBytes")


    def extract_matching_links(self, body_content: str) -> list[str]:
        """Extract URLs from body content that contain keywords or match specific patterns."""
        url_pattern = r'https?://[^\s<>"{}|\\^`\[\]]*'
        href_pattern = r'href=["\']([^"\']+)["\']'
        candidates = re.findall(url_pattern, body_content)
        candidates.extend(re.findall(href_pattern, body_content, flags=re.IGNORECASE))

        matching_urls = []
        seen_urls = set()

        for candidate in candidates:
            normalized_url = html.unescape(candidate).strip()
            if normalized_url.startswith("/"):
                normalized_url = urljoin(SZAMLAZZ_ROOT, normalized_url)

            if normalized_url in seen_urls:
                continue

            if any(re.match(pattern, normalized_url) for pattern in URL_FILTER_PATTERNS):
                matching_urls.append(normalized_url)
                seen_urls.add(normalized_url)

        return matching_urls


    def is_attachment_download_url(self, url: str) -> bool:
        parsed_url = urlparse(url)
        query = parse_qs(parsed_url.query)
        return (
            parsed_url.netloc.endswith("szamlazz.hu")
            and parsed_url.path == "/szamla/"
            and query.get("action") == ["szamlapdf"]
            and query.get("content_disp_type") == ["attachment"]
        )
    
    
    def sanitize_filename(self, value: str, fallback: str = "untitled") -> str:
        cleaned = re.sub(r'[<>:"/|?*\x00-\x1f]', "_", value).strip(" ._")
        return cleaned or fallback
    
    
    def unique_path(self, directory: Path, stem: str, suffix: str) -> Path:
        candidate = directory / f"{stem}{suffix}"
        index = 1
        while candidate.exists():
            candidate = directory / f"{stem}_{index}{suffix}"
            index += 1
        return candidate