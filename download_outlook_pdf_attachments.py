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
import socket

from pathlib import Path
from urllib.parse import quote

import urllib3.util.connection

from ms_auth import MSAuth
from mail_methods import MailMethods


def force_ipv4_only() -> None:
    urllib3.util.connection.allowed_gai_family = lambda: socket.AF_INET


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Download PDF attachments from Outlook mailboxes using Microsoft Graph."
    )
    parser.add_argument("--account", help="Mailbox address to read, e.g. user@company.com. Defaults to /me with device sign-in.")
    parser.add_argument("--folder", default="Inbox", help="Mail folder path, e.g. Inbox or Inbox/Invoices")
    parser.add_argument("--output", default="attachments", help="Directory where PDFs will be saved")
    parser.add_argument("--recursive", action="store_true", help="Scan subfolders recursively")
    parser.add_argument("--tenant-id", help="Microsoft Entra tenant id for authentication")
    parser.add_argument("--client-id", help="App registration client id")
    parser.add_argument("--client-secret", help="App registration client secret for client credentials")
    parser.add_argument("--scope", action="append", dest="scopes", help="OAuth scope for device flow. Can be repeated. Defaults to Mail.Read.")
    parser.add_argument("--days", type=int, default=7, help="Number of days to look back for emails")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.days < 1:
        print("Number of days must be at least 1")
        return 1
    output_dir = Path(args.output).expanduser().resolve()

    force_ipv4_only()

    auth_instance = MSAuth(args.tenant_id, args.client_id, args.client_secret, args.scopes)
    mail_methods = MailMethods(auth_instance.session, args.account, args.folder, args.days, args.recursive, args.output)

    
    print(f"Done. Saved {mail_methods.saved_attachments} PDF attachment(s) to {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
