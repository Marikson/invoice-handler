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
from xls_methods import XLS_FILE_PATH, XlsMethods
from pdf_parser import PDFParser


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
    parser.add_argument("--download-xls", default=True, action="store_true", help="Download the configured shared Excel file through Microsoft Graph")
    parser.add_argument("--xls-url", default=XLS_FILE_PATH, help="SharePoint or OneDrive sharing URL for the Excel file")
    parser.add_argument("--xls-output", default="xls/Nordvik.xlsx", help="Destination path for the downloaded Excel file")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.days < 1:
        print("Number of days must be at least 1")
        return 1
    output_dir = Path(args.output).expanduser().resolve()

    force_ipv4_only()

    auth_instance = MSAuth(args.tenant_id, args.client_id, args.client_secret, args.scopes)
    # mail_methods = MailMethods(auth_instance.session, args.account, args.folder, args.days, args.recursive, args.output)
    xls_methods = XlsMethods(auth_instance.session)
    xls_downloaded = False
    if args.download_xls:
      xls_downloaded = xls_methods.download_xls_file(args.xls_output, args.xls_url)

    # print(f"Done. Saved {mail_methods.saved_attachments} PDF attachment(s) to {output_dir}")
    if args.download_xls:
      xls_output = Path(args.xls_output).expanduser().resolve()
      if xls_downloaded:
        print(f"Downloaded Excel file to {xls_output}")
      else:
        print("Excel download failed")
    
  
    pdf_parser = PDFParser(output_dir)
    processed_pdf = pdf_parser()
    
    
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
