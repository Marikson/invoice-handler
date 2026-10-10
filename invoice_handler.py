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

import urllib3.util.connection

from ms_auth import MSAuth
from mail_methods import MailMethods
from xls_methods import XLS_FILE_NAME, XLS_FILE_PATH, XlsMethods
from pdf_parser import PDFParser


XLS_DATA_START_ROW = 2
XLS_DATA_COLUMNS = [3, 4, 5, 8]
XLS_INVOICE_ID_COLUMN = 3


def force_ipv4_only() -> None:
    urllib3.util.connection.allowed_gai_family = lambda: socket.AF_INET


def default_xls_output_path() -> str:
    file_name = XLS_FILE_NAME if Path(XLS_FILE_NAME).suffix else f"{XLS_FILE_NAME}.xlsx"
    return str(Path("xls") / file_name)


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
    parser.add_argument("--xls-output", default=default_xls_output_path(), help="Destination path for the downloaded Excel file")
    parser.add_argument("--xls-sheet", default="Vevői számlák", help="Worksheet name to update in the downloaded Excel file")
    parser.add_argument("--upload-xls", default=True, action="store_true", help="Upload the edited Excel file back to SharePoint after writing parsed invoice data")
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
    xls_methods = XlsMethods(auth_instance.session)
    xls_output = Path(args.xls_output).expanduser().resolve()
    xls_downloaded = False
    if args.download_xls:
      xls_downloaded = xls_methods.download_xls_file(str(xls_output), args.xls_url)

    print(f"Done. Saved {mail_methods.saved_attachments} PDF attachment(s) to {output_dir}")
    if args.download_xls:
      if xls_downloaded:
        print(f"Downloaded Excel file to {xls_output}")
      else:
        print("Excel download failed")
    
  
    pdf_parser = PDFParser(output_dir)
    invoice_data = pdf_parser()

    if xls_downloaded and invoice_data:
      next_empty_row = xls_methods.find_next_empty_row(
        str(xls_output),
        args.xls_sheet,
        min_row=XLS_DATA_START_ROW,
        columns=XLS_DATA_COLUMNS,
      )
      new_invoices, duplicate_invoices = xls_methods.filter_new_invoices(
        str(xls_output),
        args.xls_sheet,
        invoice_data,
        id_column=XLS_INVOICE_ID_COLUMN,
        min_row=XLS_DATA_START_ROW,
        next_empty_row=next_empty_row,
      )
      if duplicate_invoices:
        duplicate_ids = ", ".join(str(invoice.invoice_id or "N/A") for invoice in duplicate_invoices)
        print(f"Skipped {len(duplicate_invoices)} invoice(s) already present in Excel: {duplicate_ids}")

      if not new_invoices:
        print("No new invoice rows to write into Excel")
      else:
        if next_empty_row is None:
          print("Could not determine where to append invoice rows in Excel!")
        else:
          invoice_updates = xls_methods.build_invoice_updates(new_invoices, next_empty_row)
          if xls_methods.write_to_downloaded_xls(str(xls_output), args.xls_sheet, invoice_updates):
            print(f"Wrote {len(new_invoices)} invoice row(s) to {args.xls_sheet} in {xls_output} starting at row {next_empty_row}")
            if args.upload_xls:
              if xls_methods.upload_xls_file(str(xls_output), args.xls_url):
                print("Uploaded edited Excel file!")
              else:
                print("Excel upload failed! - Make sure xls file is not opened Excel online by a user!")
          else:
            print("Excel write failed")
    elif args.download_xls and not invoice_data:
      print("No invoice data found to write into Excel!")
    
    
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
