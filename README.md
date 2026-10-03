# Outlook PDF Attachment Downloader

This script downloads PDF attachments from Outlook mailboxes using Microsoft Graph, so it works cross-platform with Outlook desktop and Outlook on the web.

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Interactive sign-in

Register a public client app in Microsoft Entra ID, then run:

```bash
python download_outlook_pdf_attachments.py --client-id <app-client-id> --account user@company.com
```

## App-only sign-in

Use a confidential client app with application permissions:

```bash
python download_outlook_pdf_attachments.py \
  --tenant-id <tenant-id> \
  --client-id <app-client-id> \
  --client-secret <app-secret> \
  --account user@company.com
```

## Options

- `--folder` mail folder path such as `Inbox` or `Inbox/Invoices`
- `--output` destination directory for saved PDFs
- `--recursive` include subfolders
- `--download-xls` download the configured shared Excel file through Microsoft Graph
- `--xls-url` override the default SharePoint sharing URL for the Excel file
- `--xls-output` set the destination path for the downloaded Excel file

## Download The Shared Excel File

```bash
python invoice_handler.py \
  --tenant-id <tenant-id> \
  --client-id <app-client-id> \
  --client-secret <app-secret> \
  --account user@company.com \
  --download-xls \
  --xls-output attachments/Nordvik.xlsx
```

## Notes

- The script only saves attachments that look like PDFs by filename or MIME type.
- For device-code sign-in, `--client-id` is required.
- For app-only access, the app needs the appropriate Microsoft Graph mail permissions.
- Downloading the shared Excel file also requires Graph permissions that can read SharePoint or OneDrive files, such as `Files.Read.All` or `Sites.Read.All`.
