# Outlook Invoice Handler

This script downloads PDF attachments from Outlook mailboxes using Microsoft Graph, parses invoice data from the downloaded PDFs, writes new invoice rows into a shared Excel file, and uploads the edited Excel file back to SharePoint/OneDrive.

It works cross-platform with Outlook desktop and Outlook on the web because it talks to the mailbox through Microsoft Graph.

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install .
```

The command above installs runtime dependencies from package metadata and registers the CLI command:

```bash
invoice-handler --version
```

For editable development install:

```bash
pip install -e .
```

## Build wheel

```bash
python -m pip install --upgrade build
python -m build
```

Artifacts are created under `dist/` (wheel and source distribution).

## Versioning and tags

- Runtime version is defined in `version.py` as `__version__`.
- The script prints version at startup and supports `--version`.
- Tag releases with the same value as the code version, for example:

```bash
git tag v0.1.0
git push origin v0.1.0
```

## Interactive sign-in

Register a public client app in Microsoft Entra ID, then run:

```bash
python invoice_handler.py --client-id <app-client-id> --account user@company.com
```

## End-to-end workflow

Each run performs the following steps:

1. Downloads PDF attachments into the output folder.
2. Downloads the configured shared Excel workbook.
3. Parses invoice fields from PDFs (vendor, invoice id, amount, deadline).
4. Appends only new invoice rows to the selected worksheet.
5. Uploads the edited workbook back to SharePoint/OneDrive.

## App-only sign-in

Use a confidential client app with application permissions:

```bash
python invoice_handler.py \
  --tenant-id <tenant-id> \
  --client-id <app-client-id> \
  --client-secret <app-secret> \
  --account user@company.com
```

## Options

- `--folder` mail folder path such as `Inbox` or `Inbox/Invoices`
- `--output` destination directory for saved PDFs
- `--recursive` include subfolders
- `--days` number of days to look back for emails (default: `7`)
- `--download-xls` download the configured shared Excel file through Microsoft Graph (currently enabled by default)
- `--xls-url` override the default SharePoint sharing URL for the Excel file
- `--xls-output` set the destination path for the downloaded Excel file
- `--xls-sheet` worksheet name to update (default: `Vevői számlák`)
- `--local-xls` skip upload and keep the edited Excel file only on local disk

Authentication options:

- `--account` mailbox address to read (for app-only and optional for device-flow)
- `--tenant-id` Microsoft Entra tenant id
- `--client-id` app registration client id
- `--client-secret` app registration client secret (app-only flow)
- `--scope` OAuth scope for device flow, can be repeated

## Example Run

```bash
python invoice_handler.py \
  --tenant-id <tenant-id> \
  --client-id <app-client-id> \
  --client-secret <app-secret> \
  --account user@company.com \
  --days 27 \
  --folder "Inbox/Invoices" \
  --xls-output xls/Nordvik.xlsx \
  --xls-sheet "Vevői számlák"
```

Keep the edited Excel file local only (skip upload back to SharePoint/OneDrive):

```bash
python invoice_handler.py \
  --tenant-id <tenant-id> \
  --client-id <app-client-id> \
  --client-secret <app-secret> \
  --account user@company.com \
  --local-xls
```

## Notes

- The script only saves attachments that look like PDFs by filename or MIME type.
- For device-code sign-in, `--client-id` is required.
- For app-only access, the app needs the appropriate Microsoft Graph mail permissions.
- Downloading the shared Excel file also requires Graph permissions that can read SharePoint or OneDrive files, such as `Files.Read.All` or `Sites.Read.All`.
- If the output attachment folder already exists, it is removed before downloading fresh attachments.
- Parsed rows with neither invoice id nor final amount are skipped.
