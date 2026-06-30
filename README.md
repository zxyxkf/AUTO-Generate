# AUTO Generate

A local Office document generator for producing editable Word/Excel files and matching PDF exports from user-provided templates.

The app runs entirely on the user's Windows machine. It does not require a remote server.

## Features

- Generate quotation sheets, supply lists, and contracts from Office templates.
- Export both editable Office files and final PDF files.
- Insert stamp images from a local asset folder.
- Provide a browser-based form and live document preview.
- Support a portable local launcher for non-technical users.

## Privacy Notice

This repository intentionally does not include private business data.

Do not commit:

- Real customer, buyer, supplier, or company information
- Tax IDs, bank accounts, phone numbers, addresses, or contacts
- PSD stamp source files
- Stamp PNG files
- Real Word or Excel templates
- Generated Office/PDF output
- Logs or local work files

Use placeholder templates and placeholder configuration when publishing changes.

## Requirements

For development:

- Windows
- Node.js
- Microsoft Office with Word and Excel

For portable end users:

- Windows
- Microsoft Office with Word and Excel
- No separate Node.js installation is required if `runtime/node/node.exe` is included in the portable package.

## Development

Install dependencies:

```powershell
npm install
```

Run API and Vite dev server:

```powershell
npm run dev
```

Build the browser app:

```powershell
npm run build
```

Run the local single-port app:

```powershell
node start-local.js
```

Default local URL:

```text
http://127.0.0.1:5174
```

## Portable Launcher

The Windows launcher files are:

```text
start launcher: 启动工具.bat
stop launcher:  停止工具.bat
start-local.js
```

The start launcher starts the local service and opens the browser. It first tries to use:

```text
runtime/node/node.exe
```

If that file does not exist, it falls back to the system `node` command.

Startup logs are written to:

```text
logs/start.log
```

## Project Structure

```text
assets/              Local public assets. Do not commit private stamps.
dist/                Built frontend files.
excel_templates/     User-provided Excel templates. Keep real templates private.
generated/           Generated Office/PDF output. Do not commit.
logs/                Local launcher logs. Do not commit.
runtime/             Optional portable Node runtime. Do not commit large runtime binaries unless explicitly needed.
scripts/office/      Office COM automation scripts.
server/              Local Express API and theme configuration.
src/                 Frontend source.
word_templates/      User-provided Word templates. Keep real templates private.
work/                Temporary generation files. Do not commit.
```

## Configuration

Document themes are configured in:

```text
server/config.js
```

Use placeholders for public examples:

```js
example_theme: {
  label: 'Example Supplier',
  company: 'Example Company Ltd.',
  supplierInfo: 'Company: Example Company Ltd.',
  stamp: 'assets/stamps/example-official.png',
  contractStamp: 'assets/stamps/example-contract.png',
  supplyTemplate: 'excel_templates/example-supply.xlsx',
  contractTemplate: 'word_templates/example-contract.docx',
}
```

Private deployments can replace those placeholder values locally.

## Generated Files

Generated files are written under:

```text
generated/YYYYMMDD/
```

Each generation usually creates:

```text
*.docx or *.xlsx
*.pdf
```

## Notes

- The generator relies on Microsoft Office COM automation on Windows.
- Word and Excel may fail if Office is not installed, templates are open in another program, or Office modal dialogs are blocking automation.
- Keep the original templates private and operate on generated copies only.
