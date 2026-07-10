[中文](README.zh-CN.md)

# tableToExcel

A SiYuan Note plugin that exports complex tables (with merged cells, line breaks within cells) to Excel, or copies them to the clipboard for direct paste into Excel.

## Features

- **Copy to Excel** — Right-click the table block icon → "Copy to Excel", then paste directly into Excel (supports merged cells and in-cell line breaks)
- **Export to Excel** — *(Coming soon)* Export tables as .xlsx files

## Usage

1. Click the block icon (☰) on the left side of any table in the editor
2. Select **"Copy to Excel"** from the menu
3. Paste into Excel (`Ctrl+V`) — merged cells and line breaks are preserved

## Demo

| Before (SiYuan) | After (Excel) |
|:---:|:---:|
| Table with merged cells (`rowspan`/`colspan`) and `Shift+Enter` line breaks | Same structure preserved in Excel, ready for further editing |

## How it works

The plugin directly clones the table's DOM from SiYuan's WYSIWYG editor, removes proprietary attributes, and writes a clean HTML table to the clipboard in Excel-compatible format. This preserves:
- Merged cells (`colspan` / `rowspan`)
- In-cell line breaks (`Shift+Enter`)
- Cell text content (stripped of inline formatting)

## Installation

1. Download the latest release `package.zip` from [Releases](https://github.com/Koto-Sakura/tableToExcel/releases)
2. In SiYuan, go to `Settings` → `Marketplace` → `Download` → `Manual install`
3. Select the downloaded `package.zip`

Or manually extract to `{workspace}/data/plugins/tableToExcel/`.

## Development

```bash
pnpm install
pnpm run dev     # watch mode
pnpm run build   # production build
```

## Build

After `pnpm run build`, the `dist/` directory contains all plugin files. The file `package.zip` is the distributable package.

## License

[MIT](LICENSE)
