[中文](README.zh-CN.md)

# tableToExcel

A SiYuan Note plugin for exporting complex tables to Excel — supports merged cells, in-cell line breaks, and other complex table formatting.

SiYuan's built-in table copy cannot handle merged cells or line breaks properly when pasting into Excel. This plugin solves that problem.

## Features

| Feature | Description |
|---------|-------------|
| **Copy to Excel** | Right-click a table → Copy to Excel → paste into Excel (`Ctrl+V`). Fast and lightweight. Suitable for tables without merged cells. |
| **Export to Excel (.xlsx)** | Right-click a table → Export to Excel → downloads a proper `.xlsx` file. **Perfectly preserves merged cells and line breaks.** Uses SheetJS to generate a standard Excel file. |

## How to Use

1. Open any note containing a table
2. Click the **block icon (☰)** on the left side of the table
3. Choose:
   - **Copy to Excel** — then `Ctrl+V` in Excel
   - **Export to Excel** — the browser will download a `.xlsx` file

## How It Works

### Copy to Excel (Clipboard)
The plugin clones the table's DOM from SiYuan's editor, removes SiYuan-specific attributes, and writes a clean HTML table to the clipboard. Since Excel's clipboard parser does not support `rowspan`, merged cells are expanded (each cell gets its own copy of the content). Line breaks (`<br>` / `Shift+Enter`) are preserved via `mso-data-placement:same-cell`.

### Export to Excel (.xlsx)
The plugin parses the table DOM, extracts all data, and uses SheetJS (xlsx) to generate a standard `.xlsx` file. Merged cells are stored in the Excel `<mergeCells>` list, and line breaks are enabled via the `wrapText` cell style. The result is a fully compatible Excel file with all formatting intact.

## Installation

### Method 1: SiYuan Marketplace
Open SiYuan → **Settings → Marketplace → Download** → find **tableToExcel** and install.

### Method 2: Manual Install
1. Download `package.zip` from the [Releases page](https://github.com/Koto-Sakura/tableToExcel/releases)
2. Extract to `{workspace}/data/plugins/tableToExcel/`
3. Restart SiYuan or disable and re-enable the plugin in **Settings → Marketplace → Downloaded**

## License

[MIT](LICENSE)

## Credits

This plugin was developed with the assistance of AI. All code was tested by humans before release.
