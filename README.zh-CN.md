[English](README.md)

# tableToExcel — 表格转Excel

思源笔记插件，将带合并单元格、格内换行的复杂表格复制到 Excel，或导出为 .xlsx 文件。

思源自带的表格复制功能无法正确处理合并单元格和换行，粘贴到 Excel 后会格式错乱。本插件解决这个问题。

## 功能对比

| 功能 | 说明 |
|------|------|
| **复制到Excel** | 右键表格 → 复制到Excel → 在 Excel 中粘贴（`Ctrl+V`）。轻量快捷，适合无合并单元格的表格。 |
| **导出为Excel文件 (.xlsx)** | 右键表格 → 导出为Excel文件 → 下载标准的 `.xlsx` 文件。**完美保留合并单元格和格内换行。** 使用 SheetJS 生成标准 Excel 文件。 |

## 使用方法

1. 打开任意包含表格的笔记
2. 点击表格左侧的**块图标（☰）**
3. 选择：
   - **复制到Excel** — 然后在 Excel 中 `Ctrl+V` 粘贴
   - **导出为Excel文件** — 浏览器会自动下载 `.xlsx` 文件

## 实现原理

### 复制到Excel（剪贴板）
插件克隆思源编辑器中表格的 DOM，移除思源特有属性后，以 HTML 格式写入系统剪贴板。由于 Excel 的剪贴板解析器不支持 `rowspan`，合并单元格会被展开（每个格子写入相同内容）。换行（`<br>` / `Shift+Enter`）通过 `mso-data-placement:same-cell` 保留。

### 导出为 .xlsx 文件
插件解析表格 DOM，提取所有数据，使用 SheetJS（xlsx）生成标准 `.xlsx` 文件。合并单元格存入 Excel 的 `<mergeCells>` 列表，格内换行通过 `wrapText` 单元格样式开启。生成的文件与 Excel 完全兼容，所有格式完整保留。

## 安装

1. 从 [Releases 页面](https://github.com/Koto-Sakura/tableToExcel/releases) 下载 `package.zip`
2. 在思源中：**设置 → 集市 → 下载 → 手动安装**
3. 选择下载的 `package.zip`

或手动解压到 `{工作空间}/data/plugins/tableToExcel/`。

## 许可证

[MIT](LICENSE)

## 致谢

本插件由 AI（AtomCode / deepseek-v4-flash）辅助开发，所有代码均经过人工审核和测试后发布。
