[English](README.md)

# tableToExcel — 表格转Excel

思源笔记插件，将带合并单元格、格内换行的复杂表格复制到 Excel，或导出为 Excel 文件。

## 功能

- **复制到 Excel** — 右键表格块图标 →「复制到Excel」，直接粘贴到 Excel（保留合并单元格和格内换行）
- **导出为 Excel 文件** — *（开发中）* 将表格导出为 .xlsx 文件

## 使用方法

1. 点击编辑器内任意表格左侧的块图标（☰）
2. 在菜单中点击 **「复制到Excel」**
3. 在 Excel 中粘贴（`Ctrl+V`）— 合并单元格和换行均保留

## 实现原理

插件直接克隆思源 WYSIWYG 编辑器中表格的 DOM，移除思源特有属性后，以 Excel 兼容的 HTML 格式写入剪贴板。原样保留：
- 合并单元格（`colspan` / `rowspan`）
- 格内换行（`Shift+Enter`）
- 单元格文本内容

## 安装

1. 从 [Releases](https://github.com/Koto-Sakura/tableToExcel/releases) 下载最新版 `package.zip`
2. 在思源中打开 `设置` → `集市` → `下载` → `手动安装`
3. 选择下载的 `package.zip`

或手动解压到 `{工作空间}/data/plugins/tableToExcel/`。

## 开发

```bash
pnpm install
pnpm run dev     # 实时编译
pnpm run build   # 生产构建
```

构建后 `dist/` 目录包含所有插件文件，`package.zip` 为可分发的安装包。

## 许可证

[MIT](LICENSE)
