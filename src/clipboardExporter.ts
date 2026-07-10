/**
 * 复制带合并单元格的表格到剪贴板，以便直接粘贴到 Excel。
 *
 * 原理: 直接克隆思源表格的 DOM，清理多余属性后写入剪贴板。
 * 这样保留原始的换行结构（<br> / \n 文本节点），避免解析-重建过程中的信息丢失。
 */

/**
 * 清理克隆的表格 DOM，移除思源特有的属性和元素。
 */
function cleanTable(table: HTMLTableElement): void {
  // 移除所有 data-* 属性
  const allElements = table.querySelectorAll<HTMLElement>("*");
  allElements.forEach((el) => {
    // 移除 data-* 属性
    const attrs = el.attributes;
    for (let i = attrs.length - 1; i >= 0; i--) {
      const name = attrs[i].name;
      if (name.startsWith("data-") || name === "spellcheck" || name === "contenteditable" || name === "style" || name === "class") {
        el.removeAttribute(name);
      }
    }
  });

  // 移除 protyle-attr 面板
  const attrPanels = table.querySelectorAll<HTMLElement>(".protyle-attr");
  attrPanels.forEach((el) => el.remove());

  // 移除 <wbr> 元素（思源用于光标定位）
  const wbrElements = table.querySelectorAll("wbr");
  wbrElements.forEach((el) => el.remove());
}

/**
 * 从思源表格块 DOM 克隆并清理，生成 Excel 兼容的 HTML 文档。
 */
function buildExcelHtmlFromDom(blockElement: HTMLElement): string {
  const originalTable = blockElement.querySelector<HTMLTableElement>("table");
  if (!originalTable) {
    return "";
  }

  // 深克隆表格
  const clonedTable = originalTable.cloneNode(true) as HTMLTableElement;
  cleanTable(clonedTable);

  const tableHtml = clonedTable.outerHTML;

  return `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<style>
table, td, th, tr, tbody, thead {
  background-color: transparent !important;
  color: #000000 !important;
}
br { mso-data-placement: same-cell; }
</style>
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Sheet1</x:Name></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
</head>
<body>
${tableHtml}
</body>
</html>`;
}

/**
 * 将思源表格块 DOM 复制到系统剪贴板（Excel 兼容格式）。
 */
export async function copyTableBlockToClipboard(blockElement: HTMLElement): Promise<void> {
  const html = buildExcelHtmlFromDom(blockElement);
  if (!html) {
    throw new Error("未能找到表格");
  }

  try {
    const htmlBlob = new Blob([html], { type: "text/html" });
    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": htmlBlob,
      }),
    ]);
  } catch {
    // 回退到 execCommand
    await fallbackCopy(html);
  }
}

/**
 * 回退方案: 利用临时元素 + execCommand 复制 HTML。
 */
async function fallbackCopy(html: string): Promise<void> {
  const container = document.createElement("div");
  container.innerHTML = html;
  container.style.position = "fixed";
  container.style.left = "-9999px";
  container.style.top = "0";
  container.contentEditable = "true";
  document.body.appendChild(container);

  try {
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(container);
    selection?.removeAllRanges();
    selection?.addRange(range);
    document.execCommand("copy");
    selection?.removeAllRanges();
  } finally {
    document.body.removeChild(container);
  }
}
