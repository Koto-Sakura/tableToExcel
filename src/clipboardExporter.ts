/**
 * 复制带合并单元格的表格到剪贴板，以便直接粘贴到 Excel。
 *
 * 策略:
 * 1. 克隆原始 DOM 获取每个单元格的精确内容（保留 <br> / \n 文本节点）
 * 2. 解析表格结构（colspan / rowspan）确定逻辑网格
 * 3. 生成新 HTML：每行固定 colCount 个 <td>，被合并覆盖的用空 <td> 占位
 *    这样 Excel 不会因为行内 <td> 数量不一致而错列
 * 4. 合并原点保留 colspan/rowspan 属性
 */

/**
 * 获取单元格内部的 HTML 内容（清理思源特有元素后）。
 */
function getCellContent(cell: HTMLTableCellElement): string {
  const clone = cell.cloneNode(true) as HTMLTableCellElement;

  // 移除思源特有元素
  clone.querySelectorAll(".protyle-attr, wbr").forEach((el) => el.remove());

  // 移除内容可编辑区域的 contenteditable 属性（保留 div 结构）
  const editDivs = clone.querySelectorAll<HTMLElement>('[contenteditable]');
  editDivs.forEach((el) => el.removeAttribute("contenteditable"));

  return clone.innerHTML;
}

/**
 * 解析表格 DOM，返回逻辑网格结构和每个单元格的内容。
 */
interface CellInfo {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  content: string; // 单元格内部 HTML
  isTh: boolean;   // 是否为表头
}

interface ParseResult {
  cells: CellInfo[];
  rowCount: number;
  colCount: number;
}

function parseTableStructure(originalTable: HTMLTableElement): ParseResult {
  const rows = originalTable.querySelectorAll<HTMLTableRowElement>(
    ":scope > thead > tr, :scope > tbody > tr",
  );
  if (rows.length === 0) {
    return { cells: [], rowCount: 0, colCount: 0 };
  }

  const occupied: boolean[][] = [];
  const getOcc = (r: number, c: number) => occupied[r]?.[c] ?? false;
  const setOcc = (r: number, c: number) => {
    if (!occupied[r]) occupied[r] = [];
    occupied[r][c] = true;
  };

  const cells: CellInfo[] = [];
  let maxCol = 0;

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const tr = rows[rowIdx];
    const cellElements = tr.querySelectorAll<HTMLTableCellElement>(":scope > th, :scope > td");

    // 思源表格中每行都有相同数量的 <td>/<th>，被 rowspan 覆盖的位置
    // 也放空 <td> 占位（用于编辑器光标定位）。因此按顺序直接分配列号，
    // 不跳过被覆盖的列。
    let colIdx = 0;

    for (const cell of cellElements) {
      const colSpan = cell.colSpan || 1;
      const rowSpan = cell.rowSpan || 1;
      const content = getCellContent(cell);

      cells.push({ row: rowIdx, col: colIdx, rowSpan, colSpan, content, isTh: cell.tagName === "TH" });

      for (let r = rowIdx; r < rowIdx + rowSpan; r++) {
        for (let c = colIdx; c < colIdx + colSpan; c++) {
          setOcc(r, c);
        }
      }

      colIdx += colSpan;
    }

    maxCol = Math.max(maxCol, colIdx);
  }

  return { cells, rowCount: occupied.length, colCount: maxCol };
}

/**
 * 构建 Excel 兼容的 HTML 文档。
 * 每行固定 colCount 个 <td>，被合并覆盖的用空 <td> 占位。
 */
function buildExcelHtml(blockElement: HTMLElement): string {
  const originalTable = blockElement.querySelector<HTMLTableElement>("table");
  if (!originalTable) return "";

  const { cells, rowCount, colCount } = parseTableStructure(originalTable);
  if (rowCount === 0 || colCount === 0) return "";

  // 检查合并信息，辅助判断占位
  const merges = cells.filter((c) => c.colSpan > 1 || c.rowSpan > 1);
  const isOrigin = (r: number, c: number) => merges.some((m) => m.row === r && m.col === c);
  const isCovered = (r: number, c: number) => merges.some(
    (m) => r >= m.row && r < m.row + m.rowSpan && c >= m.col && c < m.col + m.colSpan,
  );

  // 构建每一行
  const rows: string[][] = []; // rows[row][col] = HTML string or null (skip)
  for (let r = 0; r < rowCount; r++) {
    rows[r] = [];
    for (let c = 0; c < colCount; c++) {
      rows[r][c] = ""; // 默认空
    }
  }

  // 展开合并单元格：将合并原点的内容填充到所有被覆盖的位置
  // 不输出 rowspan/colspan 属性（Excel 剪贴板解析器不支持 rowspan）
  for (const cell of cells) {
    const tag = cell.isTh ? "th" : "td";
    const content = cell.content || "";

    // 从合并原点扩展到所有被覆盖的位置
    for (let r = cell.row; r < cell.row + cell.rowSpan; r++) {
      for (let c = cell.col; c < cell.col + cell.colSpan; c++) {
        if (r < rowCount && c < colCount) {
          if (r === cell.row && c === cell.col) {
            // 原点：带内容
            rows[r][c] = `<${tag}>${content}</${tag}>`;
          } else if (rows[r][c] === "") {
            // 被覆盖位置：填充相同内容
            rows[r][c] = `<${tag}>${content}</${tag}>`;
          }
        }
      }
    }
  }

  // 剩余空单元格填充
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      if (rows[r][c] === "") {
        rows[r][c] = "<td></td>";
      }
    }
  }

  const tableHtml = `<table>${rows.map((row) => `<tr>${row.join("")}</tr>`).join("")}</table>`;

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
  const html = buildExcelHtml(blockElement);
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
