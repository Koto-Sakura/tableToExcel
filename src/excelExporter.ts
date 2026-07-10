/**
 * 将思源表格导出为真正的 .xlsx 文件（支持合并单元格）。
 *
 * 使用 SheetJS (xlsx) 库生成标准 Excel 文件，
 * 合并单元格通过 ws['!merges'] 原生支持。
 */

import * as XLSX from "xlsx";

export interface MergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

/**
 * 解析思源表格 DOM，返回二维数据数组和合并范围列表。
 */
function parseTable(blockElement: HTMLElement): { data: string[][]; merges: MergeRange[] } {
  const table = blockElement.querySelector<HTMLTableElement>("table");
  if (!table) {
    return { data: [], merges: [] };
  }

  const rows = table.querySelectorAll<HTMLTableRowElement>(
    ":scope > thead > tr, :scope > tbody > tr",
  );
  if (rows.length === 0) {
    return { data: [], merges: [] };
  }

  // 提取每个单元格的纯文本内容
  function getText(cell: HTMLTableCellElement): string {
    const editDiv = cell.querySelector<HTMLDivElement>('[contenteditable="true"]');
    if (!editDiv) return cell.textContent?.trim() ?? "";
    return editDiv.innerHTML
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/\u200B/g, "")
      .trim();
  }

  const occupied: boolean[][] = [];
  const getOcc = (r: number, c: number) => occupied[r]?.[c] ?? false;
  const setOcc = (r: number, c: number) => {
    if (!occupied[r]) occupied[r] = [];
    occupied[r][c] = true;
  };

  interface CellInfo {
    row: number;
    col: number;
    rowSpan: number;
    colSpan: number;
    text: string;
  }

  const cellInfos: CellInfo[] = [];
  let maxCol = 0;

  // 思源表格每行有相同数量的 <td>，按顺序分配列号
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const tr = rows[rowIdx];
    const cells = tr.querySelectorAll<HTMLTableCellElement>(":scope > th, :scope > td");
    let colIdx = 0;

    for (const cell of cells) {
      const colSpan = cell.colSpan || 1;
      const rowSpan = cell.rowSpan || 1;
      const text = getText(cell);

      cellInfos.push({ row: rowIdx, col: colIdx, rowSpan, colSpan, text });

      for (let r = rowIdx; r < rowIdx + rowSpan; r++) {
        for (let c = colIdx; c < colIdx + colSpan; c++) {
          setOcc(r, c);
        }
      }
      colIdx += colSpan;
    }
    maxCol = Math.max(maxCol, colIdx);
  }

  const rowCount = occupied.length;
  const colCount = maxCol;
  if (rowCount === 0 || colCount === 0) return { data: [], merges: [] };

  // 构建 data 数组和 merges
  const data: string[][] = Array.from({ length: rowCount }, () => Array(colCount).fill(""));
  const merges: MergeRange[] = [];

  for (const info of cellInfos) {
    data[info.row][info.col] = info.text;
    if (info.colSpan > 1 || info.rowSpan > 1) {
      merges.push({
        s: { r: info.row, c: info.col },
        e: { r: info.row + info.rowSpan - 1, c: info.col + info.colSpan - 1 },
      });
    }
  }

  return { data, merges };
}

/**
 * 将思源表格块导出为 .xlsx 文件并触发浏览器下载。
 */
export function exportTableToExcel(blockElement: HTMLElement): void {
  const { data, merges } = parseTable(blockElement);
  if (data.length === 0) {
    throw new Error("未能找到表格数据");
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  if (merges.length > 0) {
    ws["!merges"] = merges.map((m) => XLSX.utils.decode_range(
      `${XLSX.utils.encode_col(m.s.c)}${m.s.r + 1}:${XLSX.utils.encode_col(m.e.c)}${m.e.r + 1}`,
    ));
  }

  const wb = XLSX.utils.book_new(ws, "Sheet1");
  XLSX.writeFile(wb, "表格导出.xlsx", { compression: true });
}
