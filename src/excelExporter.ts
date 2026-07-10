/**
 * 将思源表格导出为 .xlsx 文件，完美支持合并单元格和格内换行。
 *
 * 使用 SheetJS 生成数据 + JSZip 修改 XML 添加 wrapText 样式
 * （开源版 SheetJS 不支持样式，需通过 XML 补丁实现）
 */

import * as XLSX from "xlsx";
import JSZip from "jszip";

export interface MergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

function parseTable(blockElement: HTMLElement): { data: string[][]; merges: MergeRange[] } {
  const table = blockElement.querySelector<HTMLTableElement>("table");
  if (!table) return { data: [], merges: [] };

  const rows = table.querySelectorAll<HTMLTableRowElement>(
    ":scope > thead > tr, :scope > tbody > tr",
  );
  if (rows.length === 0) return { data: [], merges: [] };

  function getText(cell: HTMLTableCellElement): string {
    const clone = cell.cloneNode(true) as HTMLTableCellElement;
    clone.querySelectorAll(".protyle-attr, wbr").forEach((el) => el.remove());
    return clone.innerHTML
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
 * 在 xlsx 的 styles.xml 中添加 wrapText 样式定义，
 * 并在 sheet 中为所有单元格引用该样式。
 */
async function addWrapTextStyle(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);

  // 1. 修改 styles.xml：添加一个带 wrapText 的 cellXfs
  let stylesXml = await zip.file("xl/styles.xml")?.async("string");
  if (!stylesXml) {
    // 如果没有 styles.xml，创建一个
    stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>`;
  }

  // 在 cellXfs 中添加一个带 wrapText 的样式定义
  // 查找 </cellXfs> 并插入新的 xf
  if (stylesXml.includes('xfId="0"')) {
    // 已有样式定义，追加
    const wrapTextXf = `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>`;
    stylesXml = stylesXml.replace(
      /(<cellXfs[^>]*count=")(\d+)(">)/,
      (match, prefix, count, suffix) => `${prefix}${parseInt(count) + 1}${suffix}`,
    );
    stylesXml = stylesXml.replace("</cellXfs>", `${wrapTextXf}\n</cellXfs>`);
  } else {
    // 没有 cellXfs，创建
    stylesXml = stylesXml.replace(
      "</styleSheet>",
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>\n</styleSheet>`,
    );
  }
  zip.file("xl/styles.xml", stylesXml);

  // 2. 修改 sheet1.xml：所有单元格引用样式 1（wrapText）
  // 查找 xlsx 生成的 cell 元素，添加 s="1"
  const sheetFiles = Object.keys(zip.files).filter((name) =>
    name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml")
  );

  for (const sheetPath of sheetFiles) {
    let sheetXml = await zip.file(sheetPath)?.async("string");
    if (!sheetXml) continue;

    // 给所有 <c> 元素添加 s="1"（如已有 s 属性则跳过）
    sheetXml = sheetXml.replace(/<c\s/g, '<c s="1" ');
    // 防止重复 s 属性
    sheetXml = sheetXml.replace(/s="1"\s+s="1"/g, 's="1"');
    zip.file(sheetPath, sheetXml);
  }

  // 重新生成 zip
  const newBlob = await zip.generateAsync({ type: "arraybuffer" });
  return newBlob;
}

/**
 * 导出为 .xlsx 文件并触发浏览器下载。
 * 合并单元格和格内换行完美保留。
 */
export async function exportTableToExcel(blockElement: HTMLElement): Promise<void> {
  const { data, merges } = parseTable(blockElement);
  if (data.length === 0) throw new Error("未能找到表格数据");

  const ws = XLSX.utils.aoa_to_sheet(data);

  if (merges.length > 0) {
    ws["!merges"] = merges.map((m) =>
      XLSX.utils.decode_range(
        `${XLSX.utils.encode_col(m.s.c)}${m.s.r + 1}:${XLSX.utils.encode_col(m.e.c)}${m.e.r + 1}`,
      ),
    );
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // 写出为 buffer
  const wbBuf = XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;

  // 用 JSZip 修补 XML，添加 wrapText 样式
  const patchedBuf = await addWrapTextStyle(wbBuf);

  // 触发下载
  const blob = new Blob([patchedBuf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "表格导出.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}
