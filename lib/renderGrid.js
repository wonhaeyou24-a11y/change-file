const ExcelJS = require("exceljs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, VerticalAlign, BorderStyle, ShadingType,
} = require("docx");

const FONT = "맑은 고딕";

/**
 * 범용 그리드 모델(aiTable 이 만든 JSON) → xlsx / docx.
 *
 * grid = {
 *   title, columns, widths:number[],
 *   rows: [ { cells: [ { text, fill, bold, align, colspan, rowspan } ] } ]
 * }
 * rowspan/colspan 으로 가려지는 칸은 cells 에 들어있지 않다고 가정한다.
 */

function normalize(grid) {
  const columns = grid.columns || Math.max(
    ...grid.rows.map(r => (r.cells || []).reduce((s, c) => s + (c.colspan || 1), 0))
  );
  let widths = Array.isArray(grid.widths) && grid.widths.length ? grid.widths.slice() : [];
  while (widths.length < columns) widths.push(10);
  return { title: grid.title || "", columns, widths: widths.slice(0, columns), rows: grid.rows || [] };
}

/** 각 셀에 시작 좌표(r,c)를 매겨 배치도를 만든다 */
function place(rows, columns) {
  const occ = new Set();
  const out = [];
  for (let r = 0; r < rows.length; r++) {
    const cells = rows[r].cells || [];
    let c = 0;
    const rowCells = [];
    for (const cell of cells) {
      while (occ.has(`${r},${c}`)) c++;
      const cs = Math.max(1, cell.colspan || 1);
      const rs = Math.max(1, cell.rowspan || 1);
      rowCells.push({ ...cell, r, c, cs, rs });
      for (let dr = 0; dr < rs; dr++)
        for (let dc = 0; dc < cs; dc++)
          if (dr || dc) occ.add(`${r + dr},${c + dc}`);
      c += cs;
    }
    out.push(rowCells);
  }
  return out;
}

/* ---------------- XLSX ---------------- */

const thin = { style: "thin", color: { argb: "FF000000" } };
const BORDER = { top: thin, bottom: thin, left: thin, right: thin };
const GRAY = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };

async function gridToXlsx(gridIn) {
  const grid = normalize(gridIn);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("표");
  ws.columns = grid.widths.map(w => ({ width: Math.max(5, Math.round(w * 1.15)) }));

  let rowOffset = 1;
  if (grid.title) {
    ws.mergeCells(1, 1, 1, grid.columns);
    const t = ws.getCell(1, 1);
    t.value = grid.title;
    t.font = { name: FONT, size: 12, bold: true };
    t.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    rowOffset = 2;
  }

  const placed = place(grid.rows, grid.columns);
  let maxRow = rowOffset - 1;
  for (const rowCells of placed) {
    for (const cell of rowCells) {
      const r = rowOffset + cell.r;
      const c = cell.c + 1;
      maxRow = Math.max(maxRow, r + cell.rs - 1);
      if (cell.cs > 1 || cell.rs > 1) {
        ws.mergeCells(r, c, r + cell.rs - 1, c + cell.cs - 1);
      }
      const xc = ws.getCell(r, c);
      xc.value = String(cell.text ?? "");
      xc.font = { name: FONT, size: 10, bold: !!cell.bold };
      xc.alignment = {
        horizontal: cell.align === "left" ? "left" : cell.align === "right" ? "right" : "center",
        vertical: "middle", wrapText: true,
      };
      xc.border = BORDER;
      if (cell.fill) xc.fill = GRAY;
    }
  }
  // 테두리를 병합 영역 전체에 적용
  for (let r = rowOffset; r <= maxRow; r++)
    for (let c = 1; c <= grid.columns; c++) {
      const xc = ws.getCell(r, c);
      if (!xc.border) xc.border = BORDER;
    }

  return wb.xlsx.writeBuffer();
}

/* ---------------- DOCX ---------------- */

const DB = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const CB = { top: DB, bottom: DB, left: DB, right: DB };
const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } };

async function gridToDocx(gridIn, opts = {}) {
  const grid = normalize(gridIn);
  const totalW = 9200;
  const unit = totalW / grid.widths.reduce((a, b) => a + b, 0);
  const colW = grid.widths.map(w => Math.round(w * unit));

  const placed = place(grid.rows, grid.columns);
  const docRows = placed.map(rowCells => new TableRow({
    children: rowCells.map(cell => new TableCell({
      columnSpan: cell.cs > 1 ? cell.cs : undefined,
      rowSpan: cell.rs > 1 ? cell.rs : undefined,
      width: { size: colW.slice(cell.c, cell.c + cell.cs).reduce((a, b) => a + b, 0), type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: CB,
      shading: cell.fill ? { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" } : undefined,
      margins: { top: 40, bottom: 40, left: 50, right: 50 },
      children: String(cell.text ?? "").split("\n").map(line => new Paragraph({
        alignment: cell.align === "left" ? AlignmentType.LEFT
          : cell.align === "right" ? AlignmentType.RIGHT : AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [new TextRun({ text: line, font: FONT, size: 18, bold: !!cell.bold })],
      })),
    })),
  }));

  const children = [];
  if (opts.sectionHeading) {
    children.push(new Paragraph({
      spacing: { after: 160 },
      children: [new TextRun({ text: opts.sectionHeading, font: FONT, size: 24, bold: true })],
    }));
  }
  if (grid.title) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 120 },
      children: [new TextRun({ text: grid.title, font: FONT, size: 20, bold: true })],
    }));
  }
  children.push(new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colW,
    rows: docRows,
  }));

  const doc = new Document({ sections: [{ properties: { page: PAGE }, children }] });
  return Packer.toBuffer(doc);
}

module.exports = { gridToXlsx, gridToDocx, normalize, place };
