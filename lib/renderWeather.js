const ExcelJS = require("exceljs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, VerticalAlign, BorderStyle, ShadingType,
} = require("docx");

const FONT = "맑은 고딕";

/* ------------------------------------------------------------------ *
 *  공통: 데이터 → 3열 블록 배치 계산
 * ------------------------------------------------------------------ */

function fmtValue(v) {
  return Number(v).toFixed(1); // 이미지 양식: 항상 소수점 1자리 (예: 99.0)
}

function buildLayout(data, opts = {}) {
  const years = data.years.slice().sort((a, b) => a.year - b.year);
  const n = years.length;
  const blocks = 3;
  const blockLen = Math.ceil(n / blocks);
  const startY = n ? years[0].year : "";
  const endY = n ? years[n - 1].year : "";
  const avg = n ? years.reduce((s, y) => s + y.value, 0) / n : 0;

  // blockLen 행 × 3블록. 각 행은 [연도,발생일,값, 연도,발생일,값, 연도,발생일,값]
  const grid = [];
  for (let r = 0; r < blockLen; r++) {
    const row = [];
    for (let b = 0; b < blocks; b++) {
      const item = years[b * blockLen + r];
      if (item) row.push(String(item.year), item.date || "", fmtValue(item.value));
      else row.push("", "", "");
    }
    grid.push(row);
  }

  const caption = opts.caption || "기상청 강수량 분석 결과";
  const tableNo = opts.tableNo || "";
  return {
    name: data.name || opts.name || "",
    tableNo,
    caption,
    heading: tableNo ? `<표 ${tableNo}> ${caption}` : caption,
    valueHeader: opts.valueHeader || "연별 최다\n일강수량\n(mm)",
    startY, endY, n, blockLen, blocks, grid,
    avg: fmtValue(avg),
    periodText: `${startY}년~${endY}년 (${n}년)`,
  };
}

/* ------------------------------------------------------------------ *
 *  XLSX
 * ------------------------------------------------------------------ */

const thin = { style: "thin", color: { argb: "FF000000" } };
const BORDER = { top: thin, bottom: thin, left: thin, right: thin };
const GRAY = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };

async function buildWeatherXlsx(data, opts = {}) {
  const L = buildLayout(data, opts);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("강수량분석");

  ws.columns = [
    { width: 8 }, { width: 13 }, { width: 12 },
    { width: 8 }, { width: 13 }, { width: 12 },
    { width: 8 }, { width: 13 }, { width: 12 },
  ];

  const center = (opt = {}) => ({ horizontal: "center", vertical: "middle", wrapText: true, ...opt });
  const setCell = (r, c, value, { bold = false, fill = false } = {}) => {
    const cell = ws.getCell(r, c);
    cell.value = value;
    cell.font = { name: FONT, size: 10, bold };
    cell.alignment = center();
    cell.border = BORDER;
    if (fill) cell.fill = GRAY;
    return cell;
  };

  // 1행: 표 제목
  ws.mergeCells(1, 1, 1, 9);
  const t = ws.getCell(1, 1);
  t.value = L.heading;
  t.font = { name: FONT, size: 12, bold: true };
  t.alignment = center();

  // 2행: 헤더 (3블록 반복)
  const heads = ["연도", "발생일", L.valueHeader.replace(/\n/g, " ")];
  for (let b = 0; b < 3; b++) {
    heads.forEach((h, i) => setCell(2, b * 3 + i + 1, h, { bold: true, fill: true }));
  }
  ws.getRow(2).height = 42;

  // 3행~: 데이터
  let r = 3;
  for (const gridRow of L.grid) {
    gridRow.forEach((v, i) => setCell(r, i + 1, v));
    r++;
  }

  // 요약행
  setCell(r, 1, "위치", { bold: true, fill: true });
  setCell(r, 2, L.name);
  setCell(r, 3, "기간", { bold: true, fill: true });
  ws.mergeCells(r, 4, r, 7);
  setCell(r, 4, L.periodText);
  setCell(r, 8, "누년 평균값 (mm)", { bold: true, fill: true });
  setCell(r, 9, L.avg, { bold: true });

  return wb.xlsx.writeBuffer();
}

/* ------------------------------------------------------------------ *
 *  DOCX
 * ------------------------------------------------------------------ */

const DBORDER = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const CELL_BORDERS = { top: DBORDER, bottom: DBORDER, left: DBORDER, right: DBORDER };
const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } };

function dcell(text, { bold = false, fill = false, span = 1, width } = {}) {
  return new TableCell({
    columnSpan: span,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    borders: CELL_BORDERS,
    shading: fill ? { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" } : undefined,
    margins: { top: 40, bottom: 40, left: 40, right: 40 },
    children: String(text).split("\n").map(line => new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [new TextRun({ text: line, font: FONT, size: 18, bold })],
    })),
  });
}

async function buildWeatherDocx(data, opts = {}) {
  const L = buildLayout(data, opts);
  const COLW = [700, 1150, 1050, 700, 1150, 1050, 700, 1150, 1050];
  const totalW = COLW.reduce((a, b) => a + b, 0);

  const headRun = new TableRow({
    tableHeader: true,
    children: [0, 1, 2].flatMap(b => [
      dcell("연도", { bold: true, fill: true, width: COLW[b * 3] }),
      dcell("발생일", { bold: true, fill: true, width: COLW[b * 3 + 1] }),
      dcell(L.valueHeader, { bold: true, fill: true, width: COLW[b * 3 + 2] }),
    ]),
  });

  const bodyRows = L.grid.map(gridRow => new TableRow({
    children: gridRow.map((v, i) => dcell(v, { width: COLW[i] })),
  }));

  const summaryRow = new TableRow({
    children: [
      dcell("위치", { bold: true, fill: true, width: COLW[0] }),
      dcell(L.name, { width: COLW[1] }),
      dcell("기간", { bold: true, fill: true, width: COLW[2] }),
      dcell(L.periodText, { span: 4, width: COLW[3] + COLW[4] + COLW[5] + COLW[6] }),
      dcell("누년 평균값 (mm)", { bold: true, fill: true, width: COLW[7] }),
      dcell(L.avg, { bold: true, width: COLW[8] }),
    ],
  });

  const table = new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: COLW,
    rows: [headRun, ...bodyRows, summaryRow],
  });

  const children = [];
  if (opts.sectionHeading) {
    children.push(new Paragraph({
      spacing: { after: 160 },
      children: [new TextRun({ text: opts.sectionHeading, font: FONT, size: 24, bold: true })],
    }));
  }
  if (opts.narrative !== false) {
    children.push(new Paragraph({
      spacing: { after: 200 },
      indent: { firstLine: 200 },
      children: [new TextRun({
        text: `기상청의 최근 ${L.n}개년(${L.startY}~${L.endY}년)의 최대 일강우도값을 획득하였으며, `
          + `${L.name}지점의 누년 평균값은 ${L.avg}mm인 것으로 산출되었다.`,
        font: FONT, size: 20,
      })],
    }));
  }
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({ text: L.heading, font: FONT, size: 20, bold: true })],
  }));
  children.push(table);

  const doc = new Document({ sections: [{ properties: { page: PAGE }, children }] });
  return Packer.toBuffer(doc);
}

module.exports = { buildWeatherXlsx, buildWeatherDocx, buildLayout };
