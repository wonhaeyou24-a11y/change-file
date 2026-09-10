const ExcelJS = require("exceljs");

const FONT_NAME = "맑은 고딕";
const thinBorder = { style: "thin", color: { argb: "FF000000" } };
const BORDER = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };

async function buildSlopeXlsx(rows, title) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("이력현황");

  ws.mergeCells("A1:F1");
  ws.getCell("A1").value = title || "점검진단실적";
  ws.getCell("A1").font = { name: FONT_NAME, size: 14, bold: true };

  const headers = ["No.", "점검종류", "점검기간", "점검기관", "평가등급", "점검 결과"];
  const widths = [6, 14, 16, 20, 10, 70];
  ws.columns = widths.map(w => ({ width: w }));

  const headerRow = ws.getRow(2);
  headers.forEach((h, i) => {
    const c = headerRow.getCell(i + 1);
    c.value = h;
    c.font = { name: FONT_NAME, bold: true };
    c.fill = HEADER_FILL;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = BORDER;
  });

  let r = 3;
  for (const row of rows) {
    const period = `${row.start} ~ ${row.end}`;
    const resultText = row.bullets.map(b => `-${b}`).join("\n");
    const values = [row.no, row.type, period, row.org, row.grade, resultText];
    values.forEach((v, i) => {
      const c = ws.getCell(r, i + 1);
      c.value = v;
      c.font = { name: FONT_NAME, size: 10, bold: i === 4 };
      c.border = BORDER;
      c.alignment = i === 5
        ? { horizontal: "left", vertical: "middle", wrapText: true }
        : { horizontal: "center", vertical: "middle", wrapText: true };
    });
    ws.getRow(r).height = Math.max(18, 15 * Math.max(row.bullets.length, 1));
    r++;
  }
  ws.views = [{ state: "frozen", ySplit: 2 }];
  return wb.xlsx.writeBuffer();
}

async function buildWallXlsx(rows, title) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("이력현황");

  ws.mergeCells("A1:F1");
  ws.getCell("A1").value = title || "점검진단실적";
  ws.getCell("A1").font = { name: FONT_NAME, size: 14, bold: true };

  const h1 = 2, h2 = 3;
  const top = ["No.", "수행기간", "점검·진단기관", "안전등급", "주요 점검·진단 결과", "보수·보강(안)"];
  const bottom = ["", "점검·진단구분", "책임기술자", "", "", ""];
  const widths = [6, 16, 16, 10, 60, 16];
  ws.columns = widths.map(w => ({ width: w }));

  top.forEach((t, i) => {
    const ct = ws.getCell(h1, i + 1);
    const cb = ws.getCell(h2, i + 1);
    ct.value = t;
    cb.value = bottom[i];
    for (const c of [ct, cb]) {
      c.font = { name: FONT_NAME, bold: true };
      c.fill = HEADER_FILL;
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      c.border = BORDER;
    }
    if ([1, 4, 5, 6].includes(i + 1)) {
      ws.mergeCells(h1, i + 1, h2, i + 1);
    }
  });

  let r = h2 + 1;
  for (const row of rows) {
    const rTop = r, rBot = r + 1;
    const periodTop = `${row.start.replace(/-/g, ".")}. ~`;
    const periodBot = `${row.end.replace(/-/g, ".")}.`;
    const resultText = row.result.map(b => `-${b}`).join("\n");
    const remedyText = row.remedy.map(b => `-${b}`).join("\n");

    ws.getCell(rTop, 1).value = row.no;
    ws.getCell(rTop, 2).value = `${periodTop}\n${periodBot}`;
    ws.getCell(rBot, 2).value = row.type;
    ws.getCell(rTop, 3).value = row.org;
    ws.getCell(rBot, 3).value = row.engineer;
    ws.getCell(rTop, 4).value = row.grade;
    ws.getCell(rTop, 5).value = resultText;
    ws.getCell(rTop, 6).value = remedyText;

    for (let i = 1; i <= 6; i++) {
      for (const rr of [rTop, rBot]) {
        const c = ws.getCell(rr, i);
        c.font = { name: FONT_NAME, size: 10, bold: i === 4 };
        c.border = BORDER;
        c.alignment = [5, 6].includes(i)
          ? { horizontal: "left", vertical: "middle", wrapText: true }
          : { horizontal: "center", vertical: "middle", wrapText: true };
      }
    }
    ws.mergeCells(rTop, 1, rBot, 1);
    ws.mergeCells(rTop, 4, rBot, 4);
    ws.mergeCells(rTop, 5, rBot, 5);
    ws.mergeCells(rTop, 6, rBot, 6);

    const lines = Math.max(row.result.length, row.remedy.length, 1);
    const totalH = Math.max(36, 15 * lines);
    ws.getRow(rTop).height = totalH / 2;
    ws.getRow(rBot).height = totalH / 2;
    r += 2;
  }
  ws.views = [{ state: "frozen", ySplit: h2 }];
  return wb.xlsx.writeBuffer();
}

module.exports = { buildSlopeXlsx, buildWallXlsx };
