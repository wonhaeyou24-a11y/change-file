const ExcelJS = require("exceljs");

/**
 * 원본 "점검진단실적" 엑셀(15열)을 읽어 행 배열로 반환한다.
 * 헤더는 보통 2행(1행: 제목, 2행: 실제 컬럼명)이므로 컬럼명으로 인덱스를 찾는다.
 */
async function extractRows(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];

  // 컬럼명이 들어있는 행을 찾는다 (보통 2행, 'No' 셀이 있는 행)
  let headerRowNum = null;
  let headerMap = {};
  for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
    const row = ws.getRow(r);
    const values = row.values.map(v => (v == null ? "" : String(v).trim()));
    if (values.includes("No")) {
      headerRowNum = r;
      values.forEach((v, i) => {
        if (v) headerMap[v] = i;
      });
      break;
    }
  }
  if (headerRowNum == null) {
    throw new Error("헤더 행(No 컬럼)을 찾을 수 없습니다. 원본 '점검진단실적' 엑셀 형식인지 확인해주세요.");
  }

  const col = (name) => headerMap[name];
  const rows = [];
  for (let r = headerRowNum + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const noVal = col("No") ? row.getCell(col("No")).value : null;
    if (noVal == null || noVal === "") continue;

    const cellText = (colName) => {
      const idx = col(colName);
      if (!idx) return "";
      const v = row.getCell(idx).value;
      if (v == null) return "";
      if (typeof v === "object" && v.richText) {
        return v.richText.map(t => t.text).join("");
      }
      return String(v);
    };
    const cellDate = (colName) => {
      const idx = col(colName);
      if (!idx) return "";
      const v = row.getCell(idx).value;
      if (v == null) return "";
      if (v instanceof Date) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, "0");
        const d = String(v.getDate()).padStart(2, "0");
        return `${y}-${m}-${d}`;
      }
      return String(v).slice(0, 10);
    };

    rows.push({
      no: String(noVal).trim(),
      type: cellText("점검진단구분").trim(),
      start: cellDate("시작일"),
      end: cellDate("종료일"),
      org: cellText("점검진단기관").trim(),
      engineer: cellText("책임기술자").trim(),
      grade: cellText("안전등급").trim(),
      resultRaw: cellText("주요점검진단결과"),
      remedyRaw: cellText("주요보수보강안"),
    });
  }
  return rows;
}

module.exports = { extractRows };
