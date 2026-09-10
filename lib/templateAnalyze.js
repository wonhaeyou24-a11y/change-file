const ExcelJS = require("exceljs");
const { extractData, matrixToTSV } = require("./dataExtract");

/**
 * "양식 파일"을 AI에게 전달할 형태로 변환한다.
 *  - 이미지(png/jpg/webp/gif)  → { kind:"image", mediaType, dataB64 }
 *  - PDF                        → { kind:"pdf", dataB64 }
 *  - 엑셀/csv/hwpx              → { kind:"grid", text }  (셀 배치 + 병합 정보를 텍스트로)
 */

const IMAGE_MIME = {
  ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
  ".webp": "image/webp", ".gif": "image/gif",
};

async function xlsxToGridText(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  const lines = [];
  ws.eachRow({ includeEmpty: true }, (row, rn) => {
    const cells = [];
    row.eachCell({ includeEmpty: true }, (cell, cn) => {
      let v = cell.value;
      if (v && typeof v === "object") {
        if (v.richText) v = v.richText.map(t => t.text).join("");
        else if (v.text) v = v.text;
        else if (v.result != null) v = v.result;
        else if (v instanceof Date) v = v.toISOString().slice(0, 10);
        else v = "";
      }
      if (v !== "" && v != null) cells.push(`R${rn}C${cn}="${String(v).replace(/\n/g, " ⏎ ")}"`);
    });
    if (cells.length) lines.push(cells.join("  "));
  });
  const merges = (ws.model.merges || []).join(", ");
  return `[양식 엑셀 셀 내용]\n${lines.join("\n")}\n\n[병합된 셀 범위]\n${merges || "(없음)"}`;
}

async function analyzeTemplate(buffer, originalName = "") {
  const ext = (originalName.toLowerCase().match(/\.[a-z0-9]+$/) || [""])[0];

  if (IMAGE_MIME[ext]) {
    return { kind: "image", mediaType: IMAGE_MIME[ext], dataB64: buffer.toString("base64") };
  }
  if (ext === ".pdf") {
    return { kind: "pdf", dataB64: buffer.toString("base64") };
  }
  if (ext === ".xlsx" || ext === ".xlsm") {
    return { kind: "grid", text: await xlsxToGridText(buffer) };
  }
  // .xls / .csv / .hwpx 등 → 공통 추출기로 matrix 뽑아 텍스트화
  const { matrix } = extractData(buffer, originalName);
  return { kind: "grid", text: `[양식 표 내용 (행/열)]\n${matrixToTSV(matrix)}` };
}

module.exports = { analyzeTemplate };
