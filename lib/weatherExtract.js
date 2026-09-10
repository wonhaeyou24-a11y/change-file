const ExcelJS = require("exceljs");

/**
 * 기상청(KMA) "일 최다 강수량" 자료를 읽어 연도별 행으로 반환한다.
 *
 * 지원 입력:
 *  - 기상청 사이트에서 내려받은 xls(실제로는 탭 구분 텍스트) / csv / txt
 *  - 진짜 엑셀(.xlsx)
 *
 * 기대 컬럼(헤더 이름으로 자동 인식, 순서 무관):
 *   지점 / 지점명 / 일시(=연도) / 일 최다 강수량(mm) / 일 최다 강수량 나타난날(yyyymmdd)
 * 헤더 인식에 실패하면 위 순서대로의 위치(0~4열)로 처리한다.
 */

function decodeBuffer(buffer) {
  // 우선 UTF-8, 깨지면 EUC-KR 재시도 (기상청 CSV가 EUC-KR인 경우가 있음)
  let text = buffer.toString("utf8");
  if (text.includes("�")) {
    try {
      text = new TextDecoder("euc-kr").decode(buffer);
    } catch (_) {
      /* keep utf8 */
    }
  }
  return text;
}

function splitLine(line) {
  if (line.includes("\t")) return line.split("\t");
  // 따옴표 없는 단순 CSV 가정 (기상청 자료 형식)
  return line.split(",");
}

function pick(headerCells, ...needles) {
  for (let i = 0; i < headerCells.length; i++) {
    const h = headerCells[i].replace(/\s/g, "");
    if (needles.every(n => h.includes(n))) return i;
  }
  return -1;
}

function toISODate(raw) {
  if (raw == null) return "";
  const s = String(raw).trim();
  const digits = s.replace(/[^0-9]/g, "");
  if (digits.length === 8) {
    return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  }
  // 이미 yyyy-mm-dd 또는 yyyy.mm.dd 형태
  const m = s.match(/(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  return s;
}

function rowsFromMatrix(matrix) {
  // 헤더 행 = "강수량" 또는 "지점"이 들어있는 첫 행
  let headerIdx = matrix.findIndex(cells =>
    cells.some(c => /지점|강수량|일시|연도/.test(String(c)))
  );
  if (headerIdx === -1) headerIdx = 0;

  const header = matrix[headerIdx].map(c => String(c == null ? "" : c).trim());

  let iYear = pick(header, "일시");
  if (iYear === -1) iYear = pick(header, "연도");
  if (iYear === -1) iYear = pick(header, "년");
  let iValue = pick(header, "최다", "강수량");
  if (iValue === -1) iValue = pick(header, "강수량");
  let iDate = pick(header, "나타난날");
  if (iDate === -1) iDate = pick(header, "발생일");
  if (iDate === -1) iDate = pick(header, "일자");
  const iName = pick(header, "지점명");
  let iStation = -1;
  for (let i = 0; i < header.length; i++) {
    const h = header[i].replace(/\s/g, "");
    if (h === "지점" || (h.includes("지점") && !h.includes("지점명"))) { iStation = i; break; }
  }

  // 헤더 인식 실패 시 위치 기반 폴백
  if (iYear === -1 && iValue === -1) {
    return positionalRows(matrix.slice(headerIdx + 1));
  }
  if (iYear === -1) iYear = 2;
  if (iValue === -1) iValue = 3;
  if (iDate === -1) iDate = 4;

  const out = { station: "", name: "", years: [] };
  for (let r = headerIdx + 1; r < matrix.length; r++) {
    const cells = matrix[r];
    if (!cells || cells.every(c => String(c).trim() === "")) continue;
    const yearStr = String(cells[iYear] == null ? "" : cells[iYear]).trim();
    const yearMatch = yearStr.match(/\d{4}/);
    if (!yearMatch) continue;
    const value = parseFloat(String(cells[iValue]).replace(/[^0-9.\-]/g, ""));
    if (!isFinite(value)) continue;
    if (!out.name && iName !== -1 && cells[iName]) out.name = String(cells[iName]).trim();
    if (!out.station && iStation !== -1 && cells[iStation]) out.station = String(cells[iStation]).trim();
    out.years.push({
      year: parseInt(yearMatch[0], 10),
      value,
      date: iDate !== -1 ? toISODate(cells[iDate]) : "",
    });
  }
  return out;
}

function positionalRows(dataRows) {
  const out = { station: "", name: "", years: [] };
  for (const cells of dataRows) {
    if (!cells || cells.every(c => String(c).trim() === "")) continue;
    const yearMatch = String(cells[2] ?? cells[0] ?? "").match(/\d{4}/);
    if (!yearMatch) continue;
    const value = parseFloat(String(cells[3] ?? "").replace(/[^0-9.\-]/g, ""));
    if (!isFinite(value)) continue;
    if (!out.name && cells[1]) out.name = String(cells[1]).trim();
    if (!out.station && cells[0]) out.station = String(cells[0]).trim();
    out.years.push({ year: parseInt(yearMatch[0], 10), value, date: toISODate(cells[4]) });
  }
  return out;
}

function looksLikeZip(buffer) {
  return buffer.length > 3 && buffer[0] === 0x50 && buffer[1] === 0x4b; // 'PK' → xlsx
}

async function extractWeather(buffer, originalName = "") {
  const isXlsx = looksLikeZip(buffer) || /\.xlsx$/i.test(originalName);

  let matrix;
  if (isXlsx) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    matrix = [];
    ws.eachRow({ includeEmpty: false }, row => {
      const cells = [];
      row.eachCell({ includeEmpty: true }, cell => {
        let v = cell.value;
        if (v && typeof v === "object") {
          if (v.text) v = v.text;
          else if (v.result != null) v = v.result;
          else if (v.richText) v = v.richText.map(t => t.text).join("");
          else if (v instanceof Date) v = v.toISOString().slice(0, 10);
        }
        cells.push(v == null ? "" : v);
      });
      matrix.push(cells);
    });
  } else {
    const text = decodeBuffer(buffer);
    matrix = text.split(/\r?\n/).filter(l => l.trim() !== "").map(splitLine);
  }

  const data = rowsFromMatrix(matrix);
  data.years.sort((a, b) => a.year - b.year);
  return data;
}

module.exports = { extractWeather, toISODate };
