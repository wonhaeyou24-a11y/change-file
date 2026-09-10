const XLSX = require("xlsx");
const AdmZip = require("adm-zip");

/**
 * 여러 형식의 "자료 파일"을 2차원 배열(행 → 셀 문자열)로 변환한다.
 * 지원: .xlsx .xlsm .xls .csv .txt(탭/쉼표) · .hwpx(한글, zip 기반) · .hwp(구버전 → 안내)
 *
 * 반환: { matrix: string[][], sheetName: string, kind: string }
 */

function looksLikeZip(buf) {
  return buf.length > 3 && buf[0] === 0x50 && buf[1] === 0x4b; // 'PK'
}
function looksLikeOle(buf) {
  // 구버전 .hwp / .xls(BIFF) 공통 OLE2 시그니처
  return buf.length > 8 &&
    buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0;
}

function decode(buf) {
  let t = buf.toString("utf8");
  if (t.includes("�")) {
    try { t = new TextDecoder("euc-kr").decode(buf); } catch (_) { /* keep */ }
  }
  return t;
}

function matrixFromDelimited(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
  const delim = lines[0] && lines[0].includes("\t") ? "\t" : ",";
  return lines.map(l => l.split(delim).map(c => c.trim().replace(/^"|"$/g, "")));
}

function matrixFromWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  return {
    matrix: rows.map(r => r.map(c => (c == null ? "" : String(c).trim()))),
    sheetName: wb.SheetNames[0] || "Sheet1",
  };
}

/** .hwpx: Contents/section*.xml 안의 표(<hp:tbl>) 텍스트를 추출 */
function matrixFromHwpx(buf) {
  const zip = new AdmZip(buf);
  const entries = zip.getEntries()
    .filter(e => /Contents\/section\d*\.xml$/i.test(e.entryName))
    .sort((a, b) => a.entryName.localeCompare(b.entryName));
  if (!entries.length) throw new Error("HWPX 안에서 본문(section) XML을 찾지 못했습니다.");

  const xml = entries.map(e => e.getData().toString("utf8")).join("\n");

  // 가장 큰 표 하나를 사용
  const tables = xml.match(/<hp:tbl[\s\S]*?<\/hp:tbl>/g) || [];
  if (!tables.length) {
    // 표가 없으면 문단 텍스트라도 줄 단위로
    const paras = (xml.match(/<hp:t>([\s\S]*?)<\/hp:t>/g) || [])
      .map(s => s.replace(/<[^>]+>/g, "").trim()).filter(Boolean);
    return { matrix: paras.map(p => [p]), sheetName: "hwpx" };
  }
  const tbl = tables.sort((a, b) => b.length - a.length)[0];
  const rows = tbl.match(/<hp:tr[\s\S]*?<\/hp:tr>/g) || [];
  const matrix = rows.map(tr => {
    const cells = tr.match(/<hp:tc[\s\S]*?<\/hp:tc>/g) || [];
    return cells.map(tc => {
      const texts = tc.match(/<hp:t>([\s\S]*?)<\/hp:t>/g) || [];
      return texts.map(s => s.replace(/<[^>]+>/g, "")).join("").trim();
    });
  }).filter(r => r.some(c => c !== ""));
  return { matrix, sheetName: "hwpx-table" };
}

function extractData(buffer, originalName = "") {
  const name = originalName.toLowerCase();
  const ext = (name.match(/\.[a-z0-9]+$/) || [""])[0];

  if (ext === ".hwpx" || (looksLikeZip(buffer) && ext === ".hwp")) {
    return { ...matrixFromHwpx(buffer), kind: "hwpx" };
  }
  if (ext === ".hwp" || (looksLikeOle(buffer) && ext !== ".xls")) {
    throw new Error(
      "구버전 한글 문서(.hwp)는 표 자동 추출이 불안정합니다. 한글에서 " +
      "[파일 → 다른 이름으로 저장]으로 .hwpx 또는 .xlsx 로 저장한 뒤 올려주세요."
    );
  }
  // 진짜 바이너리 엑셀(zip=xlsx, OLE=xls)만 SheetJS로. 그 외(csv/txt/탭텍스트,
  // 확장자만 .xls 인 기상청 자료 등)는 직접 UTF-8/EUC-KR 로 디코드해 구분자 파싱.
  if (looksLikeZip(buffer) || looksLikeOle(buffer)) {
    try {
      return { ...matrixFromWorkbook(buffer), kind: "workbook" };
    } catch (e) {
      return { matrix: matrixFromDelimited(decode(buffer)), sheetName: "text", kind: "text" };
    }
  }
  return { matrix: matrixFromDelimited(decode(buffer)), sheetName: "text", kind: "text" };
}

/** matrix → TSV 문자열 (AI 입력용, 빈 뒤쪽 열 정리) */
function matrixToTSV(matrix, maxRows = 2000) {
  return matrix.slice(0, maxRows).map(r => {
    const row = r.slice();
    while (row.length && row[row.length - 1] === "") row.pop();
    return row.join("\t");
  }).join("\n");
}

module.exports = { extractData, matrixToTSV };
