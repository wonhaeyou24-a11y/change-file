const form = document.getElementById("convertForm");
const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const submitBtn = document.getElementById("submitBtn");

fileInput.addEventListener("change", () => {
  fileName.textContent = fileInput.files[0] ? fileInput.files[0].name : "파일을 선택하거나 여기로 드래그하세요";
});

const formatSel = document.getElementById("format");
const fieldsDefault = document.getElementById("fields-default");
const fieldsWeather = document.getElementById("fields-weather");
const fieldsUniversal = document.getElementById("fields-universal");
const dropLabel = document.getElementById("dropLabel");
const templateInput = document.getElementById("templateInput");
const templateName = document.getElementById("templateName");

templateInput.addEventListener("change", () => {
  templateName.textContent = templateInput.files[0] ? templateInput.files[0].name : "양식 파일을 선택하세요";
});

function syncFormatUI() {
  const f = formatSel.value;
  fieldsWeather.hidden = f !== "weather";
  fieldsUniversal.hidden = f !== "universal";
  fieldsDefault.hidden = f === "weather" || f === "universal";
  if (f === "weather") {
    dropLabel.textContent = "기상청 '일 최다 강수량' 자료 (.xls / .csv / .xlsx)";
    fileInput.setAttribute("accept", ".xlsx,.xls,.csv,.txt");
  } else if (f === "universal") {
    dropLabel.textContent = "정리할 자료 파일 (.xlsx / .xls / .csv / .hwpx)";
    fileInput.setAttribute("accept", ".xlsx,.xlsm,.xls,.csv,.txt,.tsv,.hwpx,.hwp");
  } else {
    dropLabel.textContent = "원본 엑셀 파일 (.xlsx)";
    fileInput.setAttribute("accept", ".xlsx,.xls,.csv,.txt");
  }
}
formatSel.addEventListener("change", syncFormatUI);
syncFormatUI();

function showStatus(msg, type = "info") {
  statusEl.hidden = false;
  statusEl.textContent = msg;
  statusEl.className = `status ${type}`;
}

function base64ToBlobUrl(base64, mime) {
  const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
  return URL.createObjectURL(new Blob([bytes], { type: mime }));
}

const esc = s => (s ?? "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;");

function renderWeatherPreview(data) {
  const container = document.getElementById("previewTable");
  const head = `<tr>${["연도", "발생일", "연별 최다 일강수량(mm)"].map(h => `<th>${h}</th>`).join("").repeat(3)}</tr>`;
  const body = (data.grid || []).map(r =>
    `<tr>${r.map(c => `<td>${c === "" ? "&nbsp;" : esc(c)}</td>`).join("")}</tr>`).join("");
  const s = data.summary || {};
  const foot = `<tr><td><b>위치</b></td><td>${esc(s.name)}</td><td><b>기간</b></td>`
    + `<td colspan="4">${esc(s.period)}</td><td><b>누년 평균값 (mm)</b></td><td><b>${esc(s.avg)}</b></td></tr>`;
  container.innerHTML = `<table>${head}${body}${foot}</table>`;
}

/** 범용 grid 모델({rows:[{cells:[{text,fill,bold,align,colspan,rowspan}]}]}) → HTML 표 */
function renderGridPreview(grid) {
  const container = document.getElementById("previewTable");
  const rows = (grid.rows || []).map(r => {
    const cells = (r.cells || []).map(c => {
      const tag = c.fill ? "th" : "td";
      const attrs = [];
      if (c.colspan > 1) attrs.push(`colspan="${c.colspan}"`);
      if (c.rowspan > 1) attrs.push(`rowspan="${c.rowspan}"`);
      const style = [];
      if (c.align) style.push(`text-align:${c.align}`);
      if (c.bold) style.push("font-weight:700");
      if (style.length) attrs.push(`style="${style.join(";")}"`);
      return `<${tag} ${attrs.join(" ")}>${esc(c.text).replace(/\n/g, "<br>")}</${tag}>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
  const cap = grid.title ? `<caption style="font-weight:700;margin-bottom:6px">${esc(grid.title)}</caption>` : "";
  container.innerHTML = `<table>${cap}${rows}</table>`;
}

function renderPreview(rows, format) {
  const container = document.getElementById("previewTable");
  if (!rows || !rows.length) { container.innerHTML = ""; return; }

  let headers, cellsFor;
  if (format === "wall") {
    headers = ["No", "수행기간", "구분", "기관", "기술자", "등급", "점검결과", "보수보강안"];
    cellsFor = r => [r.no, `${r.start} ~ ${r.end}`, r.type, r.org, r.engineer, r.grade,
      (r.result || []).map(b => `-${b}`).join("\n"), (r.remedy || []).map(b => `-${b}`).join("\n")];
  } else {
    headers = ["No", "점검종류", "점검기간", "점검기관", "등급", "점검결과"];
    cellsFor = r => [r.no, r.type, `${r.start} ~ ${r.end}`, r.org, r.grade,
      (r.bullets || []).map(b => `-${b}`).join("\n")];
  }

  const thead = `<tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr>`;
  const tbody = rows.map(r => `<tr>${cellsFor(r).map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("");
  container.innerHTML = `<table>${thead}${tbody}</table>`;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  resultEl.hidden = true;
  statusEl.hidden = true;

  const file = fileInput.files[0];
  if (!file) return;
  const format = formatSel.value;
  const universal = format === "universal";

  const fd = new FormData();
  fd.append("file", file);
  fd.append("format", format);

  let endpoint = "/api/convert";
  if (universal) {
    endpoint = "/api/convert-universal";
    const tmpl = templateInput.files[0];
    const key = document.getElementById("apiKey").value.trim();
    if (!tmpl) { showStatus("양식 파일을 선택해주세요.", "error"); return; }
    if (!key) { showStatus("Claude API 키를 입력해주세요.", "error"); return; }
    fd.append("template", tmpl);
    fd.append("apiKey", key);
    fd.append("model", document.getElementById("model").value);
    const ins = document.getElementById("instructions").value.trim();
    if (ins) fd.append("instructions", ins);
  } else if (format === "weather") {
    const sn = document.getElementById("stationName").value.trim();
    const tn = document.getElementById("tableNo").value.trim();
    const cap = document.getElementById("caption").value.trim();
    if (sn) fd.append("stationName", sn);
    if (tn) fd.append("tableNo", tn);
    if (cap) fd.append("caption", cap);
  } else {
    const title = document.getElementById("title").value;
    if (title) fd.append("title", title);
  }

  submitBtn.disabled = true;
  showStatus(universal ? "AI가 양식을 분석해 표를 만드는 중입니다... (10~40초)" : "변환 중입니다...", "info");

  try {
    const res = await fetch(endpoint, { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "변환에 실패했습니다.");

    statusEl.hidden = true;
    resultEl.hidden = false;

    const notesEl = document.getElementById("aiNotes");
    if (universal) {
      const c = data.computed || {};
      const bits = [`모델: ${data.model}`];
      if (c.average != null) bits.push(`평균: ${c.average}`);
      if (c.count != null) bits.push(`개수: ${c.count}`);
      if (c.period) bits.push(`기간: ${c.period}`);
      document.getElementById("rowCountText").textContent = bits.join(" · ");
      notesEl.hidden = !data.notes;
      notesEl.textContent = data.notes ? `AI 메모: ${data.notes}` : "";
    } else {
      notesEl.hidden = true;
      document.getElementById("rowCountText").textContent =
        format === "weather"
          ? `${data.rowCount}개년 변환 완료 · 누년 평균값 ${data.summary.avg}mm`
          : `총 ${data.rowCount}건 변환 완료`;
    }

    const docxLink = document.getElementById("docxLink");
    const xlsxLink = document.getElementById("xlsxLink");
    docxLink.href = base64ToBlobUrl(data.files.docx.data, data.files.docx.mime);
    docxLink.setAttribute("download", data.files.docx.name);
    xlsxLink.href = base64ToBlobUrl(data.files.xlsx.data, data.files.xlsx.mime);
    xlsxLink.setAttribute("download", data.files.xlsx.name);

    if (universal) renderGridPreview(data.grid);
    else if (format === "weather") renderWeatherPreview(data);
    else renderPreview(data.preview, format);
  } catch (err) {
    showStatus(err.message, "error");
  } finally {
    submitBtn.disabled = false;
  }
});
