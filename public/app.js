const form = document.getElementById("convertForm");
const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const submitBtn = document.getElementById("submitBtn");

const modeSel = document.getElementById("mode");
const modeBasic = document.getElementById("mode-basic");
const modeUniversal = document.getElementById("mode-universal");
const basicFormatSel = document.getElementById("basicFormat");
const basicDefault = document.getElementById("basic-default");
const basicWeather = document.getElementById("basic-weather");
const dropLabel = document.getElementById("dropLabel");

const templateInput = document.getElementById("templateInput");
const templateName = document.getElementById("templateName");
const providerSel = document.getElementById("provider");
const modelSel = document.getElementById("model");
const loadModelsBtn = document.getElementById("loadModels");
const apiKeyInput = document.getElementById("apiKey");
const keyLabel = document.getElementById("keyLabel");
const keyGetHint = document.getElementById("keyGetHint");

const KEY_HINTS = {
  gemini: "키 발급: https://aistudio.google.com/apikey",
  claude: "키 발급: https://console.anthropic.com/settings/keys",
  openai: "키 발급: https://platform.openai.com/api-keys",
};
const FALLBACK_PROVIDERS = {
  gemini: { label: "Google Gemini", default: "gemini-flash-latest", models: ["gemini-flash-latest", "gemini-3.5-flash", "gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-pro"], supportsPdf: true },
  claude: { label: "Anthropic Claude", default: "claude-sonnet-5", models: ["claude-sonnet-5", "claude-opus-5", "claude-haiku-4-5"], supportsPdf: true },
  openai: { label: "OpenAI", default: "gpt-4o", models: ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"], supportsPdf: false },
};
let PROVIDERS = FALLBACK_PROVIDERS;

fileInput.addEventListener("change", () => {
  fileName.textContent = fileInput.files[0] ? fileInput.files[0].name : "파일을 선택하거나 여기로 드래그하세요";
});
templateInput.addEventListener("change", () => {
  templateName.textContent = templateInput.files[0] ? templateInput.files[0].name : "양식 파일을 선택하세요";
});

function populateProviders() {
  providerSel.innerHTML = "";
  for (const [k, v] of Object.entries(PROVIDERS)) {
    const o = document.createElement("option");
    o.value = k;
    o.textContent = v.label + (k === "gemini" ? " (기본)" : "");
    providerSel.appendChild(o);
  }
  providerSel.value = "gemini";
  syncProvider();
}
const CUSTOM = "__custom__";
function fillModelList(models, keep) {
  modelSel.innerHTML = "";
  for (const m of models) {
    const o = document.createElement("option");
    o.value = m; o.textContent = m;
    modelSel.appendChild(o);
  }
  const c = document.createElement("option");
  c.value = CUSTOM; c.textContent = "✎ 직접 입력…";
  modelSel.appendChild(c);
  if (keep && models.includes(keep)) modelSel.value = keep;
  else if (models.length) modelSel.value = models[0];
}
modelSel.addEventListener("change", () => {
  if (modelSel.value !== CUSTOM) { modelSel.dataset.prev = modelSel.value; return; }
  const m = (prompt("사용할 모델 ID를 입력하세요 (예: gemini-3.5-flash)", "") || "").trim();
  if (m) {
    if (![...modelSel.options].some(o => o.value === m)) {
      const o = document.createElement("option");
      o.value = m; o.textContent = m;
      modelSel.insertBefore(o, modelSel.lastChild);
    }
    modelSel.value = m;
    modelSel.dataset.prev = m;
  } else {
    modelSel.value = modelSel.dataset.prev || (modelSel.options[0] && modelSel.options[0].value) || "";
  }
});
function syncProvider() {
  const p = PROVIDERS[providerSel.value] || FALLBACK_PROVIDERS.gemini;
  fillModelList(p.models, p.default);
  keyLabel.textContent = p.label;
  keyGetHint.textContent = KEY_HINTS[providerSel.value] || "";
  templateInput.setAttribute(
    "accept",
    p.supportsPdf ? ".xlsx,.xlsm,.xls,.csv,.png,.jpg,.jpeg,.webp,.pdf" : ".xlsx,.xlsm,.xls,.csv,.png,.jpg,.jpeg,.webp"
  );
}
providerSel.addEventListener("change", syncProvider);

loadModelsBtn.addEventListener("click", async () => {
  const key = apiKeyInput.value.trim();
  if (!key) { showStatus(`${keyLabel.textContent} API 키를 먼저 입력해주세요.`, "error"); return; }
  loadModelsBtn.disabled = true;
  const orig = loadModelsBtn.textContent;
  loadModelsBtn.textContent = "불러오는 중...";
  try {
    const r = await fetch("/api/ai-models", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ provider: providerSel.value, apiKey: key }),
    });
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "실패");
    if (!d.models.length) throw new Error("사용 가능한 모델이 없습니다.");
    const keep = modelSel.value !== CUSTOM ? modelSel.value : "";
    fillModelList(d.models, keep);
    showStatus(`모델 ${d.models.length}개를 불러왔습니다. 아래 목록에서 선택하세요.`, "info");
  } catch (e) {
    showStatus("모델 불러오기 실패: " + e.message, "error");
  } finally {
    loadModelsBtn.disabled = false;
    loadModelsBtn.textContent = orig;
  }
});

async function loadProviders() {
  try {
    const r = await fetch("/api/providers");
    if (r.ok) PROVIDERS = await r.json();
  } catch (_) { /* keep fallback */ }
  populateProviders();
}
loadProviders();

function syncMode() {
  const universal = modeSel.value === "universal";
  modeUniversal.hidden = !universal;
  modeBasic.hidden = universal;
  if (universal) {
    dropLabel.textContent = "정리할 자료 파일 (.xlsx / .xls / .csv / .hwpx)";
    fileInput.setAttribute("accept", ".xlsx,.xlsm,.xls,.csv,.txt,.tsv,.hwpx,.hwp");
  } else {
    syncBasicFormat();
  }
}
function syncBasicFormat() {
  const f = basicFormatSel.value;
  basicWeather.hidden = f !== "weather";
  basicDefault.hidden = f === "weather";
  if (f === "weather") {
    dropLabel.textContent = "기상청 '일 최다 강수량' 자료 (.xls / .csv / .xlsx)";
    fileInput.setAttribute("accept", ".xlsx,.xls,.csv,.txt");
  } else {
    dropLabel.textContent = "원본 엑셀 파일 (.xlsx)";
    fileInput.setAttribute("accept", ".xlsx,.xls,.csv,.txt");
  }
}
modeSel.addEventListener("change", syncMode);
basicFormatSel.addEventListener("change", syncBasicFormat);
syncMode();

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
  const universal = modeSel.value === "universal";
  const basicFormat = basicFormatSel.value;

  const fd = new FormData();
  fd.append("file", file);

  let endpoint = "/api/convert";
  if (universal) {
    endpoint = "/api/convert-universal";
    const tmpl = templateInput.files[0];
    const key = document.getElementById("apiKey").value.trim();
    if (!tmpl) { showStatus("양식 파일을 선택해주세요.", "error"); return; }
    if (!key) { showStatus(`${keyLabel.textContent} API 키를 입력해주세요.`, "error"); return; }
    const model = modelSel.value === CUSTOM ? "" : modelSel.value;
    if (!model) { showStatus("AI 모델을 선택하거나 직접 입력해주세요.", "error"); return; }
    fd.append("template", tmpl);
    fd.append("provider", providerSel.value);
    fd.append("model", model);
    fd.append("apiKey", key);
    const ins = document.getElementById("instructions").value.trim();
    if (ins) fd.append("instructions", ins);
  } else {
    fd.append("format", basicFormat);
    if (basicFormat === "weather") {
      const sn = document.getElementById("stationName").value.trim();
      const cap = document.getElementById("caption").value.trim();
      if (sn) fd.append("stationName", sn);
      if (cap) fd.append("caption", cap);
    } else {
      const title = document.getElementById("title").value;
      if (title) fd.append("title", title);
    }
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
      const bits = [`${(PROVIDERS[data.provider] || {}).label || data.provider} · ${data.model}`];
      if (c.average != null) bits.push(`평균: ${c.average}`);
      if (c.count != null) bits.push(`개수: ${c.count}`);
      if (c.period) bits.push(`기간: ${c.period}`);
      document.getElementById("rowCountText").textContent = bits.join(" · ");
      notesEl.hidden = !data.notes;
      notesEl.textContent = data.notes ? `AI 메모: ${data.notes}` : "";
    } else {
      notesEl.hidden = true;
      document.getElementById("rowCountText").textContent =
        basicFormat === "weather"
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
    else if (basicFormat === "weather") renderWeatherPreview(data);
    else renderPreview(data.preview, basicFormat);
  } catch (err) {
    showStatus(err.message, "error");
  } finally {
    submitBtn.disabled = false;
  }
});
