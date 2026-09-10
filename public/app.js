const form = document.getElementById("convertForm");
const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const submitBtn = document.getElementById("submitBtn");

fileInput.addEventListener("change", () => {
  fileName.textContent = fileInput.files[0] ? fileInput.files[0].name : "파일을 선택하거나 여기로 드래그하세요";
});

function showStatus(msg, type = "info") {
  statusEl.hidden = false;
  statusEl.textContent = msg;
  statusEl.className = `status ${type}`;
}

function base64ToBlobUrl(base64, mime) {
  const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
  return URL.createObjectURL(new Blob([bytes], { type: mime }));
}

function renderPreview(rows, format) {
  const container = document.getElementById("previewTable");
  if (!rows.length) { container.innerHTML = ""; return; }

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
  const tbody = rows.map(r => `<tr>${cellsFor(r).map(c => `<td>${(c ?? "").toString()
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")}</td>`).join("")}</tr>`).join("");
  container.innerHTML = `<table>${thead}${tbody}</table>`;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  resultEl.hidden = true;
  statusEl.hidden = true;

  const file = fileInput.files[0];
  if (!file) return;

  const format = document.getElementById("format").value;
  const title = document.getElementById("title").value;

  const fd = new FormData();
  fd.append("file", file);
  fd.append("format", format);
  if (title) fd.append("title", title);

  submitBtn.disabled = true;
  showStatus("변환 중입니다...", "info");

  try {
    const res = await fetch("/api/convert", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "변환에 실패했습니다.");

    statusEl.hidden = true;
    resultEl.hidden = false;
    document.getElementById("rowCountText").textContent = `총 ${data.rowCount}건 변환 완료`;

    const docxLink = document.getElementById("docxLink");
    const xlsxLink = document.getElementById("xlsxLink");
    docxLink.href = base64ToBlobUrl(data.files.docx.data, data.files.docx.mime);
    docxLink.setAttribute("download", data.files.docx.name);
    xlsxLink.href = base64ToBlobUrl(data.files.xlsx.data, data.files.xlsx.mime);
    xlsxLink.setAttribute("download", data.files.xlsx.name);

    renderPreview(data.preview, format);
  } catch (err) {
    showStatus(err.message, "error");
  } finally {
    submitBtn.disabled = false;
  }
});
