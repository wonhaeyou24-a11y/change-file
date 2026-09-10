/**
 * 만능 양식 모드: 여러 AI 제공자로 "양식 → 최종 표(JSON)" 변환.
 *
 * 지원 제공자: gemini(기본) / claude / openai
 * API 키는 이 함수 호출 동안 메모리에서만 사용하고 로그·디스크·환경변수에 저장하지 않는다.
 */

// models 배열은 UI 추천 목록일 뿐이다. 실제로는 아래 어느 모델 ID든 입력할 수 있고,
// "모델 불러오기"(GET /api/ai-models) 로 해당 키가 쓸 수 있는 최신 목록을 받아올 수 있다.
const PROVIDERS = {
  gemini: {
    label: "Google Gemini",
    default: "gemini-flash-latest",
    models: ["gemini-flash-latest", "gemini-3.5-flash", "gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-pro", "gemini-2.5-flash"],
    supportsPdf: true,
  },
  claude: {
    label: "Anthropic Claude",
    default: "claude-sonnet-5",
    models: ["claude-sonnet-5", "claude-opus-5", "claude-haiku-4-5"],
    supportsPdf: true,
  },
  openai: {
    label: "OpenAI",
    default: "gpt-4o",
    models: ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"],
    supportsPdf: false,
  },
};

/** 제공자별 사용 가능한 모델 목록 조회 (사용자 키 사용, 미저장) */
async function listModels(provider, apiKey) {
  const key = (apiKey || "").trim();
  if (key.length < 20) throw fail("먼저 API 키를 입력해주세요.");
  try {
    if (provider === "gemini") {
      const r = await fetch("https://generativelanguage.googleapis.com/v1beta/models?pageSize=200", {
        headers: { "x-goog-api-key": key },
      });
      const d = await r.json();
      if (!r.ok) throw fail(`Gemini: ${(d.error && d.error.message) || r.status}`);
      return (d.models || [])
        .filter(m => (m.supportedGenerationMethods || []).includes("generateContent") && !/embedding|aqa|imagen/i.test(m.name))
        .map(m => m.name.replace(/^models\//, ""))
        .sort().reverse();
    }
    if (provider === "openai") {
      const r = await fetch("https://api.openai.com/v1/models", { headers: { authorization: `Bearer ${key}` } });
      const d = await r.json();
      if (!r.ok) throw fail(`OpenAI: ${(d.error && d.error.message) || r.status}`);
      return (d.data || []).map(m => m.id).filter(id => /^(gpt|o\d|chatgpt)/.test(id)).sort();
    }
    if (provider === "claude") {
      const r = await fetch("https://api.anthropic.com/v1/models?limit=100", {
        headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
      });
      const d = await r.json();
      if (!r.ok) throw fail(`Claude: ${(d.error && d.error.message) || r.status}`);
      return (d.data || []).map(m => m.id);
    }
  } catch (e) {
    if (e.userFacing) throw e;
    throw fail("모델 목록을 불러오지 못했습니다: " + (e.message || e));
  }
  throw fail("알 수 없는 제공자입니다.");
}

const SYSTEM_PROMPT = `당신은 한국 과업보고서·엔지니어링 문서의 "표 편집 엔진"이다.
입력으로 (1) 양식(빈 표의 이미지 또는 셀 배치)과 (2) 원본 자료(TSV)를 받는다.
양식의 구조를 그대로 재현하면서 원본 자료로 표를 채운 결과를 JSON 으로만 출력한다.

규칙:
- 양식의 열 구성·머리글 문구·병합 셀·좌우 분할 블록·회색 라벨 칸·합계/요약 행을 똑같이 재현한다.
- 본문 값은 원본 자료에서 가져온다. 숫자·문자를 임의로 만들거나 바꾸지 않는다.
  단, 양식의 다른 숫자가 고정 소수 자리(예: 99.0)를 쓰면 같은 자리수로 맞춘다.
- 계산 칸(합계·평균·개수·기간·최대·최소 등)은 원본 자료로부터 계산한다.
  "누년 평균값" = (해당 값들의 합) ÷ (연도 개수). 소수 자리는 표의 다른 값과 동일하게.
- 날짜는 양식 형식에 맞춘다 (예: 19960625 → 1996-06-25).
- 양식이 N개 행을 K개 좌우 블록으로 나눠 배치하면 순서대로 분배한다.
- rowspan/colspan 으로 가려지는 칸은 HTML 처럼 생략한다.
- 출력은 JSON 객체 하나. 코드펜스나 설명 문장을 붙이지 않는다.

JSON 스키마:
{
  "title": string,
  "columns": number,
  "widths": number[],
  "rows": [ { "cells": [ { "text": string, "fill": boolean, "bold": boolean,
             "align": "center"|"left"|"right", "colspan": number, "rowspan": number } ] } ],
  "computed": { "sum": number|null, "count": number|null, "average": string|null, "period": string|null },
  "notes": string
}`;

function fail(msg) {
  const e = new Error(msg);
  e.userFacing = true;
  return e;
}

function parseJsonLoose(text) {
  let t = String(text).trim();
  t = t.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "");
  const s = t.indexOf("{");
  const e = t.lastIndexOf("}");
  if (s !== -1 && e !== -1 && e > s) t = t.slice(s, e + 1);
  return JSON.parse(t);
}

/** template(analyzeTemplate 결과) + 자료 → 중립 parts 배열 */
function buildParts(template, dataTSV, instructions) {
  const parts = [];
  if (template.kind === "image") {
    parts.push({ image: { mime: template.mediaType, b64: template.dataB64 } });
    parts.push({ text: "↑ 위 이미지가 채워야 할 빈 양식이다." });
  } else if (template.kind === "pdf") {
    parts.push({ pdf: { b64: template.dataB64 } });
    parts.push({ text: "↑ 위 PDF 안의 표가 양식이다." });
  } else {
    parts.push({ text: `=== 양식 ===\n${template.text}` });
  }
  parts.push({ text: `=== 원본 자료 (TSV) ===\n${dataTSV}` });
  if (instructions && instructions.trim()) {
    parts.push({ text: `=== 추가 지시사항 ===\n${instructions.trim()}` });
  }
  parts.push({ text: "위 규칙과 스키마대로 JSON 객체 하나만 출력하라." });
  return parts;
}

/* ------------------------- Gemini ------------------------- */

async function callGemini(system, turns, apiKey, model) {
  const contents = turns.map(t => ({
    role: t.role === "model" ? "model" : "user",
    parts: t.parts.map(p => {
      if (p.text != null) return { text: p.text };
      if (p.image) return { inlineData: { mimeType: p.image.mime, data: p.image.b64 } };
      if (p.pdf) return { inlineData: { mimeType: "application/pdf", data: p.pdf.b64 } };
      return { text: "" };
    }),
  }));
  const generationConfig = { temperature: 0, maxOutputTokens: 48000, responseMimeType: "application/json" };

  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,
    {
      method: "POST",
      headers: { "content-type": "application/json", "x-goog-api-key": apiKey },
      body: JSON.stringify({ systemInstruction: { parts: [{ text: system }] }, contents, generationConfig }),
    }
  );
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const m = data && data.error && data.error.message;
    if (res.status === 400 && /API key not valid/i.test(m || "")) throw fail("Gemini API 키가 유효하지 않습니다.");
    if (res.status === 403) throw fail("이 Gemini 키로는 해당 모델을 쓸 수 없습니다.");
    if (res.status === 429) throw fail("Gemini 사용량 한도에 걸렸습니다. 잠시 후 다시 시도해주세요.");
    throw fail(`Gemini 오류: ${m || res.status}`);
  }
  const cand = data.candidates && data.candidates[0];
  if (!cand) throw fail("Gemini가 빈 응답을 보냈습니다. 다시 시도해주세요.");
  if (cand.finishReason === "MAX_TOKENS") throw fail("표가 커서 응답이 잘렸습니다. 자료 행 수를 줄여주세요.");
  return (cand.content && cand.content.parts || []).map(p => p.text || "").join("");
}

/* ------------------------- Claude ------------------------- */

async function callClaude(system, turns, apiKey, model) {
  const Anthropic = require("@anthropic-ai/sdk");
  const client = new Anthropic({ apiKey });
  const messages = turns.map(t => ({
    role: t.role === "model" ? "assistant" : "user",
    content: t.parts.map(p => {
      if (p.text != null) return { type: "text", text: p.text };
      if (p.image) return { type: "image", source: { type: "base64", media_type: p.image.mime, data: p.image.b64 } };
      if (p.pdf) return { type: "document", source: { type: "base64", media_type: "application/pdf", data: p.pdf.b64 } };
      return { type: "text", text: "" };
    }),
  }));
  let resp;
  try {
    resp = await client.messages.create({ model, max_tokens: 16000, system, messages });
  } catch (e) {
    const map = {
      401: "Claude API 키가 유효하지 않습니다.",
      403: "이 Claude 키로는 해당 모델을 쓸 수 없습니다.",
      429: "Claude 사용량 한도에 걸렸습니다. 잠시 후 다시 시도해주세요.",
      529: "Claude 서버가 혼잡합니다. 잠시 후 다시 시도해주세요.",
    };
    throw fail(map[e && e.status] || (e && e.message) || "Claude 호출 실패");
  }
  if (resp.stop_reason === "max_tokens") throw fail("표가 커서 응답이 잘렸습니다. 자료 행 수를 줄여주세요.");
  return resp.content.filter(b => b.type === "text").map(b => b.text).join("");
}

/* ------------------------- OpenAI ------------------------- */

async function callOpenAI(system, turns, apiKey, model) {
  const messages = [{ role: "system", content: system }];
  for (const t of turns) {
    const role = t.role === "model" ? "assistant" : "user";
    const content = t.parts.map(p => {
      if (p.text != null) return { type: "text", text: p.text };
      if (p.image) return { type: "image_url", image_url: { url: `data:${p.image.mime};base64,${p.image.b64}` } };
      if (p.pdf) throw fail("OpenAI는 PDF 양식을 지원하지 않습니다. 이미지(png/jpg)나 엑셀로 올려주세요.");
      return { type: "text", text: "" };
    });
    messages.push({ role, content });
  }
  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({ model, messages, temperature: 0, max_tokens: 16000, response_format: { type: "json_object" } }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const m = data && data.error && data.error.message;
    if (res.status === 401) throw fail("OpenAI API 키가 유효하지 않습니다.");
    if (res.status === 429) throw fail("OpenAI 사용량 한도에 걸렸습니다. 잠시 후 다시 시도해주세요.");
    throw fail(`OpenAI 오류: ${m || res.status}`);
  }
  const choice = data.choices && data.choices[0];
  if (!choice) throw fail("OpenAI가 빈 응답을 보냈습니다.");
  if (choice.finish_reason === "length") throw fail("표가 커서 응답이 잘렸습니다. 자료 행 수를 줄여주세요.");
  return choice.message && choice.message.content || "";
}

const DISPATCH = { gemini: callGemini, claude: callClaude, openai: callOpenAI };

/* ------------------------- 진입점 ------------------------- */

async function aiBuildTable({ provider, apiKey, model, template, dataTSV, instructions }) {
  provider = PROVIDERS[provider] ? provider : "gemini";
  const conf = PROVIDERS[provider];
  const useModel = (model && String(model).trim()) || conf.default;
  if (!apiKey || apiKey.trim().length < 20) {
    throw fail(`${conf.label} API 키를 입력해주세요.`);
  }
  if (template.kind === "pdf" && !conf.supportsPdf) {
    throw fail(`${conf.label}는 PDF 양식을 지원하지 않습니다. 이미지나 엑셀로 올려주세요.`);
  }
  const key = apiKey.trim();
  const run = DISPATCH[provider];

  const turns = [{ role: "user", parts: buildParts(template, dataTSV, instructions) }];

  for (let attempt = 0; attempt < 2; attempt++) {
    const raw = await run(SYSTEM_PROMPT, turns, key, useModel);
    try {
      const parsed = parseJsonLoose(raw);
      if (!parsed.rows || !Array.isArray(parsed.rows)) throw new Error("rows 없음");
      return { grid: parsed, provider, model: useModel };
    } catch (e) {
      if (attempt === 1) throw fail("AI 응답을 표로 해석하지 못했습니다. 다시 시도하거나 다른 모델을 선택해보세요.");
      turns.push({ role: "model", parts: [{ text: String(raw).slice(0, 8000) }] });
      turns.push({ role: "user", parts: [{ text: "스키마에 맞는 JSON 객체 하나만 다시 출력하라. 설명·코드펜스 금지." }] });
    }
  }
}

module.exports = { aiBuildTable, listModels, PROVIDERS };
