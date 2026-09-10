const Anthropic = require("@anthropic-ai/sdk");

/**
 * AI(Claude)로 "양식 → 최종 표" 변환.
 *
 * API 키는 이 함수 호출 동안 메모리에서만 사용되고, 로그/디스크/환경변수에
 * 절대 저장하지 않는다. (server.js 에서도 마찬가지)
 */

const ALLOWED_MODELS = new Set(["claude-opus-5", "claude-sonnet-5", "claude-haiku-4-5"]);

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
- 양식이 N개 행을 K개 좌우 블록으로 나눠 배치하면, 순서대로 분배한다
  (블록1 = 1~ceil(N/K)행, 블록2 = 그다음 …).
- rowspan/colspan 으로 가려지는 칸은 HTML 처럼 생략한다 (그 자리에 cell 을 넣지 않는다).
- 출력은 JSON 객체 하나. 마크다운 코드펜스나 설명 문장을 앞뒤에 붙이지 않는다.

JSON 스키마:
{
  "title": string,               // 표 위 캡션. 없으면 ""
  "columns": number,             // 열 개수
  "widths": number[],            // 길이 == columns, 상대 폭(문자 수 기준 대략치)
  "rows": [
    { "cells": [
        { "text": string,
          "fill": boolean,       // 회색 배경(머리글/라벨) 여부
          "bold": boolean,
          "align": "center" | "left" | "right",
          "colspan": number,     // 기본 1
          "rowspan": number }    // 기본 1
    ] }
  ],
  "computed": { "sum": number|null, "count": number|null, "average": string|null, "period": string|null },
  "notes": string                // 한국어. 원본 열 → 양식 칸 매핑과 가정을 2~4줄로.
}`;

function parseJsonLoose(text) {
  let t = String(text).trim();
  t = t.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "");
  const s = t.indexOf("{");
  const e = t.lastIndexOf("}");
  if (s !== -1 && e !== -1 && e > s) t = t.slice(s, e + 1);
  return JSON.parse(t);
}

function templateBlocks(template) {
  if (template.kind === "image") {
    return [{
      type: "image",
      source: { type: "base64", media_type: template.mediaType, data: template.dataB64 },
    }, { type: "text", text: "↑ 이 이미지가 채워야 할 빈 양식이다." }];
  }
  if (template.kind === "pdf") {
    return [{
      type: "document",
      source: { type: "base64", media_type: "application/pdf", data: template.dataB64 },
    }, { type: "text", text: "↑ 이 PDF 안의 표가 양식이다." }];
  }
  return [{ type: "text", text: `=== 양식 ===\n${template.text}` }];
}

async function aiBuildTable({ apiKey, model, template, dataTSV, instructions }) {
  if (!apiKey || !/^sk-/.test(apiKey.trim())) {
    const err = new Error("유효한 Claude API 키(sk-...)를 입력해주세요.");
    err.userFacing = true;
    throw err;
  }
  const useModel = ALLOWED_MODELS.has(model) ? model : "claude-opus-5";
  const client = new Anthropic({ apiKey: apiKey.trim() });

  const userContent = [
    ...templateBlocks(template),
    { type: "text", text: `=== 원본 자료 (TSV) ===\n${dataTSV}` },
  ];
  if (instructions && instructions.trim()) {
    userContent.push({ type: "text", text: `=== 추가 지시사항 ===\n${instructions.trim()}` });
  }
  userContent.push({ type: "text", text: "위 규칙과 스키마대로 JSON 만 출력하라." });

  const messages = [{ role: "user", content: userContent }];

  const call = (msgs) => client.messages.create({
    model: useModel, max_tokens: 16000, system: SYSTEM_PROMPT, messages: msgs,
  }).catch(e => {
    const status = e && e.status;
    const map = {
      401: "API 키가 유효하지 않습니다. 키를 다시 확인해주세요.",
      403: "이 API 키로는 해당 모델을 쓸 수 없습니다.",
      429: "API 사용량 한도에 걸렸습니다. 잠시 후 다시 시도해주세요.",
      529: "Claude 서버가 혼잡합니다. 잠시 후 다시 시도해주세요.",
    };
    const err = new Error(map[status] || (e && e.message) || "AI 호출에 실패했습니다.");
    err.userFacing = true;
    throw err;
  });

  let raw;
  for (let attempt = 0; attempt < 2; attempt++) {
    const resp = await call(messages);
    raw = resp.content.filter(b => b.type === "text").map(b => b.text).join("");
    if (resp.stop_reason === "max_tokens") {
      const err = new Error("표가 너무 커서 응답이 잘렸습니다. 자료 행 수를 줄이거나 다시 시도해주세요.");
      err.userFacing = true;
      throw err;
    }
    try {
      const parsed = parseJsonLoose(raw);
      if (!parsed.rows || !Array.isArray(parsed.rows)) throw new Error("rows 없음");
      return { grid: parsed, model: useModel };
    } catch (e) {
      if (attempt === 1) {
        const err = new Error("AI 응답을 표로 해석하지 못했습니다. 다시 시도해주세요.");
        err.userFacing = true;
        throw err;
      }
      messages.push({ role: "assistant", content: raw });
      messages.push({ role: "user", content: "스키마에 맞는 JSON 객체 하나만 다시 출력하라. 설명 금지." });
    }
  }
}

module.exports = { aiBuildTable, ALLOWED_MODELS };
