/**
 * AI(Claude API) 없이 순수 규칙으로 원본 셀 텍스트를 불릿 리스트로 정리한다.
 *
 * 목적은 "양식 구성(표 구조)"을 원본 이미지와 동일하게 맞추는 것이 1차 목표이므로,
 * 이 모듈은 단어/문장 내용을 절대 바꾸거나 생략하지 않는다. 유일하게 손대는 부분은
 * (1) 줄바꿈을 문장 단위로 다시 묶는 것과 (2) 그 과정에서 줄과 줄 사이에 공백 1칸을
 * 넣는 것뿐이다. 공백 위치가 100% 정확하지 않을 수 있으므로(원래 단어 중간이 끊긴
 * 경우에도 공백이 들어갈 수 있음) 최종 검토 시 띄어쓰기만 다듬으면 된다 — 글자 자체는
 * 항상 원문 그대로 보존된다.
 */

// 줄 맨 앞에 붙는 제각각의 불릿 문자
const LEADING_BULLET_RE = /^[ㆍ\-？·○●▶]+\s*/;

/**
 * 원본 셀 텍스트 하나를 정리된 문자열 배열(불릿 하나 = 배열 항목 하나)로 변환.
 * 반환값에는 불릿 기호(-)를 붙이지 않는다 (렌더링 단계에서 붙임).
 *
 * 이 데이터(안전점검 실적)에서 관찰되는 두 가지 입력 패턴을 구분해서 처리한다:
 *
 *  (A) 어느 줄에도 불릿 문자가 없는 경우 — 이 경우 줄바꿈마다 각각 독립된 항목이다
 *      (예: "산마루측구 단차 발생\n낙석방지망 상태 양호\n배수구 상태 양호" → 3개 항목).
 *
 *  (B) 일부/모든 줄에 불릿 문자(ㆍ？·○-)가 있는 경우 — 불릿 문자가 있는 줄만 "새 항목
 *      시작"으로 보고, 불릿 문자가 없는 줄은 바로 앞 항목이 엑셀 셀 너비 때문에 줄바꿈된
 *      것으로 보아 앞 항목에 공백을 넣어 이어붙인다
 *      (예: "ㆍ...부분\n적으로 관찰됨" → "...부분 적으로 관찰됨" 한 항목).
 *
 * 두 경우 모두 원문 단어/글자를 바꾸거나 생략하지 않는다 — 항목을 나누는 위치와 이어붙일
 * 때 공백을 넣는 것만 판단한다. (B)에서 앞 항목 자체가 이미 불릿 문자를 달고 줄바꿈된
 * 경우(예: 모든 줄에 불릿이 있는데 그 중 한 줄이 실제로는 이전 줄의 연속인 경우)까지는
 * 기계적으로 완벽히 구분할 수 없으므로, 이런 드문 경우는 항목이 예상보다 더 잘게 쪼개질
 * 수 있다 — 이 때도 글자는 그대로이므로 검토 시 줄바꿈만 합쳐주면 된다.
 */
function cleanCell(raw) {
  if (raw == null) return [];
  const text = String(raw).trim();
  if (!text) return [];

  const lines = text.split(/\r?\n/);
  const hasAnyMarker = lines.some(l => LEADING_BULLET_RE.test(l.trim()));

  if (!hasAnyMarker) {
    // 패턴 (A): 줄마다 독립된 항목
    const items = lines.map(l => l.trim()).filter(l => l !== "");
    return items.length ? items : [text];
  }

  // 패턴 (B): 불릿 문자가 있는 줄만 새 항목, 없는 줄은 이전 항목에 이어붙임
  const items = [];
  let buffer = null;
  const flush = () => {
    if (buffer !== null && buffer.trim() !== "") items.push(buffer.trim());
    buffer = null;
  };
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (line === "") { flush(); continue; }
    const hasMarker = LEADING_BULLET_RE.test(line);
    const stripped = line.replace(LEADING_BULLET_RE, "");
    if (hasMarker || buffer === null) {
      flush();
      buffer = stripped;
    } else {
      buffer = buffer + " " + stripped;
    }
  }
  flush();
  return items.length ? items : [text.replace(LEADING_BULLET_RE, "")];
}

function cleanRows(rows, format) {
  return rows.map(r => {
    if (format === "wall") {
      return {
        ...r,
        result: cleanCell(r.resultRaw).length ? cleanCell(r.resultRaw) : [""],
        remedy: r.remedyRaw ? cleanCell(r.remedyRaw) : [],
      };
    }
    return {
      ...r,
      bullets: cleanCell(r.resultRaw).length ? cleanCell(r.resultRaw) : [""],
    };
  });
}

module.exports = { cleanRows, cleanCell };
