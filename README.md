# 보고서 표 변환기 (웹앱)

원본 자료(엑셀 등)를 보고서용 표(Word .docx + Excel .xlsx)로 변환하는 웹앱입니다.

## 양식(모드) 종류

| 모드 | AI | 설명 |
|---|---|---|
| 절토사면형 | ✕ | 점검진단실적 → 이력현황 표 (규칙 기반) |
| 옹벽형 | ✕ | 점검진단실적 → 이력현황 표 (규칙 기반) |
| 기상청 강수량 분석 | ✕ | 기상청 "일 최다 강수량" 자료 → 연도/발생일/최다일강수량 3분할 + 누년 평균값 표 |
| **만능 양식 모드** | ○ | **양식 파일(빈 표: 엑셀·이미지·PDF)** 을 올리면 AI가 그 양식을 분석해 자료를 채워 줌 |

**내장 3개 모드는 AI를 쓰지 않습니다** (API 키 불필요). "만능 양식 모드"만 사용자가 입력한
**Claude API 키**로 동작하며, 그 키는 해당 요청을 처리하는 동안 메모리에서만 쓰이고
**로그·디스크·환경변수·배포 코드 어디에도 저장하지 않습니다**.

### 만능 양식 모드 사용법

1. 모드에서 "만능 양식 모드 (AI)" 선택
2. **양식 파일** 업로드 — 빈 표를 그린 `.xlsx` / `.xls` / `.csv`, 또는 표 **이미지**(png·jpg) / **PDF**
3. **자료 파일** 업로드 — `.xlsx` / `.xls` / `.csv` / `.hwpx` (구버전 `.hwp`는 한글에서 `.hwpx`나 `.xlsx`로 저장 후 사용)
4. 본인의 **Claude API 키**(`sk-ant-...`) 입력, 모델 선택 (Opus 5 / Sonnet 5 / Haiku 4.5)
5. (선택) 추가 지시사항 — 예: "발생일은 yyyy.mm.dd, 평균은 소수 2자리"
6. 변환하기 → AI가 `{ rows: [...] }` 형태의 표 모델을 만들고, 서버가 그대로 xlsx/docx로 렌더링

동작 순서: `lib/dataExtract.js`(자료→행렬) + `lib/templateAnalyze.js`(양식→AI 입력) →
`lib/aiTable.js`(Claude 호출) → `lib/renderGrid.js`(표 모델→파일).

- 프런트엔드: 정적 HTML/JS (`public/`)
- 백엔드: Node.js + Express (`server.js`, `lib/`)
- 엑셀 파싱/생성: `exceljs`
- Word 생성: `docx`
- 텍스트 정리: 순수 JS 규칙 (`lib/ruleClean.js`, 외부 호출 없음)

## 1. 로컬에서 먼저 실행해보기

```bash
npm install
npm start
```

브라우저에서 http://localhost:3000 접속 → 엑셀 업로드 → 양식(절토사면형/옹벽형) 선택 →
변환하기. API 키나 `.env` 설정이 필요 없습니다.

## 2. GitHub에 올리기

```bash
git init
git add .
git commit -m "init: 이력현황 표 변환 웹앱"
git branch -M main
git remote add origin <내 GitHub 저장소 URL>
git push -u origin main
```

`.gitignore`에 `node_modules/`가 이미 포함되어 있어 불필요한 파일은 올라가지 않습니다.

## 3. 배포 방법 — 원하는 환경에 맞게 고르기

상태를 서버에 저장하지 않는(stateless) 일반 Express 서버라 아래 방법 모두 같은 코드로
동작합니다. API 키 설정이 필요 없어서 배포가 한 단계 더 간단합니다.

### (A) GitHub 연동 클라우드 호스팅 — 가장 간단 (Render 기준)

1. https://render.com 가입 후 "New +" → "Web Service" → 방금 만든 GitHub 저장소 선택
2. Build Command: `npm install`, Start Command: `npm start`
3. 배포 완료 후 발급되는 URL로 바로 접속 가능. `main` 브랜치에 push할 때마다 자동 재배포됨

Railway(https://railway.app)도 같은 방식(GitHub 저장소 연결 → 자동 빌드/배포)으로 가능합니다.

`vercel.json`도 포함되어 있어 Vercel에서도 됩니다:

```bash
npm install -g vercel
vercel            # 처음엔 GitHub 저장소를 그대로 import해도 됨
vercel --prod
```

### (B) 회사 내부서버 / 인트라넷

```bash
git clone <내 GitHub 저장소 URL>
cd <폴더>
npm install --production
npm start            # 기본 3000번 포트
```

계속 띄워두려면 `pm2` 같은 프로세스 매니저를 추천합니다:

```bash
npm install -g pm2
pm2 start server.js --name inspection-history-table
pm2 save
```

사내망에서만 접근하게 하려면 방화벽/보안그룹에서 외부 인바운드를 막고, 앞단에 Nginx로
리버스 프록시 + 사내 SSO/기본인증을 붙이는 걸 권장합니다 (이 저장소에는 인증 기능이 없습니다).

Docker로 띄우고 싶다면 아래 `Dockerfile`을 추가해서 사용하세요:

```dockerfile
FROM node:20-slim
WORKDIR /app
COPY package.json ./
RUN npm install --production
COPY . .
ENV PORT=3000
EXPOSE 3000
CMD ["node", "server.js"]
```

### (C) 내 PC에서만 실행 (로컬)

위 "1. 로컬에서 먼저 실행해보기"와 동일합니다.

## 4. 환경변수

`PORT` (선택, 기본 3000) 외에는 필요한 환경변수가 없습니다.

## 5. 표 양식 추가/수정하기

`lib/renderDocx.js`, `lib/renderXlsx.js`에 열 구성/병합 방식이 정의되어 있습니다. 새로운
시설물 양식이 필요하면 `buildSlopeDocx`/`buildWallDocx` 중 더 비슷한 함수를 복사해서 열
구성만 바꾸면 됩니다. 프론트엔드(`public/index.html`의 `<select id="format">`)와
`server.js`의 `format` 분기도 같이 맞춰줘야 합니다.

텍스트 정리 규칙(줄바꿈 병합 로직)을 바꾸고 싶으면 `lib/ruleClean.js`의 `cleanCell` 함수를
수정하면 됩니다. 이 함수는:
- 어느 줄에도 불릿 문자(ㆍ？·○-)가 없으면 → 줄바꿈마다 별도 항목으로 취급
- 일부 줄에 불릿 문자가 있으면 → 불릿 문자가 있는 줄만 새 항목 시작, 없는 줄은 바로 앞
  항목에 공백을 넣어 이어붙임 (엑셀 셀 너비로 인한 줄바꿈으로 간주)

## 6. 보안 관련 참고

- 이 앱은 인증이 없습니다. 사내망 밖으로 열려면 반드시 로그인/접근 제어를 앞단에 추가하세요.
- 업로드된 엑셀과 변환 결과는 서버 디스크에 저장하지 않고 메모리에서만 처리한 뒤 응답으로
  돌려줍니다(디스크에 남지 않음).
