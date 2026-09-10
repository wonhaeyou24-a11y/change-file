const path = require("path");
const express = require("express");
const cors = require("cors");
const multer = require("multer");

const { extractRows } = require("./lib/excelExtract");
const { cleanRows } = require("./lib/ruleClean");
const { buildSlopeDocx, buildWallDocx } = require("./lib/renderDocx");
const { buildSlopeXlsx, buildWallXlsx } = require("./lib/renderXlsx");
const { extractWeather } = require("./lib/weatherExtract");
const { buildWeatherDocx, buildWeatherXlsx, buildLayout } = require("./lib/renderWeather");
const { extractData, matrixToTSV } = require("./lib/dataExtract");
const { analyzeTemplate } = require("./lib/templateAnalyze");
const { aiBuildTable } = require("./lib/aiTable");
const { gridToXlsx, gridToDocx } = require("./lib/renderGrid");

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

app.use(cors());
app.use(express.static(path.join(__dirname, "public")));

// 서버리스(Vercel 등) 환경에서는 요청마다 다른 인스턴스가 처리될 수 있어 서버 메모리에
// 파일을 캐시해뒀다가 별도 다운로드 요청으로 내려주는 방식은 신뢰할 수 없다.
// 그래서 변환 결과 파일은 base64로 인코딩해 /api/convert 응답에 바로 실어 보내고,
// 프론트엔드에서 Blob으로 바꿔 다운로드한다 (로컬/인트라넷/클라우드 어디서나 동일하게 동작).

app.post("/api/convert", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "엑셀 파일이 없습니다." });
    const format = ["wall", "weather"].includes(req.body.format) ? req.body.format : "slope";
    const title = req.body.title || "점검진단실적";

    // ── 기상청 강수량 분석 양식 ──────────────────────────────
    if (format === "weather") {
      const data = await extractWeather(req.file.buffer, req.file.originalname);
      if (!data.years.length) {
        return res.status(400).json({ error: "연도별 강수량 행을 찾을 수 없습니다. 기상청 '일 최다 강수량' 자료인지 확인해주세요." });
      }
      const opts = {
        name: req.body.stationName || data.name || "",
        tableNo: req.body.tableNo || "6.2-5",
        caption: req.body.caption || "기상청 강수량 분석 결과",
        sectionHeading: req.body.sectionHeading || "",
      };
      if (opts.name && !data.name) data.name = opts.name;
      const layout = buildLayout(data, opts);
      const [docxBuf, xlsxBuf] = await Promise.all([
        buildWeatherDocx(data, opts),
        buildWeatherXlsx(data, opts),
      ]);
      const wbase = path.parse(req.file.originalname).name || "강수량분석";
      return res.json({
        ok: true,
        format,
        rowCount: layout.n,
        summary: { name: layout.name, period: layout.periodText, avg: layout.avg },
        grid: layout.grid,
        files: {
          docx: { name: `기상청강수량분석_${wbase}.docx`, data: docxBuf.toString("base64"),
            mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
          xlsx: { name: `기상청강수량분석_${wbase}.xlsx`, data: Buffer.from(xlsxBuf).toString("base64"),
            mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        },
      });
    }

    const rawRows = await extractRows(req.file.buffer);
    if (rawRows.length === 0) {
      return res.status(400).json({ error: "점검 이력 행을 찾을 수 없습니다. 원본 엑셀 형식을 확인해주세요." });
    }

    const cleanedRows = cleanRows(rawRows, format);

    const [docxBuf, xlsxBuf] = await Promise.all([
      format === "wall" ? buildWallDocx(cleanedRows, title) : buildSlopeDocx(cleanedRows, title),
      format === "wall" ? buildWallXlsx(cleanedRows, title) : buildSlopeXlsx(cleanedRows, title),
    ]);

    const base = path.parse(req.file.originalname).name || "이력현황";

    res.json({
      ok: true,
      rowCount: cleanedRows.length,
      preview: cleanedRows.slice(0, 5),
      files: {
        docx: { name: `이력현황_${base}.docx`, data: docxBuf.toString("base64"),
          mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
        xlsx: { name: `이력현황_${base}.xlsx`, data: Buffer.from(xlsxBuf).toString("base64"),
          mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      },
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "변환 중 오류가 발생했습니다." });
  }
});

// ── 만능 양식 모드 (AI 사용) ───────────────────────────────
// API 키는 이 요청 처리 동안 메모리에서만 쓰이고 로그/디스크/환경변수에 남기지 않는다.
const uploadUniversal = upload.fields([
  { name: "file", maxCount: 1 },
  { name: "template", maxCount: 1 },
]);

app.post("/api/convert-universal", uploadUniversal, async (req, res) => {
  try {
    const dataFile = req.files && req.files.file && req.files.file[0];
    const tmplFile = req.files && req.files.template && req.files.template[0];
    if (!dataFile) return res.status(400).json({ error: "자료 파일이 없습니다." });
    if (!tmplFile) return res.status(400).json({ error: "양식 파일이 없습니다." });

    const apiKey = req.body.apiKey || "";
    const model = req.body.model || "claude-opus-5";
    const instructions = req.body.instructions || "";
    const sectionHeading = req.body.sectionHeading || "";

    const { matrix, kind } = extractData(dataFile.buffer, dataFile.originalname);
    if (!matrix.length) return res.status(400).json({ error: "자료 파일에서 표 내용을 읽지 못했습니다." });
    const dataTSV = matrixToTSV(matrix);

    const template = await analyzeTemplate(tmplFile.buffer, tmplFile.originalname);

    const { grid, model: usedModel } = await aiBuildTable({ apiKey, model, template, dataTSV, instructions });

    const [xlsxBuf, docxBuf] = await Promise.all([
      gridToXlsx(grid),
      gridToDocx(grid, { sectionHeading }),
    ]);
    const base = path.parse(dataFile.originalname).name || "변환결과";

    res.json({
      ok: true,
      format: "universal",
      model: usedModel,
      dataKind: kind,
      grid,
      notes: grid.notes || "",
      computed: grid.computed || null,
      files: {
        xlsx: { name: `양식변환_${base}.xlsx`, data: Buffer.from(xlsxBuf).toString("base64"),
          mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        docx: { name: `양식변환_${base}.docx`, data: docxBuf.toString("base64"),
          mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      },
    });
  } catch (err) {
    // 키가 로그에 남지 않도록 message 만 출력
    console.error("[universal]", err.message);
    res.status(err.userFacing ? 400 : 500).json({ error: err.message || "변환 중 오류가 발생했습니다." });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`이력현황 변환 웹앱 실행 중: http://localhost:${PORT}`);
});
