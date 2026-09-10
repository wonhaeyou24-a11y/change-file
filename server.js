const path = require("path");
const express = require("express");
const cors = require("cors");
const multer = require("multer");

const { extractRows } = require("./lib/excelExtract");
const { cleanRows } = require("./lib/ruleClean");
const { buildSlopeDocx, buildWallDocx } = require("./lib/renderDocx");
const { buildSlopeXlsx, buildWallXlsx } = require("./lib/renderXlsx");

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
    const format = req.body.format === "wall" ? "wall" : "slope";
    const title = req.body.title || "점검진단실적";

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`이력현황 변환 웹앱 실행 중: http://localhost:${PORT}`);
});
