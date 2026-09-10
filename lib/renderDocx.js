const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, VerticalAlign, BorderStyle, ShadingType,
  VerticalMergeType
} = require("docx");

const FONT = "맑은 고딕";
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const CELL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const HEADER_FILL = "D9D9D9";
const PAGE = {
  size: { width: 11906, height: 16838 }, // A4
  margin: { top: 850, bottom: 850, left: 850, right: 850 }
};

function txt(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: 20, bold: !!opts.bold });
}

async function buildDocument(children) {
  const doc = new Document({
    sections: [{ properties: { page: PAGE }, children }]
  });
  return Packer.toBuffer(doc);
}

function titleParagraph(title) {
  return new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({ text: title || "점검진단실적", bold: true, font: FONT, size: 28 })]
  });
}

// ---------------- 절토사면형 ----------------

function buildSlopeDocx(rows, title) {
  const W = { no: 600, type: 1050, period: 1500, org: 1550, grade: 900, result: 3760 };

  function headerCell(text, width) {
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: HEADER_FILL },
      verticalAlign: VerticalAlign.CENTER,
      borders: CELL_BORDERS,
      margins: { top: 60, bottom: 60, left: 60, right: 60 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [txt(text, { bold: true })] })]
    });
  }
  function centeredCell(lines, width, opts = {}) {
    const arr = Array.isArray(lines) ? lines : [lines];
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: CELL_BORDERS,
      margins: { top: 60, bottom: 60, left: 60, right: 60 },
      children: arr.map(line => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [txt(line, { bold: !!opts.bold })]
      }))
    });
  }
  function resultCell(bullets, width) {
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: CELL_BORDERS,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: bullets.map((b, i) => new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: i === bullets.length - 1 ? 0 : 40 },
        children: [txt("-" + b)]
      }))
    });
  }

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell("No.", W.no), headerCell("점검종류", W.type), headerCell("점검기간", W.period),
      headerCell("점검기관", W.org), headerCell("평가등급", W.grade), headerCell("점검 결과", W.result),
    ]
  });

  const bodyRows = rows.map(r => new TableRow({
    children: [
      centeredCell(r.no, W.no),
      centeredCell([r.type], W.type),
      centeredCell([r.start, "~", r.end], W.period),
      centeredCell([r.org], W.org),
      centeredCell([r.grade], W.grade, { bold: true }),
      resultCell(r.bullets, W.result),
    ]
  }));

  const table = new Table({
    width: { size: Object.values(W).reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: Object.values(W),
    rows: [headerRow, ...bodyRows]
  });

  return buildDocument([titleParagraph(title), table]);
}

// ---------------- 옹벽형 ----------------

function buildWallDocx(rows, title) {
  const W = { no: 550, period: 1650, org: 1350, grade: 700, result: 4110, remedy: 1000 };

  function headerCell(text, width, { merge } = {}) {
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: HEADER_FILL },
      verticalAlign: VerticalAlign.CENTER,
      borders: CELL_BORDERS,
      margins: { top: 60, bottom: 60, left: 60, right: 60 },
      verticalMerge: merge,
      children: merge === VerticalMergeType.CONTINUE ? [] : [new Paragraph({
        alignment: AlignmentType.CENTER, children: [txt(text, { bold: true })]
      })]
    });
  }
  function bodyCell(lines, width, { merge, align = AlignmentType.CENTER, bold = false, leftPad = false } = {}) {
    const arr = Array.isArray(lines) ? lines : [lines];
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      borders: CELL_BORDERS,
      margins: { top: 60, bottom: 60, left: leftPad ? 100 : 60, right: 60 },
      verticalMerge: merge,
      children: merge === VerticalMergeType.CONTINUE ? [] : arr.map((line, i) => new Paragraph({
        alignment: align,
        spacing: { after: i === arr.length - 1 ? 0 : 40 },
        children: [txt(align === AlignmentType.LEFT ? "-" + line : line, { bold })]
      }))
    });
  }

  const headerRow1 = new TableRow({
    tableHeader: true,
    children: [
      headerCell("No.", W.no, { merge: VerticalMergeType.RESTART }),
      headerCell("수행기간", W.period),
      headerCell("점검·진단기관", W.org),
      headerCell("안전등급", W.grade, { merge: VerticalMergeType.RESTART }),
      headerCell("주요 점검·진단 결과", W.result, { merge: VerticalMergeType.RESTART }),
      headerCell("보수·보강(안)", W.remedy, { merge: VerticalMergeType.RESTART }),
    ]
  });
  const headerRow2 = new TableRow({
    tableHeader: true,
    children: [
      headerCell("", W.no, { merge: VerticalMergeType.CONTINUE }),
      headerCell("점검·진단구분", W.period),
      headerCell("책임기술자", W.org),
      headerCell("", W.grade, { merge: VerticalMergeType.CONTINUE }),
      headerCell("", W.result, { merge: VerticalMergeType.CONTINUE }),
      headerCell("", W.remedy, { merge: VerticalMergeType.CONTINUE }),
    ]
  });

  const rowsOut = [headerRow1, headerRow2];
  for (const r of rows) {
    rowsOut.push(new TableRow({
      children: [
        bodyCell(r.no, W.no, { merge: VerticalMergeType.RESTART }),
        bodyCell(`${r.start.replace(/-/g, ".")}. ~\n${r.end.replace(/-/g, ".")}.`, W.period, {}),
        bodyCell(r.org, W.org, {}),
        bodyCell(r.grade, W.grade, { merge: VerticalMergeType.RESTART, bold: true }),
        bodyCell(r.result, W.result, { merge: VerticalMergeType.RESTART, align: AlignmentType.LEFT, leftPad: true }),
        bodyCell(r.remedy, W.remedy, { merge: VerticalMergeType.RESTART, align: AlignmentType.LEFT, leftPad: true }),
      ]
    }));
    rowsOut.push(new TableRow({
      children: [
        bodyCell("", W.no, { merge: VerticalMergeType.CONTINUE }),
        bodyCell(r.type, W.period, {}),
        bodyCell(r.engineer, W.org, {}),
        bodyCell("", W.grade, { merge: VerticalMergeType.CONTINUE }),
        bodyCell("", W.result, { merge: VerticalMergeType.CONTINUE }),
        bodyCell("", W.remedy, { merge: VerticalMergeType.CONTINUE }),
      ]
    }));
  }

  const table = new Table({
    width: { size: Object.values(W).reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: Object.values(W),
    rows: rowsOut
  });

  return buildDocument([titleParagraph(title), table]);
}

module.exports = { buildSlopeDocx, buildWallDocx };
