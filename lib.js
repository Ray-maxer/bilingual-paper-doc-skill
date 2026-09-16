/**
 * lib.js — reusable engine for "bilingual paper reading doc" generation.
 *
 * This module is content-agnostic: it only knows how to turn a JS data object
 * (see references/data-schema.md) into a fully-formatted, cover-to-references
 * .docx buffer with:
 *   - alternating English/Chinese paragraph pairs
 *   - embedded figures with bilingual captions
 *   - numbered display equations built as REAL Word "Equation" objects (OOXML
 *     m:oMath — editable via Word's own Equation Editor, not plain text)
 *   - inline scientific symbols (Mmax, CD, θmax, ...) auto-detected in prose
 *     and rendered with real Word super/subscript formatting
 *   - yellow "reading note" callout boxes
 *
 * Usage (see scripts/render.js for the CLI wrapper):
 *   const { renderDocument } = require("./lib.js");
 *   const data = require("./content-data.js");   // paper-specific content
 *   renderDocument(data).then(buf => fs.writeFileSync("output.docx", buf));
 *
 * IMPORTANT — after writing the .docx, you MUST run inject_mathpr.py on it
 * (see scripts/inject_mathpr.py) before it is a fully valid/well-behaved
 * Word file. docx-js does not emit the <m:mathPr> settings block that Word
 * expects once a document contains equation objects; omitting this step
 * still produces a schema-valid file but skips Word's default equation
 * font/spacing settings.
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, convertInchesToTwip, Header, Footer, PageNumber, VerticalAlign,
  Math: MathBlock, MathRun, MathSubScript, MathSuperScript, MathSubSuperScript,
  MathFraction, MathIntegral, TableOfContents,
} = require("docx");

// ======================================================================
// COLORS / STYLE CONSTANTS — tweak per-paper via data.style if desired
// ======================================================================
const DEFAULT_STYLE = {
  enColor: "1F1F1F",
  zhColor: "0B5394",
  headColor: "1B4F72",
  noteShade: "FFF2CC",
  noteBorder: "D6A700",
  noteTextColor: "7F6000",
  enFont: "Calibri",
  zhFont: "Microsoft JhengHei",
};

// ======================================================================
// EQUATION DSL -> real OOXML Equation objects (m:oMath)
// Same objects Word's own "Insert > Equation" produces — double-clickable
// and editable in Word's Equation Editor.
// ======================================================================
function toChildren(arr) {
  const out = [];
  for (const tok of arr) {
    if (typeof tok === "string") {
      if (tok.length) out.push(new MathRun(tok));
    } else {
      out.push(buildToken(tok));
    }
  }
  return out;
}
function baseChildren(base) {
  return Array.isArray(base) ? toChildren(base) : toChildren([base]);
}
function buildToken(tok) {
  switch (tok.t) {
    case "sub":
      return new MathSubScript({ children: baseChildren(tok.base), subScript: [new MathRun(tok.sub)] });
    case "sup":
      return new MathSuperScript({ children: baseChildren(tok.base), superScript: [new MathRun(tok.sup)] });
    case "subsup":
      return new MathSubSuperScript({
        children: baseChildren(tok.base),
        subScript: [new MathRun(tok.sub)],
        superScript: [new MathRun(tok.sup)],
      });
    case "frac":
      return new MathFraction({ numerator: toChildren(tok.num), denominator: toChildren(tok.den) });
    case "int":
      return new MathIntegral({
        subScript: tok.sub ? [new MathRun(tok.sub)] : undefined,
        superScript: tok.sup ? [new MathRun(tok.sup)] : undefined,
        children: toChildren(tok.base),
      });
    default:
      return new MathRun(String(tok));
  }
}
// Shorthand token builders — use these when writing equation tokens in content-data.js
const SUB = (base, sub) => ({ t: "sub", base, sub });
const SUP = (base, sup) => ({ t: "sup", base, sup });
const SUBSUP = (base, sub, sup) => ({ t: "subsup", base, sub, sup });
const FRAC = (num, den) => ({ t: "frac", num, den });
const INT = (base, sub, sup) => ({ t: "int", base, sub, sup });

// Renders a numbered display equation as a 2-column borderless table:
// left cell = centered native Word equation object, right cell = right-aligned "(n)" label.
function equationBlock(tokens, label) {
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [9000, 900],
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9000, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new MathBlock({ children: toChildren(tokens) })] })],
          }),
          new TableCell({
            width: { size: 900, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            verticalAlign: VerticalAlign.CENTER,
            borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: `(${label})`, size: 22, color: "444444", font: "Calibri" })] })],
          }),
        ],
      }),
    ],
  });
}

// ======================================================================
// INLINE SUBSCRIPT/SUPERSCRIPT DETECTION FOR PROSE
// Splits running text at known scientific-variable tokens (Mmax, CD, θmax...)
// and renders each as base + real superscript/subscript runs.
// Extend DEFAULT_INLINE_RULES per-paper via data.extraInlineRules (see schema doc).
// List longer/composite patterns first; order rarely matters since these
// patterns are not literal substrings of each other, but keep it tidy.
// ======================================================================
const DEFAULT_INLINE_RULES = [
  { re: /F̄theorymax/g, base: "F̄", sup: "theory", sub: "max" },
  { re: /Mtheorymax/g, base: "M", sup: "theory", sub: "max" },
  { re: /MLmax/g, base: "M", sup: "L", sub: "max" },
  { re: /MQmax/g, base: "M", sup: "Q", sub: "max" },
  { re: /fLmax/g, base: "f", sup: "L", sub: "max" },
  { re: /fQmax/g, base: "f", sup: "Q", sub: "max" },
  { re: /θLmax/g, base: "θ", sup: "L", sub: "max" },
  { re: /θQmax/g, base: "θ", sup: "Q", sub: "max" },
  { re: /σLmax/g, base: "σ", sup: "L", sub: "max" },
  { re: /σQmax/g, base: "σ", sup: "Q", sub: "max" },
  { re: /Mmax/g, base: "M", sub: "max" },
  { re: /fmax/g, base: "f", sub: "max" },
  { re: /θmax/g, base: "θ", sub: "max" },
  { re: /σmax/g, base: "σ", sub: "max" },
  { re: /Hpull/g, base: "H", sub: "pull" },
  { re: /F̄D/g, base: "F̄", sub: "D" },
  { re: /F̄M/g, base: "F̄", sub: "M" },
  { re: /FD/g, base: "F", sub: "D" },
  { re: /FM/g, base: "F", sub: "M" },
  { re: /CD/g, base: "C", sub: "D" },
  { re: /CM/g, base: "C", sub: "M" },
  { re: /DBH²/g, base: "DBH", sup: "2" },
  { re: /H²/g, base: "H", sup: "2" },
  { re: /a²/g, base: "a", sup: "2" },
  { re: /R²/g, base: "R", sup: "2" },
];

function tokenizeInline(text, rules) {
  let segments = [text];
  for (const rule of rules) {
    const next = [];
    for (const seg of segments) {
      if (typeof seg !== "string") { next.push(seg); continue; }
      rule.re.lastIndex = 0;
      let lastIndex = 0, m;
      while ((m = rule.re.exec(seg)) !== null) {
        if (m.index > lastIndex) next.push(seg.slice(lastIndex, m.index));
        next.push({ base: rule.base, sup: rule.sup, sub: rule.sub });
        lastIndex = m.index + m[0].length;
      }
      if (lastIndex < seg.length) next.push(seg.slice(lastIndex));
    }
    segments = next;
  }
  return segments;
}

function richRuns(text, opts, rules) {
  const { color, font, size } = opts;
  const subSize = Math.round(size * 0.72);
  const segs = tokenizeInline(text, rules);
  const runs = [];
  for (const seg of segs) {
    if (typeof seg === "string") {
      if (seg.length) runs.push(new TextRun({ text: seg, color, font, size }));
    } else {
      runs.push(new TextRun({ text: seg.base, color, font, size, italics: true }));
      if (seg.sup) runs.push(new TextRun({ text: seg.sup, color, font, size: subSize, italics: true, superScript: true }));
      if (seg.sub) runs.push(new TextRun({ text: seg.sub, color, font, size: subSize, italics: true, subScript: true }));
    }
  }
  return runs;
}

// ======================================================================
// DOCUMENT ASSEMBLY
// ======================================================================
function renderDocument(data) {
  const style = { ...DEFAULT_STYLE, ...(data.style || {}) };
  const rules = [...(data.extraInlineRules || []), ...DEFAULT_INLINE_RULES];
  const imgDir = data.imageDir || ".";

  const enPara = (text) => new Paragraph({
    spacing: { after: 60, line: 300 },
    children: richRuns(text, { color: style.enColor, font: style.enFont, size: 22 }, rules),
  });
  const zhPara = (text) => new Paragraph({
    spacing: { after: 220, line: 320 },
    indent: { left: convertInchesToTwip(0.15) },
    border: { left: { style: BorderStyle.SINGLE, size: 6, color: "9FC5E8", space: 6 } },
    children: richRuns(text, { color: style.zhColor, font: style.zhFont, size: 22 }, rules),
  });
  const h1 = (en, zh) => new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: style.headColor, space: 4 } },
    children: [
      new TextRun({ text: en, bold: true, color: style.headColor, size: 30, font: style.enFont }),
      new TextRun({ text: "  ", size: 30 }),
      new TextRun({ text: zh, bold: true, color: style.headColor, size: 30, font: style.zhFont }),
    ],
  });
  const h2 = (en, zh) => new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 140 },
    children: [
      new TextRun({ text: en, bold: true, color: style.headColor, size: 26, font: style.enFont }),
      new TextRun({ text: "  ", size: 26 }),
      new TextRun({ text: zh, bold: true, color: style.headColor, size: 26, font: style.zhFont }),
    ],
  });
  const noteBox = (title, text) => new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [new TableRow({
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: style.noteShade, color: "auto" },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 8, color: style.noteBorder },
          bottom: { style: BorderStyle.SINGLE, size: 8, color: style.noteBorder },
          left: { style: BorderStyle.SINGLE, size: 8, color: style.noteBorder },
          right: { style: BorderStyle.SINGLE, size: 8, color: style.noteBorder },
        },
        children: [
          new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `📌 ${title}`, bold: true, color: style.noteTextColor, font: style.zhFont, size: 22 })] }),
          new Paragraph({ children: richRuns(text, { color: style.noteTextColor, font: style.zhFont, size: 22 }, rules) }),
        ],
      })],
    })],
  });
  const spacer = (h = 120) => new Paragraph({ spacing: { after: h }, children: [] });
  const figureBlock = (file, w, h, capEn, capZh, num) => [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 80 },
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(imgDir, file)), transformation: { width: w, height: h } })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: `Fig. ${num}. ${capEn}`, italics: true, size: 20, font: style.enFont, color: "444444" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `圖 ${num}．${capZh}`, italics: true, size: 20, font: style.zhFont, color: "444444" })] }),
  ];
  const tableBlock = (headers, headersZh, rows, capEn, capZh, num) => {
    const nCols = headers.length;
    const headShade = style.tableHeadShade || "D9E2F3";
    const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
    const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder, insideHorizontal: cellBorder, insideVertical: cellBorder };
    const colWidth = Math.floor(10080 / nCols);

    const headerRow = new TableRow({
      tableHeader: true,
      children: headers.map((h) => new TableCell({
        width: { size: colWidth, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: headShade, color: "auto" },
        margins: { top: 80, bottom: headersZh ? 20 : 80, left: 100, right: 100 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, color: style.headColor, font: style.enFont, size: 20 })] })],
      })),
    });
    const headerZhRow = headersZh ? new TableRow({
      tableHeader: true,
      children: headersZh.map((h) => new TableCell({
        width: { size: colWidth, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: headShade, color: "auto" },
        margins: { top: 0, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, color: style.headColor, font: style.zhFont, size: 18 })] })],
      })),
    }) : null;
    const dataRows = rows.map((r) => new TableRow({
      children: r.map((cellText) => new TableCell({
        width: { size: colWidth, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: richRuns(String(cellText), { color: style.enColor, font: style.enFont, size: 20 }, rules) })],
      })),
    }));

    const out = [];
    // Academic convention: table captions sit ABOVE the table (figure captions sit below).
    if (capEn || capZh || num != null) {
      out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 40 }, children: [new TextRun({ text: `Table ${num}.${capEn ? " " + capEn : ""}`, bold: true, italics: true, size: 20, font: style.enFont, color: "444444" })] }));
      if (capZh) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: `表 ${num}．${capZh}`, italics: true, size: 20, font: style.zhFont, color: "444444" })] }));
    }
    out.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: headers.map(() => colWidth),
      borders,
      rows: [headerRow, ...(headerZhRow ? [headerZhRow] : []), ...dataRows],
    }));
    out.push(spacer(160));
    return out;
  };
  const bulletsBlock = (items) => items.map((t) => new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 100 },
    children: richRuns(t, { color: style.enColor, font: style.zhFont, size: 22 }, rules),
  }));

  const children = [];

  // --- cover ---
  const m = data.meta || {};
  if (m.citation) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: m.citation, size: 20, color: "666666" })] }));
  if (m.titleEn) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80 }, children: [new TextRun({ text: m.titleEn, bold: true, size: 32, color: style.headColor, font: style.enFont })] }));
  if (m.titleZh) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: m.titleZh, bold: true, size: 30, color: style.headColor, font: style.zhFont })] }));
  if (m.authors) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: m.authors, size: 18, color: "555555", font: style.enFont })] }));
  if (m.doiLine) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: m.doiLine, size: 18, color: "888888", font: style.enFont })] }));
  if (data.usageNote) {
    children.push(noteBox(data.usageNote.title, data.usageNote.text));
    children.push(spacer(200));
  }

  // --- table of contents (headings already use real Word heading styles, ---
  // --- so this is a native, clickable field, not a manually-typed list) ---
  if (data.toc !== false) {
    children.push(new Paragraph({ children: [], pageBreakBefore: true }));
    children.push(new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: "Table of Contents ｜ 目錄", bold: true, size: 28, color: style.headColor, font: style.enFont })],
    }));
    children.push(new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }));
    children.push(new Paragraph({ children: [], pageBreakBefore: true }));
  }

  // --- body: ordered list of blocks ---
  for (const block of data.sections || []) {
    switch (block.type) {
      case "h1": children.push(h1(block.en, block.zh)); break;
      case "h2": children.push(h2(block.en, block.zh)); break;
      case "pair": children.push(enPara(block.en), zhPara(block.zh)); break;
      case "en": children.push(enPara(block.en)); break;
      case "zh": children.push(zhPara(block.zh)); break;
      case "equation": children.push(equationBlock(block.tokens, block.label)); break;
      case "figure": children.push(...figureBlock(block.file, block.w, block.h, block.capEn, block.capZh, block.num)); break;
      case "table": children.push(...tableBlock(block.headers, block.headersZh, block.rows, block.capEn, block.capZh, block.num)); break;
      case "note": children.push(noteBox(block.title, block.text)); break;
      case "bullets": children.push(...bulletsBlock(block.items)); break;
      case "spacer": children.push(spacer(block.h)); break;
      case "pageBreak": children.push(new Paragraph({ children: [], pageBreakBefore: true })); break;
      default: throw new Error("Unknown block type: " + block.type);
    }
  }

  // --- references ---
  if (data.references && data.references.length) {
    children.push(h1(data.referencesHeading?.en || "References", data.referencesHeading?.zh || "參考文獻（原文，未翻譯）"));
    for (const r of data.references) {
      children.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: r, font: style.enFont, size: 18, color: "555555" })] }));
    }
  }

  const doc = new Document({
    title: m.titleEn || m.titleZh || undefined,
    subject: m.citation || undefined,
    creator: m.authors || undefined,
    description: m.doiLine
      ? `Bilingual (EN/繁中) reading translation. ${m.doiLine}`
      : "Bilingual (EN/繁中) reading translation.",
    features: { updateFields: true },
    sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } },
      },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: m.headerRight || "", size: 16, color: "999999" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "999999" })] })] }) },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}

module.exports = { renderDocument, SUB, SUP, SUBSUP, FRAC, INT, DEFAULT_INLINE_RULES };
