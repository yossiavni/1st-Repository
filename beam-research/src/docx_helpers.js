// Shared helpers for building the Beam research Word document.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, ImageRun, PageBreak, TableOfContents, Footer, Header, PageNumber,
  LevelFormat, ExternalHyperlink, BorderStyle,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const FIG = (f) => path.join(ROOT, "figures", f);
const BLUE = "2A78D6", INK2 = "52514E";
const W = 9360; // content width in DXA (US Letter, 1" margins)

const p = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, ...opts,
  children: Array.isArray(text) ? text : [new TextRun(text)] });
const runs = (...parts) => parts.map((x) => typeof x === "string" ? new TextRun(x) : new TextRun(x));
const b = (t) => ({ text: t, bold: true });
let numInst = 0;
const h1 = (t) => (numInst++, 0) || new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(t)] });
const h2 = (t) => (numInst++, 0) || new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const h3 = (t) => (numInst++, 0) || new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const bullet = (t, level = 0) => new Paragraph({ numbering: { reference: "bul", level }, spacing: { after: 60 },
  children: Array.isArray(t) ? t : [new TextRun(t)] });
const num = (t) => new Paragraph({ numbering: { reference: "num", level: 0, instance: numInst }, spacing: { after: 60 },
  children: Array.isArray(t) ? t : [new TextRun(t)] });
const note = (t) => new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: t, italics: true, size: 18, color: INK2 })] });
const link = (text, url) => new ExternalHyperlink({ link: url, children: [new TextRun({ text, style: "Hyperlink" })] });

function img(file, wIn, caption) {
  const buf = fs.readFileSync(FIG(file));
  const wpx = buf.readUInt32BE(16), hpx = buf.readUInt32BE(20);
  const w = wIn * 96, h = Math.round(w * hpx / wpx);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: w, height: h } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: caption, italics: true, size: 18, color: INK2 })] }),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "D0CFC8" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, i, hdr) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: hdr ? { fill: "DBE9F9", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t), bold: hdr, size: 18 })] })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, i, false)) }))],
  });
}
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });


module.exports = { p, runs, b, h1, h2, h3, bullet, num, note, link, img, table, gap, BLUE, INK2, W };
