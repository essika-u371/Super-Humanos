const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ShadingType, Header, Footer, PageNumber, LevelFormat, TabStopType } = require("docx");

const BLACK = "0D0D0C", YELLOW = "FCC800", GOLD = "8A6B00", MUTED = "56564F", LINE = "DEDED9", BG = "F7F7F5";
const HF = "Barlow Condensed", BF = "Atkinson Hyperlegible Next";

const run = (t, o = {}) => new TextRun({ text: t, font: BF, size: 21, color: BLACK, ...o });
const P = (t, o = {}) => new Paragraph({ spacing: { after: 120, line: 300 }, ...o, children: Array.isArray(t) ? t : [run(t, o.run || {})] });
const H1 = (t) => new Paragraph({ spacing: { before: 420, after: 140 }, keepNext: true, children: [new TextRun({ text: t.toUpperCase(), font: HF, size: 40, bold: true, color: BLACK })] });
const H2 = (t) => new Paragraph({ spacing: { before: 280, after: 100 }, keepNext: true, children: [new TextRun({ text: t.toUpperCase(), font: HF, size: 28, bold: true, color: BLACK })] });
const K = (t) => new Paragraph({ spacing: { before: 60, after: 60 }, keepNext: true, children: [new TextRun({ text: "› ", font: HF, size: 20, bold: true, color: YELLOW }), new TextRun({ text: t.toUpperCase(), font: HF, size: 20, bold: true, color: GOLD, characterSpacing: 40 })] });
const B = (t) => new Paragraph({ numbering: { reference: "chev", level: 0 }, spacing: { after: 80, line: 290 }, children: Array.isArray(t) ? t : [run(t)] });
const bold = (t) => run(t, { bold: true });

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thin = { style: BorderStyle.SINGLE, size: 6, color: LINE };
const CONTENT_DXA = 9972; // carta 12240 menos márgenes de 1134 a cada lado
function table(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const dxa = widths.map((w) => Math.round(w / total * CONTENT_DXA));
  dxa[dxa.length - 1] = CONTENT_DXA - dxa.slice(0, -1).reduce((a, b) => a + b, 0);
  const cell = (t, i, isHead) => new TableCell({
    width: { size: dxa[i], type: WidthType.DXA },
    shading: isHead ? { type: ShadingType.CLEAR, fill: BLACK, color: BLACK } : undefined,
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    borders: { top: thin, bottom: thin, left: noBorder, right: noBorder },
    children: (Array.isArray(t) ? t : [t]).map((line) => new Paragraph({ spacing: { after: 40, line: 260 }, children: [isHead ? new TextRun({ text: String(line).toUpperCase(), font: HF, size: 19, bold: true, color: YELLOW }) : (line && line.children ? line : run(String(line), { size: 19 }))] })),
  });
  return new Table({
    width: { size: CONTENT_DXA, type: WidthType.DXA },
    columnWidths: dxa,
    layout: "fixed",
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }), ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false)) }))],
  });
}
const box = (title, lines) => new Table({
  width: { size: CONTENT_DXA, type: WidthType.DXA },
  columnWidths: [CONTENT_DXA],
  layout: "fixed",
  rows: [new TableRow({ children: [new TableCell({
    width: { size: CONTENT_DXA, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: BG, color: BG }, margins: { top: 140, bottom: 140, left: 200, right: 200 },
    borders: { top: noBorder, bottom: noBorder, left: { style: BorderStyle.SINGLE, size: 36, color: YELLOW }, right: noBorder },
    children: [new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: title.toUpperCase(), font: HF, size: 24, bold: true, color: BLACK })] }), ...lines.map((l) => new Paragraph({ spacing: { after: 60, line: 280 }, children: [run(l, { size: 20 })] }))],
  })] })],
});
const spacer = () => new Paragraph({ spacing: { after: 120 }, children: [] });

module.exports = { fs, path, Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ShadingType, Header, Footer, PageNumber, LevelFormat, TabStopType, BLACK, YELLOW, GOLD, MUTED, LINE, BG, HF, BF, run, P, H1, H2, K, B, bold, table, box, spacer };
