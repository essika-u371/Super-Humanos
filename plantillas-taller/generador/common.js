const path = require("path");
const B = require("./brand.js");
const { applyTheme } = require("./apply_theme.js");

async function build(fileName, footer, fn) {
  const pres = new B.pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: B.THEME.headFontFace, bodyFontFace: B.THEME.bodyFontFace };
  pres.author = "Pivot Power";
  pres.company = "Pivot Power";
  pres.title = fileName.replace(".pptx", "");
  pres.lang = "es-CO";
  B.defineLayouts(pres, footer);
  await fn(pres);
  const out = path.join(__dirname, "out", fileName);
  await pres.writeFile({ fileName: out });
  await applyTheme(out, B.THEME);
  console.log("ok", out);
}

/* Portada común */
function cover(pres, section, { encuentro, kicker, duracion, notas }) {
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: section });
  B.wordmark(s, B.M, 0.55, { size: 26 });
  B.kicker(s, kicker, B.M, 2.0, true, 10);
  s.addText("[NOMBRE DEL TALLER]", { x: B.M, y: 2.35, w: 11.5, h: 1.5, fontFace: B.THEME.headFontFace, fontSize: 66, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
  s.addText("[CONCEPTO DEL MES]", { x: B.M, y: 3.85, w: 11.5, h: 0.7, fontFace: B.THEME.headFontFace, fontSize: 32, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
  // Cinta del encuentro
  B.ribbon(s, encuentro, B.M + 0.3, 4.85, 5.2, { yellow: true, fontSize: 20, h: 0.55 });
  s.addText(duracion, { x: 6.4, y: 4.85, w: 3, h: 0.55, fontFace: B.THEME.bodyFontFace, fontSize: 16, color: "C4C4BE", margin: 0, valign: "middle", isTextBox: true });
  // Datos
  s.addText([
    { text: "[FECHA]", options: { bold: true, color: "FFFFFF" } },
    { text: "  ·  [HORA]", options: { color: "C4C4BE" } },
    { text: "  ·  Teams", options: { color: "C4C4BE" } },
  ], { x: B.M, y: 6.1, w: 7, h: 0.4, fontFace: B.THEME.bodyFontFace, fontSize: 16, margin: 0, valign: "middle", isTextBox: true });
  s.addText([
    { text: "Docente  ", options: { color: "C4C4BE" } },
    { text: "[DOCENTE]", options: { bold: true, color: "FFFFFF" } },
  ], { x: B.M, y: 6.5, w: 7, h: 0.4, fontFace: B.THEME.bodyFontFace, fontSize: 16, margin: 0, valign: "middle", isTextBox: true });
  // Recorrido de los seis talleres a la derecha
  const items = [1, 2, 3, 4, 5, 6].map((n) => ({ label: "T" + n, done: n === 1 }));
  B.chevronBar(s, items, 8.3, 6.2, 4.45, { dark: true, h: 0.42, fontSize: 12 });
  s.addText("Recorrido · marca en amarillo el taller actual", { x: 8.3, y: 6.68, w: 4.45, h: 0.3, fontFace: B.THEME.bodyFontFace, fontSize: 10, color: "9A9A93", margin: 0, isTextBox: true });
  s.addNotes(notas);
  return s;
}

module.exports = { build, cover, B };
