/* Campus Pivot · piel Super Humanos · base compartida para las plantillas de taller */
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");

const THEME = {
  name: "Campus Pivot · Super Humanos",
  headFontFace: "Barlow Condensed",
  bodyFontFace: "Atkinson Hyperlegible Next",
  colors: {
    dk1: "0D0D0C", // black · text, block
    lt1: "FFFFFF", // white · surface
    dk2: "3D3D38", // neutral-700 · block-track
    lt2: "F7F7F5", // neutral-50 · bg
    accent1: "FCC800", // yellow-400 · action, highlight
    accent2: "8A6B00", // yellow-700 · accent text on white
    accent3: "56564F", // neutral-600 · text-muted
    accent4: "15702F", // success
    accent5: "1D5FBF", // info
    accent6: "B54708", // warning
    hlink: "0D0D0C",
    folHlink: "56564F",
  },
};
const HEX = THEME.colors;

/* Íconos del sistema (bundle.js de Campus Pivot) */
const ICONS = {
  "chevron-right": '<path d="M9 5l7 7-7 7"/>',
  "chevron-left": '<path d="M15 5l-7 7 7 7"/>',
  "chevrons-up": '<path d="M6 13l6-6 6 6"/><path d="M6 19l6-6 6 6"/>',
  play: '<path d="M8 5v14l11-7z"/>',
  live: '<rect x="2.5" y="6" width="13" height="12" rx="2"/><path d="M15.5 10.5l6-3.5v10l-6-3.5"/>',
  camera: '<path d="M3.5 8h3.5l2-3h6l2 3h3.5v11.5h-17z"/><circle cx="12" cy="13" r="3.8"/>',
  upload: '<path d="M12 16V4"/><path d="M7 9l5-5 5 5"/><path d="M4 20h16"/>',
  check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.5 2"/>',
  calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  star: '<path d="M12 3.5l2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 16.9l-5.2 2.7 1-5.8-4.3-4.1 5.9-.9z"/>',
  trophy: '<path d="M8 4h8v5a4 4 0 0 1-8 0z"/><path d="M8 6H4.5v1.5A3.5 3.5 0 0 0 8 11M16 6h3.5v1.5A3.5 3.5 0 0 1 16 11M12 13v4M8 21h8M9.5 17h5v4h-5z"/>',
  home: '<path d="M3.5 11L12 4l8.5 7"/><path d="M6 9.5V20h12V9.5"/><path d="M10 20v-5h4v5"/>',
  route: '<path d="M4 6.5l5-2.5 6 2.5 5-2.5v13.5l-5 2.5-6-2.5-5 2.5z"/><path d="M9 4v13.5M15 6.5V20"/>',
  user: '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 3.6-6.5 8-6.5s8 2.5 8 6.5"/>',
  users: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c0-3.5 3-5.5 6.5-5.5s6.5 2 6.5 5.5"/><path d="M16 4.5a3.5 3.5 0 0 1 0 7M18 14.6c2.3.6 3.5 2.4 3.5 5.4"/>',
  alert: '<path d="M12 3.5l9.5 17h-19z"/><path d="M12 10v4.5M12 17.5v.01"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.01"/>',
  "wifi-off": '<path d="M3 3l18 18"/><path d="M8.5 16.2a5 5 0 0 1 6.2-.6M5 12.6a10 10 0 0 1 4.6-2.4M19 12.6a10 10 0 0 0-2.6-1.8M2 8.8a15 15 0 0 1 4.3-2.6M22 8.8A15 15 0 0 0 11 5M12 20h.01"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>',
  bulb: '<path d="M9 18h6M10 21h4"/><path d="M12 3a6 6 0 0 0-4 10.5c.7.7 1 1.5 1 2.5h6c0-1 .3-1.8 1-2.5A6 6 0 0 0 12 3z"/>',
  rocket: '<path d="M14 4.5c2.8-1 6-1.2 6-1.2s-.2 3.2-1.2 6L12 16l-4-4z"/><circle cx="15.2" cy="8.8" r="1.4"/><path d="M8 12l-3.5.5L3 15l4 .5M12 16l-.5 3.5L9 21l-.5-4M6.5 17.5L4 20"/>',
  tool: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3.5 17.5l3 3 5.8-5.8a4 4 0 0 0 5.4-5.4l-2.4 2.4-2.6-.4-.4-2.6z"/>',
  clipboard: '<rect x="5" y="4.5" width="14" height="16.5" rx="2"/><path d="M9 3h6v3H9zM9 13l2 2 4-4"/>',
  chat: '<path d="M4 5h16v11H10l-6 4.5z"/><path d="M8 9.5h8M8 12.5h5"/>',
  lock: '<rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7.5a4 4 0 0 1 8 0V11"/>',
  refresh: '<path d="M20 12a8 8 0 1 1-2.3-5.7"/><path d="M20 4v5h-5"/>',
  "id-card": '<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2.4"/><path d="M5.6 16.2c.8-1.5 2-2 3.4-2s2.6.5 3.4 2M14.5 10h4M14.5 14h3"/>',
  "circle-empty": '<circle cx="12" cy="12" r="8"/>',
  "circle-half": '<circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 1 0 16z" fill="currentColor"/>',
  "circle-check": '<circle cx="12" cy="12" r="9"/><path d="M7.5 12.5l3 3 6-6.5"/>',
  menu: '<path d="M4 7h16M4 12h16M4 17h16"/>',
  medal: '<circle cx="12" cy="15" r="5.5"/><path d="M8.5 10.5L6 3h4l2 5 2-5h4l-2.5 7.5"/><path d="M12 12.5l.9 1.8 2 .3-1.4 1.4.3 2-1.8-.9-1.8.9.3-2-1.4-1.4 2-.3z"/>',
};

const cache = {};
/* variant: 'ink' (negro), 'yellow', 'white', 'circle' (círculo negro + trazo amarillo), 'circle-yellow' (círculo amarillo + trazo negro) */
async function iconPng(name, variant = "circle") {
  const key = name + "|" + variant;
  if (cache[key]) return cache[key];
  const body = ICONS[name];
  if (!body) throw new Error("icono desconocido: " + name);
  let stroke, bg = "", inner;
  if (variant === "ink") stroke = "#0D0D0C";
  else if (variant === "yellow") stroke = "#FCC800";
  else if (variant === "white") stroke = "#FFFFFF";
  else if (variant === "circle") { stroke = "#FCC800"; bg = '<circle cx="24" cy="24" r="24" fill="#0D0D0C"/>'; }
  else if (variant === "circle-yellow") { stroke = "#0D0D0C"; bg = '<circle cx="24" cy="24" r="24" fill="#FCC800"/>'; }
  else if (variant === "circle-white") { stroke = "#0D0D0C"; bg = '<circle cx="24" cy="24" r="23" fill="#FFFFFF" stroke="#0D0D0C" stroke-width="2"/>'; }
  if (bg) {
    inner = `<g transform="translate(12 12)" fill="none" stroke="${stroke}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" color="${stroke}">${body}</g>`;
  } else {
    inner = `<g transform="translate(12 12)" fill="none" stroke="${stroke}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" color="${stroke}">${body}</g>`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" width="256" height="256">${bg}${inner}</svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  cache[key] = "image/png;base64," + buf.toString("base64");
  return cache[key];
}

/* Geometría */
const W = 13.333, H = 7.5, M = 0.6;
const CONTENT_W = W - 2 * M;

/* Layouts comunes */
function defineLayouts(pres, footerText) {
  const C = pres.SchemeColor;
  const footer = (dark) => ([
    { text: { text: footerText, options: { x: M, y: H - 0.45, w: 8, h: 0.3, fontFace: THEME.headFontFace, fontSize: 11, bold: true, charSpacing: 2, color: dark ? C.accent1 : C.text2, margin: 0, valign: "middle" } } },
  ]);
  const slideNum = (dark) => ({ x: W - M - 0.6, y: H - 0.45, w: 0.6, h: 0.3, fontFace: THEME.headFontFace, fontSize: 11, bold: true, color: dark ? C.accent1 : C.text2, align: "right", margin: 0, valign: "middle" });

  // Portada oscura
  pres.defineSlideMaster({ title: "PORTADA", background: { color: HEX.dk1 }, objects: [] });
  // Divisor de sección oscuro
  pres.defineSlideMaster({ title: "SECCION", background: { color: HEX.dk1 }, objects: [], slideNumber: slideNum(true) });
  // Título y contenido, claro
  pres.defineSlideMaster({
    title: "CONTENIDO_CLARO", background: { color: HEX.lt1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.45, w: CONTENT_W, h: 0.9, fontFace: THEME.headFontFace, fontSize: 36, bold: true, color: C.text1, margin: 0, valign: "top", align: "left", fit: "shrink" }, text: "Título de la diapositiva" } },
      ...footer(false),
    ],
    slideNumber: slideNum(false),
  });
  // Título y contenido, oscuro
  pres.defineSlideMaster({
    title: "CONTENIDO_OSCURO", background: { color: HEX.dk1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.45, w: CONTENT_W, h: 0.9, fontFace: THEME.headFontFace, fontSize: 36, bold: true, color: C.accent1, margin: 0, valign: "top", align: "left", fit: "shrink" }, text: "Título de la diapositiva" } },
      ...footer(true),
    ],
    slideNumber: slideNum(true),
  });
  // Pantalla libre (consultorio, demostración)
  pres.defineSlideMaster({ title: "LIBRE_CLARO", background: { color: HEX.lt1 }, objects: [...footer(false)], slideNumber: slideNum(false) });
  pres.defineSlideMaster({ title: "LIBRE_OSCURO", background: { color: HEX.dk1 }, objects: [...footer(true)], slideNumber: slideNum(true) });
}

/* Piezas */
function title(slide, text) {
  slide.addText(text.toUpperCase(), { placeholder: "title" });
}

/* Antetítulo tipo etiqueta (cp-ribbon--label): chevrón amarillo + texto */
function kicker(slide, text, x, y, dark, w = 6) {
  const C = slide._pres ? slide._pres.SchemeColor : null;
  slide.addShape("chevron", { x, y: y + 0.04, w: 0.16, h: 0.2, fill: { color: HEX.accent1 }, line: { color: HEX.accent1, width: 0 } });
  slide.addText(text.toUpperCase(), { x: x + 0.24, y, w, h: 0.28, fontFace: THEME.headFontFace, fontSize: 13, bold: true, charSpacing: 3, color: dark ? "FCC800" : "8A6B00", margin: 0, valign: "middle", isTextBox: true });
}

/* Cinta de sección (cp-ribbon): banda con cola en V y punta en chevrón */
function ribbon(slide, text, x, y, w, opts = {}) {
  const yellow = !!opts.yellow;
  const h = opts.h || 0.62;
  slide.addShape("chevron", { x: x - 0.28, y: y + 0.11, w: 0.22, h: h - 0.22, fill: { color: yellow ? HEX.dk1 : HEX.accent1 }, line: { width: 0, color: yellow ? HEX.dk1 : HEX.accent1 } });
  slide.addShape("chevron", { x, y, w, h, fill: { color: yellow ? HEX.accent1 : HEX.dk1 }, line: { width: 0, color: yellow ? HEX.accent1 : HEX.dk1 } });
  const off = h * 0.55 + 0.05;
  slide.addText(text.toUpperCase(), { x: x + off, y, w: w - off - 0.4, h, fontFace: THEME.headFontFace, fontSize: opts.fontSize || 22, bold: true, italic: true, color: yellow ? "0D0D0C" : "FFFFFF", margin: 0, valign: "middle", isTextBox: true, fit: "shrink" });
}

/* Barra de chevrones (ChevronProgress): segmentos con rótulo arriba y texto debajo */
function chevronBar(slide, items, x, y, w, opts = {}) {
  // items: [{label, sub, done}] ; done -> amarillo, else riel 3D3D38 (sobre negro) o DEDED9 (sobre blanco)
  const dark = !!opts.dark;
  const gap = 0.06;
  const h = opts.h || 0.5;
  const segW = (w - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const sx = x + i * (segW + gap);
    const shape = i === 0 ? "homePlate" : "chevron";
    const fillC = it.done ? HEX.accent1 : (dark ? HEX.dk2 : "DEDED9");
    slide.addShape(shape, { x: sx, y, w: segW, h, fill: { color: fillC }, line: { width: 0, color: fillC } });
    if (it.label) {
      slide.addText(String(it.label).toUpperCase(), { x: sx + (i === 0 ? 0.12 : 0.25), y, w: segW - 0.45, h, fontFace: THEME.headFontFace, fontSize: opts.fontSize || 14, bold: true, color: it.done ? "0D0D0C" : (dark ? "FFFFFF" : "0D0D0C"), margin: 0, valign: "middle", isTextBox: true, fit: "shrink" });
    }
    if (it.sub) {
      slide.addText(it.sub, { x: sx, y: y + h + 0.08, w: segW, h: 0.7, fontFace: THEME.bodyFontFace, fontSize: opts.subSize || 12, color: dark ? "C4C4BE" : "56564F", margin: 0, valign: "top", isTextBox: true });
    }
  });
}

/* Tarjeta (ContentCard): superficie con borde, radio 8 px ≈ 0.11" */
function card(slide, x, y, w, h, opts = {}) {
  const dark = !!opts.dark;
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.11, fill: { color: dark ? "161614" : (opts.fill || "FFFFFF") }, line: { color: dark ? "3D3D38" : (opts.border || "DEDED9"), width: 1.25 } });
}

/* Bloque de prompt: caja negra, chip amarillo PROMPT, texto blanco */
function promptBlock(slide, text, x, y, w, h, opts = {}) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.11, fill: { color: HEX.dk1 }, line: { color: HEX.dk1, width: 0 } });
  const label = (opts.label || "Prompt").toUpperCase();
  const chipW = 0.55 + 0.115 * label.length;
  slide.addShape("homePlate", { x: x + 0.25, y: y + 0.22, w: chipW, h: 0.32, fill: { color: HEX.accent1 }, line: { color: HEX.accent1, width: 0 } });
  slide.addText(label, { x: x + 0.33, y: y + 0.22, w: chipW - 0.25, h: 0.32, fontFace: THEME.headFontFace, fontSize: 13, bold: true, charSpacing: 2, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true });
  slide.addText(text, { x: x + 0.25, y: y + 0.68, w: w - 0.5, h: h - 0.9, fontFace: THEME.bodyFontFace, fontSize: opts.fontSize || 15, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 6 });
}

/* Ícono en círculo + texto (fila icono + texto) */
async function iconRow(slide, iconName, head, body, x, y, w, opts = {}) {
  const dark = !!opts.dark;
  const d = opts.size || 0.6;
  const png = await iconPng(iconName, dark ? "circle-yellow" : "circle");
  slide.addImage({ data: png, x, y, w: d, h: d });
  slide.addText(head.toUpperCase(), { x: x + d + 0.2, y: y - 0.02, w: w - d - 0.2, h: 0.34, fontFace: THEME.headFontFace, fontSize: opts.headSize || 18, bold: true, color: dark ? "FFFFFF" : "0D0D0C", margin: 0, valign: "top", isTextBox: true });
  if (body) slide.addText(body, { x: x + d + 0.2, y: y + 0.32, w: w - d - 0.2, h: opts.bodyH || 0.8, fontFace: THEME.bodyFontFace, fontSize: opts.bodySize || 14, color: dark ? "C4C4BE" : "3D3D38", margin: 0, valign: "top", isTextBox: true });
}

/* Número grande (stat callout) */
function bigNumber(slide, num, label, x, y, w, opts = {}) {
  const dark = !!opts.dark;
  slide.addText(String(num), { x, y, w, h: 1.0, fontFace: THEME.headFontFace, fontSize: opts.size || 60, bold: true, color: dark ? "FCC800" : "0D0D0C", margin: 0, valign: "bottom", isTextBox: true, align: opts.align || "left" });
  slide.addText(label, { x, y: y + 1.02, w, h: 0.5, fontFace: THEME.bodyFontFace, fontSize: 14, color: dark ? "C4C4BE" : "56564F", margin: 0, valign: "top", isTextBox: true, align: opts.align || "left" });
}

/* Logotipo tipográfico SUPER HUMANOS (BrandMark) */
function wordmark(slide, x, y, opts = {}) {
  const size = opts.size || 28;
  const dark = opts.dark !== false;
  const wA = size * 0.042, wB = size * 0.062, h = size * 0.02;
  slide.addText("SUPER", { x, y, w: wA, h, fontFace: THEME.headFontFace, fontSize: size, bold: true, color: dark ? "FFFFFF" : "0D0D0C", margin: 0, valign: "middle", isTextBox: true, wrap: false });
  slide.addShape("rect", { x: x + wA, y: y + h * 0.1, w: wB, h: h * 0.8, fill: { color: HEX.accent1 }, line: { width: 0, color: HEX.accent1 } });
  slide.addText("HUMANOS", { x: x + wA + 0.06, y, w: wB, h, fontFace: THEME.headFontFace, fontSize: size, bold: true, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true, wrap: false });
}

/* Texto cuerpo estándar */
function body(slide, text, x, y, w, h, opts = {}) {
  const dark = !!opts.dark;
  slide.addText(text, { x, y, w, h, fontFace: THEME.bodyFontFace, fontSize: opts.fontSize || 16, color: dark ? "F7F7F5" : "0D0D0C", margin: 0, valign: opts.valign || "top", isTextBox: true, paraSpaceAfter: opts.paraSpaceAfter || 8, bullet: opts.bullet, align: opts.align });
}

function bullets(slide, items, x, y, w, h, opts = {}) {
  const dark = !!opts.dark;
  slide.addText(items.map((t, i) => ({ text: t, options: { bullet: { code: "25B8" }, breakLine: i < items.length - 1 } })), { x, y, w, h, fontFace: THEME.bodyFontFace, fontSize: opts.fontSize || 16, color: dark ? "F7F7F5" : "0D0D0C", margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: opts.paraSpaceAfter || 8 });
}

/* Guía de uso: tabla de marcadores */
function guideSlide(pres, fileCode, markers, rules, sectionTitle) {
  const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle });
  title(s, "Guía de uso de esta plantilla");
  kicker(s, "Nombre del archivo · " + fileCode, M, 1.35, false, 8);
  const rows = [[
    { text: "Marcador", options: { bold: true, fill: { color: "0D0D0C" }, color: "FCC800", fontFace: THEME.headFontFace, fontSize: 13 } },
    { text: "Qué va ahí", options: { bold: true, fill: { color: "0D0D0C" }, color: "FCC800", fontFace: THEME.headFontFace, fontSize: 13 } },
  ]];
  markers.forEach((m) => rows.push([{ text: m[0], options: { fontFace: THEME.bodyFontFace, fontSize: 12, bold: true } }, { text: m[1], options: { fontFace: THEME.bodyFontFace, fontSize: 12 } }]));
  s.addTable(rows, { x: M, y: 1.8, w: 6.3, colW: [2.1, 4.2], border: { type: "solid", color: "DEDED9", pt: 0.75 }, fill: { color: "FFFFFF" }, color: "0D0D0C", margin: 0.06, valign: "middle", rowH: 0.34 });
  ribbon(s, "Reglas que no cambian", 7.55, 1.75, 5.1, { fontSize: 18, h: 0.5 });
  bullets(s, rules, 7.3, 2.5, 5.4, 4.2, { fontSize: 13, paraSpaceAfter: 6 });
  s.addNotes("Última diapositiva de la plantilla. Se borra antes de presentar. Nombre del archivo: " + fileCode + ". El pie de página se edita una sola vez por archivo en Vista > Patrón de diapositivas.");
  return s;
}

/* Biblioteca de íconos */
async function iconLibrarySlide(pres, sectionTitle) {
  const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle });
  title(s, "Biblioteca de íconos");
  kicker(s, "Copia y pega el ícono en la diapositiva que lo necesite", M, 1.35, false, 8);
  const names = Object.keys(ICONS);
  const cols = 11, d = 0.62, cw = CONTENT_W / cols;
  for (let i = 0; i < names.length; i++) {
    const r = Math.floor(i / cols), c = i % cols;
    const x = M + c * cw + (cw - d) / 2, y = 1.85 + r * 1.55;
    s.addImage({ data: await iconPng(names[i], "circle"), x, y, w: d, h: d });
    s.addImage({ data: await iconPng(names[i], "ink"), x: x + 0.05, y: y + d + 0.05, w: 0.28, h: 0.28 });
    s.addImage({ data: await iconPng(names[i], "circle-yellow"), x: x + 0.36, y: y + d + 0.05, w: 0.28, h: 0.28 });
    s.addText(names[i], { x: M + c * cw, y: y + d + 0.38, w: cw, h: 0.25, fontFace: THEME.bodyFontFace, fontSize: 9, color: "56564F", align: "center", margin: 0, isTextBox: true });
  }
  s.addNotes("Treinta y tres íconos del sistema Campus Pivot en tres variantes: círculo negro con trazo amarillo, trazo negro suelto y círculo amarillo con trazo negro. Son imágenes PNG de 512 px. El sprite SVG original vive en el design system.");
  return s;
}

module.exports = { pptxgen, THEME, HEX, ICONS, iconPng, W, H, M, CONTENT_W, defineLayouts, title, kicker, ribbon, chevronBar, card, promptBlock, iconRow, bigNumber, wordmark, body, bullets, guideSlide, iconLibrarySlide };
