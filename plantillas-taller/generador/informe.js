const { build, B } = require("./common.js");
const HF = B.THEME.headFontFace, BF = B.THEME.bodyFontFace;
const LOGO_D = __dirname + "/pivot-power-logo-oscuro.png", LOGO_L = __dirname + "/pivot-power-logo-claro.png";
const LR = 546 / 2020;

function tbl(s, head, rows, x, y, w, colW, opts = {}) {
  const fs = opts.fontSize || 12;
  const H = head.map((t) => ({ text: t, options: { bold: true, fontFace: HF, fontSize: fs + 1, color: "0D0D0C", fill: { color: "FCC800" }, valign: "middle" } }));
  const R = rows.map((r, i) => r.map((t, j) => ({ text: String(t), options: { fontFace: BF, fontSize: fs, color: "0D0D0C", bold: !!(opts.boldLast && i === rows.length - 1), fill: { color: (opts.boldLast && i === rows.length - 1) ? "F7F7F5" : "FFFFFF" }, align: j === 0 ? "left" : (opts.alignRest || "center"), valign: "middle" } })));
  s.addTable([H, ...R], { x, y, w, colW, border: { type: "solid", pt: 0.75, color: "DEDED9" }, rowH: opts.rowH || 0.32, margin: 0.06 });
}
function section(pres, n, t, sub) {
  const s = pres.addSlide({ masterName: "SECCION" });
  s.addText(n, { x: B.M, y: 1.9, w: 3, h: 1.2, fontFace: HF, fontSize: 96, bold: true, color: "FCC800", margin: 0, isTextBox: true });
  B.ribbon(s, t, B.M + 0.3, 3.3, 9, { yellow: true, fontSize: 34, h: 0.85 });
  B.body(s, sub, B.M, 4.5, 10, 1, { dark: true, fontSize: 18 });
}
function note(s, t, dark) { s.addText(t, { x: B.M, y: 6.55, w: 12, h: 0.3, fontFace: BF, fontSize: 10.5, color: dark ? "9A9A93" : "56564F", margin: 0, isTextBox: true }); }

build("Super Humanos · Informe de avance · G&J · 10 nov 2026.pptx", "SUPER HUMANOS · INFORME DE AVANCE · G&J EMPRESAS DE ACERO", async (pres) => {
  /* 1 · Portada */
  let s = pres.addSlide({ masterName: "PORTADA" });
  B.wordmark(s, B.M, 0.6, { size: 26 });
  s.addImage({ path: LOGO_D, x: 10.15, y: 0.5, w: 2.55, h: 2.55 * LR });
  B.kicker(s, "Informe de avance para el equipo de G&J", B.M, 2.2, true, 10);
  s.addText("ASÍ VA SUPER HUMANOS", { x: B.M, y: 2.55, w: 12, h: 1.3, fontFace: HF, fontSize: 66, bold: true, color: "FCC800", margin: 0, isTextBox: true });
  s.addText("Diagnóstico, retos de los jueves, análisis y próximas sesiones", { x: B.M, y: 3.85, w: 12, h: 0.6, fontFace: HF, fontSize: 28, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
  B.ribbon(s, "Etapa 1 · comercial y administrativo", B.M + 0.3, 4.85, 6.2, { yellow: true, fontSize: 20, h: 0.55 });
  s.addText([{ text: "10 de noviembre de 2026", options: { bold: true, color: "FFFFFF" } }, { text: "  ·  Pivot Power", options: { color: "C4C4BE" } }], { x: B.M, y: 6.2, w: 8, h: 0.4, fontFace: BF, fontSize: 16, margin: 0, isTextBox: true });

  /* 2 · Resumen */
  s = pres.addSlide({ masterName: "CONTENIDO_OSCURO" });
  B.title(s, "El programa en cuatro cifras");
  const nums = [["79 %", "terminó el diagnóstico: 536 de 680 personas"], ["67 %", "envió al menos un reto: 454 de 676"], ["34 %", "envió 4 retos o más: 227 personas"], ["50", "candidatos a formador interno salen del diagnóstico"]];
  nums.forEach(([n, l], i) => B.bigNumber(s, n, l, B.M + i * 3.05, 1.7, 2.8, { dark: true, size: 64 }));
  B.kicker(s, "Lo que sigue", B.M, 3.95, true);
  B.bullets(s, [
    "Taller 1: consultorio el jueves 12 de noviembre y segunda sesión el 18 (administrativo) y el 19 (comercial).",
    "Taller 2 en diciembre: arranca el viernes 4 con el perfil comercial.",
    "Cerrar el diagnóstico de G&J Ferreterías, que concentra el faltante, y mover la participación en retos de esa empresa.",
  ], B.M, 4.35, 12, 2, { dark: true, fontSize: 17 });
  note(s, "Fuentes: plataforma Pipolearning y consolidado de retos, corte del 9 de noviembre de 2026.", true);

  /* 3 · Avance del proyecto */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "Avance del proyecto");
  B.chevronBar(s, [
    { label: "Alistamiento", sub: "Plan, cronograma, plataforma y docentes", done: true },
    { label: "Diagnóstico", sub: "79 % terminado · análisis por perfil", done: true },
    { label: "Taller 1", sub: "Sesión 1 el 23 de octubre · cierra el 19 de noviembre", done: true },
    { label: "Talleres 2 a 6", sub: "Diciembre a abril", done: false },
    { label: "Cierre", sub: "Proyectos, Banco de Retos y diagnóstico final", done: false },
  ], B.M, 1.55, 12.1, { h: 0.6, fontSize: 15, subSize: 12 });
  B.card(s, B.M, 3.15, 5.9, 3.25); B.card(s, B.M + 6.2, 3.15, 5.9, 3.25);
  B.kicker(s, "Hecho", B.M + 0.25, 3.3, false);
  B.bullets(s, [
    "Diagnóstico aplicado en la plataforma, con análisis por perfil y grupos de nivel.",
    "Docentes asignados: Mateo Peralta en comercial y José Cormane en administrativo, con su paquete de preparación.",
    "Los 12 talleres y sus 36 encuentros cargados en la plataforma.",
    "Retos de los jueves 1 a 6 publicados; 454 personas ya participaron.",
    "Propuesta de política de uso de IA preparada para TIC: fases, información por cuenta y licencias.",
  ], B.M + 0.25, 3.7, 5.4, 2.6, { fontSize: 13.5, paraSpaceAfter: 5 });
  B.kicker(s, "En curso", B.M + 6.45, 3.3, false);
  B.bullets(s, [
    "Segunda sesión y consultorio del Taller 1.",
    "Respuesta de TIC sobre herramientas, cuentas y dominios autorizados.",
    "Revisión de las 16 postulaciones a formador interno, con el visto bueno de cada líder.",
    "Cierre del diagnóstico con las 144 personas que faltan.",
    "Definición de premios y canje de puntos por empresa.",
  ], B.M + 6.45, 3.7, 5.4, 2.6, { fontSize: 13.5, paraSpaceAfter: 5 });

  /* 4 · Sección diagnóstico */
  section(pres, "01", "Avance del diagnóstico", "Quiénes respondieron, por empresa, y qué tanto falta.");

  /* 5 · Avance por empresa */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "536 de 680 personas terminaron el diagnóstico");
  tbl(s, ["Empresa", "Personas", "Diagnóstico"], [
    ["Sigmasteel", "55", "95 %"], ["Colmena", "79", "95 %"], ["G&J Soluciones", "82", "82 %"], ["G&J Ramírez", "20", "80 %"],
    ["Almasa", "81", "79 %"], ["Almapanel", "4", "75 %"], ["G&J Ferreterías", "355", "72 %"], ["Total", "676", "79 %"],
  ], B.M, 1.55, 6.2, [3.2, 1.5, 1.5], { boldLast: true, rowH: 0.42, fontSize: 13 });
  B.bigNumber(s, "62 % → 79 %", "Cobertura del 14 de octubre al 9 de noviembre", 7.4, 1.4, 5.3, { size: 48 });
  B.bullets(s, [
    "Sigmasteel y Colmena casi cerraron: tuvieron un responsable haciendo seguimiento persona por persona.",
    "G&J Ferreterías tiene más de la mitad de la población y es la que más falta.",
    "Con el 79 %, los grupos y los resultados por perfil ya son estables.",
  ], 7.4, 3.2, 5.3, 3, { fontSize: 14.5 });
  note(s, "Fuente: plataforma, 9 de noviembre de 2026. El total por empresa usa la base de retos (676); la plataforma reporta 680 personas activas.");

  /* 6 · Sección análisis */
  section(pres, "02", "Lo que dice el diagnóstico", "Análisis de 482 personas con índice completo, corte del 23 de octubre.");

  /* 7 · Índices */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "El uso de la IA es lo que más hay que mover");
  tbl(s, ["Eje", "Peso", "Comercial", "Administrativo"], [
    ["Casos de su propio trabajo", "40 %", "67,6", "88,1"], ["Herramientas digitales", "35 %", "80,4", "85,7"], ["Uso actual de IA", "25 %", "56,0", "57,1"],
    ["Índice de dominio", "", "69,2", "79,5"], ["Mentalidad (aparte)", "", "76,2", "77,2"],
  ], B.M, 1.55, 7.2, [3.4, 0.9, 1.45, 1.45], { rowH: 0.45, fontSize: 13.5 });
  B.card(s, 8.1, 1.55, 4.6, 4.6, { fill: "F7F7F5" });
  B.kicker(s, "Lectura", 8.35, 1.75, false);
  B.bullets(s, [
    "Los dos perfiles resuelven bien su trabajo y manejan las herramientas digitales.",
    "El uso de IA queda 20 a 30 puntos por debajo: es el eje donde el programa tiene más espacio para mostrar avance.",
    "Cada perfil tiene casos distintos; el índice sirve para comparar personas dentro del mismo perfil y medir el avance al cierre.",
  ], 8.35, 2.15, 4.15, 3.9, { fontSize: 14 });
  note(s, "Escala de 0 a 100. Mismo método del primer corte, verificado sobre las 415 personas de ese corte.");

  /* 8 · Situaciones */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "Las situaciones donde más se traban");
  B.kicker(s, "Comercial · 150 personas", B.M, 1.5, false);
  tbl(s, ["Situación", "Acierto"], [["15 solicitudes de cotización por WhatsApp un lunes", "30 %"], ["Dos horas cada lunes revisando 40 correos", "47 %"], ["Presentar resultados del trimestre en 10 minutos", "63 %"], ["Pedido retenido por cartera y el cliente lo necesita el jueves", "75 %"]], B.M, 1.9, 5.9, [4.8, 1.1], { rowH: 0.5, fontSize: 12.5 });
  B.kicker(s, "Administrativo · 332 personas", B.M + 6.2, 1.5, false);
  tbl(s, ["Situación", "Acierto"], [["Presentar resultados del trimestre en 10 minutos", "75 %"], ["El ingreso de cada persona exige 30 firmas en papel", "77 %"], ["Construir un documento entre cuatro personas", "79 %"], ["Consolidar el reporte de cinco archivos cada lunes", "84 %"]], B.M + 6.2, 1.9, 5.9, [4.8, 1.1], { rowH: 0.5, fontSize: 12.5 });
  B.card(s, B.M, 4.75, 12.1, 1.55, { fill: "F7F7F5" });
  B.body(s, [
    { text: "Comercial: ", options: { bold: true } }, { text: "organizar el volumen repetitivo que llega por canales desordenados. El Taller 1 y el Taller 3 trabajan exactamente esos dos casos.", options: { breakLine: true } },
    { text: "Administrativo: ", options: { bold: true } }, { text: "comunicar resultados y rediseñar procesos, más que ejecutarlos. El Taller 1 abre por ahí y el trámite de las 30 firmas es el caso del Taller 4." },
  ], B.M + 0.25, 4.95, 11.6, 1.2, { fontSize: 14 });

  /* 9 · Cómo trabajan */
  s = pres.addSlide({ masterName: "CONTENIDO_OSCURO" });
  B.title(s, "Cómo trabaja la gente hoy");
  const st = [["47 %", "usó la IA una o dos veces o ninguna en el último mes"], ["51 %", "usa la IA con cuenta personal"], ["35 %", "no ha firmado la política de IA o no sabía que existía"], ["69 %", "lleva sus pendientes en papel o marcando correos"], ["32 %", "dedica más de 5 horas a la semana a tareas repetitivas"], ["38 %", "hace informes que también hace otra área"]];
  st.forEach(([n, l], i) => B.bigNumber(s, n, l, B.M + (i % 3) * 4.1, 1.5 + Math.floor(i / 3) * 2.35, 3.7, { dark: true, size: 54 }));
  note(s, "Sobre 482 personas con diagnóstico completo, corte del 23 de octubre. 33 % reporta herramientas de IA bloqueadas total o parcialmente desde la red.", true);

  /* 10 · Grupos y formadores */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "Grupos de nivel y formadores internos");
  tbl(s, ["Grupo", "Comercial", "Administrativo"], [["G1 Arranque", "38", "82"], ["G2 Base", "36", "84"], ["G3 Avance", "38", "82"], ["G4 Referentes", "38", "84"]], B.M, 1.55, 6, [2.6, 1.7, 1.7], { rowH: 0.45, fontSize: 13.5 });
  B.body(s, "Los grupos se arman por cuartiles dentro de cada perfil. G1 marca el ritmo del Taller 1 y tiene el consultorio reservado. G4 apoya a dos compañeros durante la práctica.", B.M, 4.0, 6, 1.6, { fontSize: 14 });
  B.bigNumber(s, "50", "candidatos a formador interno: 17 comerciales y 33 administrativos", 7.2, 1.4, 5.5, { size: 72 });
  B.bullets(s, [
    "Criterio: nivel en el cuartil superior de su perfil, mentalidad de 80 o más y uso de IA al menos una vez por semana.",
    "Por empresa: G&J Ferreterías 19, G&J Soluciones 10, Almasa 6, Colmena 5, G&J Ramírez 5, Sigmasteel 4, Almapanel 1.",
    "En la plataforma hay 16 postulaciones esperando el visto bueno del líder.",
  ], 7.2, 3.1, 5.5, 3.2, { fontSize: 13.5 });

  /* 11 · Sección retos */
  section(pres, "03", "Retos de los jueves", "Participación en los retos 1 a 6, por empresa y por constancia.");

  /* 12 · Participación por reto */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "Participación por reto");
  s.addChart(pres.ChartType.bar, [{ name: "% de inscritos", labels: ["Reto 1", "Reto 2", "Reto 3", "Reto 4", "Reto 5", "Reto 6"], values: [51, 43, 41, 34, 40, 18] }], {
    x: B.M, y: 1.45, w: 7.2, h: 4.9, barDir: "col", chartColors: ["FCC800"], showValue: true, dataLabelFormatCode: '0"%"', dataLabelFontFace: HF, dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelColor: "0D0D0C",
    catAxisLabelFontFace: HF, catAxisLabelFontSize: 14, catAxisLabelColor: "0D0D0C", valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 60, valAxisMinVal: 0, barGapWidthPct: 55,
  });
  B.card(s, 8.2, 1.55, 4.5, 4.75, { fill: "F7F7F5" });
  B.kicker(s, "Lectura", 8.45, 1.75, false);
  B.bullets(s, [
    "El reto 1 llegó a la mitad de los inscritos. La participación se sostiene alrededor del 40 % hasta el reto 5.",
    "El reto 6 lleva una semana publicado: los retos 3 a 6 siguen abiertos hasta el 26 de noviembre.",
    "Comercial arrancó más alto en el reto 1 (57 %) y administrativo sostiene mejor el reto 5 (41 %).",
  ], 8.45, 2.15, 4.05, 4, { fontSize: 14 });
  note(s, "Fuente: consolidado de retos, 9 de noviembre de 2026. Retos 1 y 2 por correo; retos 3 a 6 en la plataforma y por correo. Base: 676 inscritos.");

  /* 13 · Constancia y empresas */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "Constancia y participación por empresa");
  s.addChart(pres.ChartType.bar, [{ name: "Personas", labels: ["6 retos", "5", "4", "3", "2", "1", "Ninguno"], values: [40, 92, 95, 71, 80, 76, 222] }], {
    x: B.M, y: 1.45, w: 5.2, h: 4.5, barDir: "bar", chartColors: ["0D0D0C"], showValue: true, dataLabelFontFace: HF, dataLabelFontSize: 13, dataLabelFontBold: true,
    catAxisLabelFontFace: HF, catAxisLabelFontSize: 13, catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 40,
  });
  tbl(s, ["Empresa", "Reto 1", "Reto 5", "4 o más", "Ninguno"], [
    ["Sigmasteel", "80 %", "62 %", "36", "5"], ["Colmena", "51 %", "54 %", "41", "11"], ["G&J Soluciones", "59 %", "48 %", "35", "18"], ["G&J Ramírez", "40 %", "50 %", "6", "8"],
    ["Almasa", "36 %", "44 %", "22", "35"], ["G&J Ferreterías", "49 %", "29 %", "86", "144"], ["Total", "51 %", "40 %", "227", "222"],
  ], 6.2, 1.55, 6.5, [2.3, 1.05, 1.05, 1.05, 1.05], { boldLast: true, rowH: 0.42, fontSize: 12.5 });
  B.body(s, "Dos de cada tres personas que no han enviado ningún reto son de G&J Ferreterías.", 6.2, 5.2, 6.5, 0.8, { fontSize: 14 });
  note(s, "Almapanel (4 personas) incluida en el total.");

  /* 14 · Sección cronograma */
  section(pres, "04", "Próximas sesiones", "Encuentros en vivo por Teams, de noviembre a abril.");

  /* 15 · Cronograma */
  s = pres.addSlide({ masterName: "CONTENIDO_CLARO" });
  B.title(s, "Cronograma de las próximas sesiones");
  tbl(s, ["Taller", "Perfil", "Sesión 1", "Consultorio · 4:00 p. m.", "Sesión 2"], [
    ["1 · La IA no sabe nada de tu cliente", "Comercial", "Vie 23 oct · hecha", "Jue 12 nov", "Jue 19 nov · 7:30"],
    ["1 · La IA no consulta tus sistemas", "Administrativo", "Vie 23 oct · hecha", "Jue 12 nov · 2:00 p. m.", "Mié 18 nov · 8:00"],
    ["2 · Tu cartera tiene cosas que contarte", "Comercial", "Vie 4 dic · 7:30", "Mié 9 dic", "Vie 11 dic · 7:30"],
    ["2 · Tres días armando el reporte", "Administrativo", "Jue 10 dic · 8:00", "Mar 15 dic", "Jue 17 dic · 8:00"],
    ["3 · El pedido no se demora trabajando", "Comercial", "Mié 13 ene · 7:30", "Lun 18 ene", "Mié 20 ene · 7:30"],
    ["3 · De la bandeja al tablero", "Administrativo", "Mié 13 ene · 8:00", "Vie 15 ene", "Mié 20 ene · 8:00"],
  ], B.M, 1.5, 12.1, [3.8, 1.6, 2.2, 2.3, 2.2], { rowH: 0.5, fontSize: 12 });
  tbl(s, ["Talleres 4 a 6", "Comercial", "Administrativo"], [
    ["4 · Contexto antes que modelo", "Vie 5 feb · Vie 12 feb", "Mié 10 feb · Mié 17 feb"],
    ["5 · Humano + IA", "Vie 5 mar · Vie 12 mar", "Mié 10 mar · Mié 17 mar"],
    ["6 · Mide, mejora, demuestra", "Vie 9 abr · Vie 16 abr", "Mié 14 abr · Mié 21 abr"],
  ], B.M, 5.15, 12.1, [4.5, 3.8, 3.8], { rowH: 0.36, fontSize: 11.5 });
  note(s, "Fuente: encuentros cargados en la plataforma al 10 de noviembre. Comercial a las 7:30 a. m. y administrativo a las 8:00 a. m., salvo donde se indica.");

  /* 16 · Pedidos */
  s = pres.addSlide({ masterName: "CONTENIDO_OSCURO" });
  B.title(s, "Lo que necesitamos de G&J");
  const asks = [
    ["users", "Cerrar el diagnóstico", "Un responsable por empresa, empezando por G&J Ferreterías, para las 144 personas que faltan."],
    ["chat", "Convocar el Taller 1", "Invitaciones al consultorio del 12 de noviembre y a la segunda sesión del 18 y el 19."],
    ["lock", "Herramientas autorizadas", "Respuesta de TIC sobre la propuesta de política: herramientas, cuentas y dominios habilitados."],
    ["medal", "Formadores y premios", "Visto bueno de los líderes a las 16 postulaciones y definición de premios por empresa."],
  ];
  for (let i = 0; i < asks.length; i++) {
    const [ic, h, b] = asks[i];
    await B.iconRow(s, ic, h, b, B.M + (i % 2) * 6.15, 1.7 + Math.floor(i / 2) * 2.1, 5.8, { dark: true, size: 0.7, headSize: 20, bodySize: 15, bodyH: 1.2 });
  }
});
