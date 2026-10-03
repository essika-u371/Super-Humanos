const { build, cover, B } = require("./common.js");
const { M, CONTENT_W, W } = B;

const SEC = "Consultorio";

build("Plantilla taller · Consultorio.pptx", "Super Humanos · [PERFIL] · Taller [N] · Consultorio", async (pres) => {
  pres.addSection({ title: SEC });

  /* 1 · Portada */
  cover(pres, SEC, {
    encuentro: "Consultorio · hora abierta de dudas",
    kicker: "[PERFIL] · Taller [N] de 6",
    duracion: "60 minutos · 4:00 a 5:00 p. m.",
    notas: "Portada del consultorio. Cambia los marcadores. La hora es fija: 4:00 a 5:00 p. m. Si cambia la regla, cámbiala aquí y en la diapositiva 3.",
  });

  /* 2 · Para qué sirve esta hora */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Para qué sirve esta hora");
    B.kicker(s, "Entre el encuentro 1 y el encuentro 2", M, 1.35, false);
    const rows = [["chat", "Resolver tu duda con tu caso", "Traes la tarea que elegiste para «Mi proceso» o el reto de la semana, y la trabajamos en pantalla."], ["users", "Aprender de la duda de otros", "La mayoría de las dudas se repiten. Lo que se resuelve para una persona sirve para su equipo."], ["clipboard", "Llegar listo al encuentro 2", "Sales con tu formato «Mi proceso» más avanzado y con el prompt ajustado a tu tarea."]];
    for (let i = 0; i < 3; i++) await B.iconRow(s, rows[i][0], rows[i][1], rows[i][2], M, 1.95 + i * 1.3, 8.2, { size: 0.7, headSize: 22, bodySize: 16, bodyH: 0.9 });
    B.card(s, 9.4, 1.95, 3.33, 3.75, { fill: "0D0D0C", border: "0D0D0C" });
    s.addText("SIN", { x: 9.4, y: 2.25, w: 3.33, h: 0.6, fontFace: B.THEME.headFontFace, fontSize: 40, bold: true, color: "FCC800", align: "center", margin: 0, valign: "middle", isTextBox: true });
    B.body(s, "asistencia obligatoria\npuntos\ngrabación de tu cámara\npreguntas tontas", 9.6, 2.95, 2.95, 2.6, { dark: true, fontSize: 16, align: "center" });
    s.addNotes("Para qué sirve. Tres razones para venir y la tarjeta negra con lo que el consultorio no exige. Este contenido no cambia entre talleres.");
  }

  /* 3 · Agenda */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Cómo va la hora");
    B.kicker(s, "[FECHA CONSULTORIO] · 4:00 a 5:00 p. m. · Teams", M, 1.35, false, 9);
    B.chevronBar(s, [
      { label: "0–5", sub: "Apertura\nTres preguntas para arrancar", done: true },
      { label: "5–45", sub: "Dudas y resolución en vivo\nUna por una, con la herramienta en pantalla", done: true },
      { label: "45–55", sub: "Avance de «Mi proceso»\nDónde debería ir cada persona hoy", done: true },
      { label: "55–60", sub: "Cierre\nQué traer al encuentro 2", done: true },
    ], M, 1.95, CONTENT_W, { h: 0.55, fontSize: 15, subSize: 13 });
    B.ribbon(s, "El orden lo ponen las dudas que lleguen", M + 0.3, 4.2, 6.2, { fontSize: 18, h: 0.5 });
    B.body(s, "Si nadie trae una duda, el docente abre con un caso del taller y lo resuelve en vivo. Si llegan más dudas de las que caben, las que quedan pendientes pasan al encuentro 2.", M, 4.95, 8.5, 1.3, { fontSize: 16 });
    s.addNotes("Agenda del consultorio. Los minutos son orientativos: la hora se adapta a las dudas que lleguen.");
  }

  /* 4 · Preguntas de apertura */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Para arrancar");
    B.kicker(s, "Tres preguntas, por si nadie habla", M, 1.35, true);
    const qs = ["[¿Qué tarea elegiste para «Mi proceso» y en qué paso vas?]", "[¿Qué le pediste a la herramienta esta semana y qué te devolvió?]", "[¿Qué parte del reto te tomó más de cinco minutos?]"];
    for (let i = 0; i < 3; i++) {
      const y = 1.95 + i * 1.45;
      s.addShape("homePlate", { x: M, y: y + 0.1, w: 0.95, h: 0.8, fill: { color: "FCC800" }, line: { width: 0, color: "FCC800" } });
      s.addText(String(i + 1), { x: M, y: y + 0.1, w: 0.8, h: 0.8, fontFace: B.THEME.headFontFace, fontSize: 36, bold: true, color: "0D0D0C", align: "center", margin: 0, valign: "middle", isTextBox: true });
      s.addText(qs[i], { x: M + 1.25, y, w: 10.8, h: 1.0, fontFace: B.THEME.headFontFace, fontSize: 30, bold: true, color: "FFFFFF", margin: 0, valign: "middle", isTextBox: true, fit: "shrink" });
    }
    s.addNotes("Preguntas de apertura. Tres como máximo. Escríbelas para el taller: las de la plantilla son ejemplos.");
  }

  /* 5 · Tablero de dudas */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Tablero de dudas");
    B.kicker(s, "Se llena en vivo · el moderador anota mientras el docente resuelve", M, 1.35, false, 10);
    const hdr = (t) => ({ text: t, options: { bold: true, fill: { color: "0D0D0C" }, color: "FCC800", fontFace: B.THEME.headFontFace, fontSize: 14 } });
    const cell = (t, b) => ({ text: t, options: { fontFace: B.THEME.bodyFontFace, fontSize: 13, bold: !!b, color: "0D0D0C" } });
    const rows = [[hdr("#"), hdr("Duda"), hdr("Quién"), hdr("Estado")]];
    const estados = ["Resuelta", "Pendiente", "Va al encuentro 2", "", "", ""];
    for (let i = 0; i < 6; i++) rows.push([cell(String(i + 1), true), cell(i < 3 ? "[Duda en una línea]" : ""), cell(i < 3 ? "[Nombre · área]" : ""), cell(estados[i], true)]);
    s.addTable(rows, { x: M, y: 1.85, w: CONTENT_W, colW: [0.6, 7.3, 2.4, 1.83], border: { type: "solid", color: "DEDED9", pt: 0.75 }, fill: { color: "FFFFFF" }, margin: 0.08, valign: "middle", rowH: 0.52 });
    // leyenda de estados
    const leg = [["circle-check", "Resuelta"], ["circle-half", "Pendiente"], ["chevron-right", "Va al encuentro 2"]];
    for (let i = 0; i < 3; i++) {
      s.addImage({ data: await B.iconPng(leg[i][0], "ink"), x: M + i * 2.6, y: 5.75, w: 0.3, h: 0.3 });
      s.addText(leg[i][1], { x: M + i * 2.6 + 0.38, y: 5.75, w: 2.1, h: 0.3, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "56564F", margin: 0, valign: "middle", isTextBox: true });
    }
    s.addNotes("Tablero de dudas. Se llena en vivo. Cada estado lleva palabra e ícono, sin depender del color. Si hay más de seis dudas, duplica la diapositiva.");
  }

  /* 6 · Resolución en vivo (pantalla libre) */
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: SEC });
    B.kicker(s, "Resolviendo · [DUDA EN UNA LÍNEA]", M, 0.5, false, 10);
    s.addShape("roundRect", { x: M, y: 0.95, w: CONTENT_W, h: 5.7, rectRadius: 0.11, fill: { color: "F7F7F5" }, line: { color: "9A9A93", width: 1.25, dashType: "dash" } });
    s.addImage({ data: await B.iconPng("live", "circle"), x: W / 2 - 0.35, y: 3.1, w: 0.7, h: 0.7 });
    s.addText("Pantalla compartida con la herramienta", { x: M, y: 3.95, w: CONTENT_W, h: 0.5, fontFace: B.THEME.bodyFontFace, fontSize: 16, color: "56564F", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Resolución en vivo. Cambia el texto del antetítulo por la duda que estás resolviendo y comparte pantalla. El recuadro punteado es solo una guía.");
  }

  /* 6b · Resolución con prompt arriba */
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: SEC });
    B.kicker(s, "Resolviendo · [DUDA EN UNA LÍNEA]", M, 0.5, false, 10);
    B.promptBlock(s, "[PROMPT QUE SE ESTÁ PROBANDO]", M, 0.95, CONTENT_W, 1.9, { fontSize: 15 });
    s.addShape("roundRect", { x: M, y: 3.1, w: CONTENT_W, h: 3.55, rectRadius: 0.11, fill: { color: "F7F7F5" }, line: { color: "9A9A93", width: 1.25, dashType: "dash" } });
    s.addText("Resultado o pantalla compartida", { x: M, y: 4.65, w: CONTENT_W, h: 0.5, fontFace: B.THEME.bodyFontFace, fontSize: 16, color: "56564F", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Variante con el prompt visible arriba, para que la gente lo copie mientras el docente lo prueba.");
  }

  /* 7 · Avance de Mi proceso */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "¿Dónde vas con «Mi proceso»?");
    B.kicker(s, "Minuto 45 al 55 · revisa tu formato mientras hablamos", M, 1.35, false, 9);
    const cols = [["clipboard", "1 · La tarea", "Ya elegiste una tarea que se repite y la escribiste con nombre, área y frecuencia.", "Hoy deberías tenerla"], ["clock", "2 · Cuánto toma", "Ya la mediste al menos una vez: minutos por vez y veces por semana.", "Hoy deberías ir en esto"], ["bulb", "3 · La mejora", "Qué parte podría apoyar la IA y qué resultado esperas. Se termina en el encuentro 2.", "Se cierra el [FECHA E2]"]];
    const cw = (CONTENT_W - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = M + i * (cw + 0.3);
      B.card(s, x, 1.95, cw, 4.2, { fill: i === 1 ? "FFF9E0" : "FFFFFF", border: i === 1 ? "FCC800" : "DEDED9" });
      s.addImage({ data: await B.iconPng(cols[i][0], "circle"), x: x + 0.3, y: 2.25, w: 0.7, h: 0.7 });
      s.addText(cols[i][1].toUpperCase(), { x: x + 0.3, y: 3.1, w: cw - 0.6, h: 0.45, fontFace: B.THEME.headFontFace, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true });
      B.body(s, cols[i][2], x + 0.3, 3.6, cw - 0.6, 1.5, { fontSize: 15 });
      s.addShape("homePlate", { x: x + 0.3, y: 5.4, w: cw - 0.6, h: 0.42, fill: { color: "0D0D0C" }, line: { width: 0, color: "0D0D0C" } });
      s.addText(cols[i][3].toUpperCase(), { x: x + 0.45, y: 5.4, w: cw - 1.0, h: 0.42, fontFace: B.THEME.headFontFace, fontSize: 13, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true, charSpacing: 1 });
    }
    s.addNotes("Avance de Mi proceso. La tarjeta resaltada en amarillo claro es donde debería ir la mayoría hoy. Muévela según el calendario del taller.");
  }

  /* 8 · Errores frecuentes */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Errores frecuentes de este taller");
    B.kicker(s, "Lo que el docente vigila · del temario detallado", M, 1.35, true, 9);
    const errs = [["[Error 1: qué hace la gente y qué sale mal]", "[Qué hacer en cambio]"], ["[Error 2]", "[Qué hacer en cambio]"], ["[Error 3]", "[Qué hacer en cambio]"], ["[Error 4]", "[Qué hacer en cambio]"]];
    const cw = (CONTENT_W - 0.3) / 2;
    for (let i = 0; i < 4; i++) {
      const x = M + (i % 2) * (cw + 0.3), y = 1.9 + Math.floor(i / 2) * 2.25;
      B.card(s, x, y, cw, 2.0, { dark: true });
      s.addImage({ data: await B.iconPng("alert", "circle-yellow"), x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55 });
      s.addText(errs[i][0], { x: x + 1.05, y: y + 0.25, w: cw - 1.35, h: 0.75, fontFace: B.THEME.headFontFace, fontSize: 20, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
      s.addImage({ data: await B.iconPng("check", "circle"), x: x + 0.3, y: y + 1.1, w: 0.55, h: 0.55 });
      B.body(s, errs[i][1], x + 1.05, y + 1.1, cw - 1.35, 0.75, { dark: true, fontSize: 15 });
    }
    s.addNotes("Errores frecuentes. Entre tres y cuatro, del apartado «Lo que el docente debe vigilar» del temario. Cada error va con qué hacer en cambio.");
  }

  /* 9 · Cierre */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Qué traes al encuentro 2");
    await B.iconRow(s, "clipboard", "Tu formato «Mi proceso»", "Con la tarea, la medición y una primera idea de mejora.", M, 1.95, 5.8, { dark: true, size: 0.7, headSize: 22, bodySize: 16, bodyH: 0.9 });
    await B.iconRow(s, "upload", "El reto de la semana entregado", "Hasta el miércoles en la plataforma.", M, 3.25, 5.8, { dark: true, size: 0.7, headSize: 22, bodySize: 16, bodyH: 0.9 });
    await B.iconRow(s, "chat", "Las dudas que quedaron pendientes", "Las del tablero marcadas «Va al encuentro 2».", M, 4.55, 5.8, { dark: true, size: 0.7, headSize: 22, bodySize: 16, bodyH: 0.9 });
    B.card(s, 7.3, 1.95, 5.43, 3.9, { dark: true });
    s.addImage({ data: await B.iconPng("calendar", "circle-yellow"), x: 7.65, y: 2.3, w: 0.8, h: 0.8 });
    s.addText("[FECHA E2]", { x: 7.65, y: 3.25, w: 4.8, h: 0.9, fontFace: B.THEME.headFontFace, fontSize: 40, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("[HORA] · Teams", { x: 7.65, y: 4.15, w: 4.8, h: 0.5, fontFace: B.THEME.headFontFace, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
    B.body(s, "Encuentro 2 · aplicación a tu tarea", 7.65, 4.8, 4.8, 0.8, { dark: true, fontSize: 16 });
    s.addNotes("Cierre. Tres cosas que la persona trae al encuentro 2 y la fecha en grande.");
  }

  /* 10 · Guía de uso */
  B.guideSlide(pres, "SH_[PERFIL]_T[N]_C_[AAAAMMDD]", [
    ["[PERFIL]", "COM o ADM en el nombre del archivo; Comercial o Administrativo en el texto"],
    ["[N]", "Número del taller, 1 a 6"],
    ["[NOMBRE DEL TALLER]", "Del temario detallado, tal como está escrito"],
    ["[CONCEPTO DEL MES]", "El concepto de IA del mes"],
    ["[FECHA CONSULTORIO] [FECHA E2]", "Con día de la semana: «miércoles 14 de octubre»"],
    ["[DUDA EN UNA LÍNEA]", "Se cambia en vivo en las pantallas de resolución"],
    ["[DOCENTE]", "Nombre completo"],
    ["[Error 1…4]", "Del apartado «Lo que el docente debe vigilar» del temario"],
  ], [
    "Titulares en Barlow Condensed y en mayúsculas. Cuerpo en Atkinson Hyperlegible Next, en tipo oración.",
    "Negro #0D0D0C y amarillo #FCC800. Sobre amarillo el texto va en negro. El amarillo como texto solo va sobre negro. Sobre blanco, el acento de texto es #8A6B00.",
    "El consultorio se vive en pantalla compartida: la presentación acompaña, la herramienta protagoniza.",
    "Sin logos de las empresas: la pieza sirve igual en las cinco compañías.",
    "Pie de página: edítalo una sola vez en Vista > Patrón de diapositivas.",
    "Si falta una fuente, PowerPoint usa Arial Narrow y Arial. Instala las dos fuentes antes de presentar.",
    "Esta diapositiva y la biblioteca de íconos se borran antes de presentar.",
  ], SEC);

  /* 11 · Biblioteca de íconos */
  await B.iconLibrarySlide(pres, SEC);
});
