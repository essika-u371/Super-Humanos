const { build, cover, B } = require("./common.js");
const { M, CONTENT_W, W } = B;

const SEC = "Encuentro 1";

build("Plantilla taller · Encuentro 1 · concepto y práctica.pptx", "Super Humanos · [PERFIL] · Taller [N] · Encuentro 1", async (pres) => {
  pres.addSection({ title: SEC });

  /* 1 · Portada */
  cover(pres, SEC, {
    encuentro: "Encuentro 1 · concepto y práctica",
    kicker: "[PERFIL] · Taller [N] de 6",
    duracion: "90 minutos",
    notas: "Portada. Cambia los marcadores entre corchetes. El título admite hasta 40 caracteres en una línea; si pasa a dos, baja a 56 pt. El recorrido de abajo marca en amarillo el taller actual: cambia el color del chevrón en vez de borrarlo.",
  });

  /* 2 · Agenda */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Agenda de hoy");
    B.kicker(s, "90 minutos · seis momentos", M, 1.35, false);
    B.chevronBar(s, [
      { label: "0–5", sub: "Conexión\nBienvenida y pregunta rápida en el chat", done: true },
      { label: "5–15", sub: "Hábito de apertura\nUna historia real de G&J", done: true },
      { label: "15–40", sub: "Concepto de IA\n[CONCEPTO DEL MES]", done: true },
      { label: "40–60", sub: "Demostración en vivo\nCaso: [CASO DE G&J]", done: true },
      { label: "60–80", sub: "Práctica guiada\nCon el prompt de partida", done: true },
      { label: "80–90", sub: "«Mi proceso» y reto\nQué traes al encuentro 2", done: true },
    ], M, 1.95, CONTENT_W, { h: 0.55, fontSize: 15, subSize: 13 });
    await B.iconRow(s, "clipboard", "Qué te llevas hoy", "Un concepto, una demostración resuelta y tu primera práctica con el formato «Mi proceso».", M, 4.3, 5.6);
    await B.iconRow(s, "calendar", "Lo que sigue", "Consultorio el [FECHA CONSULTORIO] a las 4:00 p. m. y encuentro 2 el [FECHA E2].", 7.1, 4.3, 5.6);
    await B.iconRow(s, "star", "Puntos de hoy", "Asistencia completa: [PUNTOS ASISTENCIA] puntos. Reto de la semana: [PUNTOS RETO] puntos.", M, 5.5, 5.6);
    await B.iconRow(s, "live", "Grabación", "Queda en la plataforma al día siguiente.", 7.1, 5.5, 5.6);
    s.addNotes("Agenda. Los minutos son la distribución por defecto; ajústalos en los chevrones si el taller lo necesita. Cada texto bajo el chevrón admite dos líneas cortas.");
  }

  /* 3 · Hábito de apertura */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Hábito de apertura");
    B.kicker(s, "10 minutos · minuto 5 al 15", M, 1.35, true);
    s.addText("[HÁBITO]", { x: M, y: 1.8, w: 5.6, h: 1.4, fontFace: B.THEME.headFontFace, fontSize: 54, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    B.body(s, "[En una frase: qué hace una persona que tiene este hábito y por qué le sirve para que la IA le dé resultado.]", M, 3.3, 5.4, 1.6, { dark: true, fontSize: 18 });
    // Historia
    B.card(s, 6.6, 1.8, 6.1, 4.6, { dark: true });
    B.ribbon(s, "Una historia real de G&J", 6.95, 2.05, 4.6, { fontSize: 16, h: 0.46 });
    B.body(s, "[HISTORIA REAL DE G&J: qué pasó, en qué área, qué se perdió o se ganó. Máximo 70 palabras. Sale de las mesas de trabajo del 10 y 11 de septiembre.]", 6.95, 2.75, 5.4, 2.3, { dark: true, fontSize: 16 });
    s.addText("“[Cita textual de la persona que contó la historia]”", { x: 6.95, y: 5.2, w: 5.4, h: 0.95, fontFace: B.THEME.headFontFace, fontSize: 20, bold: true, italic: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true });
    s.addNotes("Hábito de apertura. El hábito en una o dos palabras a la izquierda. La historia a la derecha: 70 palabras como máximo. La cita es opcional; si no hay, borra el texto y deja el aire.");
  }

  /* 4 · Concepto de IA */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "[Concepto del mes]");
    B.kicker(s, "Concepto de IA · minuto 15 al 40", M, 1.35, false);
    B.body(s, "[Definición en una frase, con palabras de la empresa. Máximo 25 palabras.]", M, 1.85, CONTENT_W, 0.9, { fontSize: 22 });
    const ideas = [["bulb", "Idea 1", "[Qué es, en lenguaje de oficina. Máximo 30 palabras.]"], ["tool", "Idea 2", "[Dónde se aplica en el rol [PERFIL]. Máximo 30 palabras.]"], ["check", "Idea 3", "[Qué cambia cuando se usa bien. Máximo 30 palabras.]"]];
    const cw = (CONTENT_W - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = M + i * (cw + 0.3);
      B.card(s, x, 3.0, cw, 3.4);
      s.addImage({ data: await B.iconPng(ideas[i][0], "circle"), x: x + 0.3, y: 3.3, w: 0.7, h: 0.7 });
      s.addText(ideas[i][1].toUpperCase(), { x: x + 0.3, y: 4.15, w: cw - 0.6, h: 0.4, fontFace: B.THEME.headFontFace, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true });
      B.body(s, ideas[i][2], x + 0.3, 4.6, cw - 0.6, 1.6, { fontSize: 15 });
    }
    s.addNotes("Concepto de IA. Una definición y tres ideas. Si el concepto necesita más de tres ideas, duplica la diapositiva; no agregues una cuarta tarjeta.");
  }

  /* 5 · Elemento ORBITA IA */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "El elemento del método que trabajamos hoy");
    B.kicker(s, "Marco ORBITA IA · seis elementos, uno por taller", M, 1.35, false);
    B.card(s, M, 1.85, 5.4, 4.55, { fill: "F7F7F5", border: "DEDED9" });
    s.addImage({ data: await B.iconPng("route", "circle"), x: M + 0.35, y: 2.2, w: 0.85, h: 0.85 });
    s.addText("[ELEMENTO ORBITA IA]", { x: M + 0.35, y: 3.2, w: 4.7, h: 1.0, fontFace: B.THEME.headFontFace, fontSize: 34, bold: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    B.body(s, "[En qué consiste, en dos frases. Qué pregunta le hace a cada tarea del rol.]", M + 0.35, 4.3, 4.7, 1.9, { fontSize: 16 });
    const els = ["Uso responsable y riesgo", "Seis destinos del trabajo", "Jerarquía de transformación", "Contexto antes que modelo", "Matriz humano–IA", "Línea base antes de intervenir"];
    for (let i = 0; i < 6; i++) {
      const y = 1.85 + i * 0.76;
      const cur = i === 0;
      s.addShape(i === 0 ? "homePlate" : "chevron", { x: 6.55, y, w: 0.55, h: 0.6, fill: { color: cur ? "FCC800" : "DEDED9" }, line: { width: 0, color: cur ? "FCC800" : "DEDED9" } });
      s.addText("T" + (i + 1), { x: 6.62, y, w: 0.4, h: 0.6, fontFace: B.THEME.headFontFace, fontSize: 14, bold: true, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true });
      s.addText(els[i].toUpperCase(), { x: 7.3, y, w: 5.4, h: 0.6, fontFace: B.THEME.headFontFace, fontSize: cur ? 22 : 18, bold: true, color: cur ? "0D0D0C" : "56564F", margin: 0, valign: "middle", isTextBox: true });
    }
    s.addNotes("Elemento del método. Pinta de amarillo el chevrón del taller actual y pon su nombre en la tarjeta de la izquierda. La lista de la derecha no cambia entre talleres.");
  }

  /* 6 · Demostración en vivo */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Demostración en vivo");
    B.kicker(s, "Caso de G&J · [CASO DE G&J] · minuto 40 al 60", M, 1.35, false, 9);
    const steps = ["[Paso 1: qué recibe el docente y de dónde sale]", "[Paso 2: qué le pide a la herramienta]", "[Paso 3: qué revisa en la respuesta]", "[Paso 4: qué corrige o ajusta]", "[Paso 5: qué entrega y a quién]"];
    for (let i = 0; i < steps.length; i++) {
      const y = 1.9 + i * 0.9;
      s.addShape("ellipse", { x: M, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: "0D0D0C" }, line: { width: 0, color: "0D0D0C" } });
      s.addText(String(i + 1), { x: M, y: y + 0.05, w: 0.55, h: 0.55, fontFace: B.THEME.headFontFace, fontSize: 20, bold: true, color: "FCC800", align: "center", margin: 0, valign: "middle", isTextBox: true });
      B.body(s, steps[i], M + 0.75, y, 5.0, 0.8, { fontSize: 15, valign: "middle" });
    }
    B.promptBlock(s, "[PROMPT DE LA DEMOSTRACIÓN. Escríbelo completo, tal como se pega en la herramienta: rol, contexto, tarea y formato de salida.]", 6.8, 1.9, 5.9, 4.5);
    s.addNotes("Demostración. Cinco pasos como máximo; si el caso necesita más, usa la variante de pantalla dividida que sigue. El prompt va completo: la gente lo copia de la grabación.");
  }

  /* 6b · Demostración · pantalla dividida */
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: SEC });
    B.kicker(s, "Demostración · [CASO DE G&J]", M, 0.5, false, 9);
    B.promptBlock(s, "[PROMPT]", M, 0.95, 5.4, 5.7);
    s.addShape("roundRect", { x: 6.3, y: 0.95, w: 6.4, h: 5.7, rectRadius: 0.11, fill: { color: "F7F7F5" }, line: { color: "9A9A93", width: 1.25, dashType: "dash" } });
    s.addImage({ data: await B.iconPng("camera", "circle"), x: 9.15, y: 3.0, w: 0.7, h: 0.7 });
    s.addText("Captura del resultado\no pantalla compartida", { x: 6.6, y: 3.85, w: 5.8, h: 0.8, fontFace: B.THEME.bodyFontFace, fontSize: 15, color: "56564F", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Variante de la demostración. A la izquierda el prompt; a la derecha, la captura del resultado o el espacio para compartir pantalla. Borra el recuadro punteado si vas a compartir pantalla en vivo.");
  }

  /* 7 · Práctica guiada */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Práctica guiada");
    B.kicker(s, "Minuto 60 al 80 · individual o en parejas", M, 1.35, false);
    // Instrucciones
    B.card(s, M, 1.85, 4.1, 4.55);
    s.addImage({ data: await B.iconPng("tool", "circle"), x: M + 0.3, y: 2.15, w: 0.6, h: 0.6 });
    s.addText("INSTRUCCIONES", { x: M + 1.05, y: 2.2, w: 2.9, h: 0.5, fontFace: B.THEME.headFontFace, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true });
    B.bullets(s, ["[Qué tarea resuelve cada persona]", "[Con qué herramienta]", "[Qué entrega al final y dónde]"], M + 0.3, 2.95, 3.5, 3.2, { fontSize: 15 });
    // Prompt
    B.promptBlock(s, "[PROMPT DE PARTIDA. La persona lo copia y lo ajusta a su caso.]", 4.95, 1.85, 4.9, 4.55, { label: "Prompt de partida" });
    // Tiempo
    B.card(s, 10.1, 1.85, 2.63, 4.55, { fill: "0D0D0C", border: "0D0D0C" });
    s.addImage({ data: await B.iconPng("clock", "circle-yellow"), x: 10.9, y: 2.2, w: 1.0, h: 1.0 });
    s.addText("[20]", { x: 10.1, y: 3.3, w: 2.63, h: 1.1, fontFace: B.THEME.headFontFace, fontSize: 72, bold: true, color: "FCC800", align: "center", margin: 0, valign: "middle", isTextBox: true });
    s.addText("MINUTOS", { x: 10.1, y: 4.4, w: 2.63, h: 0.4, fontFace: B.THEME.headFontFace, fontSize: 18, bold: true, color: "FFFFFF", align: "center", margin: 0, isTextBox: true, charSpacing: 3 });
    s.addText("Al terminar, dos personas muestran su resultado", { x: 10.35, y: 4.95, w: 2.15, h: 1.2, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "C4C4BE", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Práctica guiada. Tres instrucciones como máximo. El tiempo de la derecha es el que ve la gente; ponle el cronómetro de Teams encima si lo usas.");
  }

  /* 8 · Mi proceso */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Formato «Mi proceso»");
    B.kicker(s, "Lo llenas entre hoy y el encuentro 2 · lo traes el [FECHA E2]", M, 1.35, false, 10);
    const cols = [["clipboard", "1 · La tarea", "Una tarea tuya que se repite. Con nombre, área y cada cuánto la haces."], ["clock", "2 · Cuánto toma", "Mídela esta semana: minutos por vez y veces por semana."], ["bulb", "3 · La mejora", "Qué parte podría apoyar la IA y qué resultado esperas."]];
    const cw = (CONTENT_W - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = M + i * (cw + 0.3);
      B.card(s, x, 1.95, cw, 3.3);
      s.addImage({ data: await B.iconPng(cols[i][0], "circle"), x: x + 0.3, y: 2.25, w: 0.7, h: 0.7 });
      s.addText(cols[i][1].toUpperCase(), { x: x + 0.3, y: 3.1, w: cw - 0.6, h: 0.45, fontFace: B.THEME.headFontFace, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true });
      B.body(s, cols[i][2], x + 0.3, 3.6, cw - 0.6, 1.5, { fontSize: 15 });
    }
    B.ribbon(s, "Dónde está el formato: [ENLACE O SECCIÓN DE LA PLATAFORMA]", M + 0.3, 5.6, 8.5, { fontSize: 16, h: 0.5 });
    s.addNotes("Mi proceso. Los tres campos del formato en el orden en que se llenan. En la cinta de abajo va el lugar exacto donde la persona lo encuentra.");
  }

  /* 9 · Reto de la semana */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Reto de la semana");
    B.ribbon(s, "Jueves de Superpoderes · Reto [NÚMERO DEL RETO]", M + 0.3, 1.4, 6.6, { yellow: true, fontSize: 18, h: 0.5 });
    s.addText("[TEXTO DEL RETO EN UNA FRASE]", { x: M, y: 2.2, w: 7.4, h: 1.6, fontFace: B.THEME.headFontFace, fontSize: 40, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    B.body(s, "[Dos líneas con el detalle: qué herramienta, qué entregas y qué evidencia subes.]", M, 3.9, 7.0, 1.2, { dark: true, fontSize: 17 });
    await B.iconRow(s, "upload", "Sube tu evidencia", "Una captura o un archivo en la plataforma, hasta el miércoles.", 8.4, 2.2, 4.3, { dark: true, bodySize: 14, bodyH: 0.9 });
    await B.iconRow(s, "star", "[PUNTOS RETO] puntos", "Suman al entregar. Sin calificación individual.", 8.4, 3.6, 4.3, { dark: true, bodySize: 14, bodyH: 0.9 });
    await B.iconRow(s, "clock", "Cinco minutos", "Es una tarea corta. Si te toma más, pregúntalo en el consultorio.", 8.4, 5.0, 4.3, { dark: true, bodySize: 14, bodyH: 0.9 });
    s.addNotes("Reto de la semana. El texto del reto sale del calendario de los 28 retos. Una frase en grande, dos líneas de detalle. Los puntos vienen del manual de puntajes.");
  }

  /* 10 · Consultorio */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Consultorio");
    B.kicker(s, "Una hora abierta de dudas, entre los dos encuentros", M, 1.35, false, 9);
    B.card(s, M, 1.9, 5.2, 4.5, { fill: "0D0D0C", border: "0D0D0C" });
    s.addImage({ data: await B.iconPng("calendar", "circle-yellow"), x: M + 0.35, y: 2.25, w: 0.8, h: 0.8 });
    s.addText("[FECHA CONSULTORIO]", { x: M + 0.35, y: 3.2, w: 4.5, h: 0.9, fontFace: B.THEME.headFontFace, fontSize: 40, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("4:00 a 5:00 p. m. · Teams", { x: M + 0.35, y: 4.1, w: 4.5, h: 0.5, fontFace: B.THEME.headFontFace, fontSize: 24, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
    B.body(s, "El enlace llega por correo y por WhatsApp desde Gestión Humana.", M + 0.35, 4.8, 4.5, 1.3, { dark: true, fontSize: 15 });
    await B.iconRow(s, "chat", "Trae una duda real", "Sobre la tarea que elegiste para «Mi proceso» o sobre el reto de la semana.", 6.5, 2.0, 6.2, { bodyH: 0.9 });
    await B.iconRow(s, "users", "Es abierto", "Entras y sales cuando quieras. Sin asistencia obligatoria y sin puntos.", 6.5, 3.4, 6.2, { bodyH: 0.9 });
    await B.iconRow(s, "live", "Queda grabado", "Si no puedes ir, ves la grabación al día siguiente en la plataforma.", 6.5, 4.8, 6.2, { bodyH: 0.9 });
    s.addNotes("Consultorio. La fecha en grande. Los tres textos de la derecha no cambian entre talleres salvo que cambie la regla del consultorio.");
  }

  /* 11 · Cierre */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Lo que suma hoy y lo que sigue");
    B.bigNumber(s, "[+00]", "puntos por asistir al encuentro completo", M, 1.8, 3.6, { dark: true });
    B.bigNumber(s, "[+00]", "puntos por el reto de la semana", 4.5, 1.8, 3.6, { dark: true });
    B.bigNumber(s, "[+00]", "puntos por entregar «Mi proceso»", 8.4, 1.8, 3.9, { dark: true });
    B.chevronBar(s, [
      { label: "Hoy · E1", sub: "Concepto y práctica", done: true },
      { label: "[FECHA CONSULTORIO]", sub: "Consultorio · 4:00 p. m.", done: false },
      { label: "[FECHA E2]", sub: "Encuentro 2 · aplicación", done: false },
      { label: "Evaluación", sub: "10 preguntas · una semana", done: false },
    ], M, 4.2, CONTENT_W, { dark: true, h: 0.55, fontSize: 15, subSize: 13 });
    s.addText("Nos vemos el [FECHA E2]", { x: M, y: 5.7, w: 8, h: 0.6, fontFace: B.THEME.headFontFace, fontSize: 32, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true });
    s.addNotes("Cierre. Tres cifras de puntos tomadas del manual de puntajes. La barra muestra el recorrido del mes con el primer chevrón en amarillo.");
  }

  /* 12 · Guía de uso */
  B.guideSlide(pres, "SH_[PERFIL]_T[N]_E1_[AAAAMMDD]", [
    ["[PERFIL]", "COM o ADM en el nombre del archivo; Comercial o Administrativo en el texto"],
    ["[N]", "Número del taller, 1 a 6"],
    ["[NOMBRE DEL TALLER]", "Del temario detallado, tal como está escrito"],
    ["[CONCEPTO DEL MES]", "El concepto de IA del mes, igual para los dos perfiles"],
    ["[HÁBITO]", "El hábito de apertura del taller"],
    ["[CASO DE G&J]", "Caso de la guía de casos, con su número"],
    ["[FECHA] [FECHA E2] [FECHA CONSULTORIO]", "Con día de la semana: «viernes 9 de octubre»"],
    ["[DOCENTE]", "Nombre completo"],
    ["[PUNTOS …]", "Del manual de puntajes y reconocimiento"],
  ], [
    "Titulares en Barlow Condensed y en mayúsculas. Cuerpo en Atkinson Hyperlegible Next, en tipo oración.",
    "Negro #0D0D0C y amarillo #FCC800. Sobre amarillo el texto va en negro. El amarillo como texto solo va sobre negro. Sobre blanco, el acento de texto es #8A6B00.",
    "Una idea por diapositiva y máximo 60 palabras de cuerpo.",
    "Sin logos de las empresas: la pieza sirve igual en las cinco compañías.",
    "Pie de página: edítalo una sola vez en Vista > Patrón de diapositivas.",
    "Si falta una fuente, PowerPoint usa Arial Narrow y Arial. Instala las dos fuentes antes de presentar.",
    "Esta diapositiva y la biblioteca de íconos se borran antes de presentar.",
  ], SEC);

  /* 13 · Biblioteca de íconos */
  await B.iconLibrarySlide(pres, SEC);
});
