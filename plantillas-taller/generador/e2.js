const { build, cover, B } = require("./common.js");
const { M, CONTENT_W, W } = B;

const SEC = "Encuentro 2";

build("Plantilla taller · Encuentro 2 · aplicación.pptx", "Super Humanos · [PERFIL] · Taller [N] · Encuentro 2", async (pres) => {
  pres.addSection({ title: SEC });

  /* 1 · Portada */
  cover(pres, SEC, {
    encuentro: "Encuentro 2 · aplicación",
    kicker: "[PERFIL] · Taller [N] de 6",
    duracion: "90 minutos",
    notas: "Portada del encuentro 2. Cambia los marcadores. El recorrido de abajo marca en amarillo el taller actual.",
  });

  /* 2 · Agenda */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Agenda de hoy");
    B.kicker(s, "90 minutos · cinco momentos", M, 1.35, false);
    B.chevronBar(s, [
      { label: "0–10", sub: "Repaso del reto\nQué salió, dos ejemplos", done: true },
      { label: "10–40", sub: "Aplicación\nCada persona con su «Mi proceso»", done: true },
      { label: "40–70", sub: "Casos compartidos\nTres o cuatro, con retroalimentación", done: true },
      { label: "70–80", sub: "Evaluación corta\nDiez preguntas en la plataforma", done: true },
      { label: "80–90", sub: "Cierre\nPuntos y siguiente taller", done: true },
    ], M, 1.95, CONTENT_W, { h: 0.55, fontSize: 15, subSize: 13 });
    await B.iconRow(s, "clipboard", "Qué necesitas tener a mano", "Tu formato «Mi proceso» con la tarea y la medición, y la herramienta abierta.", M, 4.3, 5.6);
    await B.iconRow(s, "trophy", "Con qué sales hoy", "Una tarea tuya medida y una mejora propuesta. Eso es lo que cuenta el programa.", 7.1, 4.3, 5.6);
    await B.iconRow(s, "star", "Puntos de hoy", "Asistencia: [PUNTOS ASISTENCIA]. «Mi proceso» entregado: [PUNTOS MI PROCESO]. Evaluación aprobada: [PUNTOS EVALUACIÓN].", M, 5.5, 5.6);
    await B.iconRow(s, "live", "Grabación", "Queda en la plataforma al día siguiente.", 7.1, 5.5, 5.6);
    s.addNotes("Agenda del encuentro 2. Los minutos son la distribución por defecto. Los puntos salen del manual de puntajes.");
  }

  /* 3 · Repaso del reto */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Qué salió del reto");
    B.ribbon(s, "Reto [NÚMERO DEL RETO] · [TEXTO DEL RETO EN UNA FRASE]", M + 0.3, 1.4, 8.5, { yellow: true, fontSize: 18, h: 0.5 });
    B.bigNumber(s, "[00]", "personas lo entregaron", M, 2.2, 3.2, { dark: true });
    B.bigNumber(s, "[00 %]", "del grupo", 3.6, 2.2, 2.6, { dark: true });
    const cw = (W - M - 6.3 - 0.3) / 2;
    for (let i = 0; i < 2; i++) {
      const x = 6.3 + i * (cw + 0.3);
      B.card(s, x, 2.2, cw, 3.9, { dark: true });
      s.addImage({ data: await B.iconPng("star", "circle-yellow"), x: x + 0.3, y: 2.5, w: 0.55, h: 0.55 });
      s.addText("[NOMBRE · ÁREA]", { x: x + 1.0, y: 2.5, w: cw - 1.3, h: 0.55, fontFace: B.THEME.headFontFace, fontSize: 16, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true });
      B.body(s, "[Qué hizo y qué le devolvió la herramienta. Dos o tres líneas. Si hay captura, va aquí en lugar del texto.]", x + 0.3, 3.25, cw - 0.6, 2.6, { dark: true, fontSize: 15 });
    }
    B.body(s, "[Una línea sobre lo que más se repitió en las evidencias.]", M, 4.6, 5.4, 1.4, { dark: true, fontSize: 17 });
    s.addNotes("Repaso del reto. Dos cifras de participación y dos ejemplos de personas reales, con su permiso. Si hay capturas, reemplazan el texto de la tarjeta.");
  }

  /* 4 · Aplicación */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Aplicación a tu tarea");
    B.kicker(s, "Minuto 10 al 40 · trabajo individual con «Mi proceso»", M, 1.35, false, 9);
    B.card(s, M, 1.85, 4.1, 4.55);
    s.addImage({ data: await B.iconPng("tool", "circle"), x: M + 0.3, y: 2.15, w: 0.6, h: 0.6 });
    s.addText("QUÉ HACES", { x: M + 1.05, y: 2.2, w: 2.9, h: 0.5, fontFace: B.THEME.headFontFace, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true });
    B.bullets(s, ["Abre tu formato «Mi proceso» en el campo 3, la mejora.", "Prueba la mejora con la herramienta usando el prompt de apoyo.", "Anota el nuevo tiempo o el nuevo resultado.", "Deja listo tu caso para compartirlo en tres líneas."], M + 0.3, 2.95, 3.5, 3.3, { fontSize: 14, paraSpaceAfter: 6 });
    B.promptBlock(s, "[PROMPT DE APOYO. Parte del prompt de la demostración, adaptado para que cada persona pegue su tarea.]", 4.95, 1.85, 4.9, 4.55, { label: "Prompt de apoyo" });
    B.card(s, 10.1, 1.85, 2.63, 4.55, { fill: "0D0D0C", border: "0D0D0C" });
    s.addImage({ data: await B.iconPng("clock", "circle-yellow"), x: 10.9, y: 2.2, w: 1.0, h: 1.0 });
    s.addText("[30]", { x: 10.1, y: 3.3, w: 2.63, h: 1.1, fontFace: B.THEME.headFontFace, fontSize: 72, bold: true, color: "FCC800", align: "center", margin: 0, valign: "middle", isTextBox: true });
    s.addText("MINUTOS", { x: 10.1, y: 4.4, w: 2.63, h: 0.4, fontFace: B.THEME.headFontFace, fontSize: 18, bold: true, color: "FFFFFF", align: "center", margin: 0, isTextBox: true, charSpacing: 3 });
    s.addText("El docente pasa por el chat y por los micrófonos abiertos", { x: 10.35, y: 4.95, w: 2.15, h: 1.2, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "C4C4BE", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Aplicación. Cuatro pasos y el prompt de apoyo. El tiempo es el que ve la gente.");
  }

  /* 5 · Tarjeta de caso compartido */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Caso compartido [1 de 4]");
    B.kicker(s, "Minuto 40 al 70 · cada caso en cinco minutos, con retroalimentación", M, 1.35, false, 10);
    B.card(s, M, 1.85, 4.3, 4.55, { fill: "F7F7F5", border: "DEDED9" });
    s.addImage({ data: await B.iconPng("user", "circle"), x: M + 0.3, y: 2.15, w: 0.7, h: 0.7 });
    s.addText("[NOMBRE]", { x: M + 1.15, y: 2.15, w: 3.0, h: 0.4, fontFace: B.THEME.headFontFace, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true });
    s.addText("[Área · empresa]", { x: M + 1.15, y: 2.5, w: 3.0, h: 0.35, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "56564F", margin: 0, valign: "middle", isTextBox: true });
    B.kicker(s, "La tarea", M + 0.3, 3.1, false, 3.5);
    B.body(s, "[Qué tarea, cada cuánto, para quién.]", M + 0.3, 3.4, 3.7, 1.0, { fontSize: 15 });
    B.kicker(s, "La mejora propuesta", M + 0.3, 4.45, false, 3.5);
    B.body(s, "[Qué parte apoya la IA y cómo.]", M + 0.3, 4.75, 3.7, 1.4, { fontSize: 15 });
    // Antes y después
    const bx = 5.2, bw = 3.5;
    B.card(s, bx, 1.85, bw, 2.15, { fill: "FFFFFF", border: "DEDED9" });
    B.kicker(s, "Antes", bx + 0.3, 2.1, false, 2.5);
    s.addText("[00 min]", { x: bx + 0.3, y: 2.45, w: bw - 0.6, h: 0.9, fontFace: B.THEME.headFontFace, fontSize: 44, bold: true, color: "56564F", margin: 0, valign: "middle", isTextBox: true });
    s.addText("por vez · [veces] por semana", { x: bx + 0.3, y: 3.35, w: bw - 0.6, h: 0.4, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "56564F", margin: 0, isTextBox: true });
    B.card(s, bx, 4.25, bw, 2.15, { fill: "0D0D0C", border: "0D0D0C" });
    B.kicker(s, "Después", bx + 0.3, 4.5, true, 2.5);
    s.addText("[00 min]", { x: bx + 0.3, y: 4.85, w: bw - 0.6, h: 0.9, fontFace: B.THEME.headFontFace, fontSize: 44, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true });
    s.addText("por vez · [veces] por semana", { x: bx + 0.3, y: 5.75, w: bw - 0.6, h: 0.4, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "C4C4BE", margin: 0, isTextBox: true });
    // Retroalimentación
    const rx = 9.0, rw = W - M - rx;
    B.ribbon(s, "Retroalimentación", rx + 0.3, 1.85, rw - 0.3, { fontSize: 18, h: 0.5 });
    await B.iconRow(s, "check", "Lo que está bien", "[Una línea]", rx, 2.6, rw, { size: 0.5, headSize: 16, bodySize: 14, bodyH: 0.7 });
    await B.iconRow(s, "alert", "Lo que hay que revisar", "[Una línea: qué validar antes de usar la salida]", rx, 3.75, rw, { size: 0.5, headSize: 16, bodySize: 14, bodyH: 0.7 });
    await B.iconRow(s, "chevron-right", "Siguiente paso", "[Una línea]", rx, 4.9, rw, { size: 0.5, headSize: 16, bodySize: 14, bodyH: 0.7 });
    s.addNotes("Tarjeta de caso. Duplica esta diapositiva por cada caso: tres o cuatro por encuentro. Las cifras de antes y después salen del formato «Mi proceso» de la persona.");
  }

  /* 6 · Comparación antes y después del grupo */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Antes y después");
    B.kicker(s, "Lo que midió el grupo · minutos por tarea", M, 1.35, false);
    s.addChart("bar", [
      { name: "Antes", labels: ["[Caso 1]", "[Caso 2]", "[Caso 3]", "[Caso 4]"], values: [45, 30, 60, 20] },
      { name: "Después", labels: ["[Caso 1]", "[Caso 2]", "[Caso 3]", "[Caso 4]"], values: [15, 12, 25, 8] },
    ], {
      x: M, y: 1.85, w: 8.0, h: 4.55, barDir: "bar", barGrouping: "clustered",
      chartColors: ["9A9A93", "FCC800"], showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: "0D0D0C",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: "0D0D0C",
      catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 12, catAxisLabelColor: "0D0D0C", catGridLine: { style: "none" },
      valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: 11, valAxisLabelColor: "56564F", valGridLine: { color: "DEDED9", size: 0.75 }, valAxisLineShow: false,
      showTitle: false,
    });
    B.card(s, 8.9, 1.85, 3.83, 4.55, { fill: "0D0D0C", border: "0D0D0C" });
    B.bigNumber(s, "[00 %]", "menos tiempo, en promedio, en los casos compartidos", 9.2, 2.3, 3.2, { dark: true, size: 54 });
    B.body(s, "Cambia los datos del gráfico con clic derecho > Editar datos. Los nombres de los casos van en la primera columna.", 9.2, 4.6, 3.2, 1.5, { dark: true, fontSize: 12 });
    s.addNotes("Antes y después. Gráfico nativo de PowerPoint: se edita con clic derecho > Editar datos. Serie «Antes» en gris, «Después» en amarillo.");
  }

  /* 7 · Evaluación corta */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Evaluación corta");
    B.kicker(s, "Minuto 70 al 80 · en la plataforma", M, 1.35, false);
    B.bigNumber(s, "10", "preguntas de opción múltiple", M, 1.9, 3.2);
    B.bigNumber(s, "70 %", "para aprobar: siete de diez", 4.1, 1.9, 3.2);
    B.bigNumber(s, "[PUNTOS]", "puntos al aprobar", 7.6, 1.9, 3.6);
    await B.iconRow(s, "route", "Dónde", "[Sección de la plataforma donde está la evaluación del taller [N].]", M, 4.1, 5.8, { bodyH: 0.8 });
    await B.iconRow(s, "clock", "Hasta cuándo", "Abre hoy y cierra el [FECHA LÍMITE EVALUACIÓN]. Si no apruebas, repasas la grabación y lo intentas de nuevo.", 7.1, 4.1, 5.6, { bodyH: 0.9 });
    await B.iconRow(s, "info", "Qué evalúa", "Lo que se vio en los dos encuentros: el concepto, cuándo aplicarlo y qué verificar antes de usar la salida.", M, 5.3, 5.8, { bodyH: 0.9 });
    s.addNotes("Evaluación corta. Las reglas vienen del plan: diez preguntas, 70 % para aprobar, reintento permitido. La fecha límite es una semana después del encuentro 2.");
  }

  /* 8 · Encuesta de satisfacción */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "Una sola pregunta");
    s.addText("¿Qué tanto le recomendarías este taller a un compañero de tu área?", { x: M, y: 1.7, w: CONTENT_W, h: 1.2, fontFace: B.THEME.headFontFace, fontSize: 36, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    // Escala 0–10 (Scale): casillas con el 10 en amarillo
    const n = 11, gap = 0.1, bw = (CONTENT_W - gap * (n - 1)) / n;
    for (let i = 0; i < n; i++) {
      const x = M + i * (bw + gap);
      const hot = i >= 9;
      s.addShape("rect", { x, y: 3.3, w: bw, h: 0.9, fill: { color: hot ? "FCC800" : "262623" }, line: { color: hot ? "FCC800" : "3D3D38", width: 1 } });
      s.addText(String(i), { x, y: 3.3, w: bw, h: 0.9, fontFace: B.THEME.headFontFace, fontSize: 28, bold: true, color: hot ? "0D0D0C" : "FFFFFF", align: "center", margin: 0, valign: "middle", isTextBox: true });
    }
    s.addText("Nada", { x: M, y: 4.3, w: 2, h: 0.35, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "C4C4BE", margin: 0, isTextBox: true });
    s.addText("Totalmente", { x: W - M - 2, y: 4.3, w: 2, h: 0.35, fontFace: B.THEME.bodyFontFace, fontSize: 13, color: "C4C4BE", margin: 0, isTextBox: true, align: "right" });
    await B.iconRow(s, "chat", "Responde en el chat o en la encuesta de Teams", "Un número del 0 al 10. Toma diez segundos. La meta del programa es 8 o más.", M, 5.0, 9.0, { dark: true, bodyH: 0.8 });
    s.addNotes("Encuesta de una pregunta. Se responde al cierre del taller. La escala es solo visual: la respuesta va por el chat o por la encuesta de Teams.");
  }

  /* 9 · Puntos del mes */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: SEC });
    B.title(s, "Lo que sumó este taller");
    B.kicker(s, "Del manual de puntajes y reconocimiento", M, 1.35, false);
    const rows = [["live", "Encuentro 1", "[PUNTOS]"], ["live", "Encuentro 2", "[PUNTOS]"], ["upload", "Retos de la semana", "[PUNTOS] × 4"], ["clipboard", "«Mi proceso» entregado", "[PUNTOS]"], ["check", "Evaluación aprobada", "[PUNTOS]"]];
    for (let i = 0; i < rows.length; i++) {
      const y = 1.9 + i * 0.85;
      s.addImage({ data: await B.iconPng(rows[i][0], "circle"), x: M, y: y + 0.08, w: 0.55, h: 0.55 });
      s.addText(rows[i][1].toUpperCase(), { x: M + 0.8, y, w: 4.5, h: 0.7, fontFace: B.THEME.headFontFace, fontSize: 20, bold: true, color: "0D0D0C", margin: 0, valign: "middle", isTextBox: true });
      s.addText(rows[i][2], { x: 5.4, y, w: 1.8, h: 0.7, fontFace: B.THEME.headFontFace, fontSize: 24, bold: true, color: "8A6B00", margin: 0, valign: "middle", isTextBox: true, align: "right" });
      if (i < rows.length - 1) s.addShape("line", { x: M, y: y + 0.78, w: 6.6, h: 0, line: { color: "DEDED9", width: 0.75 } });
    }
    B.card(s, 8.0, 1.9, 4.73, 4.2, { fill: "0D0D0C", border: "0D0D0C" });
    s.addImage({ data: await B.iconPng("trophy", "circle-yellow"), x: 8.35, y: 2.25, w: 0.8, h: 0.8 });
    B.bigNumber(s, "[000]", "puntos posibles en el taller [N]", 8.35, 3.1, 4.0, { dark: true, size: 60 });
    B.body(s, "Tu acumulado y tu puesto están en la plataforma. Los premios los define Gestión Humana.", 8.35, 4.85, 4.0, 1.1, { dark: true, fontSize: 14 });
    s.addNotes("Puntos del mes. Las cifras salen del manual de puntajes. Si el manual cambia, cambia aquí y en la agenda.");
  }

  /* 10 · Siguiente taller */
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: SEC });
    B.title(s, "El próximo taller");
    B.kicker(s, "Taller [N+1] de 6 · [MES]", M, 1.35, true);
    s.addText("[NOMBRE DEL SIGUIENTE TALLER]", { x: M, y: 1.8, w: 7.2, h: 1.3, fontFace: B.THEME.headFontFace, fontSize: 40, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("[CONCEPTO DEL SIGUIENTE MES]", { x: M, y: 3.2, w: 7.2, h: 0.6, fontFace: B.THEME.headFontFace, fontSize: 26, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
    B.body(s, "[Dos líneas: qué va a poder hacer la persona al terminar ese taller.]", M, 3.9, 7.0, 1.2, { dark: true, fontSize: 17 });
    B.chevronBar(s, [
      { label: "[FECHA E1]", sub: "Encuentro 1", done: true },
      { label: "[FECHA C]", sub: "Consultorio", done: false },
      { label: "[FECHA E2]", sub: "Encuentro 2", done: false },
    ], 8.2, 1.9, 4.53, { dark: true, h: 0.55, fontSize: 13, subSize: 12 });
    await B.iconRow(s, "calendar", "La invitación llega un mes antes", "Por Outlook desde Gestión Humana. Si no la ves, escríbele a tu líder.", 8.2, 3.4, 4.53, { dark: true, bodySize: 13, bodyH: 0.9 });
    await B.iconRow(s, "upload", "Mientras tanto, el reto del jueves", "Cada jueves a las 7:00 a. m. Cuatro retos refuerzan lo que viste hoy.", 8.2, 4.7, 4.53, { dark: true, bodySize: 13, bodyH: 0.9 });
    s.addNotes("Siguiente taller. Nombre, concepto y las tres fechas del próximo mes, del cronograma.");
  }

  /* 11 · Cierre */
  {
    const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: SEC });
    B.wordmark(s, M, 0.55, { size: 26 });
    s.addText("Gracias", { x: M, y: 2.0, w: 11, h: 1.5, fontFace: B.THEME.headFontFace, fontSize: 80, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true });
    s.addText("Taller [N] completo. Sigue el reto del jueves.", { x: M, y: 3.5, w: 11, h: 0.7, fontFace: B.THEME.headFontFace, fontSize: 30, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
    await B.iconRow(s, "chat", "superhumanos@gyj.com.co", "Dudas, sugerencias y evidencias que no cargaron. Respuesta cada semana, a más tardar el viernes.", M, 4.7, 7.5, { dark: true, bodyH: 0.9 });
    const items = [1, 2, 3, 4, 5, 6].map((n) => ({ label: "T" + n, done: n === 1 }));
    B.chevronBar(s, items, 8.3, 6.2, 4.45, { dark: true, h: 0.42, fontSize: 12 });
    s.addNotes("Cierre. El correo del programa y el recorrido. Pinta de amarillo los talleres completados.");
  }

  /* 12 · Guía de uso */
  B.guideSlide(pres, "SH_[PERFIL]_T[N]_E2_[AAAAMMDD]", [
    ["[PERFIL]", "COM o ADM en el nombre del archivo; Comercial o Administrativo en el texto"],
    ["[N]", "Número del taller, 1 a 6"],
    ["[NOMBRE DEL TALLER]", "Del temario detallado, tal como está escrito"],
    ["[CONCEPTO DEL MES]", "El concepto de IA del mes"],
    ["[NÚMERO DEL RETO] [TEXTO DEL RETO]", "Del calendario de los 28 retos"],
    ["[NOMBRE] [ÁREA]", "Casos compartidos: con permiso de la persona"],
    ["[FECHA …]", "Con día de la semana: «viernes 16 de octubre»"],
    ["[PUNTOS …]", "Del manual de puntajes y reconocimiento"],
    ["[N+1] [MES]", "Siguiente taller y su mes, del cronograma"],
  ], [
    "Titulares en Barlow Condensed y en mayúsculas. Cuerpo en Atkinson Hyperlegible Next, en tipo oración.",
    "Negro #0D0D0C y amarillo #FCC800. Sobre amarillo el texto va en negro. El amarillo como texto solo va sobre negro. Sobre blanco, el acento de texto es #8A6B00.",
    "La tarjeta de caso se duplica por cada caso: tres o cuatro por encuentro.",
    "El gráfico de antes y después es nativo: clic derecho > Editar datos.",
    "Sin logos de las empresas: la pieza sirve igual en las cinco compañías.",
    "Pie de página: edítalo una sola vez en Vista > Patrón de diapositivas.",
    "Si falta una fuente, PowerPoint usa Arial Narrow y Arial. Instala las dos fuentes antes de presentar.",
    "Esta diapositiva y la biblioteca de íconos se borran antes de presentar.",
  ], SEC);

  /* 13 · Biblioteca de íconos */
  await B.iconLibrarySlide(pres, SEC);
});
