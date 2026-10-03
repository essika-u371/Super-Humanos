const { build, B } = require("./common.js");
const { M, CONTENT_W, W, H } = B;
const HF = B.THEME.headFontFace, BF = B.THEME.bodyFontFace;

const T = "Título de la diapositiva";
const SUB = "Antetítulo o contexto en una línea";
const P = "Texto de cuerpo. Dos o tres líneas, en tipo oración, con una sola idea.";
const P2 = "Segundo párrafo si hace falta. Máximo 60 palabras por diapositiva.";

build("Plantilla base Super Humanos · estilos de diapositiva.pptx", "Super Humanos · [Perfil] · [Taller]", async (pres) => {
  /* ===== PORTADAS ===== */
  pres.addSection({ title: "Portadas y divisores" });
  {
    const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portadas y divisores" });
    B.wordmark(s, M, 0.55, { size: 26 });
    B.kicker(s, "Antetítulo", M, 2.3, true, 10);
    s.addText("TÍTULO GRANDE", { x: M, y: 2.65, w: 11.5, h: 1.5, fontFace: HF, fontSize: 66, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("SUBTÍTULO EN UNA LÍNEA", { x: M, y: 4.15, w: 11.5, h: 0.7, fontFace: HF, fontSize: 32, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
    s.addText("Fecha · hora · lugar", { x: M, y: 6.3, w: 7, h: 0.4, fontFace: BF, fontSize: 16, color: "C4C4BE", margin: 0, valign: "middle", isTextBox: true });
    s.addNotes("Portada oscura. Título en amarillo, hasta 40 caracteres en una línea. Subtítulo en blanco.");
  }
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: "Portadas y divisores" });
    B.wordmark(s, M, 0.55, { size: 26, dark: false });
    B.kicker(s, "Antetítulo", M, 2.3, false, 10);
    s.addText("TÍTULO GRANDE", { x: M, y: 2.65, w: 11.5, h: 1.5, fontFace: HF, fontSize: 66, bold: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("SUBTÍTULO EN UNA LÍNEA", { x: M, y: 4.15, w: 11.5, h: 0.7, fontFace: HF, fontSize: 32, bold: true, color: "8A6B00", margin: 0, valign: "top", isTextBox: true });
    s.addText("Fecha · hora · lugar", { x: M, y: 6.3, w: 7, h: 0.4, fontFace: BF, fontSize: 16, color: "56564F", margin: 0, valign: "middle", isTextBox: true });
    s.addNotes("Portada clara. Misma estructura, fondo blanco. El acento de texto sobre blanco es #8A6B00.");
  }
  {
    const s = pres.addSlide({ masterName: "SECCION", sectionTitle: "Portadas y divisores" });
    s.addShape("chevron", { x: M, y: 2.3, w: 1.6, h: 2.2, fill: { color: "FCC800" }, line: { width: 0, color: "FCC800" } });
    s.addText("01", { x: 2.5, y: 2.2, w: 3, h: 1.0, fontFace: HF, fontSize: 60, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true });
    s.addText("NOMBRE DE LA SECCIÓN", { x: 2.5, y: 3.2, w: 10, h: 1.3, fontFace: HF, fontSize: 54, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("Una línea que diga qué viene en esta parte.", { x: 2.5, y: 4.5, w: 9, h: 0.5, fontFace: BF, fontSize: 18, color: "C4C4BE", margin: 0, isTextBox: true });
    s.addNotes("Divisor de sección. Número grande, nombre y una línea. Úsalo para cambiar de momento dentro del encuentro.");
  }
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: "Portadas y divisores" });
    B.ribbon(s, "Cinta de sección", M + 0.3, 3.0, 7.0, { fontSize: 36, h: 1.0 });
    s.addText("Una línea que diga qué viene en esta parte.", { x: M, y: 4.3, w: 9, h: 0.5, fontFace: BF, fontSize: 18, color: "56564F", margin: 0, isTextBox: true });
    s.addNotes("Divisor claro con cinta negra. Texto en cursiva, hasta 24 caracteres.");
  }

  /* ===== TEXTO ===== */
  pres.addSection({ title: "Texto" });
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Texto" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, false, 10);
    B.body(s, P + "\n\n" + P2, M, 1.95, 8.0, 3.5, { fontSize: 18 });
    s.addNotes("Título y cuerpo. El estilo más simple: un título, un antetítulo y hasta dos párrafos cortos.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: "Texto" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, true, 10);
    B.body(s, P + "\n\n" + P2, M, 1.95, 8.0, 3.5, { dark: true, fontSize: 18 });
    s.addNotes("Título y cuerpo, versión oscura.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Texto" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, false, 10);
    B.bullets(s, ["Primer punto, en una línea.", "Segundo punto, en una línea.", "Tercer punto, en una línea.", "Cuarto punto si hace falta."], M, 1.95, 7.5, 3.5, { fontSize: 18, paraSpaceAfter: 10 });
    s.addNotes("Lista. Hasta cinco puntos de una línea. Las viñetas son chevrones.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Texto" });
    B.title(s, T);
    B.kicker(s, "Columna izquierda", M, 1.35, false, 5.5);
    B.body(s, P + "\n\n" + P2, M, 1.75, 5.8, 4.0, { fontSize: 16 });
    B.kicker(s, "Columna derecha", 6.95, 1.35, false, 5.5);
    B.body(s, P + "\n\n" + P2, 6.95, 1.75, 5.8, 4.0, { fontSize: 16 });
    s.addNotes("Dos columnas. Sirve para texto e imagen, o para dos ideas que se comparan.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: "Texto" });
    B.title(s, "Una sola frase");
    s.addText("La idea más importante de la diapositiva, escrita en una frase de hasta quince palabras.", { x: M, y: 2.1, w: CONTENT_W, h: 2.4, fontFace: HF, fontSize: 48, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true, fit: "shrink" });
    s.addNotes("Una sola frase. Para la idea central o una pregunta al grupo.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Texto" });
    B.title(s, "Cita");
    s.addShape("chevron", { x: M, y: 2.2, w: 0.5, h: 0.7, fill: { color: "FCC800" }, line: { width: 0, color: "FCC800" } });
    s.addText("“Lo que dijo la persona, con sus palabras, en dos o tres líneas.”", { x: M + 0.8, y: 1.95, w: 10.5, h: 1.8, fontFace: HF, fontSize: 34, bold: true, italic: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    s.addText("Nombre · cargo · empresa", { x: M + 0.8, y: 3.9, w: 8, h: 0.4, fontFace: BF, fontSize: 15, color: "56564F", margin: 0, isTextBox: true });
    s.addNotes("Cita o historia. Texto en Barlow cursiva. La atribución va en cuerpo.");
  }

  /* ===== TARJETAS Y FILAS ===== */
  pres.addSection({ title: "Tarjetas y filas" });
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Tarjetas y filas" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, false, 10);
    const ic = ["bulb", "tool", "check"];
    const cw = (CONTENT_W - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = M + i * (cw + 0.3);
      B.card(s, x, 1.95, cw, 4.3);
      s.addImage({ data: await B.iconPng(ic[i], "circle"), x: x + 0.3, y: 2.25, w: 0.7, h: 0.7 });
      s.addText("TÍTULO DE LA TARJETA", { x: x + 0.3, y: 3.1, w: cw - 0.6, h: 0.45, fontFace: HF, fontSize: 22, bold: true, color: "0D0D0C", margin: 0, valign: "top", isTextBox: true });
      B.body(s, P, x + 0.3, 3.6, cw - 0.6, 2.4, { fontSize: 15 });
    }
    s.addNotes("Tres tarjetas. Para tres ideas, tres pasos o tres opciones. Cambia el ícono desde la biblioteca.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_OSCURO", sectionTitle: "Tarjetas y filas" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, true, 10);
    const ic = ["clipboard", "clock", "bulb", "star"];
    const cw = (CONTENT_W - 0.3) / 2;
    for (let i = 0; i < 4; i++) {
      const x = M + (i % 2) * (cw + 0.3), y = 1.9 + Math.floor(i / 2) * 2.25;
      B.card(s, x, y, cw, 2.0, { dark: true });
      s.addImage({ data: await B.iconPng(ic[i], "circle-yellow"), x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6 });
      s.addText("TÍTULO DE LA TARJETA", { x: x + 1.1, y: y + 0.3, w: cw - 1.4, h: 0.5, fontFace: HF, fontSize: 20, bold: true, color: "FFFFFF", margin: 0, valign: "middle", isTextBox: true });
      B.body(s, P, x + 1.1, y + 0.9, cw - 1.4, 1.0, { dark: true, fontSize: 14 });
    }
    s.addNotes("Cuatro tarjetas en rejilla, versión oscura.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Tarjetas y filas" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, false, 10);
    const ic = ["chat", "users", "live", "upload"];
    for (let i = 0; i < 4; i++) await B.iconRow(s, ic[i], "Encabezado de la fila", P, M + (i % 2) * 6.3, 1.95 + Math.floor(i / 2) * 1.9, 5.8, { bodyH: 1.0 });
    s.addNotes("Filas de ícono y texto. Dos columnas de dos filas. Hasta seis filas por diapositiva.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Tarjetas y filas" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, false, 10);
    const steps = ["Primer paso en una línea.", "Segundo paso en una línea.", "Tercer paso en una línea.", "Cuarto paso en una línea.", "Quinto paso en una línea."];
    for (let i = 0; i < steps.length; i++) {
      const y = 1.9 + i * 0.9;
      s.addShape("ellipse", { x: M, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: "0D0D0C" }, line: { width: 0, color: "0D0D0C" } });
      s.addText(String(i + 1), { x: M, y: y + 0.05, w: 0.55, h: 0.55, fontFace: HF, fontSize: 20, bold: true, color: "FCC800", align: "center", margin: 0, valign: "middle", isTextBox: true });
      B.body(s, steps[i], M + 0.75, y, 7.5, 0.8, { fontSize: 16, valign: "middle" });
    }
    s.addNotes("Paso a paso numerado. De tres a seis pasos.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Tarjetas y filas" });
    B.title(s, T);
    B.kicker(s, "Línea de tiempo o agenda", M, 1.35, false, 10);
    B.chevronBar(s, [
      { label: "Paso 1", sub: "Qué pasa aquí\nen dos líneas", done: true },
      { label: "Paso 2", sub: "Qué pasa aquí\nen dos líneas", done: true },
      { label: "Paso 3", sub: "Qué pasa aquí\nen dos líneas", done: true },
      { label: "Paso 4", sub: "Qué pasa aquí\nen dos líneas", done: false },
      { label: "Paso 5", sub: "Qué pasa aquí\nen dos líneas", done: false },
    ], M, 1.95, CONTENT_W, { h: 0.55, fontSize: 15, subSize: 13 });
    B.body(s, P, M, 3.6, 8, 1.0, { fontSize: 16 });
    s.addNotes("Barra de chevrones. Amarillo para lo hecho o lo actual, gris para lo que falta. Hasta seis segmentos.");
  }

  /* ===== DATOS ===== */
  pres.addSection({ title: "Datos" });
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Datos" });
    B.title(s, T);
    B.kicker(s, SUB, M, 1.35, false, 10);
    B.bigNumber(s, "42", "qué mide esta cifra", M, 1.9, 3.6);
    B.bigNumber(s, "85 %", "qué mide esta cifra", 4.6, 1.9, 3.6);
    B.bigNumber(s, "3×", "qué mide esta cifra", 8.6, 1.9, 3.6);
    B.body(s, P, M, 4.2, 9, 1.2, { fontSize: 16 });
    s.addNotes("Cifras grandes. Hasta tres por diapositiva, con su etiqueta debajo.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Datos" });
    B.title(s, "Antes y después");
    B.kicker(s, SUB, M, 1.35, false, 10);
    const cw = (CONTENT_W - 0.3) / 2;
    B.card(s, M, 1.95, cw, 4.3, { fill: "FFFFFF", border: "DEDED9" });
    B.kicker(s, "Antes", M + 0.3, 2.2, false, 3);
    s.addText("00", { x: M + 0.3, y: 2.55, w: cw - 0.6, h: 1.0, fontFace: HF, fontSize: 54, bold: true, color: "56564F", margin: 0, valign: "middle", isTextBox: true });
    B.body(s, P, M + 0.3, 3.7, cw - 0.6, 2.3, { fontSize: 15 });
    B.card(s, M + cw + 0.3, 1.95, cw, 4.3, { fill: "0D0D0C", border: "0D0D0C" });
    B.kicker(s, "Después", M + cw + 0.6, 2.2, true, 3);
    s.addText("00", { x: M + cw + 0.6, y: 2.55, w: cw - 0.6, h: 1.0, fontFace: HF, fontSize: 54, bold: true, color: "FCC800", margin: 0, valign: "middle", isTextBox: true });
    B.body(s, P, M + cw + 0.6, 3.7, cw - 0.6, 2.3, { dark: true, fontSize: 15 });
    s.addNotes("Comparación. Gris para el antes, amarillo sobre negro para el después.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Datos" });
    B.title(s, "Gráfico");
    B.kicker(s, SUB, M, 1.35, false, 10);
    s.addChart("bar", [
      { name: "Serie A", labels: ["Uno", "Dos", "Tres", "Cuatro"], values: [45, 30, 60, 20] },
      { name: "Serie B", labels: ["Uno", "Dos", "Tres", "Cuatro"], values: [15, 12, 25, 8] },
    ], {
      x: M, y: 1.85, w: 8.0, h: 4.55, barDir: "bar", barGrouping: "clustered",
      chartColors: ["9A9A93", "FCC800"], showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: "0D0D0C",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: "0D0D0C",
      catAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 12, catAxisLabelColor: "0D0D0C", catGridLine: { style: "none" },
      valAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: 11, valAxisLabelColor: "56564F", valGridLine: { color: "DEDED9", size: 0.75 }, valAxisLineShow: false,
      showTitle: false,
    });
    B.body(s, "Qué muestra el gráfico, en dos o tres líneas. Los datos se cambian con clic derecho > Editar datos.", 8.9, 1.95, 3.83, 3.0, { fontSize: 15 });
    s.addNotes("Gráfico nativo. Serie principal en amarillo, secundaria en gris. Clic derecho > Editar datos.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Datos" });
    B.title(s, "Tabla");
    B.kicker(s, SUB, M, 1.35, false, 10);
    const hdr = (t) => ({ text: t, options: { bold: true, fill: { color: "0D0D0C" }, color: "FCC800", fontFace: HF, fontSize: 14 } });
    const cell = (t, b) => ({ text: t, options: { fontFace: BF, fontSize: 13, bold: !!b, color: "0D0D0C" } });
    const rows = [[hdr("Columna"), hdr("Columna"), hdr("Columna"), hdr("Columna")]];
    for (let i = 0; i < 5; i++) rows.push([cell("Fila " + (i + 1), true), cell("Dato"), cell("Dato"), cell("Dato")]);
    s.addTable(rows, { x: M, y: 1.95, w: CONTENT_W, colW: [3.0, 3.04, 3.04, 3.04], border: { type: "solid", color: "DEDED9", pt: 0.75 }, fill: { color: "FFFFFF" }, margin: 0.08, valign: "middle", rowH: 0.5 });
    s.addNotes("Tabla. Encabezado negro con texto amarillo, filas con borde gris. Hasta ocho filas.");
  }

  /* ===== HERRAMIENTA Y PANTALLA ===== */
  pres.addSection({ title: "Herramienta y pantalla" });
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Herramienta y pantalla" });
    B.title(s, "Prompt");
    B.kicker(s, SUB, M, 1.35, false, 10);
    B.promptBlock(s, "Escribe aquí el prompt completo, tal como se pega en la herramienta: rol, contexto, tarea y formato de salida.", M, 1.95, 7.2, 4.3);
    B.body(s, "A la derecha, qué hace el prompt o qué revisar en la respuesta.", 8.1, 1.95, 4.6, 3.0, { fontSize: 15 });
    s.addNotes("Bloque de prompt. Caja negra con chip amarillo. El texto va completo para que la gente lo copie de la grabación.");
  }
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: "Herramienta y pantalla" });
    B.kicker(s, "Qué se está mostrando", M, 0.5, false, 10);
    s.addShape("roundRect", { x: M, y: 0.95, w: CONTENT_W, h: 5.7, rectRadius: 0.11, fill: { color: "F7F7F5" }, line: { color: "9A9A93", width: 1.25, dashType: "dash" } });
    s.addImage({ data: await B.iconPng("live", "circle"), x: W / 2 - 0.35, y: 3.1, w: 0.7, h: 0.7 });
    s.addText("Pantalla compartida o captura", { x: M, y: 3.95, w: CONTENT_W, h: 0.5, fontFace: BF, fontSize: 16, color: "56564F", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Pantalla libre. Para compartir la herramienta o pegar una captura. Borra el recuadro punteado al pegar.");
  }
  {
    const s = pres.addSlide({ masterName: "LIBRE_CLARO", sectionTitle: "Herramienta y pantalla" });
    B.kicker(s, "Qué se está mostrando", M, 0.5, false, 10);
    B.promptBlock(s, "Prompt", M, 0.95, 5.4, 5.7);
    s.addShape("roundRect", { x: 6.3, y: 0.95, w: 6.4, h: 5.7, rectRadius: 0.11, fill: { color: "F7F7F5" }, line: { color: "9A9A93", width: 1.25, dashType: "dash" } });
    s.addImage({ data: await B.iconPng("camera", "circle"), x: 9.15, y: 3.0, w: 0.7, h: 0.7 });
    s.addText("Resultado", { x: 6.6, y: 3.85, w: 5.8, h: 0.5, fontFace: BF, fontSize: 16, color: "56564F", align: "center", margin: 0, isTextBox: true });
    s.addNotes("Prompt y resultado, lado a lado.");
  }
  {
    const s = pres.addSlide({ masterName: "LIBRE_OSCURO", sectionTitle: "Herramienta y pantalla" });
    s.addShape("rect", { x: 0, y: 0, w: W * 0.58, h: H, fill: { color: "262623" }, line: { width: 0, color: "262623" } });
    s.addImage({ data: await B.iconPng("camera", "circle-yellow"), x: W * 0.29 - 0.4, y: H / 2 - 0.4, w: 0.8, h: 0.8 });
    s.addText("Imagen a sangre", { x: 0, y: H / 2 + 0.5, w: W * 0.58, h: 0.4, fontFace: BF, fontSize: 14, color: "9A9A93", align: "center", margin: 0, isTextBox: true });
    s.addText("TÍTULO SOBRE LA IMAGEN", { x: W * 0.58 + 0.5, y: 2.2, w: W * 0.42 - 1.1, h: 1.6, fontFace: HF, fontSize: 40, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true, fit: "shrink" });
    B.body(s, P, W * 0.58 + 0.5, 3.9, W * 0.42 - 1.1, 2.0, { dark: true, fontSize: 16 });
    s.addNotes("Imagen a media sangre con titular. Reemplaza el rectángulo gris por la foto: clic derecho > Cambiar imagen.");
  }

  /* ===== CIERRE ===== */
  pres.addSection({ title: "Cierre" });
  {
    const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Cierre" });
    B.wordmark(s, M, 0.55, { size: 26 });
    s.addText("Gracias", { x: M, y: 2.0, w: 11, h: 1.5, fontFace: HF, fontSize: 80, bold: true, color: "FCC800", margin: 0, valign: "top", isTextBox: true });
    s.addText("Una línea de cierre.", { x: M, y: 3.5, w: 11, h: 0.7, fontFace: HF, fontSize: 30, bold: true, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true });
    await B.iconRow(s, "chat", "superhumanos@gyj.com.co", "Correo del programa.", M, 4.7, 7.5, { dark: true, bodyH: 0.6 });
    s.addNotes("Cierre. Gracias y el correo del programa.");
  }

  /* ===== GUÍA Y BIBLIOTECA ===== */
  pres.addSection({ title: "Guía y biblioteca" });
  {
    const s = pres.addSlide({ masterName: "CONTENIDO_CLARO", sectionTitle: "Guía y biblioteca" });
    B.title(s, "Cómo usar esta plantilla");
    B.kicker(s, "Se borra antes de presentar, junto con la biblioteca de íconos", M, 1.35, false, 10);
    B.bullets(s, [
      "Duplica la diapositiva del estilo que necesites y cambia el texto. Cada una tiene en las notas para qué sirve.",
      "Para una diapositiva vacía con el pie y el número: Inicio > Nueva diapositiva y elige CONTENIDO_CLARO, CONTENIDO_OSCURO, LIBRE_CLARO o LIBRE_OSCURO.",
      "Los íconos están en la última diapositiva: copia y pega el que necesites.",
      "El pie de página se cambia una vez por archivo en Vista > Patrón de diapositivas.",
      "Nombre del archivo: SH_[PERFIL]_T[taller]_E[encuentro]_[AAAAMMDD].",
    ], M, 1.95, 6.2, 4.3, { fontSize: 14, paraSpaceAfter: 8 });
    B.ribbon(s, "Lo que no cambia", 7.55, 1.9, 5.1, { fontSize: 18, h: 0.5 });
    B.bullets(s, [
      "Titulares en Barlow Condensed, en mayúsculas. Cuerpo en Atkinson Hyperlegible Next, en tipo oración.",
      "Negro #0D0D0C y amarillo #FCC800. Sobre amarillo, texto negro. Amarillo como texto solo sobre negro. Sobre blanco, el acento es #8A6B00.",
      "Una idea por diapositiva y hasta 60 palabras de cuerpo.",
      "Sin logos de las empresas.",
      "Si falta una fuente, PowerPoint usa Arial Narrow y Arial. Instala las dos antes de presentar.",
    ], 7.3, 2.6, 5.4, 3.8, { fontSize: 13, paraSpaceAfter: 6 });
    s.addNotes("Guía de uso. Se borra antes de presentar.");
  }
  await B.iconLibrarySlide(pres, "Guía y biblioteca");
});
