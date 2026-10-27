const { fs, path, Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer, PageNumber, LevelFormat, AlignmentType, TabStopType, BLACK, GOLD, MUTED, HF, BF, run, P, H1, K, B, bold, table, box, spacer } = require("./docx-brand.js");

const pct = (a, b) => `${Math.round(a / b * 100)} %`;

const PERFILES = {
  comercial: {
    nombre: "comercial", NOMBRE: "COMERCIAL",
    cobertura: "150 de 201 personas terminaron el diagnóstico, 75 % del perfil, y las 150 tienen índice completo. La base del perfil pasó de 180 a 201 personas desde el corte anterior.",
    empresas: "G&J Ferreterías aporta 115 de las 150 respuestas; siguen Colmena (12), Almasa (8), Sigmasteel (8), G&J Soluciones (4) y Almapanel (3).",
    indices: [["Casos situacionales", "40 %", "67,6"], ["Herramientas digitales", "35 %", "80,4"], ["Uso actual de IA", "25 %", "56,0"], ["Índice de dominio", "100 %", "69,2"], ["Índice de mentalidad", "aparte", "76,2"]],
    lecturaIndices: "El uso de IA es el eje más bajo del perfil. El eje situacional quedó en 67,6 porque los casos comerciales del instrumento resultaron exigentes: dos de los diez ítems bajaron del 50 % de acierto.",
    items: [
      ["S_1", "15 solicitudes de cotización por WhatsApp un lunes", "30,0 %"],
      ["S_5", "Dos horas cada lunes revisando 40 correos del fin de semana", "46,7 %"],
      ["D_5", "Estructura para presentar resultados del trimestre en 10 minutos", "63,3 %"],
      ["S_4", "Pedido retenido por cartera y el cliente lo necesita el jueves", "75,3 %"],
      ["D_1", "Construir un documento entre cuatro personas", "76,7 %"],
      ["D_4", "Lista con el mismo cliente escrito de formas distintas", "85,3 %"],
      ["S_2", "Cliente monta 300 toneladas y lo proyectado eran 100", "88,0 %"],
      ["D_3", "500 ventas del mes y hay que filtrar por UNE", "88,0 %"],
      ["D_2", "Piden por WhatsApp la cédula y el teléfono de un cliente", "88,7 %"],
      ["S_3", "Cotización en 10 minutos sin alcanzar a validar el dato", "98,0 %"],
    ],
    lecturaItems: "Los tres ítems más bajos comparten la misma raíz: un volumen de trabajo repetitivo que llega por un canal desordenado. El equipo resuelve bien los casos de criterio y de relación con el cliente, y se queda cuando la pregunta es cómo organizar el flujo.",
    n: 150,
    frecuencia: [["Ninguna vez", 33], ["Una o dos veces", 40], ["Una vez por semana", 18], ["Varias veces por semana", 46], ["Todos los días", 13]],
    seguimiento: [["En un cuaderno o en papel", 84], ["En el correo, marcando los mensajes", 27], ["En un archivo propio de Excel", 19], ["En una herramienta de tareas compartida", 11], ["Los recuerda, sin anotarlos", 9]],
    tiempo: [["Menos de 2 horas", 46], ["Entre 2 y 5 horas", 50], ["Entre 5 y 10 horas", 35], ["Más de 10 horas", 19]],
    trabajo: [
      "62 comerciales hacen informes que también hace otra área y 41 creen que sí.",
      "CRM: 91 de 150 lo usan en más del 75 % de su trabajo. 27 podrían enseñarlo y 97 lo manejan bien. 16 se traban y 10 no lo conocen: 26 personas para nivelar antes del Taller 4.",
      "Política de IA: 52 comerciales, cerca del 35 %, no la han firmado o no sabían que existía. 86 usan la IA con cuenta personal y 17 con la de la empresa.",
    ],
    condiciones: "39 comerciales trabajan desde el celular o sin dispositivo en su puesto. 49 se conectarían a una sesión en vivo desde el celular y 32 tienen conexión intermitente o solo datos. 60 reportan herramientas de IA bloqueadas total o parcialmente desde la red de la empresa.",
    mentalidad: "Las afirmaciones más altas del perfil: resuelve con la otra área antes de escalar (4,53 sobre 5) y busca la causa cuando algo falla (4,45). Las más bajas: usar la información del sistema sin rehacerla en un archivo propio (3,51; 35 personas responden 1 o 2) y tener escrito lo que sabe de su trabajo (3,68; 23 personas).",
    reconocimiento: "En comercial, el bono o premio llega al 76 %, el crecimiento o ascenso al 50 % y el tiempo libre al 39 %. Úsalo al anunciar los puntos en la apertura y el cierre.",
    grupos: [["G1 Arranque", "menos de 59,5", "38"], ["G2 Base", "59,5 a 71,7", "36"], ["G3 Avance", "71,8 a 80,0", "38"], ["G4 Referentes", "80,1 o más", "38"]],
    tramos: "Con los cortes absolutos del instrumento: 7 en Chispa, 33 en Impulso, 72 en Fuerza y 38 en Grandeza.",
    formadores: "17 de los 50 candidatos a formador interno son comerciales.",
    talleres: [
      "Taller 1: el caso de la demostración es el de las 15 cotizaciones por WhatsApp. Es el ítem con 30,0 % de acierto y aparece también en lo que la gente escribió que espera de la IA.",
      "Taller 1: avanza al ritmo de quien nunca la abrió. 73 de 150, casi la mitad, usaron IA una o dos veces o ninguna en el último mes.",
      "Taller 2: la clasificación de tareas arranca con las horas que cada persona declaró. 54 comerciales dedican más de 5 horas semanales a tareas de pasos iguales.",
      "Taller 3: la bandeja de 40 correos del lunes es el caso de flujo. Es el segundo ítem más bajo, con 46,7 %.",
      "Taller 4: veinte minutos de nivelación de CRM antes del mapa de fuentes, con los 27 que podrían enseñarlo como apoyo.",
      "Taller 2: bloque de cierre para contar los resultados del mes en diez minutos: qué pasó, por qué, qué propongo, qué necesito. Es el tercer ítem más bajo, con 63,3 %.",
      "Todas las prácticas tienen versión para celular: el prompt se copia desde el kit y la evidencia es una captura.",
    ],
  },
  administrativo: {
    nombre: "administrativo", NOMBRE: "ADMINISTRATIVO",
    cobertura: "332 de 479 personas terminaron el diagnóstico, 69 % del perfil, y las 332 tienen índice completo.",
    empresas: "G&J Ferreterías aporta 106 de las 332 respuestas; siguen Colmena (62), G&J Soluciones (58), Almasa (50), Sigmasteel (46) y G&J Ramírez (10). El grueso de quienes faltan sigue en G&J Ferreterías.",
    indices: [["Casos situacionales", "40 %", "88,1"], ["Herramientas digitales", "35 %", "85,7"], ["Uso actual de IA", "25 %", "57,1"], ["Índice de dominio", "100 %", "79,5"], ["Índice de mentalidad", "aparte", "77,2"]],
    lecturaIndices: "El uso de IA queda treinta puntos por debajo de los otros dos ejes. El perfil resuelve bien los casos de su trabajo y maneja las herramientas digitales; lo que falta es el uso de IA con método.",
    items: [
      ["D_5", "Estructura para presentar resultados del trimestre en 10 minutos", "74,7 %"],
      ["S_5", "El ingreso de cada persona exige 30 firmas en papel", "76,8 %"],
      ["D_1", "Construir un documento entre cuatro personas", "78,6 %"],
      ["S_1", "Dos horas cada lunes consolidando el reporte de cinco archivos", "84,0 %"],
      ["D_4", "Lista con el mismo cliente escrito de formas distintas", "89,5 %"],
      ["S_2", "Dos informes con cifras distintas de devoluciones del mismo mes", "91,6 %"],
      ["D_3", "500 ventas del mes y hay que filtrar por UNE", "91,6 %"],
      ["S_3", "Informe en 20 minutos sin alcanzar a verificar una cifra", "92,2 %"],
      ["D_2", "Piden por WhatsApp la cédula y el teléfono de un cliente", "94,3 %"],
      ["S_4", "PQR que necesita información de producción de otra sede", "96,1 %"],
    ],
    lecturaItems: "Ningún ítem baja del 74 %. Los dos más bajos piden comunicar y rediseñar un proceso: estructurar una presentación corta y repensar un trámite de 30 firmas. El taller gana más trabajando el criterio de qué contar y qué eliminar que explicando herramientas.",
    n: 332,
    frecuencia: [["Ninguna vez", 70], ["Una o dos veces", 84], ["Una vez por semana", 44], ["Varias veces por semana", 96], ["Todos los días", 38]],
    seguimiento: [["En un cuaderno o en papel", 117], ["En el correo, marcando los mensajes", 103], ["En un archivo propio de Excel", 55], ["Los recuerda, sin anotarlos", 34], ["En una herramienta de tareas compartida", 23]],
    tiempo: [["Menos de 2 horas", 112], ["Entre 2 y 5 horas", 120], ["Entre 5 y 10 horas", 59], ["Más de 10 horas", 41]],
    trabajo: [
      "121 administrativos hacen informes que también hace otra área y 69 creen que sí.",
      "Registro del trabajo: 127 usan el ERP UnoE, 80 archivos propios de Excel y 65 el correo. 219 responden que el CRM no aplica a su cargo.",
      "Política de IA: 115 administrativos, cerca del 35 %, no la han firmado o no sabían que existía. 160 usan la IA con cuenta personal y 57 con la de la empresa.",
      "32 administrativos dicen que lo que saben de su trabajo no está escrito de forma que otra persona pueda continuarlo, y 57 rehacen en un archivo propio la información que ya está en el sistema.",
    ],
    condiciones: "24 administrativos trabajan desde el celular o sin dispositivo en su puesto y 18 comparten computador. 53 tienen conexión intermitente o solo datos del celular. 100 reportan herramientas de IA bloqueadas total o parcialmente desde la red de la empresa.",
    mentalidad: "Las afirmaciones más altas del perfil: busca la causa cuando algo falla (4,58 sobre 5) y resuelve con la otra área antes de escalar (4,48). Las más bajas: usar la información del sistema sin rehacerla en un archivo propio (3,64; 57 personas responden 1 o 2) y tener escrito lo que sabe de su trabajo (3,79; 32 personas).",
    reconocimiento: "En administrativo, el crecimiento o ascenso llega al 66 %, el bono o premio al 62 %, una experiencia para el área al 22 % y la certificación al 17 %. Úsalo al anunciar los puntos en la apertura y el cierre.",
    grupos: [["G1 Arranque", "menos de 74,1", "82"], ["G2 Base", "74,1 a 82,9", "84"], ["G3 Avance", "83,0 a 89,9", "82"], ["G4 Referentes", "90 o más", "84"]],
    tramos: "Con los cortes absolutos del instrumento: 6 en Chispa, 23 en Impulso, 105 en Fuerza y 198 en Grandeza.",
    formadores: "33 de los 50 candidatos a formador interno son administrativos.",
    talleres: [
      "Taller 1: abre por la estructura de comunicar resultados, el ítem más bajo con 74,7 %. El resumen de cinco puntos de la demostración se convierte en lo que se presenta en diez minutos.",
      "Taller 1: avanza al ritmo de quien nunca la abrió. 154 de 332, el 46 %, usaron IA una o dos veces o ninguna en el último mes.",
      "Taller 2: cada persona anota cuánto le toma el reporte del mes antes de tocarlo. 100 administrativos dedican más de 5 horas semanales a tareas de pasos iguales.",
      "Taller 3: el tablero de pendientes parte de cómo hacen seguimiento hoy: 220 en papel o marcando correos, 23 en una herramienta compartida.",
      "Taller 4: inventario de reportes duplicados del área y guía de una página para lo que vive en la cabeza de alguien.",
      "Taller 6: el trámite de 30 firmas es el caso de rediseño. Es el segundo ítem más bajo, con 76,8 %.",
    ],
  },
};

function build(p) {
  const c = [];
  const logo = path.join(__dirname, "..", "logo", "pivot-power-logo-claro.png");
  c.push(new Paragraph({ spacing: { before: 600, after: 500 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(logo), transformation: { width: 230, height: Math.round(230 * 546 / 2020) } })] }));
  c.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: "SUPER HUMANOS", font: HF, size: 28, bold: true, color: GOLD, characterSpacing: 60 })] }));
  c.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "RESUMEN DEL DIAGNÓSTICO", font: HF, size: 64, bold: true, color: BLACK })] }));
  c.push(new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: `PERFIL ${p.NOMBRE} · LO QUE TRAE EL GRUPO AL PRIMER TALLER`, font: HF, size: 32, bold: true, color: BLACK })] }));
  c.push(P("Diagnóstico inicial con corte del 23 de octubre de 2026 · instrumento de 38 preguntas · resultados agregados del perfil", { run: { color: MUTED } }));
  c.push(spacer());
  c.push(box("Para qué sirve este documento", [
    `Resume lo que el diagnóstico encontró en el perfil ${p.nombre} y lo convierte en decisiones para tus talleres. Trae solo cifras agregadas. Los resultados por persona quedan con la dirección del programa.`,
  ]));

  c.push(H1("1 · Quién respondió"));
  c.push(P(p.cobertura));
  c.push(P(p.empresas));

  c.push(H1("2 · Los índices del perfil"));
  c.push(P("El índice de dominio combina tres ejes ponderados. La mentalidad se mide aparte, en escala de 1 a 5 llevada a base 100."));
  c.push(table(["Eje", "Peso", "Resultado"], p.indices, [5, 2, 3]));
  c.push(P(p.lecturaIndices));
  c.push(P("El instrumento usa casos distintos para cada perfil. El índice sirve para comparar personas dentro del perfil y para medir a cada persona contra sí misma al cierre.", { run: { color: MUTED } }));

  c.push(H1("3 · Qué resolvió bien y qué resolvió mal"));
  c.push(P("Acierto por situación, de menor a mayor. Las que quedan por debajo del 75 % son el insumo directo de tus casos de clase."));
  c.push(table(["Ítem", "Situación planteada", "Acierto"], p.items, [1.2, 7, 1.8]));
  c.push(P(p.lecturaItems));

  c.push(H1("4 · Cómo usa la IA hoy"));
  c.push(table(["Uso de IA en el último mes", "Personas", "%"], p.frecuencia.map(([a, b]) => [a, String(b), pct(b, p.n)]), [6, 2, 2]));
  c.push(P("Los usos se concentran en buscar información y redactar; analizar datos y resumir documentos quedan muy por detrás. ChatGPT es la herramienta más usada, seguida de Gemini."));

  c.push(H1("5 · Cómo trabaja hoy"));
  c.push(K("Seguimiento de pendientes"));
  c.push(table(["Dónde lleva sus pendientes", "Personas", "%"], p.seguimiento.map(([a, b]) => [a, String(b), pct(b, p.n)]), [6, 2, 2]));
  c.push(spacer());
  c.push(K("Horas semanales en tareas de pasos siempre iguales"));
  c.push(table(["Tiempo declarado", "Personas", "%"], p.tiempo.map(([a, b]) => [a, String(b), pct(b, p.n)]), [6, 2, 2]));
  c.push(spacer());
  p.trabajo.forEach((t) => c.push(B(t)));

  c.push(H1("6 · Condiciones para los encuentros"));
  c.push(P(p.condiciones));
  c.push(P("La grabación en ecert es parte del diseño del taller. Cada práctica funciona en celular y la herramienta del encuentro es una sola, la autorizada por TIC."));

  c.push(H1("7 · Mentalidad y lo que espera de la IA"));
  c.push(P(p.mentalidad));
  c.push(P("386 personas del programa escribieron qué esperan de la IA. Los temas que más se repiten: ahorrar tiempo, organizar tareas y pendientes, atención al cliente y cotizaciones, automatizar lo repetitivo y analizar datos."));
  c.push(P([run("Una respuesta para citar en la apertura: "), bold("«Me gustaría aligerar las tareas, ganando tiempo y hacer uso adecuado del mismo. No me gusta que la IA reemplace mi pensar, análisis y criterio.»"), run(" El programa enseña a delegar la ejecución repetitiva y a quedarse con el criterio.")]));
  c.push(P(p.reconocimiento));

  c.push(H1("8 · Los grupos del perfil"));
  c.push(P("Los grupos se arman con percentiles dentro del perfil para que queden parejos."));
  c.push(table(["Grupo", "Corte de dominio", "Personas"], p.grupos, [3.5, 4, 2.5]));
  c.push(P(p.tramos));
  c.push(P(`Los G4 Referentes apoyan a dos compañeros cada uno durante la práctica. ${p.formadores} La asignación es provisional hasta que cierre la convocatoria.`));

  c.push(H1("9 · Qué cambia en tus talleres"));
  p.talleres.forEach((t) => c.push(B(t)));
  c.push(P("El detalle taller por taller está en «Ajustes a los talleres · perfil " + p.nombre + "», sección 4.", { run: { color: MUTED } }));

  c.push(spacer());
  c.push(box("Advertencia de lectura", [
    "Las cifras corresponden a quienes terminaron el diagnóstico hasta el 23 de octubre: 482 de 680 personas, 71 % del programa. Respondieron primero las empresas con seguimiento cercano y las personas con computador y conexión estables. Las 191 que faltan pueden mover los cortes de los grupos.",
  ]));

  const doc = new Document({
    creator: "Pivot Power", title: `Resumen del diagnóstico · perfil ${p.nombre}`,
    styles: { default: { document: { run: { font: BF, size: 21, color: BLACK } } } },
    numbering: { config: [{ reference: "chev", levels: [{ level: 0, format: LevelFormat.BULLET, text: "›", alignment: AlignmentType.LEFT, style: { run: { font: HF, bold: true, color: GOLD, size: 24 }, paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
    sections: [{
      properties: { page: { margin: { top: 1100, bottom: 1100, left: 1134, right: 1134 } } },
      headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: `SUPER HUMANOS · RESUMEN DEL DIAGNÓSTICO · PERFIL ${p.NOMBRE}`, font: HF, size: 17, bold: true, color: MUTED, characterSpacing: 30 })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9600 }], children: [new TextRun({ text: "Pivot Power · 27 de octubre de 2026", font: BF, size: 16, color: MUTED }), new TextRun({ text: "\t", font: BF, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 18, bold: true, color: BLACK })] })] }) },
      children: c,
    }],
  });
  const out = path.join(__dirname, "out", `Resumen del diagnóstico · perfil ${p.nombre}.docx`);
  return Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ok", out, buf.length); });
}
Promise.all(Object.values(PERFILES).map(build));
