const { fs, path, Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer, PageNumber, LevelFormat, AlignmentType, TabStopType, BLACK, GOLD, MUTED, HF, BF, run, P, H1, K, B, bold, table, box, spacer } = require("./docx-brand.js");

const pct = (a, b) => `${Math.round(a / b * 100)} %`;

const PERFILES = {
  comercial: {
    nombre: "comercial", NOMBRE: "COMERCIAL",
    cobertura: "135 de 180 personas respondieron, 75 % del perfil. 133 tienen índice completo.",
    empresas: "G&J Ferreterías aporta 100 de las 133 respuestas; siguen Colmena (12), Sigmasteel (8), Almasa (7), Almapanel (3) y G&J Soluciones (3).",
    indices: [["Casos situacionales", "40 %", "66,6"], ["Herramientas digitales", "35 %", "80,2"], ["Uso actual de IA", "25 %", "54,3"], ["Índice de dominio", "100 %", "68,3"], ["Índice de mentalidad", "aparte", "76,3"]],
    lecturaIndices: "El uso de IA es el eje más bajo del perfil. El eje situacional quedó en 66,6 porque los casos comerciales del instrumento resultaron exigentes: dos de los diez ítems bajaron del 50 % de acierto.",
    items: [
      ["S_1", "15 solicitudes de cotización por WhatsApp un lunes", "27,8 %"],
      ["S_5", "Dos horas cada lunes revisando 40 correos del fin de semana", "46,6 %"],
      ["D_5", "Estructura para presentar resultados del trimestre en 10 minutos", "60,9 %"],
      ["S_4", "Pedido retenido por cartera y el cliente lo necesita el jueves", "73,7 %"],
      ["D_1", "Construir un documento entre cuatro personas", "74,4 %"],
      ["D_4", "Lista con el mismo cliente escrito de formas distintas", "86,5 %"],
      ["S_2", "Cliente monta 300 toneladas y lo proyectado eran 100", "87,2 %"],
      ["D_3", "500 ventas del mes y hay que filtrar por UNE", "88,0 %"],
      ["D_2", "Piden por WhatsApp la cédula y el teléfono de un cliente", "91,0 %"],
      ["S_3", "Cotización en 10 minutos sin alcanzar a validar el dato", "97,7 %"],
    ],
    lecturaItems: "Los tres ítems más bajos comparten la misma raíz: un volumen de trabajo repetitivo que llega por un canal desordenado. El equipo resuelve bien los casos de criterio y de relación con el cliente, y se queda cuando la pregunta es cómo organizar el flujo.",
    n: 133,
    frecuencia: [["Ninguna vez", 32], ["Una o dos veces", 35], ["Una vez por semana", 16], ["Varias veces por semana", 38], ["Todos los días", 12]],
    seguimiento: [["En un cuaderno o en papel", 75], ["En el correo, marcando los mensajes", 22], ["En un archivo propio de Excel", 19], ["En una herramienta de tareas compartida", 11], ["Los recuerda, sin anotarlos", 6]],
    tiempo: [["Menos de 2 horas", 39], ["Entre 2 y 5 horas", 42], ["Entre 5 y 10 horas", 33], ["Más de 10 horas", 19]],
    trabajo: [
      "57 comerciales hacen informes que también hace otra área, o creen que sí.",
      "CRM: 79 de 135 lo usan en más del 75 % de su trabajo. 23 podrían enseñarlo y 85 lo manejan bien. 15 se traban y 10 no lo conocen: 25 personas para nivelar antes del Taller 4.",
      "Política de IA: 42 comerciales, cerca del 31 %, no la han firmado o no sabían que existía.",
    ],
    condiciones: "31 comerciales trabajan sin computador propio, más de la mitad de los 52 del programa. Parte del grupo se conecta desde el celular.",
    reconocimiento: "En comercial, el bono o premio llega al 76 % y el tiempo libre al 40 %. Úsalo al anunciar los puntos en la apertura y el cierre.",
    grupos: [["G1 Arranque", "menos de 59,0", "32"], ["G2 Base", "59,0 a 70,3", "34"], ["G3 Avance", "70,4 a 79,7", "33"], ["G4 Referentes", "80 o más", "34"]],
    tramos: "Con los cortes absolutos del instrumento: 7 en Chispa, 31 en Impulso, 63 en Fuerza y 32 en Grandeza.",
    formadores: "16 de los 43 candidatos a formador interno son comerciales.",
    talleres: [
      "Taller 1: el caso de la demostración es el de las 15 cotizaciones por WhatsApp. Es el ítem con 27,8 % de acierto y aparece también en lo que la gente escribió que espera de la IA.",
      "Taller 1: avanza al ritmo de quien nunca la abrió. La mitad del grupo usó IA una o dos veces o ninguna en el último mes.",
      "Taller 2: la clasificación de tareas arranca con las horas que cada persona declaró. 52 comerciales dedican más de 5 horas semanales a tareas de pasos iguales.",
      "Taller 3: la bandeja de 40 correos del lunes es el caso de flujo. Es el segundo ítem más bajo, con 46,6 %.",
      "Taller 4: veinte minutos de nivelación de CRM antes del mapa de fuentes, con los 23 que podrían enseñarlo como apoyo.",
      "Taller 2: bloque de cierre para contar los resultados del mes en diez minutos: qué pasó, por qué, qué propongo, qué necesito. Es el tercer ítem más bajo, con 60,9 %.",
      "Todas las prácticas tienen versión para celular: el prompt se copia desde el kit y la evidencia es una captura.",
    ],
  },
  administrativo: {
    nombre: "administrativo", NOMBRE: "ADMINISTRATIVO",
    cobertura: "282 de 496 personas respondieron, 57 % del perfil. Todas tienen índice completo.",
    empresas: "G&J Ferreterías aporta 80 de las 282 respuestas; siguen Colmena (58), G&J Soluciones (45), Almasa (45), Sigmasteel (44) y G&J Ramírez (10). El grueso de quienes faltan está en G&J Ferreterías.",
    indices: [["Casos situacionales", "40 %", "88,7"], ["Herramientas digitales", "35 %", "86,2"], ["Uso actual de IA", "25 %", "56,8"], ["Índice de dominio", "100 %", "79,9"], ["Índice de mentalidad", "aparte", "77,1"]],
    lecturaIndices: "El uso de IA queda treinta puntos por debajo de los otros dos ejes. El perfil resuelve bien los casos de su trabajo y maneja las herramientas digitales; lo que falta es el uso de IA con método.",
    items: [
      ["D_5", "Estructura para presentar resultados del trimestre en 10 minutos", "74,8 %"],
      ["S_5", "El ingreso de cada persona exige 30 firmas en papel", "78,4 %"],
      ["D_1", "Construir un documento entre cuatro personas", "79,8 %"],
      ["S_1", "Dos horas cada lunes consolidando el reporte de cinco archivos", "86,2 %"],
      ["D_4", "Lista con el mismo cliente escrito de formas distintas", "89,4 %"],
      ["S_2", "Dos informes con cifras distintas de devoluciones del mismo mes", "91,5 %"],
      ["S_3", "Informe en 20 minutos sin alcanzar a verificar una cifra", "91,5 %"],
      ["D_3", "500 ventas del mes y hay que filtrar por UNE", "92,9 %"],
      ["D_2", "Piden por WhatsApp la cédula y el teléfono de un cliente", "94,3 %"],
      ["S_4", "PQR que necesita información de producción de otra sede", "96,1 %"],
    ],
    lecturaItems: "Ningún ítem baja del 74 %. Los dos más bajos piden comunicar y rediseñar un proceso: estructurar una presentación corta y repensar un trámite de 30 firmas. El taller gana más trabajando el criterio de qué contar y qué eliminar que explicando herramientas.",
    n: 282,
    frecuencia: [["Ninguna vez", 59], ["Una o dos veces", 69], ["Una vez por semana", 38], ["Varias veces por semana", 81], ["Todos los días", 35]],
    seguimiento: [["En un cuaderno o en papel", 98], ["En el correo, marcando los mensajes", 87], ["En un archivo propio de Excel", 49], ["Los recuerda, sin anotarlos", 26], ["En una herramienta de tareas compartida", 22]],
    tiempo: [["Menos de 2 horas", 87], ["Entre 2 y 5 horas", 106], ["Entre 5 y 10 horas", 53], ["Más de 10 horas", 36]],
    trabajo: [
      "103 administrativos hacen informes que también hace otra área, o creen que sí.",
      "Registro del trabajo: 104 personas usan el ERP UnoE y 72 archivos propios de Excel. 184 responden que el CRM no aplica a su cargo.",
      "Política de IA: 114 administrativos, cerca del 40 %, no la han firmado o no sabían que existía.",
      "48 personas del programa dicen que lo que saben de su trabajo no está escrito de forma que otra persona pueda continuarlo, y 79 rehacen en un archivo propio la información que ya está en el sistema.",
    ],
    condiciones: "21 administrativos trabajan sin computador propio. En el programa, 66 personas tienen conexión intermitente o solo datos del celular y 142 reportan bloqueos totales o parciales a herramientas de IA desde la red.",
    reconocimiento: "En administrativo, el crecimiento o ascenso llega al 66 % y la certificación al 17 %. Úsalo al anunciar los puntos en la apertura y el cierre.",
    grupos: [["G1 Arranque", "menos de 74,4", "70"], ["G2 Base", "74,4 a 83,2", "70"], ["G3 Avance", "83,3 a 89,9", "70"], ["G4 Referentes", "90 o más", "72"]],
    tramos: "Con los cortes absolutos del instrumento: 4 en Chispa, 19 en Impulso, 87 en Fuerza y 172 en Grandeza.",
    formadores: "27 de los 43 candidatos a formador interno son administrativos.",
    talleres: [
      "Taller 1: abre por la estructura de comunicar resultados, el ítem más bajo con 74,8 %. El resumen de cinco puntos de la demostración se convierte en lo que se presenta en diez minutos.",
      "Taller 1: avanza al ritmo de quien nunca la abrió. 45 % del grupo usó IA una o dos veces o ninguna en el último mes.",
      "Taller 2: cada persona anota cuánto le toma el reporte del mes antes de tocarlo. 89 administrativos dedican más de 5 horas semanales a tareas de pasos iguales.",
      "Taller 3: el tablero de pendientes parte de cómo hacen seguimiento hoy: 185 en papel o marcando correos, 22 en una herramienta compartida.",
      "Taller 4: inventario de reportes duplicados del área y guía de una página para lo que vive en la cabeza de alguien.",
      "Taller 6: el trámite de 30 firmas es el caso de rediseño. Es el segundo ítem más bajo, con 78,4 %.",
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
  c.push(P("Diagnóstico inicial con corte del 14 de octubre de 2026 · instrumento de 38 preguntas · resultados agregados del perfil", { run: { color: MUTED } }));
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
  c.push(P("En el programa, 50 % usa la IA con cuenta personal y 15 % con la cuenta de la empresa. ChatGPT llega al 74 % de quienes la usan y Gemini al 47 %. Los usos se concentran en buscar información y redactar; analizar datos llega al 26 % y resumir documentos al 19 %."));

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
  c.push(P("Las afirmaciones más altas del programa: la gente busca la causa cuando algo falla (4,53 sobre 5) y resuelve con la otra área antes de escalar (4,48). Las más bajas: usar la información del sistema sin rehacerla en un archivo propio (3,61) y tener escrito lo que sabe de su trabajo (3,76)."));
  c.push(P("329 personas escribieron qué esperan de la IA: ahorrar tiempo (38 %), organizar tareas y pendientes (23 %), atención al cliente, cotizaciones y ventas (14 %), automatizar lo repetitivo (11 %) y analizar datos (10 %)."));
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
    "Las cifras corresponden a quienes respondieron hasta el 14 de octubre. Respondieron primero las empresas con seguimiento cercano y las personas con computador y conexión estables. Al cerrar la convocatoria es razonable esperar índices algo más bajos y más personas sin dispositivo propio.",
  ]));

  const doc = new Document({
    creator: "Pivot Power", title: `Resumen del diagnóstico · perfil ${p.nombre}`,
    styles: { default: { document: { run: { font: BF, size: 21, color: BLACK } } } },
    numbering: { config: [{ reference: "chev", levels: [{ level: 0, format: LevelFormat.BULLET, text: "›", alignment: AlignmentType.LEFT, style: { run: { font: HF, bold: true, color: GOLD, size: 24 }, paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
    sections: [{
      properties: { page: { margin: { top: 1100, bottom: 1100, left: 1134, right: 1134 } } },
      headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: `SUPER HUMANOS · RESUMEN DEL DIAGNÓSTICO · PERFIL ${p.NOMBRE}`, font: HF, size: 17, bold: true, color: MUTED, characterSpacing: 30 })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9600 }], children: [new TextRun({ text: "Pivot Power · 21 de octubre de 2026", font: BF, size: 16, color: MUTED }), new TextRun({ text: "\t", font: BF, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 18, bold: true, color: BLACK })] })] }) },
      children: c,
    }],
  });
  const out = path.join(__dirname, "out", `Resumen del diagnóstico · perfil ${p.nombre}.docx`);
  return Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ok", out, buf.length); });
}
Promise.all(Object.values(PERFILES).map(build));
