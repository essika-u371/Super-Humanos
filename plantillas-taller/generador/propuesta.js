const { fs, path, Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer, PageNumber, LevelFormat, AlignmentType, TabStopType, BLACK, YELLOW, GOLD, MUTED, HF, BF, run, P, H1, H2, K, B, bold, table, box, spacer } = require("./docx-brand.js");

const children = [];

/* Portada */
children.push(new Paragraph({ spacing: { before: 600, after: 500 }, children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(__dirname, "..", "logo", "pivot-power-logo-claro.png")), transformation: { width: 230, height: Math.round(230 * 546 / 2020) } })] }));
children.push(new Paragraph({ spacing: { before: 0, after: 100 }, children: [new TextRun({ text: "SUPER HUMANOS", font: HF, size: 28, bold: true, color: GOLD, characterSpacing: 60 })] }));
children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "PROPUESTA PARA LA POLÍTICA DE USO DE IA", font: HF, size: 64, bold: true, color: BLACK })] }));
children.push(new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: "FASES DE ADOPCIÓN, REGLA DE INFORMACIÓN POR CUENTA Y LICENCIAS POR ACTIVIDAD", font: HF, size: 32, bold: true, color: BLACK })] }));
children.push(P("Para el área de TIC y Gestión del Talento Humano de G&J Empresas de Acero. Complementa la política de uso de inteligencia artificial y gestión de riesgo, código TH-PT-003, versión 1, que rige para las siete empresas del grupo.", { run: { color: MUTED } }));
children.push(P("Pivot Power · 15 de octubre de 2026 · versión para discusión", { run: { color: MUTED } }));
children.push(spacer());
children.push(box("Qué propone este documento", [
  "Una adopción en tres fases, amarradas al calendario del programa: aprendizaje con herramientas gratuitas, trabajo con información interna en cuenta corporativa y licencias pagas asignadas por actividad.",
  "Una regla de información por tipo de cuenta que cualquier persona aplica en diez segundos. Es la misma tabla que se enseña en el Taller 1 y la que se propone como anexo de la política.",
  "Precisiones de texto a ocho numerales de la política, con la redacción propuesta, para que la regla que firma la gente sea la que practica en los talleres.",
  "Cinco niveles de licencia con un criterio de asignación por actividad medida, y lo que el programa necesita de TIC con fechas.",
]));

/* 1 · Punto de partida */
children.push(H1("1 · Punto de partida"));
children.push(P("El diagnóstico del programa, con 415 personas de 676, deja los datos que esta propuesta usa para dimensionar cada decisión."));
children.push(table(["Dato", "Cifra", "Qué implica para la política"], [
  ["Uso actual de IA", "55,8 sobre 100. 22 % no usó IA en el último mes, 25 % una o dos veces.", "La mayoría arranca desde cero. Una regla simple vale más que una completa."],
  ["Cuenta con la que usan IA", "50 % cuenta personal, 15 % cuenta de la empresa, 17 % declara no usarla.", "Prohibir la cuenta personal sin alternativa deja a la mitad por fuera de la regla el primer día."],
  ["Acceso desde la red", "142 personas con bloqueos totales o parciales a herramientas de IA; 19 % no lo ha intentado.", "La lista de dominios por fase es la primera decisión de TIC."],
  ["Trabajo repetitivo", "141 personas dedican más de 5 horas semanales a tareas de pasos siempre iguales; cerca de 1.900 horas semanales en total.", "Ahí está el retorno de una licencia paga, y el dato para asignarla."],
  ["Registro del trabajo", "104 administrativos registran en el ERP UnoE; 72 en archivos propios de Excel. 79 de 135 comerciales usan el CRM por encima del 75 %.", "Las licencias de nivel 1 y 3 van a estos grupos."],
  ["Informes duplicados", "160 personas reportan informes que también hace otra área.", "Material para pilotos entre áreas: la ruta de piloto de la sección 6."],
  ["Política de IA", "240 la firmaron. 37 % no la firmó o no sabía que existía.", "El Taller 1 abre con la política y cierra con la firma."],
], [2.2, 4.0, 3.8]));

/* 2 · Fases */
children.push(H1("2 · Tres fases de adopción"));
children.push(P("Cada fase define qué herramienta, con qué cuenta, para qué información y qué hace TIC. Las fechas siguen el cronograma del programa."));
children.push(table(["", "Fase 1 · Sandbox de aprendizaje", "Fase 2 · Cuenta corporativa", "Fase 3 · Licencias por actividad"], [
  ["Cuándo", "Octubre a diciembre de 2026 · talleres 1 a 3", "Enero a marzo de 2027 · talleres 4 a 6", "Abril de 2027 en adelante · Banco de Retos y etapa 2"],
  ["Qué se aprende", "Instrucción completa, clasificar tareas, limpiar bases, ordenar flujos", "Trabajar con la información de la empresa: procedimientos, CRM, procesos del área", "Construir mejoras medidas y pilotos"],
  ["Herramienta", "ChatGPT, Gemini y Copilot Chat en versión gratuita", "Copilot Chat con la cuenta de Microsoft 365 como herramienta por defecto", "Licencias pagas según el nivel de la sección 5"],
  ["Cuenta", "Personal, con la opción de entrenamiento apagada", "Corporativa, con protección de datos empresarial", "Corporativa"],
  ["Información permitida", "Material del kit e información propia sin datos de personas ni cifras confidenciales", "Interna sin datos personales. Confidencial solo por categoría autorizada", "Según el alcance de cada licencia"],
  ["Qué hace TIC", "Abre los dominios y confirma por escrito la regla de cuenta personal", "Verifica el plan de Microsoft 365 de cada empresa y resuelve las 20 personas sin correo corporativo", "Asigna licencias con la evidencia del formato «Mi proceso» y aprueba pilotos"],
], [1.6, 2.8, 2.8, 2.8]));
children.push(spacer());
children.push(K("Por qué arrancar con herramientas gratuitas"));
children.push(P("Habilitar licencias pagas para 700 personas antes de que sepan usarlas gasta el presupuesto en quien todavía no lo necesita. El sandbox enseña el método con material preparado y sin riesgo para la información de la empresa. La licencia llega cuando la persona ya midió su tarea y la mejora está demostrada."));
children.push(K("Dos configuraciones que hacen seguro el sandbox"));
children.push(B("En ChatGPT gratuito las conversaciones entrenan el modelo por defecto. Se apaga en Configuración, Controles de datos, «Mejorar el modelo para todos». Es el primer paso del Taller 1, antes de escribir la primera instrucción."));
children.push(B("Nada de la empresa entra con cuenta personal: ni listados, ni cifras, ni nombres. El kit del programa trae los archivos de práctica, construidos sobre los casos reales de las mesas de trabajo con datos ficticios."));

/* 3 · Regla de información */
children.push(H1("3 · Regla de información por tipo de cuenta"));
children.push(P("La regla se define por la información y no por la herramienta. Se propone como anexo 1 de la política y es la tabla que se enseña en el Taller 1 de los dos perfiles."));
children.push(table(["Tipo de información", "Cuenta personal gratuita", "Copilot Chat corporativo", "Licencia paga corporativa"], [
  ["Pública y material de formación: normas publicadas, fichas técnicas públicas, casos del kit", "Sí", "Sí", "Sí"],
  ["Interna sin datos personales: procedimientos, plantillas, actas con cargos, indicadores agregados", "No", "Sí", "Sí"],
  ["Confidencial: precios, márgenes, cartera, nómina, información financiera y estratégica", "No", "Solo las categorías que TIC autorice por escrito", "Sí, dentro del alcance de la licencia"],
  ["Datos de identificación y contacto de clientes, proveedores y colaboradores", "No", "No, salvo anonimizados", "Solo en los sistemas de la empresa: CRM, ERP, Microsoft 365"],
], [3.4, 2.0, 2.3, 2.3]));
children.push(spacer());
children.push(K("Qué significa anonimizar"));
children.push(P("Reemplazar, antes de pegar o cargar, el nombre completo, el número de identificación, el teléfono, el correo, la dirección y el cargo de cualquier persona por una etiqueta: «Cliente A», «Proveedor 2», «Coordinador del área». Toma dos minutos y es la primera práctica del Taller 1."));
children.push(K("Categorías preautorizadas"));
children.push(P("Para que la verificación con TIC del numeral 7.3 sea aplicable a 700 personas, se propone preautorizar: documentos internos sin datos personales en Copilot Chat corporativo; mensajes y correos anonimizados en cualquier herramienta autorizada; indicadores agregados sin nombres. Se consulta a TIC solo para información confidencial y para cargar archivos completos."));

/* 4 · Precisiones */
children.push(H1("4 · Precisiones a la política, numeral por numeral"));
children.push(P("Ocho ajustes de texto. Cada uno conserva la intención del borrador y agrega la parte operativa que faltaba."));
children.push(table(["Numeral", "Texto vigente", "Redacción propuesta"], [
  ["2 · Alcance", "«Aplica a la totalidad del personal de GYJ Ferreterías S.A.»", "«Aplica a la totalidad del personal de las empresas del grupo: GYJ Ferreterías, G&J Soluciones, G&J Ramírez, Almasa, Colmena, Sigmasteel y Almapanel.»"],
  ["4 · Definiciones", "Define «sistema de IA interno» sin ruta para crearlo.", "Agregar «Piloto de IA: solución construida por un equipo interno sobre un proceso propio, con ficha de caso, revisión de TIC y medición, según la ruta del anexo 3.»"],
  ["5.4 · Transparencia", "«Debe ser identificable y, cuando aplique, informado.»", "«En documentos internos no requiere aviso. En comunicaciones a clientes y proveedores requiere revisión humana antes del envío. En decisiones que afecten a personas requiere revisión documentada por el líder del proceso.»"],
  ["7.1 · Se permite", "«Consultar al área de TIC sobre las herramientas autorizadas.»", "«Usar las herramientas del anexo 2 según su nivel y con la cuenta indicada. Una herramienta nueva se solicita a TIC con el formato del anexo 2 y queda autorizada por escrito.»"],
  ["7.2 · Cuentas personales", "«Utilizar cuentas o licencias personales... salvo autorización expresa del área de Sistemas.»", "«Con cuenta personal solo se trabaja información pública y material de formación, con la opción de entrenamiento del proveedor desactivada. Toda información interna o confidencial se trabaja en la cuenta corporativa.»"],
  ["7.3 · Cargue de información", "«Antes de cargar cualquier archivo... verificar... que la plataforma esté autorizada.»", "«Los documentos internos sin datos personales se cargan en la herramienta corporativa sin consulta previa. La información confidencial y los archivos con datos de terceros requieren autorización de TIC por categoría, según el anexo 1.»"],
  ["8 · Dudas y reportes", "«Consultar previamente con el área de Sistemas o con su jefe inmediato.»", "«Consultar al formador interno de su área o al canal [correo de TIC]. Los cargues accidentales se reportan al mismo canal el mismo día.»"],
  ["9 · Líderes", "«Validar los resultados obtenidos con herramientas de IA antes de su uso en los procesos de su equipo.»", "«Cada persona revisa lo que produce con IA antes de usarlo. El líder valida las salidas de riesgo alto: lo que se envía a clientes y proveedores, las cifras financieras y lo que afecta a personas.»"],
  ["10 · Capacitación", "«Jornadas periódicas de capacitación» con cuatro contenidos mínimos.", "«El programa Super Humanos cubre los cuatro contenidos en el Taller 1 de cada perfil. La firma de la política se registra en ese taller. Quien ingrese después recibe la inducción con el mismo contenido.»"],
], [1.7, 3.6, 4.7]));

/* 5 · Licencias */
children.push(H1("5 · Licencias por nivel de actividad"));
children.push(P("Cinco niveles. El criterio de asignación es la actividad medida en el formato «Mi proceso»: horas semanales repetitivas, sistema donde vive la información y tarea mejorada. La jerarquía no asigna licencias."));
children.push(table(["Nivel", "Quién", "Herramienta", "Precio de lista por persona y mes", "Personas estimadas"], [
  ["0 · Base", "Todo el personal del programa", "Copilot Chat con cuenta de Microsoft 365, más gratuitas para información pública", "Incluido en las suscripciones elegibles de Microsoft 365", "700"],
  ["1 · Asistente en las aplicaciones", "Quien arma reportes en Excel, Word y Outlook: contabilidad, SGI, talento humano, líderes de área", "Microsoft 365 Copilot", "USD 30. USD 18 en el plan Business para empresas de hasta 300 usuarios", "60 a 90"],
  ["2 · Documentos y conocimiento", "Gestión documental, archivo, calidad: quien responde preguntas sobre procedimientos", "Copilot sobre SharePoint, o NotebookLM Enterprise", "Dentro del nivel 1, o USD 25", "20 a 40"],
  ["3 · Comercial con CRM", "Asesores con uso del CRM por encima del 75 %", "Copilot con el export del CRM; Agentforce cuando entre Salesforce", "Por definir con Salesforce", "79"],
  ["4 · Constructores", "Formadores internos y equipos del Banco de Retos", "ChatGPT Business o Claude Team, más Copilot Studio o Power Automate para pilotos", "USD 20 a 25 por asiento. Ambos excluyen los datos del entrenamiento", "50 a 60"],
], [1.5, 2.6, 2.5, 2.2, 1.2]));
children.push(spacer());
children.push(box("Sobre los precios", [
  "Son precios de lista en dólares, tomados de guías de terceros con corte de octubre de 2026. Cambian con frecuencia y varían por país. Antes de presupuestar se confirman contra las páginas oficiales de Microsoft, OpenAI, Google y Anthropic.",
]));
children.push(K("Escenario de costo para la fase 3"));
children.push(P("Con 75 personas en el nivel 1 al precio Business, 30 en el nivel 2 cubiertas por el nivel 1 y 55 asientos en el nivel 4 a USD 22, el costo mensual queda cerca de USD 2.600, unos USD 31.000 al año. El nivel 0 cubre a las 700 sin costo adicional. El nivel 3 depende de la decisión de Salesforce."));
children.push(K("Tres observaciones"));
children.push(B("El plan Business de Copilot cuesta USD 18 en lugar de USD 30 para organizaciones de hasta 300 usuarios. Seis de las siete empresas están por debajo de ese tamaño. Si cada empresa licencia por separado, el nivel 1 cuesta 40 % menos."));
children.push(B("El nivel 0 depende de un dato pendiente: qué plan de Microsoft 365 tiene cada empresa y si Copilot Chat con protección empresarial aparece para todas las cuentas, incluidas las 20 personas sin correo corporativo."));
children.push(B("El nivel 4 produce la evidencia. Los formadores internos y los equipos del Banco construyen las soluciones. Un asiento pago desde enero deja medido, antes de abril, cuánto rinde una licencia frente a la herramienta gratuita."));

/* 6 · Ruta de piloto */
children.push(H1("6 · Ruta de piloto"));
children.push(P("Se propone como anexo 3 de la política. Es el camino para que una mejora del formato «Mi proceso» o del Banco de Retos se convierta en un sistema interno autorizado."));
children.push(table(["Paso", "Qué entrega el equipo", "Quién revisa", "Tiempo"], [
  ["1 · Ficha del caso", "Problema, proceso, datos que usa, quién valida, qué indicador cambia y la línea base medida", "Dirección del programa", "Una semana"],
  ["2 · Revisión de TIC", "Clasificación de los datos según el anexo 1, herramienta propuesta, accesos necesarios", "TIC", "Una semana"],
  ["3 · Piloto medido", "Prueba en el proceso real con un grupo acotado, durante cuatro semanas, con el antes y el después", "Líder del proceso", "Cuatro semanas"],
  ["4 · Adopción", "Resultado medido, decisión de escalar, licencia o desarrollo que requiere", "TIC y Gestión Humana", "Comité quincenal"],
], [1.8, 4.4, 2.2, 1.6]));

/* 7 · Dominios */
children.push(H1("7 · Dominios por fase"));
children.push(table(["Fase", "Dominios que se habilitan", "Desde"], [
  ["1 · Sandbox", "chatgpt.com · gemini.google.com · copilot.microsoft.com", "21 de octubre de 2026"],
  ["2 · Cuenta corporativa", "Los de la fase 1 más m365.cloud.microsoft y los de SharePoint y OneDrive del grupo. NotebookLM solo si entra en el nivel 2", "13 de enero de 2027"],
  ["3 · Licencias", "Los que requiera cada licencia aprobada", "Abril de 2027"],
], [2.2, 5.6, 2.2]));

/* 8 · Qué necesita el programa de TIC */
children.push(H1("8 · Lo que el programa necesita de TIC"));
children.push(table(["Qué", "Para qué", "Fecha"], [
  ["Dominios de la fase 1 abiertos en la red de las siete empresas", "Encuentro 1 del Taller 1 administrativo", "Lunes 19 de octubre"],
  ["Confirmación escrita de la regla de cuenta personal de la sección 3", "Enseñarla en el Taller 1 con respaldo", "Lunes 19 de octubre"],
  ["Canal de consulta y reporte: correo o canal de Teams", "Decírselo a la gente en el Taller 1", "Miércoles 21 de octubre"],
  ["Plan de Microsoft 365 por empresa y estado de Copilot Chat", "Dimensionar el nivel 0 y el precio del nivel 1", "Viernes 13 de noviembre"],
  ["Cuentas corporativas para las 20 personas sin correo", "Fase 2 completa", "Martes 1 de diciembre"],
  ["Fecha de entrada de Salesforce", "Definir el nivel 3", "Martes 1 de diciembre"],
  ["Pilotos del nivel 4 aprobados: formadores y equipos del Banco", "Evidencia para la decisión de abril", "Miércoles 13 de enero"],
], [4.0, 3.6, 2.4]));
children.push(spacer());
children.push(box("Cómo seguimos", [
  "Pivot Power presenta esta propuesta a TIC y Gestión Humana en una sesión de 45 minutos. De ahí sale la versión 2 de la política con los tres anexos, y el Taller 1 enseña exactamente esa versión.",
]));

const doc = new Document({
  creator: "Pivot Power", title: "Propuesta para la política de uso de IA · fases, regla de información y licencias",
  styles: { default: { document: { run: { font: BF, size: 21, color: BLACK } } } },
  numbering: { config: [{ reference: "chev", levels: [{ level: 0, format: LevelFormat.BULLET, text: "›", alignment: AlignmentType.LEFT, style: { run: { font: HF, bold: true, color: GOLD, size: 24 }, paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1100, bottom: 1100, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: "SUPER HUMANOS · PROPUESTA PARA LA POLÍTICA DE USO DE IA", font: HF, size: 17, bold: true, color: MUTED, characterSpacing: 30 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9600 }], children: [new TextRun({ text: "Pivot Power · versión para discusión · 15 de octubre de 2026", font: BF, size: 16, color: MUTED }), new TextRun({ text: "\t", font: BF, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 18, bold: true, color: BLACK })] })] }) },
    children,
  }],
});

const out = path.join(__dirname, "out", "Propuesta para la política de uso de IA · fases, información y licencias.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ok", out, buf.length); });
