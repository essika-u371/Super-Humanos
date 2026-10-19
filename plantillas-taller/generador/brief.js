const { fs, path, Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, LevelFormat, AlignmentType, TabStopType, BLACK, YELLOW, GOLD, MUTED, HF, BF, run, P, H1, H2, K, B, bold, table, box, spacer } = require("./docx-brand.js");

const PERFILES = {
  comercial: {
    nombre: "comercial", NOMBRE: "COMERCIAL", docente: "Mateo Peralta", DOCENTE: "MATEO PERALTA",
    horario: "viernes de 7:30 a 9:00 a.m.", ventana: "entre los días 1 y 20 de cada mes",
    talleres: [
      ["T1", "IA para el asesor comercial", "IA generativa, prompting y uso responsable", "Aprendizaje continuo", "La IA no sabe nada de tu cliente (todavía)", "5 a 8"],
      ["T2", "Inteligencia de mercado con IA", "Qué hace la IA y qué hago yo: tareas y datos", "Dato compartido", "Tu cartera tiene cosas que contarte", "9 a 12"],
      ["T3", "Venta consultiva y experiencia del cliente", "Flujos de trabajo con IA", "Recorrido completo", "El pedido no se demora trabajando, se demora esperando", "13 a 16"],
      ["T4", "Gestión comercial basada en datos", "Contexto antes que modelo", "Conocimiento documentado", "Lo que vive en tu cabeza no le sirve a nadie más", "17 a 20"],
      ["T5", "Atención al cliente con IA en canales digitales", "Humano + IA: verificación, responsabilidad y adopción", "Colaboración entre áreas", "La IA redacta, tú respondes", "21 a 24"],
      ["T6", "Proyecto comercial con IA", "Proyecto de mejora: medir, rediseñar y demostrar", "Prevención", "Mide, mejora, demuestra", "25 a 28"],
    ],
    fechas: [
      ["T1", "Vie 23 oct", "Mié 28 oct", "Vie 30 oct", "Vie 6 nov"],
      ["T2", "Vie 13 nov", "Mié 18 nov", "Vie 20 nov", "Vie 27 nov"],
      ["T3", "Vie 4 dic", "Mié 9 dic", "Vie 11 dic", "Vie 18 dic"],
      ["T4", "Mié 13 ene", "Vie 15 ene", "Mié 20 ene", "Mié 27 ene"],
      ["T5", "Vie 5 feb", "Mié 10 feb", "Vie 12 feb", "Vie 19 feb"],
      ["T6", "Vie 5 mar", "Mié 10 mar", "Vie 12 mar", "Vie 19 mar"],
    ],
    notasFechas: [
      "El Taller 1 y el Taller 2 quedaron fuera de la ventana comercial del día 20. La dirección del programa gestiona la excepción con Carolina Rivera; las fechas se confirman por correo antes de enviar las invitaciones.",
      "En enero los dos encuentros caen en miércoles y el consultorio en viernes, porque el único viernes antes del día 20 es el 15.",
    ],
    entregas: [
      ["T1", "Lun 19 oct", "Mié 21 oct", "Jue 22 oct", "Vie 23 oct"],
      ["T2", "Vie 16 oct", "Vie 23 oct", "Vie 30 oct", "Vie 13 nov"],
      ["T3", "Vie 13 nov", "Vie 20 nov", "Vie 27 nov", "Vie 4 dic"],
      ["T4", "Mié 23 dic", "Mié 23 dic", "Mié 6 ene", "Mié 13 ene"],
      ["T5", "Vie 15 ene", "Vie 22 ene", "Vie 29 ene", "Vie 5 feb"],
      ["T6", "Vie 12 feb", "Vie 19 feb", "Vie 26 feb", "Vie 5 mar"],
    ],
    notaEntregas: "El Taller 1 y el Taller 2 se preparan en paralelo: el guion del T2 vence el mismo viernes 16 y sus materiales el 23. Las fechas del T1 las fijó la dirección del programa para esta semana; de T2 en adelante rigen las de la hoja «Preparación de materiales» del cronograma.",
    diag: [
      "Más de la mitad del grupo nunca ha abierto una herramienta de IA o la usó una o dos veces en el último mes. El Taller 1 avanza al ritmo del más lento, con una sola herramienta y sin jerga.",
      "31 comerciales trabajan sin computador propio y varios se conectan desde el celular. Cada práctica tiene una versión para celular: el prompt se copia desde el kit y la evidencia es una captura.",
      "25 comerciales se traban con el CRM o no lo conocen. El Taller 4 abre con veinte minutos de nivelación antes del mapa de fuentes.",
      "52 comerciales dedican más de 5 horas semanales a tareas repetitivas. La clasificación de tareas del Taller 2 arranca con las horas que cada persona declaró.",
      "La bandeja de 40 correos del lunes es el caso de flujo del Taller 3: distinguir trabajo de espera y armar la solicitud completa desde ahí.",
      "106 referentes del grupo G4 apoyan a dos compañeros cada uno durante la práctica. Los 43 candidatos a formador interno empiezan ahí su rol.",
    ],
    seccionAjustes: "sección 4",
  },
  administrativo: {
    nombre: "administrativo", NOMBRE: "ADMINISTRATIVO", docente: "José Cormane", DOCENTE: "JOSÉ CORMANE",
    horario: "miércoles de 8:00 a 9:30 a.m.", ventana: "entre los días 10 y 25 de cada mes",
    talleres: [
      ["T1", "IA para la gestión administrativa", "IA generativa, prompting y uso responsable", "Aprendizaje continuo", "La IA no consulta tus sistemas (todavía)", "5 a 8"],
      ["T2", "Datos para decidir con IA", "Qué hace la IA y qué hago yo: tareas y datos", "Dato compartido", "Tres días armando el reporte, cero analizándolo", "9 a 12"],
      ["T3", "Productividad digital con IA: de la bandeja al tablero", "Flujos de trabajo con IA", "Recorrido completo", "De la bandeja al tablero", "13 a 16"],
      ["T4", "Procesos internos con IA", "Contexto antes que modelo", "Conocimiento documentado", "Pregúntale al procedimiento, no a tu memoria", "17 a 20"],
      ["T5", "Adopción de la IA en el equipo", "Humano + IA: verificación, responsabilidad y adopción", "Colaboración entre áreas", "La IA prepara, tú decides", "21 a 24"],
      ["T6", "Proyecto de mejora con IA", "Proyecto de mejora: medir, rediseñar y demostrar", "Prevención", "Mide, mejora, demuestra", "25 a 28"],
    ],
    fechas: [
      ["T1", "Mié 21 oct", "Vie 23 oct", "Mié 28 oct", "Mié 4 nov"],
      ["T2", "Mié 11 nov", "Mar 17 nov", "Mié 18 nov", "Mié 25 nov"],
      ["T3", "Jue 10 dic", "Lun 14 dic", "Mié 16 dic", "Vie 18 dic"],
      ["T4", "Jue 14 ene", "Lun 18 ene", "Jue 21 ene", "Jue 28 ene"],
      ["T5", "Mié 10 feb", "Lun 15 feb", "Mié 17 feb", "Mié 24 feb"],
      ["T6", "Mié 10 mar", "Lun 15 mar", "Mié 17 mar", "Mié 24 mar"],
    ],
    notasFechas: [
      "El encuentro 2 del Taller 1 cae el 28 de octubre, tres días fuera de la ventana administrativa del día 25. La dirección del programa gestiona la excepción con Carolina Rivera.",
      "El consultorio del Taller 2 es un martes porque el lunes 16 de noviembre es festivo. El Taller 3 abre en jueves 10 de diciembre y el Taller 4 en jueves 14 y 21 de enero, después del regreso de vacaciones.",
    ],
    entregas: [
      ["T1", "Vie 16 oct", "Lun 19 oct", "Mar 20 oct", "Mié 21 oct"],
      ["T2", "Mié 21 oct", "Mié 28 oct", "Mié 4 nov", "Mié 11 nov"],
      ["T3", "Jue 19 nov", "Jue 26 nov", "Jue 3 dic", "Jue 10 dic"],
      ["T4", "Mié 16 dic", "Mié 23 dic", "Mié 6 ene", "Jue 14 ene"],
      ["T5", "Mié 20 ene", "Mié 27 ene", "Mié 3 feb", "Mié 10 feb"],
      ["T6", "Mié 17 feb", "Mié 24 feb", "Mié 3 mar", "Mié 10 mar"],
    ],
    notaEntregas: "El Taller 1 arranca en seis días. Sus fechas de entrega las fijó la dirección del programa para esta semana; de T2 en adelante rigen las de la hoja «Preparación de materiales» del cronograma.",
    diag: [
      "Más de la mitad del grupo nunca ha abierto una herramienta de IA o la usó una o dos veces en el último mes. El Taller 1 avanza al ritmo del más lento, con una sola herramienta y sin jerga.",
      "89 administrativos dedican más de 5 horas semanales a tareas repetitivas. En el Taller 2 cada persona anota cuánto le toma el reporte del mes antes de tocarlo.",
      "103 administrativos hacen informes que también hace otra área, o creen que sí. El Taller 4 levanta el inventario de reportes del área.",
      "48 personas dicen que su trabajo está sin escribir. El ejercicio de la guía de una página del Taller 4 y el reto 19 se conectan con ellas.",
      "La estructura de comunicar abre el Taller 1: el resumen de cinco puntos se convierte en lo que se presenta en diez minutos: qué pasó, por qué, qué proponemos, qué necesitamos.",
      "106 referentes del grupo G4 apoyan a dos compañeros cada uno durante la práctica. Los 43 candidatos a formador interno empiezan ahí su rol.",
    ],
    seccionAjustes: "sección 5",
  },
};

function build(p) {
  const children = [];
  /* Portada */
  children.push(new Paragraph({ spacing: { before: 1200, after: 100 }, children: [new TextRun({ text: "SUPER HUMANOS", font: HF, size: 28, bold: true, color: GOLD, characterSpacing: 60 })] }));
  children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "PAQUETE DEL DOCENTE", font: HF, size: 72, bold: true, color: BLACK })] }));
  children.push(new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: `PERFIL ${p.NOMBRE} · ${p.DOCENTE}`, font: HF, size: 34, bold: true, color: BLACK })] }));
  children.push(P("Pivot Power · 15 de octubre de 2026 · fechas según el cronograma ajustado del 15 de octubre", { run: { color: MUTED } }));
  children.push(spacer());
  children.push(box("Para qué sirve este documento", [
    `Abre el paquete con el que preparas los seis talleres del perfil ${p.nombre}. Trae tus fechas, lo que entregas y cuándo, cómo se arma cada encuentro, las reglas que condicionan el contenido y la lista de archivos de la carpeta. Cuando dos documentos difieran en una fecha, vale la de este documento y la del cronograma.`,
  ]));

  /* 1 · Fechas */
  children.push(H1("1 · Tus fechas"));
  children.push(P(`Encuentros ${p.horario} por Teams. Consultorio de 4:00 a 5:00 p.m. el día indicado. La evaluación corta de diez preguntas abre al cerrar el encuentro 2 y cierra en la fecha de la última columna. La empresa programa al perfil ${p.nombre} ${p.ventana}.`));
  children.push(table(["Taller", "Encuentro 1", "Consultorio", "Encuentro 2", "Cierre de la evaluación"], p.fechas, [1.2, 2, 2, 2, 2.4]));
  p.notasFechas.forEach((n) => children.push(B(n)));

  /* 2 · Entregas */
  children.push(H1("2 · Qué entregas y cuándo"));
  children.push(P("Por cada taller entregas cinco piezas. La líder instruccional las revisa y aprueba; después pasan al kit del participante y a ecert."));
  children.push(B([bold("Presentación del encuentro 1"), run(", de 18 a 24 diapositivas, sobre la plantilla del encuentro 1.")]));
  children.push(B([bold("Presentación del encuentro 2"), run(", de 10 a 14 diapositivas, sobre la plantilla del encuentro 2.")]));
  children.push(B([bold("Guion de la demostración en vivo"), run(", paso a paso, con el caso de G&J elegido de la guía de casos y el archivo del kit que vas a usar en pantalla.")]));
  children.push(B([bold("Ejercicio de práctica guiada"), run(" con sus instrucciones, el prompt de partida y la versión para celular.")]));
  children.push(B([bold("Tres o cuatro prompts para el kit del participante"), run(" y las diez preguntas de la evaluación con su clave.")]));
  children.push(spacer());
  children.push(K("Fechas de entrega por taller"));
  children.push(table(["Taller", "Guion de la sesión", "Materiales completos", "Prueba técnica", "Encuentro 1"], p.entregas, [1.2, 2.2, 2.4, 2, 2]));
  children.push(P(p.notaEntregas, { run: { color: MUTED } }));
  children.push(P("El guion incluye el caso de la demostración, la lista de archivos del kit y el prompt de partida. Los materiales completos son las dos presentaciones, el ejercicio, los prompts y la evaluación. En la prueba técnica recorres la demostración completa en Teams con la herramienta autorizada, desde la red de la empresa."));

  /* 3 · Estructura */
  children.push(H1("3 · Cómo se arma cada encuentro"));
  children.push(K("Encuentro 1 · 90 minutos · concepto, demostración y práctica"));
  children.push(table(["Sección", "Minutos", "Diapositivas", "Qué contiene"], [
    ["Apertura", "0 – 5", "2", "Bienvenida, pregunta rápida al chat y recordatorio de puntos."],
    ["El hábito del mes", "5 – 15", "3 a 4", "El hábito contado con una historia real de G&J. Termina con una pregunta abierta al grupo."],
    ["El concepto", "15 – 45", "8 a 12", "La idea del mes: qué es, cómo funciona, dónde se aplica en el rol. Una idea por diapositiva."],
    ["La demostración", "45 – 65", "1 o 2", "Pantalla compartida con la herramienta. Un caso de la guía de casos, con los errores incluidos."],
    ["La práctica", "65 – 85", "2", "Instrucciones del ejercicio y el prompt de partida. El resto es trabajo del grupo, con los referentes apoyando."],
    ["El cierre", "85 – 90", "2", "El reto de la semana, el formato «Mi proceso» y la fecha del consultorio."],
  ], [2, 1.2, 1.4, 5.4]));
  children.push(spacer());
  children.push(K("Encuentro 2 · 90 minutos · aplicación a la tarea propia"));
  children.push(table(["Sección", "Minutos", "Diapositivas", "Qué contiene"], [
    ["Repaso", "0 – 10", "3", "Qué salió del reto de la semana y del consultorio. Dos o tres ejemplos del grupo."],
    ["Aplicación", "10 – 45", "3 a 4", "Instrucciones del trabajo con el formato «Mi proceso» y una plantilla llena como ejemplo."],
    ["Casos compartidos", "45 – 70", "1", "Tres o cuatro personas muestran su pantalla. Tú retroalimentas."],
    ["Cierre", "70 – 90", "3", "Puntos del mes, evaluación abierta, encuesta y anuncio del siguiente taller."],
  ], [2, 1.2, 1.4, 5.4]));
  children.push(spacer());
  children.push(K("Consultorio · 60 minutos"));
  children.push(P("Espacio abierto entre los dos encuentros. La gente entra con su caso, pregunta y sale cuando tiene la respuesta. Sin diapositivas; prepara dos o tres casos frecuentes para abrir la conversación si nadie llega con el suyo."));
  children.push(K("Las siete reglas del contenido"));
  ["Una idea por diapositiva. Si hay dos, son dos diapositivas.",
   "Cada concepto aterriza en un caso de G&J de la guía de casos, con su dato.",
   "La demostración se hace en vivo, con la herramienta en pantalla y los errores incluidos.",
   "Todo se puede hacer con herramientas gratuitas. Lo que requiere licencia se menciona como horizonte.",
   "Cada taller dice en voz alta qué se verifica y qué se conserva como decisión humana.",
   "El número de ahorro lo pone cada persona al medir su tarea.",
   "Lenguaje llano: «instrucción» en lugar de «prompt engineering», «clase» en lugar de «sesión sincrónica».",
  ].forEach((t) => children.push(B(t)));

  /* 4 · Talleres */
  children.push(H1(`4 · Los seis talleres del perfil ${p.nombre}`));
  children.push(P("El temario detallado trae cada taller completo: objetivo, competencias, temáticas, casos, herramientas, ejercicios, evaluación y errores que vigilar. La guía de contenido dice cómo se arma: bloques, diapositivas y lo que no puede faltar de T2 a T6. Esta tabla resume el mapa y conecta cada taller con su sesión de Teams y con los retos de los jueves que lo refuerzan."));
  children.push(table(["Taller", "Título · concepto del mes", "Hábito de apertura", "Nombre de la sesión en Teams", "Retos"], p.talleres.map((t) => [t[0], `${t[1]}. ${t[2]}`, t[3], `Super Humanos · Taller ${t[0].slice(1)} · ${t[4]}`, t[5]]), [0.9, 3.4, 1.8, 3.1, 0.8]));
  children.push(P("Los retos de los jueves salen cada jueves a las 7:00 a.m. para los dos perfiles y se suben hasta el miércoles siguiente. La lista completa está en la hoja «Jueves de Superpoderes» del cronograma. El encuentro 2 abre con lo que salió del reto de la semana.", { run: { color: MUTED } }));

  /* 5 · Reglas */
  children.push(H1("5 · Reglas que condicionan el contenido"));
  children.push(K("Política de uso de IA de G&J"));
  children.push(P("El Taller 1 enseña la política como la regla de G&J en las siete empresas. Tres puntos entran en todos los talleres:"));
  children.push(B([bold("Una sola herramienta, la autorizada. "), run("TIC define por escrito qué herramientas se usan, con qué cuenta y desde qué red. Mientras llega la lista, el programa trabaja con las herramientas gratuitas de la fase 1 de la propuesta de política: chatgpt.com, gemini.google.com y copilot.microsoft.com. Las presentaciones mencionan «la herramienta autorizada» y se ajustan cuando TIC responda.")]));
  children.push(B([bold("Ningún dato personal de clientes ni de colaboradores entra a una herramienta de IA. "), run("Cada demostración y cada práctica usa el archivo del kit o un archivo propio al que la persona le quitó nombres, cédulas, correos, teléfonos y montos. El paso de anonimizar se muestra en pantalla en el Taller 1 y se repite en cada taller.")]));
  children.push(B([bold("Toda cifra se verifica. "), run("Cada demostración termina diciendo en voz alta qué se verifica contra la fuente antes de usar el resultado. Los líderes validan los resultados de su equipo.")]));
  children.push(P("El numeral 10 de la política exige cuatro contenidos de capacitación. Los cuatro van en la apertura del encuentro 1 del Taller 1, en diez minutos, y la firma pendiente se resuelve en el mismo espacio con Gestión Humana."));
  children.push(K("Lo que dice el diagnóstico de tu perfil"));
  children.push(P(`Resultados parciales con corte del 14 de octubre, 415 personas con índice completo. Los ajustes taller por taller están en «Ajustes a los talleres», ${p.seccionAjustes}. Los que más pesan en tu preparación:`));
  p.diag.forEach((t) => children.push(B(t)));
  children.push(box("Los resultados individuales del diagnóstico son de uso interno", [
    "El paquete trae los resultados agregados por perfil. Los nombres, las empresas y los puntajes de cada persona quedan con la dirección del programa. Cuando necesites un dato por grupo o por sede, pídelo a Jessica Peña.",
  ]));

  /* 6 · Plantillas */
  children.push(H1("6 · Plantillas y marca"));
  children.push(P("La carpeta «Plantillas» trae cuatro archivos de PowerPoint en formato 16:9 con el sistema de diseño Campus Pivot."));
  children.push(table(["Archivo", "Para qué lo usas"], [
    ["Plantilla taller · Encuentro 1 · concepto y práctica", "Presentación del encuentro 1. Trae las secciones de la guía en orden, con diapositivas de ejemplo para el hábito, el concepto, la demostración, la práctica y el cierre."],
    ["Plantilla taller · Encuentro 2 · aplicación", "Presentación del encuentro 2: repaso, aplicación con «Mi proceso», casos compartidos y cierre."],
    ["Plantilla taller · Consultorio", "Una portada y dos o tres diapositivas de apoyo para abrir el consultorio."],
    ["Plantilla base Super Humanos · estilos de diapositiva", "Catálogo de estilos: portadas, separadores, texto, tarjetas, bloques de prompt, cifras grandes, íconos. De aquí copias el estilo que necesites."],
  ], [3.6, 6.4]));
  children.push(spacer());
  children.push(K("Reglas de marca"));
  ["Negro 0D0D0C y amarillo FCC800. Títulos en Barlow Condensed en mayúsculas; cuerpo en Atkinson Hyperlegible Next. Las dos fuentes van instaladas con la plantilla.",
   "Sobre amarillo el texto siempre es negro. El amarillo como color de texto solo va sobre negro; sobre blanco se usa el dorado 8A6B00.",
   "Los bloques de prompt van en el estilo de la plantilla base, con el texto completo del prompt para que la persona lo copie desde el kit.",
   "Las capturas de pantalla de la herramienta llevan el nombre de la herramienta y la fecha; los datos que aparecen vienen del archivo del kit.",
  ].forEach((t) => children.push(B(t)));

  /* 7 · Archivos */
  children.push(H1("7 · Los archivos de la carpeta"));
  children.push(table(["Archivo", "Qué es y cuándo lo usas"], [
    [`Paquete del docente · perfil ${p.nombre}`, "Este documento."],
    ["Cronograma Super Humanos · etapa 1", "Todas las fechas del programa: encuentros, retos, preparación de materiales, decisiones y pendientes. Vale la versión del 15 de octubre."],
    [`Sesiones de Teams · perfil ${p.nombre}`, "Tus 18 reuniones con nombre y descripción, tal como las recibe la gente en la invitación. El nombre de la sesión abre la portada de cada presentación."],
    ["Temario detallado de los 12 talleres", "Contenido completo de cada taller. Lee los seis de tu perfil y los seis del otro, para que los dos programas se hablen."],
    ["Guía de contenido por taller", "Bloques, minutos y diapositivas de cada encuentro; mapa del Taller 1; lo que no puede faltar de T2 a T6; checklist de entrega. Las fechas que trae están superadas por este documento."],
    ["Guía de casos de G&J", "Los casos reales de las mesas de trabajo, con su dato. Cada demostración toma uno de aquí."],
    ["Kit del docente", "Formato «Mi proceso», estructura de la evaluación, prompts base y piezas que se repiten en todos los talleres."],
    ["Manual de puntajes y reconocimiento", "Cómo suman puntos la asistencia, los retos, la evaluación y «Mi proceso». Lo que anuncias en la apertura y en el cierre."],
    ["Ajustes a los talleres · diagnóstico y política de IA", "Lo que cambia en cada taller a partir del diagnóstico y de la política. Sección 3 para todos; " + p.seccionAjustes + " para tu perfil; sección 6 para los retos."],
    ["Política de uso de IA de G&J · TH-PT-003", "El documento que la gente firmó. Base del Taller 1."],
    ["Propuesta para la política de uso de IA", "Lo que Pivot Power le propuso a TIC: tres fases de herramientas, regla de información por cuenta y licencias por nivel. Contexto para saber con qué herramientas cuenta cada fase."],
    ["Plantillas", "Los cuatro archivos de PowerPoint de la sección 6."],
  ], [3.6, 6.4]));

  /* 8 · Canales */
  children.push(H1("8 · Con quién resuelves qué"));
  children.push(table(["Tema", "Persona"], [
    ["Revisión y aprobación de guiones, presentaciones y evaluaciones. Dudas de contenido y de metodología.", "Jessica Peña · coordinación metodológica · Pivot Power"],
    ["Cursos en ecert, carga de materiales y evaluaciones, grabaciones.", "Carlos Bravo · ecert"],
    ["Reuniones de Teams, invitaciones, listas de asistencia y premios.", "Gestión Humana de G&J"],
    ["Herramientas autorizadas, cuentas y acceso desde la red de la empresa.", "TIC de G&J, a través de la dirección del programa"],
    ["Dudas de los participantes entre encuentros.", "superhumanos@gyj.com.co · respuesta semanal, a más tardar el viernes"],
  ], [5.4, 4.6]));
  children.push(spacer());
  children.push(box("Primer paso", [
    `Lee el Taller 1 de tu perfil en el temario y en la guía de contenido, elige el caso de la demostración en la guía de casos y envía el guion a Jessica Peña el ${p.entregas[0][1].replace("Vie ", "viernes ").replace("Lun ", "lunes ")} de octubre. La prueba técnica del Taller 1 se agenda en Teams en la fecha de la tabla de la sección 2.`,
  ]));

  const doc = new Document({
    creator: "Pivot Power", title: `Paquete del docente · perfil ${p.nombre}`,
    styles: { default: { document: { run: { font: BF, size: 21, color: BLACK } } } },
    numbering: { config: [{ reference: "chev", levels: [{ level: 0, format: LevelFormat.BULLET, text: "›", alignment: AlignmentType.LEFT, style: { run: { font: HF, bold: true, color: GOLD, size: 24 }, paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
    sections: [{
      properties: { page: { margin: { top: 1100, bottom: 1100, left: 1134, right: 1134 } } },
      headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: `SUPER HUMANOS · PAQUETE DEL DOCENTE · PERFIL ${p.NOMBRE}`, font: HF, size: 17, bold: true, color: MUTED, characterSpacing: 30 })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9600 }], children: [new TextRun({ text: "Pivot Power · 15 de octubre de 2026", font: BF, size: 16, color: MUTED }), new TextRun({ text: "\t", font: BF, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 18, bold: true, color: BLACK })] })] }) },
      children,
    }],
  });
  const out = path.join(__dirname, "out", `Paquete del docente · perfil ${p.nombre}.docx`);
  return Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ok", out, buf.length); });
}

Promise.all(Object.values(PERFILES).map(build));
