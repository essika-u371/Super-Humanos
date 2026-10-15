const { fs, path, Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, LevelFormat, AlignmentType, TabStopType, BLACK, YELLOW, GOLD, MUTED, HF, BF, run, P, H1, H2, K, B, bold, table, box, spacer } = require("./docx-brand.js");

/* Fila de ajuste: Qué cambia · Dónde · Por qué */
const adj = (rows) => table(["Qué cambia", "Dónde se cambia", "Por qué · dato"], rows, [4.2, 2.6, 3.2]);

const children = [];

/* Portada */
children.push(new Paragraph({ spacing: { before: 1200, after: 100 }, children: [new TextRun({ text: "SUPER HUMANOS", font: HF, size: 28, bold: true, color: GOLD, characterSpacing: 60 })] }));
children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "AJUSTES A LOS TALLERES", font: HF, size: 72, bold: true, color: BLACK })] }));
children.push(new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: "A PARTIR DEL DIAGNÓSTICO PARCIAL Y DE LA POLÍTICA DE USO DE IA", font: HF, size: 34, bold: true, color: BLACK })] }));
children.push(P("Diagnóstico con corte del 14 de octubre de 2026: 415 personas con índice completo, 62 % de la base. Política de uso de inteligencia artificial y gestión de riesgo de GYJ Ferreterías S.A., código TH-PT-003, versión 1.", { run: { color: MUTED } }));
children.push(P("Pivot Power · documento de trabajo para la dirección del programa y los dos docentes · 15 de octubre de 2026 · fechas según el cronograma ajustado del 15 de octubre", { run: { color: MUTED } }));
children.push(spacer());
children.push(box("Qué hace este documento", [
  "Cruza tres fuentes: los resultados del diagnóstico, la política de IA que firmó la gente y el temario detallado de los doce talleres con su guía de contenido.",
  "Propone los cambios taller por taller, con el lugar exacto del temario o de la guía donde se aplican y el dato que los sustenta.",
  "Separa lo que se cambia hoy de lo que depende de una decisión de G&J. Las decisiones pendientes van en la sección 7.",
  "Con las fechas corridas, el Taller 1 administrativo arranca el miércoles 21 de octubre y el comercial el viernes 23. Los ajustes del Taller 1 entran completos en los dos encuentros, en el kit del participante y en los retos.",
]));

/* 1 · Datos */
children.push(H1("1 · Lo que dicen los datos"));
children.push(P("Nueve hechos del diagnóstico que tocan el contenido. Cada uno trae el taller donde se resuelve."));
children.push(table(["Hecho", "Dato", "Taller que lo toma"], [
  ["El uso de IA es el eje bajo del índice", "55,8 sobre 100, veinticinco puntos por debajo de los otros dos ejes. 22 % no usó IA en el último mes, 25 % una o dos veces. 50 % usa cuenta personal y 15 % cuenta de la empresa.", "Taller 1 de los dos perfiles. Define el ritmo y la regla de cuenta."],
  ["Comercial se traba con el volumen por canal desordenado", "15 cotizaciones por WhatsApp: 27,8 % de acierto. 40 correos del lunes: 46,6 %. Presentar resultados en 10 minutos: 60,9 %.", "T1 y T3 comercial. T2 comercial para el cierre de resultados."],
  ["Administrativo resuelve bien ejecutar y flojea en comunicar y rediseñar", "Ningún ítem bajo 74 %. Los dos más bajos: estructura para presentar resultados, 74,8 %, y el trámite de 30 firmas, 78,4 %.", "T1 y T2 administrativo para comunicar. T4 para rediseñar."],
  ["Los pendientes viven en papel", "67 % en cuaderno o marcando correos, 8 % en la memoria. 33 personas de 417 usan una herramienta compartida: 11 comerciales y 22 administrativas.", "T3 administrativo tal como está. T3 comercial recibe un bloque nuevo."],
  ["Informes duplicados entre áreas", "38 % dice que hace informes que también hace otra área; otro 22 % cree que sí. 57 comerciales y 103 administrativos.", "T4 administrativo y Banco de Retos."],
  ["Rehacer en archivo propio lo que ya está en el sistema", "79 personas lo hacen siempre o casi siempre. 72 administrativos registran su trabajo en archivos propios de Excel.", "T4 de los dos perfiles: la fuente oficial manda."],
  ["El propio trabajo está sin documentar", "48 personas responden que lo que saben no está escrito de forma que alguien pudiera continuarlo.", "T4 de los dos perfiles y reto 19."],
  ["La política de IA está a medias", "240 la firmaron. 119 no sabían que existía y 37 saben que existe y no la han firmado. 37 % en total.", "T1 de los dos perfiles abre con la política y cierra con la firma."],
  ["Las condiciones de conexión limitan la práctica", "142 personas con bloqueos totales o parciales a herramientas de IA. 52 sin computador propio, 31 de ellas comerciales. 71 se conectan desde el celular. 66 con conexión inestable.", "Todos los talleres: práctica con versión para celular y grabación obligatoria."],
], [2.6, 4.4, 3.0]));
children.push(spacer());
children.push(P([bold("Los grupos quedaron parejos. "), run("Con percentiles dentro de cada perfil hay cuatro grupos de 32 a 34 comerciales y de 70 a 72 administrativos. G1 Arranque suma 102 personas: es el grupo que marca el ritmo del Taller 1 y el que más necesita el consultorio. G4 Referentes suma 106 y de ahí salen los 43 candidatos a formador interno.")]));
children.push(P([bold("Lo que la gente pide. "), run("329 personas escribieron qué esperan de la IA: ahorrar tiempo (38 %), organizar pendientes (23 %), cotizaciones y ventas (14 %), automatizar lo repetitivo (11 %). Una respuesta marca el límite que ellos mismos ponen: «No me gusta que la IA reemplace mi pensar, análisis y criterio». Esa frase abre el encuentro 1 de los dos perfiles.")]));

/* 2 · Política */
children.push(H1("2 · Lo que exige la política y dónde toca el temario"));
children.push(P("La política TH-PT-003 fija reglas que el temario asumía de otra forma. Seis de ellas cambian cómo se enseña y con qué se practica."));
children.push(table(["Regla de la política", "Dónde toca el temario", "Ajuste"], [
  ["7.2 · Las cuentas personales de herramientas de IA quedan prohibidas para información de la empresa, salvo autorización expresa de Sistemas.", "Sección 3, «Herramientas del programa»: ChatGPT, Gemini y NotebookLM «gratis con cuenta». El 50 % de quienes respondieron usa cuenta personal.", "La regla del programa pasa a ser: con cuenta personal solo se trabaja con el material del kit, sin datos de la empresa ni de personas. Con información real, solo la herramienta y la cuenta que TIC autorice por escrito."],
  ["7.2 y 7.3 · Nombres, cédulas, teléfonos, correos, direcciones y cargos de clientes, proveedores y colaboradores nunca entran a una herramienta no autorizada. Tampoco listados ni bases.", "Demostración T1 comercial con un WhatsApp real. T2 comercial con un export real de cartera. T3 administrativo con una bandeja real. T4 comercial con la nota de voz de una visita. T5 comercial con un reclamo real.", "Cada demostración usa el archivo del kit, anonimizado, o un archivo propio al que la persona le quitó los datos de identificación. La técnica de anonimizar en dos minutos se enseña en el Taller 1."],
  ["6 y 7.3 · Antes de cargar un archivo o documento a una plataforma de IA hay que verificar con TIC que esté autorizada.", "T4 de los dos perfiles con NotebookLM y procedimientos propios. Reto 17.", "NotebookLM entra solo si TIC lo autoriza. Mientras tanto, el texto del procedimiento, sin datos personales, se pega en la herramienta autorizada."],
  ["5.5 y 6 · Validación humana obligatoria de todo resultado antes de usarlo o comunicarlo. 9 · Los líderes validan los resultados de su equipo antes de usarlos.", "Regla «toda cifra se verifica», ya en el temario. Matriz de responsabilidad del T5.", "La regla se sostiene y se cita con el número de la política. En T5 el líder entra a la matriz como aprobador, porque la política le asigna ese rol."],
  ["5.4 · El uso de IA en contenidos o decisiones que afecten a clientes o colaboradores debe ser identificable y, cuando aplique, informado.", "T5 comercial, respuestas al cliente. T5 administrativo, procesos del área.", "Se agrega la pregunta «¿hay que informar que esto lo apoyó una IA?» al banco de respuestas y a la matriz."],
  ["10 · La capacitación debe incluir principios éticos, riesgos de seguridad y datos, herramientas autorizadas y canales de consulta y reporte.", "Bloque «El marco de uso responsable» del T1.", "El bloque se reorganiza con esos cuatro contenidos, en ese orden. Así el Taller 1 cuenta como la capacitación que la política exige."],
], [3.3, 3.3, 3.4]));
children.push(spacer());
children.push(box("Alcance de la política", [
  "El documento está emitido por GYJ Ferreterías S.A. La dirección del programa confirmó el 14 de octubre que rige para las siete empresas. El Taller 1 la enseña como la regla de G&J en todas.",
]));

/* 3 · Transversales */
children.push(H1("3 · Ajustes que aplican a los doce talleres"));
children.push(K("A · Herramienta y cuenta"));
children.push(P("La sección 3 del temario, «Herramientas del programa», se reescribe. Texto propuesto para el párrafo de apertura y la regla:"));
children.push(box("Texto nuevo · sección 3 del temario", [
  "El programa usa las herramientas que TIC autoriza por escrito, con la cuenta que la política de IA indica. La lista vigente va en el kit del docente y se actualiza cuando TIC la cambie.",
  "Con una cuenta personal solo se trabaja con el material del kit: casos y archivos preparados por el programa, sin datos de la empresa ni de personas. Toda tarea con información real se hace en la herramienta y la cuenta autorizadas.",
  "Antes de pegar cualquier texto o cargar cualquier archivo, la persona le quita nombres, cédulas, teléfonos, correos, direcciones y cargos. Esa es la primera práctica del Taller 1.",
]));
children.push(K("B · Material del kit en lugar de archivos reales"));
children.push(P("Cada taller trae su archivo de práctica preparado por el programa, construido sobre los casos de las mesas, con datos ficticios. Lista mínima:"));
children.push(B("Comercial: un WhatsApp con 15 solicitudes de cotización de 5 a 20 ítems; una bandeja de 40 correos del lunes; un export de cartera con el mismo cliente escrito de tres formas; la nota de voz de una visita; un reclamo por retraso."));
children.push(B("Administrativo: tres archivos de la misma estadística con formatos distintos; una bandeja con derechos de petición y plazos; un procedimiento de vinculación; el flujo del ingreso con 30 firmas; dos informes del mismo mes con cifras distintas."));
children.push(P("En el temario, cada ejercicio que dice «con un archivo propio» o «con una base propia» pasa a decir «con el archivo del kit o con uno propio sin datos de personas»."));
children.push(K("C · Dos ritmos en el mismo encuentro"));
children.push(B("G1 Arranque, 102 personas: el bloque «Primer contacto guiado» del T1 se extiende y el consultorio del Taller 1 se reserva para ellas. El docente las nombra al inicio y las invita en persona."));
children.push(B("G4 Referentes, 106 personas: durante la práctica apoyan a dos compañeros cada una, como ya dice el temario. Los 43 candidatos a formador interno empiezan ahí su rol."));
children.push(K("D · Celular, bloqueos y conexión"));
children.push(B("Cada práctica tiene una versión para celular: el prompt se copia desde el kit en lugar de escribirlo, y la evidencia es una captura."));
children.push(B("52 personas sin computador propio, 31 comerciales, necesitan sala con equipo o una pareja de trabajo. Gestión Humana asigna el lugar antes de cada encuentro 1."));
children.push(B("La grabación en la plataforma deja de ser opcional: 66 personas reportan conexión inestable."));
children.push(B("Hasta que TIC confirme la lista de dominios habilitados, el docente lleva un plan B para los 142 con bloqueos: la demostración se sigue en pantalla compartida y la práctica se hace desde el celular con datos."));
children.push(K("E · El formato «Mi proceso» mide horas semanales"));
children.push(P("El diagnóstico estima 1.900 horas semanales de trabajo repetitivo en 417 personas. Para que esa cifra sirva al cierre, el formato pide minutos por vez y veces por semana, y el reporte mensual suma las horas liberadas. El campo ya existe en el formato; lo que cambia es que el docente lo exige lleno en cada encuentro 2."));

/* 4 · Comercial */
children.push(H1("4 · Ajustes por taller · comercial"));
children.push(H2("Taller 1 · IA para el asesor comercial"));
children.push(P("Encuentro 1 el viernes 23 de octubre, consultorio el miércoles 28, encuentro 2 el viernes 30. Los ajustes entran en los dos encuentros, en el kit y en la evaluación.", { run: { color: MUTED } }));
children.push(adj([
  ["El caso de clase es el de las 15 cotizaciones por WhatsApp: demostración en el encuentro 1 y primera instrucción reutilizable de cada persona en el encuentro 2, antes de pasar a una tarea propia.", "Temario, bloques «Casos del rol comercial» y «Mi primera tarea con IA». Guía, demostración y «Encuentro 2 · aplicación».", "Ítem S_1 con 27,8 % de acierto, el más bajo del diagnóstico. Aparece también en las respuestas abiertas."],
  ["El segundo caso de clase es ordenar los 40 correos del lunes: extraer qué pide cada uno, quién y para cuándo, en una tabla.", "Temario, bloque «Casos del rol comercial»: «ordenar los pendientes del lunes» pasa de mención a ejercicio.", "Ítem S_5 con 46,6 %."],
  ["Apertura del encuentro 1 con la política: los cuatro contenidos que exige el numeral 10, en diez minutos, y la firma pendiente resuelta en el mismo espacio con Gestión Humana.", "Temario, bloque «El marco de uso responsable de G&J». Guía, bloque «Uso responsable».", "63 comerciales de 135 sin firmar o sin saber de la política."],
  ["La regla de cuenta se enseña con el dato: «la mitad de ustedes usa cuenta personal; con ella solo se practica con el material del kit».", "Mismo bloque.", "Política 7.2. 50 % con cuenta personal."],
  ["Anonimizar en dos minutos: antes de pegar el WhatsApp, se reemplazan nombre y teléfono por «Cliente A». Se demuestra y se exige en toda práctica.", "Temario, bloque «Cómo se le habla a una IA», como paso previo. Guía, demostración, paso 2.", "Política 7.3. El mensaje real de la demostración trae nombre y número del cliente."],
  ["Evaluación: las preguntas 6 y 7 se escriben sobre casos de la política: un listado de clientes, la cédula de un comprador, el precio de una cotización, una ficha técnica pública.", "Guía, «Qué debe medir la evaluación», filas 6 y 7.", "Política 7.2 y 7.3."],
]));
children.push(H2("Taller 2 · Inteligencia de mercado con IA"));
children.push(adj([
  ["Se agrega un bloque de cierre: de la tabla al mensaje. Cómo contar los resultados del mes en diez minutos: qué pasó, por qué, qué propongo, qué necesito.", "Temario, temáticas del T2 comercial, nuevo bloque antes de «Compartirlo». Guía, «Lo que no puede faltar».", "Ítem D_5 con 60,9 % en comercial. El administrativo ya lo tiene en su T2."],
  ["El export de cartera de la demostración es el del kit, con clientes ficticios. El ejercicio «cada quien limpia una base propia» pasa a «una base propia sin nombres de clientes o la del kit».", "Guía, demostración y ejercicio del T2 comercial.", "Política 7.3: listados de clientes no entran a herramientas no autorizadas."],
  ["La clasificación de tareas con los seis destinos arranca con las horas semanales que cada persona declaró en el diagnóstico.", "Temario, bloque de los seis destinos.", "52 comerciales dedican más de 5 horas semanales a tareas repetitivas."],
]));
children.push(H2("Taller 3 · Venta consultiva y experiencia del cliente"));
children.push(adj([
  ["Bloque nuevo: del cuaderno al tablero de cotizaciones. Qué convierte una solicitud en pendiente con responsable y fecha, y cómo verla en una herramienta compartida con el líder.", "Temario, temáticas del T3 comercial, después de «la solicitud completa». Guía, «Lo que no puede faltar».", "75 de 133 comerciales llevan pendientes en papel; 11 usan una herramienta compartida. Carolina pidió este taller."],
  ["La bandeja de 40 correos del lunes es el caso de flujo: distinguir trabajo de espera y armar la solicitud completa desde ahí.", "Temario, bloque «recorrido del pedido».", "Ítem S_5 con 46,6 %."],
  ["El caso del pedido retenido por cartera se trabaja completo: causa, fecha nueva y alternativa, con la persona de cartera en el ejercicio.", "Ya está en el temario. Sube a demostración principal.", "Ítem S_4 con 73,7 %."],
  ["Construir un documento entre cuatro personas: coautoría en Word o Excel compartido en lugar de versiones por correo.", "Temario, bloque «la solicitud completa», como herramienta.", "Ítem D_1 con 74,4 %."],
]));
children.push(H2("Taller 4 · Gestión comercial basada en datos"));
children.push(adj([
  ["Nivelación de CRM: antes del mapa de fuentes, veinte minutos con los 25 que se traban o no lo conocen, y los que «podrían enseñarlo» como apoyo.", "Temario, bloque «El CRM como memoria del área».", "25 comerciales se traban o no conocen el CRM; 23 dicen que podrían enseñarlo. Salesforce entra después."],
  ["La nota de voz de la visita se registra sin nombre ni teléfono del cliente; el registro completo va al CRM, que es el sistema autorizado.", "Guía, demostración del T4 comercial.", "Política 7.3."],
  ["Mensaje de apertura: la fuente oficial manda. Si el dato está en el CRM o en la lista de precios, se usa de ahí; el archivo propio es la excepción.", "Temario, bloque «mapa de fuentes».", "79 personas rehacen en archivo propio lo que está en el sistema."],
]));
children.push(H2("Taller 5 · Atención al cliente en canales digitales"));
children.push(adj([
  ["La matriz de responsabilidad incluye al líder como aprobador de las respuestas que salen al cliente, con el número de la política.", "Temario y guía, bloque de la matriz.", "Política 9: los líderes validan los resultados antes de su uso."],
  ["Se agrega al banco de respuestas la pregunta de transparencia: cuándo se informa que una respuesta tuvo apoyo de IA.", "Guía, «Lo que no puede faltar».", "Política 5.4."],
  ["El reclamo de la demostración es el del kit; las respuestas base de cada persona se escriben sin datos del cliente.", "Guía, demostración y ejercicio.", "Política 7.3."],
]));
children.push(H2("Taller 6 · Proyecto comercial con IA"));
children.push(adj([
  ["La línea base de cada proyecto sale del formato «Mi proceso» acumulado: horas semanales antes y después.", "Temario, bloque «cómo medir la línea base».", "1.900 horas semanales estimadas como denominador del programa."],
  ["Los informes que comercial hace y otra área también hace son los primeros candidatos al Banco de Retos.", "Temario, bloque «de la queja al reto».", "57 comerciales reportan informes duplicados."],
]));

/* 5 · Administrativo */
children.push(H1("5 · Ajustes por taller · administrativo"));
children.push(H2("Taller 1 · IA para la gestión administrativa"));
children.push(P("Encuentro 1 el miércoles 21 de octubre, consultorio el viernes 23, encuentro 2 el miércoles 28. Los ajustes entran en los dos encuentros, en el kit y en la evaluación.", { run: { color: MUTED } }));
children.push(adj([
  ["El taller abre por la estructura de comunicar: el resumen de cinco puntos de la demostración se convierte en lo que se presenta en diez minutos: qué pasó, por qué, qué proponemos, qué necesitamos. En el encuentro 2 cada persona lo hace con un documento suyo.", "Temario, bloques «Casos del rol administrativo» y «Mi primera tarea con IA». Guía, demostración y «Encuentro 2».", "Ítem D_5 con 74,8 %, el más bajo del perfil."],
  ["El documento normativo de la demostración es un texto público: la Ley 1581 de 2012 o la Circular 002 de 2024 de la SIC, que la política cita. Así la demostración enseña la norma y evita pegar un documento interno en una herramienta pública.", "Guía, demostración del T1 administrativo, paso 1.", "Política 7.2: información interna fuera de plataformas públicas."],
  ["Las notas del comité para el acta se escriben con cargos en lugar de nombres completos.", "Guía, demostración, paso 4. Temario, bloque «Casos del rol administrativo».", "Política 7.3: el nombre completo y el cargo juntos identifican a la persona."],
  ["Apertura del encuentro 1 con la política, los cuatro contenidos del numeral 10, y la firma pendiente resuelta con Gestión Humana.", "Temario, bloque «El marco de uso responsable».", "114 administrativos de 282 sin firmar o sin saber de la política."],
  ["Niveles de riesgo con los datos del rol: nómina, datos de terceros, información financiera, documentos de vinculación. Cada ejemplo se clasifica con la política en pantalla.", "Ya está en la guía. Se cita el numeral 7.3.", "Política 7.2 y 7.3."],
]));
children.push(H2("Taller 2 · Datos para decidir con IA"));
children.push(adj([
  ["Los tres archivos de la demostración son del kit. El ejercicio «cada quien consolida sus propios archivos» pasa a «archivos propios sin datos de personas o los del kit».", "Guía, demostración y ejercicio.", "Política 7.3. Las estadísticas del área traen nombres de colaboradores y clientes."],
  ["El informe de una página se trabaja como presentación de diez minutos, con la misma estructura del T1.", "Temario, bloque «El informe de una página».", "Ítem D_5 con 74,8 %."],
  ["La consolidación manual se mide en horas: cada persona anota cuánto le toma el reporte del mes antes de tocarlo.", "Temario, bloque «La consolidación manual».", "89 administrativos dedican más de 5 horas semanales a tareas repetitivas."],
]));
children.push(H2("Taller 3 · Productividad digital: de la bandeja al tablero"));
children.push(adj([
  ["Este taller sube de prioridad y se presenta como el taller de seguimiento que pidió Gestión Humana. La demostración termina con el tablero del equipo visible para el líder.", "Temario, bloque «el tablero del equipo». Guía, demostración.", "185 de 282 administrativos llevan pendientes en papel o en el correo; 26 de memoria; 22 con herramienta compartida."],
  ["La bandeja de la demostración es la del kit, con derechos de petición y plazos. Las bandejas propias se trabajan con los nombres reemplazados por área.", "Guía, demostración.", "Política 7.3."],
  ["Construir un documento entre cuatro personas: coautoría en Microsoft 365, con control de versiones, como parte del tablero.", "Temario, bloque «el tablero personal».", "Ítem D_1 con 79,8 %."],
]));
children.push(H2("Taller 4 · Procesos internos con IA"));
children.push(adj([
  ["El caso del ingreso con 30 firmas se trabaja aquí como ejercicio de rediseño: qué firma aporta, cuál se elimina, cuál se vuelve digital. En el T6 se mide.", "Temario, sección 4, tabla de casos: «Administrativo · talleres 4 y 6». Guía, «Lo que no puede faltar».", "Ítem S_5 con 78,4 %, segundo más bajo del perfil."],
  ["El inventario de reportes del área se hace en clase, con la lista de informes que cada persona produce y quién más los produce.", "Temario, bloque «Informes duplicados». Sube de mención a ejercicio.", "103 administrativos reportan informes duplicados; otros 22 % creen que sí."],
  ["Mensaje de apertura: la fuente oficial manda. Qué dato vive en el ERP UnoE y qué dato se rehace en Excel sin necesidad.", "Temario, bloque «mapa de fuentes».", "79 personas rehacen en archivo propio; 72 registran en Excel propio; 104 en UnoE."],
  ["Preguntarle a los procedimientos propios: con NotebookLM solo si TIC lo autoriza. Mientras tanto, el texto del procedimiento sin datos personales se pega en la herramienta autorizada.", "Temario, bloque «preguntarle a los procedimientos». Sección 3.", "Política 6 y 7.3."],
  ["Documentar el propio trabajo: el ejercicio de la guía de una página se conecta con las 48 personas que dicen que lo suyo está sin escribir.", "Temario, bloque «formularios repetitivos» y reto 19.", "Ítem de mentalidad con 48 respuestas en 1 o 2."],
]));
children.push(H2("Taller 5 · Adopción de la IA en el equipo"));
children.push(adj([
  ["La matriz de responsabilidad se arma sobre un proceso del ERP, con el líder como aprobador.", "Temario y guía, bloque de la matriz.", "184 administrativos responden que el CRM no aplica; su sistema es UnoE. Política 9."],
  ["Transparencia: cuándo se informa a un colaborador o a un tercero que un resultado tuvo apoyo de IA.", "Guía, «Lo que no puede faltar».", "Política 5.4."],
  ["Cómo enseñarle a un compañero en cinco minutos: lo hacen los candidatos a formador interno con alguien de G1.", "Temario, bloque «cómo enseñarle a un compañero».", "27 candidatos administrativos; 70 personas en G1."],
]));
children.push(H2("Taller 6 · Proyecto de mejora con IA"));
children.push(adj([
  ["La línea base sale del formato «Mi proceso» acumulado, en horas semanales. El trámite de 30 firmas, ya rediseñado en el T4, se mide aquí.", "Temario, bloque «cómo medir la línea base».", "1.900 horas semanales estimadas."],
  ["Los informes duplicados entre áreas son los primeros proyectos del Banco de Retos.", "Temario, bloque «de la queja al reto».", "38 % afirmativo más 22 % «creo que sí»."],
]));

/* 6 · Retos */
children.push(H1("6 · Retos de los jueves que cambian"));
children.push(table(["Reto", "Como está", "Como queda", "Por qué"], [
  ["5 · 22 de octubre", "Pídele a la IA un resumen de la política de IA de G&J y verifica que sea fiel.", "Lee la política y escribe tres datos de tu trabajo que nunca pegarías en una IA. Después pregúntale a la IA qué datos nunca debería recibir y compara las dos listas.", "La política es un documento interno. Pegarla en una herramienta pública con cuenta personal contradice su numeral 7.2."],
  ["10 · 26 de noviembre", "Limpia con IA una lista con nombres repetidos o mal escritos.", "Limpia la lista del kit, o una propia a la que le quitaste nombres y contactos.", "Política 7.3."],
  ["17 · 28 de enero", "Dale a la IA un documento oficial de tu área y hazle tres preguntas.", "Dale a la IA un documento oficial de tu área sin datos de personas, en la herramienta autorizada, y hazle tres preguntas.", "Política 6 y 7.3."],
  ["22 · 4 de marzo", "Prepara con IA el acta de una reunión con decisiones, responsables y fechas.", "Igual, con los responsables escritos por cargo.", "Política 7.3."],
], [1.6, 2.6, 3.4, 2.4]));

/* 7 · Decisiones */
children.push(H1("7 · Decisiones pendientes"));
children.push(P("Cinco decisiones que condicionan los ajustes. Las dos primeras tienen fecha de esta semana."));
children.push(table(["Decisión", "Quién", "Para cuándo", "Qué pasa si no se toma"], [
  ["Lista escrita de herramientas y cuentas autorizadas: qué herramienta, con qué cuenta, desde qué red. Si NotebookLM entra o no.", "TIC, con Gestión Humana", "21 de octubre, encuentro 1 administrativo", "El programa sigue con cuenta personal y material del kit en todas las prácticas con información real."],
  ["Autorización de Sistemas para que el programa use cuentas personales con el material del kit, por escrito.", "TIC", "20 de octubre", "El docente lo pide en voz alta y queda sin respaldo."],
  ["Reemplazo del reto 5 del 22 de octubre.", "Jessica Peña", "17 de octubre", "Sale el reto que pide pegar la política en una IA pública."],
  ["Comunicar los grupos G1 a G4 ahora o al cerrar la convocatoria.", "Jessica Peña", "17 de octubre", "La asignación cambia cuando entren las 259 pendientes."],
  ["El taller de seguimiento que pidió Carolina: queda cubierto por el T3 administrativo y el bloque nuevo del T3 comercial, o se hace aparte.", "Jessica Peña con Carolina Rivera", "Antes del 6 de noviembre", "Se programa un taller adicional sin presupuesto."],
], [3.4, 1.8, 2.0, 2.8]));

/* 8 · Ediciones */
children.push(H1("8 · Qué se edita en cada documento"));
children.push(P("Lista de ediciones para aplicar una vez tomadas las decisiones de la sección 7. Cada documento conserva su nombre; la versión anterior pasa a «09 · Versiones anteriores» con fecha y la palabra «superada»."));
children.push(table(["Documento", "Ediciones"], [
  ["Temario detallado de los 12 talleres", "Sección 3 reescrita. Sección 4: el caso de 30 firmas pasa a «talleres 4 y 6». T1, T2, T3, T4 y T5 de los dos perfiles con los bloques de las secciones 4 y 5 de este documento. Sección 7 con los cuatro retos cambiados. Todo ejercicio «con archivo propio» con la coletilla «sin datos de personas o con el del kit»."],
  ["Guía de contenido por taller", "Fechas del T1: 23 y 30 de octubre comercial, 21 y 28 administrativo. T2 comercial: 13 y 20 de noviembre. T1 de los dos perfiles: demostración con el archivo del kit y el paso de anonimizar; evaluación con las preguntas 6 y 7 sobre la política. T2 a T5: «Lo que no puede faltar» con los puntos nuevos. Sección 6, estándares: regla de cuenta y de datos."],
  ["Kit del docente", "Anexo con la política completa, la lista de datos que nunca entran a una IA y la lista de herramientas autorizadas cuando TIC la entregue. El dato de cuenta personal del diagnóstico. Los archivos del kit de la sección 3B."],
  ["Guía de casos de G&J", "Cada caso que trae datos de personas lleva la marca «usar con el archivo del kit»."],
  ["Cronograma", "La versión ajustada del 15 de octubre reemplaza a la de Drive; la anterior pasa a 09 como superada. En la hoja Jueves de Superpoderes, los retos 5, 10, 17 y 22 con el texto nuevo. En Decisiones y Roles, los retos terminan el 29 de abril y el rol instruccional lo asume la dirección."],
  ["Manual de bienvenida para participantes", "Un párrafo con la regla de cuenta y de datos, en las palabras de la política."],
], [3.0, 7.0]));
children.push(spacer());
children.push(box("Advertencia sobre los datos", [
  "Todas las cifras corresponden al 62 % que respondió, sesgado hacia las empresas con seguimiento cercano y las personas con computador y conexión estables. Al cerrar la convocatoria, los ítems bajos del perfil comercial se mantienen como insumo; los cortes de los grupos se recalculan.",
]));

const doc = new Document({
  creator: "Pivot Power", title: "Ajustes a los talleres a partir del diagnóstico y la política de IA",
  styles: { default: { document: { run: { font: BF, size: 21, color: BLACK } } } },
  numbering: { config: [{ reference: "chev", levels: [{ level: 0, format: LevelFormat.BULLET, text: "›", alignment: AlignmentType.LEFT, style: { run: { font: HF, bold: true, color: GOLD, size: 24 }, paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1100, bottom: 1100, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: "SUPER HUMANOS · AJUSTES A LOS TALLERES · DIAGNÓSTICO Y POLÍTICA DE IA", font: HF, size: 17, bold: true, color: MUTED, characterSpacing: 30 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9600 }], children: [new TextRun({ text: "Pivot Power · documento de trabajo · 15 de octubre de 2026", font: BF, size: 16, color: MUTED }), new TextRun({ text: "\t", font: BF, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 18, bold: true, color: BLACK })] })] }) },
    children,
  }],
});

const out = path.join(__dirname, "out", "Ajustes a los talleres · diagnóstico y política de IA.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ok", out, buf.length); });
