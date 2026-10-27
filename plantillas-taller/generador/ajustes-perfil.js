const { fs, path, Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, LevelFormat, AlignmentType, TabStopType, BLACK, YELLOW, GOLD, MUTED, HF, BF, run, P, H1, H2, K, B, bold, table, box, spacer } = require("./docx-brand.js");

/* Fila de ajuste: Qué cambia · Dónde · Por qué */
const adj = (rows) => table(["Qué cambia", "Dónde se cambia", "Por qué · dato"], rows, [4.2, 2.6, 3.2]);

const PERFIL = process.env.PERFIL; const C = PERFIL === "comercial";
const X = (com, adm) => (C ? com : adm);
const children = [];

/* Portada */
children.push(new Paragraph({ spacing: { before: 1200, after: 100 }, children: [new TextRun({ text: "SUPER HUMANOS", font: HF, size: 28, bold: true, color: GOLD, characterSpacing: 60 })] }));
children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "AJUSTES A LOS TALLERES", font: HF, size: 72, bold: true, color: BLACK })] }));
children.push(new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: "PERFIL " + PERFIL.toUpperCase() + " · A PARTIR DEL DIAGNÓSTICO PARCIAL Y DE LA POLÍTICA DE USO DE IA", font: HF, size: 34, bold: true, color: BLACK })] }));
children.push(P("Diagnóstico con corte del 23 de octubre de 2026, resultados del perfil " + PERFIL + ". Política de uso de inteligencia artificial y gestión de riesgo de GYJ Ferreterías S.A., código TH-PT-003, versión 1.", { run: { color: MUTED } }));
children.push(P("Pivot Power · documento para el docente del perfil " + PERFIL + " · 27 de octubre de 2026 · fechas según el cronograma ajustado del 19 de octubre", { run: { color: MUTED } }));
children.push(spacer());
children.push(box("Qué hace este documento", [
  "Cruza tres fuentes: los resultados del diagnóstico, la política de IA que firmó la gente y el temario de los seis talleres del perfil con su guía de contenido.",
  "Propone los cambios taller por taller, con el lugar exacto del temario o de la guía donde se aplican y el dato que los sustenta.",
  X("El Taller 1 arranca el viernes 23 de octubre a las 7:30. Los ajustes del Taller 1 entran completos en los dos encuentros, en el kit del participante y en los retos.", "El Taller 1 arranca el viernes 23 de octubre a las 8:00. Los ajustes del Taller 1 entran completos en los dos encuentros, en el kit del participante y en los retos."),
]));

/* 1 · Datos */
children.push(H1("1 · Lo que dicen los datos"));
children.push(P(`Las cifras del perfil ${PERFIL} están en «Resumen del diagnóstico · perfil ${PERFIL}», que acompaña este documento. Cada ajuste de las secciones 3 y 4 cita el dato que lo justifica.`));

/* 2 · Política */
children.push(H1("2 · Lo que exige la política y dónde toca el temario"));
children.push(P("La política TH-PT-003 fija reglas que el temario asumía de otra forma. Seis de ellas cambian cómo se enseña y con qué se practica."));
children.push(table(["Regla de la política", "Dónde toca el temario", "Ajuste"], [
  ["7.2 · Las cuentas personales de herramientas de IA quedan prohibidas para información de la empresa, salvo autorización expresa de Sistemas.", "Sección 3, «Herramientas del programa»: ChatGPT, Gemini y NotebookLM «gratis con cuenta». El 50 % de quienes respondieron usa cuenta personal.", "La regla del programa pasa a ser: con cuenta personal solo se trabaja con el material del kit, sin datos de la empresa ni de personas. Con información real, solo la herramienta y la cuenta que TIC autorice por escrito."],
  ["7.2 y 7.3 · Nombres, cédulas, teléfonos, correos, direcciones y cargos de clientes, proveedores y colaboradores nunca entran a una herramienta no autorizada. Tampoco listados ni bases.", X("Demostración del T1 con un WhatsApp real. T2 con un export real de cartera. T4 con la nota de voz de una visita. T5 con un reclamo real.", "Demostración del T3 con una bandeja real de correo. T2 y T4 con estadísticas y procedimientos del área."), "Cada demostración usa el archivo del kit, anonimizado, o un archivo propio al que la persona le quitó los datos de identificación. La técnica de anonimizar en dos minutos se enseña en el Taller 1."],
  ["6 y 7.3 · Antes de cargar un archivo o documento a una plataforma de IA hay que verificar con TIC que esté autorizada.", "T4 con NotebookLM y procedimientos propios. Reto 17.", "NotebookLM entra solo si TIC lo autoriza. Mientras tanto, el texto del procedimiento, sin datos personales, se pega en la herramienta autorizada."],
  ["5.5 y 6 · Validación humana obligatoria de todo resultado antes de usarlo o comunicarlo. 9 · Los líderes validan los resultados de su equipo antes de usarlos.", "Regla «toda cifra se verifica», ya en el temario. Matriz de responsabilidad del T5.", "La regla se sostiene y se cita con el número de la política. En T5 el líder entra a la matriz como aprobador, porque la política le asigna ese rol."],
  ["5.4 · El uso de IA en contenidos o decisiones que afecten a clientes o colaboradores debe ser identificable y, cuando aplique, informado.", X("T5, respuestas al cliente.", "T5, procesos del área."), "Se agrega la pregunta «¿hay que informar que esto lo apoyó una IA?» al banco de respuestas y a la matriz."],
  ["10 · La capacitación debe incluir principios éticos, riesgos de seguridad y datos, herramientas autorizadas y canales de consulta y reporte.", "Bloque «El marco de uso responsable» del T1.", "El bloque se reorganiza con esos cuatro contenidos, en ese orden. Así el Taller 1 cuenta como la capacitación que la política exige."],
], [3.3, 3.3, 3.4]));
children.push(spacer());
children.push(box("Alcance de la política", [
  "El documento está emitido por GYJ Ferreterías S.A. La dirección del programa confirmó el 14 de octubre que rige para las siete empresas. El Taller 1 la enseña como la regla de G&J en todas.",
]));

/* 3 · Transversales */
children.push(H1("3 · Ajustes que aplican a los seis talleres"));
children.push(K("A · Herramienta y cuenta"));
children.push(P("La sección 3 del temario, «Herramientas del programa», se reescribe. Texto propuesto para el párrafo de apertura y la regla:"));
children.push(box("Texto nuevo · sección 3 del temario", [
  "El programa usa las herramientas que TIC autoriza por escrito, con la cuenta que la política de IA indica. La lista vigente va en el kit del docente y se actualiza cuando TIC la cambie.",
  "Con una cuenta personal solo se trabaja con el material del kit: casos y archivos preparados por el programa, sin datos de la empresa ni de personas. Toda tarea con información real se hace en la herramienta y la cuenta autorizadas.",
  "Antes de pegar cualquier texto o cargar cualquier archivo, la persona le quita nombres, cédulas, teléfonos, correos, direcciones y cargos. Esa es la primera práctica del Taller 1.",
]));
children.push(K("B · Material del kit en lugar de archivos reales"));
children.push(P("Cada taller trae su archivo de práctica preparado por el programa, construido sobre los casos de las mesas, con datos ficticios. Lista mínima:"));
if (C) children.push(B("Un WhatsApp con 15 solicitudes de cotización de 5 a 20 ítems; una bandeja de 40 correos del lunes; un export de cartera con el mismo cliente escrito de tres formas; la nota de voz de una visita; un reclamo por retraso."));
if (!C) children.push(B("Tres archivos de la misma estadística con formatos distintos; una bandeja con derechos de petición y plazos; un procedimiento de vinculación; el flujo del ingreso con 30 firmas; dos informes del mismo mes con cifras distintas."));
children.push(P("En el temario, cada ejercicio que dice «con un archivo propio» o «con una base propia» pasa a decir «con el archivo del kit o con uno propio sin datos de personas»."));
children.push(K("C · Dos ritmos en el mismo encuentro"));
children.push(B(X("G1 Arranque, 38 personas:", "G1 Arranque, 82 personas:") + " el bloque «Primer contacto guiado» del T1 se extiende y el consultorio del Taller 1 se reserva para ellas. El docente las nombra al inicio y las invita en persona."));
children.push(B(X("G4 Referentes, 38 personas: durante la práctica apoyan a dos compañeros cada una, como ya dice el temario. Los 17 candidatos comerciales a formador interno empiezan ahí su rol.", "G4 Referentes, 84 personas: durante la práctica apoyan a dos compañeros cada una, como ya dice el temario. Los 33 candidatos administrativos a formador interno empiezan ahí su rol.")));
children.push(K("D · Celular, bloqueos y conexión"));
children.push(B("Cada práctica tiene una versión para celular: el prompt se copia desde el kit en lugar de escribirlo, y la evidencia es una captura."));
children.push(B(X("39 comerciales sin computador propio", "24 administrativos sin computador propio") + " necesitan sala con equipo o una pareja de trabajo. Gestión Humana asigna el lugar antes de cada encuentro 1."));
children.push(B("La grabación en la plataforma deja de ser opcional: " + X("32 comerciales", "53 administrativos") + " reportan conexión intermitente o solo datos del celular."));
children.push(B("Hasta que TIC confirme la lista de dominios habilitados, el docente lleva un plan B para los 142 con bloqueos: la demostración se sigue en pantalla compartida y la práctica se hace desde el celular con datos."));
children.push(K("E · El formato «Mi proceso» mide horas semanales"));
children.push(P("El diagnóstico estima cerca de 2.200 horas semanales de trabajo repetitivo en las 482 personas que respondieron. Para que esa cifra sirva al cierre, el formato pide minutos por vez y veces por semana, y el reporte mensual suma las horas liberadas. El campo ya existe en el formato; lo que cambia es que el docente lo exige lleno en cada encuentro 2."));

if (C) {
/* 4 · Comercial */
children.push(H1("4 · Ajustes por taller · comercial"));
children.push(H2("Taller 1 · IA para el asesor comercial"));
children.push(P("Encuentro 1 el viernes 23 de octubre, consultorio el miércoles 28, encuentro 2 el viernes 30. Los ajustes entran en los dos encuentros, en el kit y en la evaluación.", { run: { color: MUTED } }));
children.push(adj([
  ["El caso de clase es el de las 15 cotizaciones por WhatsApp: demostración en el encuentro 1 y primera instrucción reutilizable de cada persona en el encuentro 2, antes de pasar a una tarea propia.", "Temario, bloques «Casos del rol comercial» y «Mi primera tarea con IA». Guía, demostración y «Encuentro 2 · aplicación».", "Ítem S_1 con 30,0 % de acierto, el más bajo del diagnóstico. Aparece también en las respuestas abiertas."],
  ["El segundo caso de clase es ordenar los 40 correos del lunes: extraer qué pide cada uno, quién y para cuándo, en una tabla.", "Temario, bloque «Casos del rol comercial»: «ordenar los pendientes del lunes» pasa de mención a ejercicio.", "Ítem S_5 con 46,7 %."],
  ["Apertura del encuentro 1 con la política: los cuatro contenidos que exige el numeral 10, en diez minutos, y la firma pendiente resuelta en el mismo espacio con Gestión Humana.", "Temario, bloque «El marco de uso responsable de G&J». Guía, bloque «Uso responsable».", "52 comerciales de 150 sin firmar o sin saber de la política."],
  ["La regla de cuenta se enseña con el dato: «la mitad de ustedes usa cuenta personal; con ella solo se practica con el material del kit».", "Mismo bloque.", "Política 7.2. 57 % de los comerciales usa cuenta personal."],
  ["Anonimizar en dos minutos: antes de pegar el WhatsApp, se reemplazan nombre y teléfono por «Cliente A». Se demuestra y se exige en toda práctica.", "Temario, bloque «Cómo se le habla a una IA», como paso previo. Guía, demostración, paso 2.", "Política 7.3. El mensaje real de la demostración trae nombre y número del cliente."],
  ["Evaluación: las preguntas 6 y 7 se escriben sobre casos de la política: un listado de clientes, la cédula de un comprador, el precio de una cotización, una ficha técnica pública.", "Guía, «Qué debe medir la evaluación», filas 6 y 7.", "Política 7.2 y 7.3."],
]));
children.push(H2("Taller 2 · Inteligencia de mercado con IA"));
children.push(adj([
  ["Se agrega un bloque de cierre: de la tabla al mensaje. Cómo contar los resultados del mes en diez minutos: qué pasó, por qué, qué propongo, qué necesito.", "Temario, temáticas del T2 comercial, nuevo bloque antes de «Compartirlo». Guía, «Lo que no puede faltar».", "Ítem D_5 con 63,3 %."],
  ["El export de cartera de la demostración es el del kit, con clientes ficticios. El ejercicio «cada quien limpia una base propia» pasa a «una base propia sin nombres de clientes o la del kit».", "Guía, demostración y ejercicio del T2 comercial.", "Política 7.3: listados de clientes no entran a herramientas no autorizadas."],
  ["La clasificación de tareas con los seis destinos arranca con las horas semanales que cada persona declaró en el diagnóstico.", "Temario, bloque de los seis destinos.", "54 comerciales dedican más de 5 horas semanales a tareas repetitivas."],
]));
children.push(H2("Taller 3 · Venta consultiva y experiencia del cliente"));
children.push(adj([
  ["Bloque nuevo: del cuaderno al tablero de cotizaciones. Qué convierte una solicitud en pendiente con responsable y fecha, y cómo verla en una herramienta compartida con el líder.", "Temario, temáticas del T3 comercial, después de «la solicitud completa». Guía, «Lo que no puede faltar».", "84 de 150 comerciales llevan pendientes en papel; 11 usan una herramienta compartida. Carolina pidió este taller."],
  ["La bandeja de 40 correos del lunes es el caso de flujo: distinguir trabajo de espera y armar la solicitud completa desde ahí.", "Temario, bloque «recorrido del pedido».", "Ítem S_5 con 46,7 %."],
  ["El caso del pedido retenido por cartera se trabaja completo: causa, fecha nueva y alternativa, con la persona de cartera en el ejercicio.", "Ya está en el temario. Sube a demostración principal.", "Ítem S_4 con 75,3 %."],
  ["Construir un documento entre cuatro personas: coautoría en Word o Excel compartido en lugar de versiones por correo.", "Temario, bloque «la solicitud completa», como herramienta.", "Ítem D_1 con 76,7 %."],
]));
children.push(H2("Taller 4 · Gestión comercial basada en datos"));
children.push(adj([
  ["Nivelación de CRM: antes del mapa de fuentes, veinte minutos con los 26 que se traban o no lo conocen, y los que «podrían enseñarlo» como apoyo.", "Temario, bloque «El CRM como memoria del área».", "26 comerciales se traban o no conocen el CRM; 27 dicen que podrían enseñarlo. Salesforce entra después."],
  ["La nota de voz de la visita se registra sin nombre ni teléfono del cliente; el registro completo va al CRM, que es el sistema autorizado.", "Guía, demostración del T4 comercial.", "Política 7.3."],
  ["Mensaje de apertura: la fuente oficial manda. Si el dato está en el CRM o en la lista de precios, se usa de ahí; el archivo propio es la excepción.", "Temario, bloque «mapa de fuentes».", "35 comerciales rehacen en archivo propio lo que está en el sistema."],
]));
children.push(H2("Taller 5 · Atención al cliente en canales digitales"));
children.push(adj([
  ["La matriz de responsabilidad incluye al líder como aprobador de las respuestas que salen al cliente, con el número de la política.", "Temario y guía, bloque de la matriz.", "Política 9: los líderes validan los resultados antes de su uso."],
  ["Se agrega al banco de respuestas la pregunta de transparencia: cuándo se informa que una respuesta tuvo apoyo de IA.", "Guía, «Lo que no puede faltar».", "Política 5.4."],
  ["El reclamo de la demostración es el del kit; las respuestas base de cada persona se escriben sin datos del cliente.", "Guía, demostración y ejercicio.", "Política 7.3."],
]));
children.push(H2("Taller 6 · Proyecto comercial con IA"));
children.push(adj([
  ["La línea base de cada proyecto sale del formato «Mi proceso» acumulado: horas semanales antes y después.", "Temario, bloque «cómo medir la línea base».", "Cerca de 2.200 horas semanales estimadas como denominador del programa."],
  ["Los informes que comercial hace y otra área también hace son los primeros candidatos al Banco de Retos.", "Temario, bloque «de la queja al reto».", "62 comerciales reportan informes duplicados y 41 creen que sí."],
]));

}
if (!C) {
/* 5 · Administrativo */
children.push(H1("4 · Ajustes por taller · administrativo"));
children.push(H2("Taller 1 · IA para la gestión administrativa"));
children.push(P("Encuentro 1 el viernes 23 de octubre, consultorio el martes 27, encuentro 2 el miércoles 28. Los ajustes entran en los dos encuentros, en el kit y en la evaluación.", { run: { color: MUTED } }));
children.push(adj([
  ["El taller abre por la estructura de comunicar: el resumen de cinco puntos de la demostración se convierte en lo que se presenta en diez minutos: qué pasó, por qué, qué proponemos, qué necesitamos. En el encuentro 2 cada persona lo hace con un documento suyo.", "Temario, bloques «Casos del rol administrativo» y «Mi primera tarea con IA». Guía, demostración y «Encuentro 2».", "Ítem D_5 con 74,7 %, el más bajo del perfil."],
  ["El documento normativo de la demostración es un texto público: la Ley 1581 de 2012 o la Circular 002 de 2024 de la SIC, que la política cita. Así la demostración enseña la norma y evita pegar un documento interno en una herramienta pública.", "Guía, demostración del T1 administrativo, paso 1.", "Política 7.2: información interna fuera de plataformas públicas."],
  ["Las notas del comité para el acta se escriben con cargos en lugar de nombres completos.", "Guía, demostración, paso 4. Temario, bloque «Casos del rol administrativo».", "Política 7.3: el nombre completo y el cargo juntos identifican a la persona."],
  ["Apertura del encuentro 1 con la política, los cuatro contenidos del numeral 10, y la firma pendiente resuelta con Gestión Humana.", "Temario, bloque «El marco de uso responsable».", "115 administrativos de 332 sin firmar o sin saber de la política."],
  ["Niveles de riesgo con los datos del rol: nómina, datos de terceros, información financiera, documentos de vinculación. Cada ejemplo se clasifica con la política en pantalla.", "Ya está en la guía. Se cita el numeral 7.3.", "Política 7.2 y 7.3."],
]));
children.push(H2("Taller 2 · Datos para decidir con IA"));
children.push(adj([
  ["Los tres archivos de la demostración son del kit. El ejercicio «cada quien consolida sus propios archivos» pasa a «archivos propios sin datos de personas o los del kit».", "Guía, demostración y ejercicio.", "Política 7.3. Las estadísticas del área traen nombres de colaboradores y clientes."],
  ["El informe de una página se trabaja como presentación de diez minutos, con la misma estructura del T1.", "Temario, bloque «El informe de una página».", "Ítem D_5 con 74,7 %."],
  ["La consolidación manual se mide en horas: cada persona anota cuánto le toma el reporte del mes antes de tocarlo.", "Temario, bloque «La consolidación manual».", "100 administrativos dedican más de 5 horas semanales a tareas repetitivas."],
]));
children.push(H2("Taller 3 · Productividad digital: de la bandeja al tablero"));
children.push(adj([
  ["Este taller sube de prioridad y se presenta como el taller de seguimiento que pidió Gestión Humana. La demostración termina con el tablero del equipo visible para el líder.", "Temario, bloque «el tablero del equipo». Guía, demostración.", "220 de 332 administrativos llevan pendientes en papel o en el correo; 34 de memoria; 23 con herramienta compartida."],
  ["La bandeja de la demostración es la del kit, con derechos de petición y plazos. Las bandejas propias se trabajan con los nombres reemplazados por área.", "Guía, demostración.", "Política 7.3."],
  ["Construir un documento entre cuatro personas: coautoría en Microsoft 365, con control de versiones, como parte del tablero.", "Temario, bloque «el tablero personal».", "Ítem D_1 con 78,6 %."],
]));
children.push(H2("Taller 4 · Procesos internos con IA"));
children.push(adj([
  ["El caso del ingreso con 30 firmas se trabaja aquí como ejercicio de rediseño: qué firma aporta, cuál se elimina, cuál se vuelve digital. En el T6 se mide.", "Temario, sección 4, tabla de casos: «Administrativo · talleres 4 y 6». Guía, «Lo que no puede faltar».", "Ítem S_5 con 76,8 %, segundo más bajo del perfil."],
  ["El inventario de reportes del área se hace en clase, con la lista de informes que cada persona produce y quién más los produce.", "Temario, bloque «Informes duplicados». Sube de mención a ejercicio.", "121 administrativos reportan informes duplicados; otros 69 creen que sí."],
  ["Mensaje de apertura: la fuente oficial manda. Qué dato vive en el ERP UnoE y qué dato se rehace en Excel sin necesidad.", "Temario, bloque «mapa de fuentes».", "57 administrativos rehacen en archivo propio; 80 registran en Excel propio; 127 en UnoE."],
  ["Preguntarle a los procedimientos propios: con NotebookLM solo si TIC lo autoriza. Mientras tanto, el texto del procedimiento sin datos personales se pega en la herramienta autorizada.", "Temario, bloque «preguntarle a los procedimientos». Sección 3.", "Política 6 y 7.3."],
  ["Documentar el propio trabajo: el ejercicio de la guía de una página se conecta con los 32 administrativos que dicen que lo suyo está sin escribir.", "Temario, bloque «formularios repetitivos» y reto 19.", "Ítem de mentalidad con 32 administrativos en 1 o 2."],
]));
children.push(H2("Taller 5 · Adopción de la IA en el equipo"));
children.push(adj([
  ["La matriz de responsabilidad se arma sobre un proceso del ERP, con el líder como aprobador.", "Temario y guía, bloque de la matriz.", "219 administrativos responden que el CRM no aplica; su sistema es UnoE. Política 9."],
  ["Transparencia: cuándo se informa a un colaborador o a un tercero que un resultado tuvo apoyo de IA.", "Guía, «Lo que no puede faltar».", "Política 5.4."],
  ["Cómo enseñarle a un compañero en cinco minutos: lo hacen los candidatos a formador interno con alguien de G1.", "Temario, bloque «cómo enseñarle a un compañero».", "33 candidatos administrativos; 82 personas en G1."],
]));
children.push(H2("Taller 6 · Proyecto de mejora con IA"));
children.push(adj([
  ["La línea base sale del formato «Mi proceso» acumulado, en horas semanales. El trámite de 30 firmas, ya rediseñado en el T4, se mide aquí.", "Temario, bloque «cómo medir la línea base».", "Cerca de 2.200 horas semanales estimadas."],
  ["Los informes duplicados entre áreas son los primeros proyectos del Banco de Retos.", "Temario, bloque «de la queja al reto».", "36 % afirmativo más 21 % «creo que sí»."],
]));

}
/* 6 · Retos */
children.push(H1("5 · Retos de los jueves que cambian"));
children.push(table(["Reto", "Como está", "Como queda", "Por qué"], [
  ["5 · 22 de octubre", "Pídele a la IA un resumen de la política de IA de G&J y verifica que sea fiel.", "Lee la política y escribe tres datos de tu trabajo que nunca pegarías en una IA. Después pregúntale a la IA qué datos nunca debería recibir y compara las dos listas.", "La política es un documento interno. Pegarla en una herramienta pública con cuenta personal contradice su numeral 7.2."],
  ["10 · 26 de noviembre", "Limpia con IA una lista con nombres repetidos o mal escritos.", "Limpia la lista del kit, o una propia a la que le quitaste nombres y contactos.", "Política 7.3."],
  ["17 · 28 de enero", "Dale a la IA un documento oficial de tu área y hazle tres preguntas.", "Dale a la IA un documento oficial de tu área sin datos de personas, en la herramienta autorizada, y hazle tres preguntas.", "Política 6 y 7.3."],
  ["22 · 4 de marzo", "Prepara con IA el acta de una reunión con decisiones, responsables y fechas.", "Igual, con los responsables escritos por cargo.", "Política 7.3."],
], [1.6, 2.6, 3.4, 2.4]));


const doc = new Document({
  creator: "Pivot Power", title: "Ajustes a los talleres a partir del diagnóstico y la política de IA",
  styles: { default: { document: { run: { font: BF, size: 21, color: BLACK } } } },
  numbering: { config: [{ reference: "chev", levels: [{ level: 0, format: LevelFormat.BULLET, text: "›", alignment: AlignmentType.LEFT, style: { run: { font: HF, bold: true, color: GOLD, size: 24 }, paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1100, bottom: 1100, left: 1134, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: "SUPER HUMANOS · AJUSTES A LOS TALLERES · PERFIL " + PERFIL.toUpperCase() + "", font: HF, size: 17, bold: true, color: MUTED, characterSpacing: 30 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 9600 }], children: [new TextRun({ text: "Pivot Power · documento de trabajo · 27 de octubre de 2026", font: BF, size: 16, color: MUTED }), new TextRun({ text: "\t", font: BF, size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: HF, size: 18, bold: true, color: BLACK })] })] }) },
    children,
  }],
});

const out = path.join(__dirname, "out", `Ajustes a los talleres · perfil ${PERFIL}.docx`);
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("ok", out, buf.length); });
