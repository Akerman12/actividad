// Genera docs/Informe-Taller-ToolCalling.docx a partir de lo que hay en el repo:
// código real (src/, pruebas.http) y capturas reales de docs/evidencias/capturas.
// Uso (una vez): npm install   |   Generar: npm run generar
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, AlignmentType, ImageRun, Footer, PageNumber, LevelFormat,
  TableOfContents, PageBreak, VerticalAlign,
} = require('docx');

const RAIZ = path.resolve(__dirname, '..', '..');
const CAPTURAS = path.join(RAIZ, 'docs', 'evidencias', 'capturas');
const SALIDA = path.join(RAIZ, 'docs', 'Informe-Taller-ToolCalling.docx');

const AUTOR = 'Sebastian Montes Olivera';
const FECHA = '18 de noviembre de 2026';
const REPO_URL = process.env.REPO_URL || 'https://github.com/SebastianMontes-Dev/toolcalling-spring-ai-taller';

const AZUL = '1F3864';
const GRIS_BORDE = 'BFBFBF';
const GRIS_CODIGO = 'F2F2F2';
const ANCHO = 9026; // A4 con márgenes de 1"

// ---------- lectura ----------
const leer = (...p) => fs.readFileSync(path.join(...p), 'utf8').replace(/\r/g, '').replace(/^﻿/, '');
const codigo = (rel) => leer(RAIZ, rel).replace(/\n+$/, '').split('\n');

function dimensionesPng(buf) {
  if (buf.slice(0, 8).toString('hex') !== '89504e470d0a1a0a') return null;
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}

// ---------- bloques de documento ----------
const t = (text, o = {}) => new TextRun({ text, ...o });
const parrafo = (children, o = {}) =>
  new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: typeof children === 'string' ? [t(children)] : children });
const h1 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(s)] });
const h2 = (s) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(s)] });
const punto = (children) =>
  new Paragraph({ numbering: { reference: 'vinetas', level: 0 }, spacing: { after: 60 }, children: typeof children === 'string' ? [t(children)] : children });
const negrita = (s) => t(s, { bold: true });
const espacio = () => new Paragraph({ spacing: { after: 160 }, children: [] });

function bloqueCodigo(lineas, titulo) {
  const out = [];
  if (titulo) out.push(new Paragraph({ keepNext: true, spacing: { before: 120, after: 40 }, children: [t(titulo, { italics: true, size: 20, color: '595959' })] }));
  lineas.forEach((l, i) => {
    out.push(new Paragraph({
      keepLines: true,
      keepNext: lineas.length <= 25 && i < lineas.length - 1,
      spacing: { before: 0, after: 0, line: 252 },
      indent: { left: 200 },
      shading: { type: ShadingType.CLEAR, fill: GRIS_CODIGO, color: 'auto' },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 6 } },
      children: [t(l === '' ? ' ' : l, { font: 'Consolas', size: 17 })],
    }));
  });
  out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  return out;
}

const borde = { style: BorderStyle.SINGLE, size: 4, color: GRIS_BORDE };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
function celda(contenido, ancho, { cabecera = false, mono = false, centro = false } = {}) {
  const runs = typeof contenido === 'string'
    ? [t(contenido, { bold: cabecera, color: cabecera ? 'FFFFFF' : undefined, font: mono ? 'Consolas' : undefined, size: mono ? 18 : 20 })]
    : contenido;
  return new TableCell({
    width: { size: ancho, type: WidthType.DXA },
    borders: bordes,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    shading: cabecera ? { type: ShadingType.CLEAR, fill: AZUL, color: 'auto' } : undefined,
    children: [new Paragraph({ alignment: centro ? AlignmentType.CENTER : undefined, spacing: { after: 0 }, children: runs })],
  });
}
function tabla(cabeceras, filas, anchos, opciones = {}) {
  const suma = anchos.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: suma, type: WidthType.DXA },
    columnWidths: anchos,
    rows: [
      new TableRow({ tableHeader: true, children: cabeceras.map((c, i) => celda(c, anchos[i], { cabecera: true })) }),
      ...filas.map((f) => new TableRow({
        cantSplit: true,
        children: f.map((c, i) => celda(c, anchos[i], { mono: (opciones.mono || []).includes(i), centro: (opciones.centro || []).includes(i) })),
      })),
    ],
  });
}

// Recuadro visible cuando falta una captura: nunca se inventa evidencia.
function recuadroFalta(archivo, descripcion) {
  const b = { style: BorderStyle.DASHED, size: 8, color: 'C00000' };
  return new Table({
    width: { size: ANCHO, type: WidthType.DXA },
    columnWidths: [ANCHO],
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: ANCHO, type: WidthType.DXA },
        borders: { top: b, bottom: b, left: b, right: b },
        margins: { top: 240, bottom: 240, left: 200, right: 200 },
        shading: { type: ShadingType.CLEAR, fill: 'FDF2F2', color: 'auto' },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [t(`[Falta la captura: ${archivo}]`, { bold: true, color: 'C00000' })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [t(descripcion, { italics: true, size: 20, color: '7F7F7F' })] }),
        ],
      })],
    })],
  });
}

const faltantes = [];
let nFigura = 0;
function captura(base, descripcion) {
  nFigura += 1;
  const pie = new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 200 }, children: [t(`Figura ${nFigura}. ${descripcion}`, { italics: true, size: 19, color: '595959' })] });
  const f = path.join(CAPTURAS, `${base}.png`);
  if (!fs.existsSync(f)) {
    faltantes.push(base);
    return [recuadroFalta(`${base}.png`, descripcion), pie];
  }
  const buf = fs.readFileSync(f);
  const d = dimensionesPng(buf);
  if (!d) throw new Error(`No se pudo leer el tamaño de ${f}`);
  const esc = Math.min(1, 600 / d.w);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER, keepNext: true, spacing: { after: 0 },
      children: [new ImageRun({ type: 'png', data: buf, transformation: { width: Math.round(d.w * esc), height: Math.round(d.h * esc) }, altText: { title: descripcion, description: descripcion, name: base } })],
    }),
    pie,
  ];
}

const hijos = [];
const add = (...x) => x.flat().forEach((e) => hijos.push(e));

// ---------- Portada ----------
add(
  new Paragraph({ spacing: { before: 2200, after: 200 }, children: [t('TALLER PRÁCTICO', { size: 26, bold: true, color: '7F7F7F', characterSpacing: 60 })] }),
  new Paragraph({ spacing: { after: 160 }, children: [t('Implementación básica de Tool Calling con Spring AI', { size: 48, bold: true, color: AZUL })] }),
  new Paragraph({
    spacing: { after: 700 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 12 } },
    children: [t('Asignatura: Desarrollo Web Avanzado', { size: 28, color: '595959' })],
  }),
  tabla(['Dato', 'Detalle'], [
    ['Autor', AUTOR],
    ['Fecha', FECHA],
    ['Tecnologías', 'Java 21, Spring Boot 3.5.16, Spring AI 1.1.8, Maven, VS Code'],
    ['Proveedor del modelo', 'Groq (API compatible con OpenAI), modelo openai/gpt-oss-120b'],
    ['Repositorio', REPO_URL],
  ], [2400, ANCHO - 2400]),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ spacing: { after: 200 }, children: [t('Contenido', { size: 36, bold: true, color: AZUL })] }),
  new TableOfContents('Contenido', { hyperlink: true, headingStyleRange: '1-2' }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ---------- 1. Propósito ----------
add(
  h1('1. Propósito y alcance'),
  parrafo('Tool Calling permite que un modelo de lenguaje solicite a la aplicación la ejecución de un método previamente autorizado. El modelo no ejecuta código Java ni accede a la aplicación en general: solo conoce las herramientas que el desarrollador decide exponer. En este taller se construyó una pequeña API académica en la que el programador define las herramientas y el LLM decide cuándo pedir una y con qué argumentos.'),
  parrafo('Este informe reúne las evidencias que pide la guía: estructura del proyecto, código de las herramientas y del controlador, archivo pruebas.http, respuestas y consola de ejecución, las cuatro actividades y las preguntas de análisis. Todas las capturas proceden de una ejecución real hecha el 18 de noviembre de 2026.'),
  tabla(['Herramienta', 'Entrada', 'Resultado', 'Origen'], [
    ['calcularPromedio', 'nota1, nota2, nota3', 'Promedio numérico', 'Guía'],
    ['estadoAcademico', 'promedio', 'APROBADO / REPROBADO (mínimo 3.0)', 'Guía'],
    ['calcularNotaFinal', '3 notas y 3 porcentajes', 'Nota definitiva ponderada', 'Actividad 1'],
    ['estadoAsistencia', 'porcentaje de asistencia', 'CUMPLE_ASISTENCIA / EN_RIESGO / REPRUEBA_POR_INASISTENCIA', 'Actividad 2'],
  ], [2300, 2300, 3426, 1000], { mono: [0] }),
  espacio(),
);

// ---------- 2. Entorno y diferencias ----------
add(
  h1('2. Entorno y diferencias respecto a la guía'),
  parrafo('La guía indica que debe verificarse la versión de las dependencias, las propiedades y el modelo antes de ejecutar. Estos son los ajustes realizados y el motivo de cada uno:'),
  tabla(['Aspecto', 'Guía', 'Esta implementación', 'Motivo'], [
    ['Java', '25 (si Spring Boot lo soporta)', '21 (LTS)', 'Es el JDK instalado y Spring Boot 3.5 lo soporta.'],
    ['Spring Boot / Spring AI', 'Versión vigente en clase', '3.5.16 / 1.1.8', 'Combinación compatible; la anotación @Tool existe desde Spring AI 1.0.'],
    ['IDE', 'IntelliJ IDEA', 'VS Code (Extension Pack for Java)', 'Preferencia del autor; la ejecución, la consola y el cliente .http son equivalentes.'],
    ['Modelo', '<MODELO_CON_TOOL_CALLING>', 'openai/gpt-oss-120b', 'llama-3.3-70b-versatile ya no aparece en la API de Groq (consulta a /models el 18/11/2026); este sí soporta herramientas.'],
    ['Clave de API', 'Variable de entorno en el IDE', 'Variable GROQ_API_KEY', 'En VS Code se carga con envFile desde un archivo local fuera del repositorio. Nunca se versiona.'],
    ['pruebas.http', 'Cuerpo pegado a las cabeceras', 'Con una línea en blanco', 'REST Client exige separar cabeceras y cuerpo con una línea en blanco.'],
    ['Extras', '—', 'Pruebas JUnit y CI', 'Se añadieron 16 pruebas automatizadas y un flujo de GitHub Actions.'],
  ], [1500, 2000, 2200, 3326]),
  espacio(),
);

// ---------- 3. Estructura ----------
add(
  h1('3. Estructura del proyecto'),
  parrafo('El proyecto Maven sigue la estructura de la guía (ToolCallingApplication, controller/ChatController, tools/HerramientasAcademicas, application.properties y pruebas.http). Además contiene la carpeta de pruebas y la documentación.'),
  ...captura('01-estructura-proyecto', 'Estructura del proyecto toolcalling abierta en VS Code.'),
  ...bloqueCodigo([
    'org.springframework.boot : spring-boot-starter-web',
    'org.springframework.ai   : spring-ai-starter-model-openai   (BOM spring-ai-bom 1.1.8)',
  ], 'Dependencias principales (pom.xml)'),
);

// ---------- 4. Herramientas ----------
const herr = codigo('src/main/java/com/ejemplo/toolcalling/tools/HerramientasAcademicas.java');
add(
  h1('4. Código de HerramientasAcademicas.java'),
  parrafo('Cada método anotado con @Tool es una capacidad que Spring AI presenta al modelo. La descripción es la que ayuda al LLM a entender para qué sirve y cuándo pedirla; los parámetros del método son los datos que el modelo debe extraer del mensaje del usuario. Cada herramienta imprime una línea que empieza con ">>> TOOL": es la evidencia de que el método Java se ejecutó en la aplicación.'),
  ...bloqueCodigo(herr, 'src/main/java/com/ejemplo/toolcalling/tools/HerramientasAcademicas.java'),
  ...captura('02-herramientas-academicas-1', 'HerramientasAcademicas.java: herramientas de la guía y comienzo de calcularNotaFinal.'),
  ...captura('03-herramientas-academicas-2', 'HerramientasAcademicas.java: calcularNotaFinal (validación de porcentajes) y estadoAsistencia.'),
);

// ---------- 5. Controlador y configuración ----------
add(
  h1('5. Controlador, configuración y pruebas .http'),
  h2('5.1 ChatController.java'),
  parrafo('La instrucción central es .tools(herramientas): registra las herramientas disponibles para esa interacción. El controlador nunca invoca calcularPromedio() ni estadoAsistencia(); la selección se delega al modelo.'),
  ...bloqueCodigo(codigo('src/main/java/com/ejemplo/toolcalling/controller/ChatController.java'), 'src/main/java/com/ejemplo/toolcalling/controller/ChatController.java'),
  ...captura('04-chat-controller', 'ChatController.java en VS Code.'),
  h2('5.2 application.properties'),
  parrafo('La clave no está escrita en el archivo: se lee de la variable de entorno GROQ_API_KEY. El modelo puede cambiarse con GROQ_MODEL sin tocar el código.'),
  ...bloqueCodigo(codigo('src/main/resources/application.properties'), 'src/main/resources/application.properties'),
  ...captura('05-application-properties', 'application.properties con la clave leída desde variable de entorno.'),
  h2('5.3 pruebas.http'),
  parrafo('Contiene las cuatro pruebas de la guía y las preguntas de las actividades 3 y 4 (se ejecutan con el enlace "Send Request" de la extensión REST Client).'),
  ...captura('06-pruebas-http', 'Archivo pruebas.http con los enlaces Send Request de REST Client.'),
  ...bloqueCodigo(codigo('pruebas.http'), 'pruebas.http (contenido completo)'),
);

// ---------- 6. Ejecución ----------
add(
  h1('6. Ejecución de la aplicación'),
  parrafo('La aplicación se lanzó desde VS Code (Run and Debug). Inició sin errores en 1,56 segundos y escucha en el puerto 8080. La consola de depuración se mantuvo visible durante todas las pruebas, ya que allí se comprueba cuándo se ejecuta una herramienta.'),
  ...captura('07-aplicacion-iniciada', 'Spring Boot 3.5.16 iniciado: "Tomcat started on port 8080" y "Started ToolCallingApplication in 1.56 seconds".'),
);

// ---------- 7. Pruebas de la guía ----------
add(
  h1('7. Pruebas de la guía'),
  parrafo('Resumen de lo observado en las cuatro pruebas. Los textos de respuesta se transcriben de las capturas; el modelo puede redactar de forma distinta en otra ejecución, pero los valores y la herramienta elegida son los que muestra la consola.'),
  tabla(['Prueba', 'Mensaje enviado', 'Herramienta ejecutada (consola)', 'Respuesta HTTP 200'], [
    ['1', 'Calcula el promedio de las notas 4.0, 3.5 y 4.5', 'calcularPromedio(4.0, 3.5, 4.5)', 'El promedio ... es 4.0.'],
    ['2', 'Tengo las notas 2.5, 2.8 y 3.2. Calcula mi promedio.', 'calcularPromedio(2.5, 2.8, 3.2)', 'El promedio de tus notas es 2.83 (aproximadamente).'],
    ['3', 'Un estudiante tiene promedio 2.8. ¿Cuál es su estado académico?', 'estadoAcademico(2.8)', 'El estudiante está REPROBADO.'],
    ['4', 'Explícame qué significa promedio académico.', 'Ninguna', 'Explicación conceptual (2 914 bytes).'],
  ], [900, 3000, 2500, 2626]),
  espacio(),
  h2('Prueba 1 - Cálculo mediante una herramienta'),
  parrafo('Para 4.0, 3.5 y 4.5 el resultado matemático es 4.0. Además de la respuesta HTTP, la consola muestra ">>> TOOL calcularPromedio ejecutada con: 4.0, 3.5, 4.5": el resultado no provino solo de un texto del LLM, la aplicación ejecutó el método Java.'),
  ...captura('08-prueba1-calcular-promedio', 'Prueba 1: respuesta (derecha) y línea ">>> TOOL calcularPromedio" en la consola (abajo).'),
  h2('Prueba 2 - Otro promedio'),
  parrafo('El método devolvió 2.8333333333333335; el modelo redondeó a 2.83 al redactar su respuesta. La consola registra los argumentos exactos que el modelo extrajo del texto.'),
  ...captura('09-prueba2-otro-promedio', 'Prueba 2: calcularPromedio con 2.5, 2.8 y 3.2.'),
  h2('Prueba 3 - Selección de otra herramienta'),
  parrafo('Ante un promedio ya calculado, el modelo no volvió a calcular: eligió estadoAcademico con 2.8 y la lógica Java devolvió REPROBADO.'),
  ...captura('10-prueba3-estado-academico', 'Prueba 3: el modelo selecciona estadoAcademico(2.8) y la respuesta es REPROBADO.'),
  h2('Prueba 4 - Pregunta que no requiere herramienta'),
  parrafo('El modelo explicó el concepto directamente. En la consola no aparece ninguna línea nueva con ">>> TOOL": registrar las herramientas no significa ejecutarlas obligatoriamente.'),
  ...captura('11-prueba4-sin-tool', 'Prueba 4: respuesta conceptual sin que la consola registre una herramienta nueva.'),
);

// ---------- 8. Actividades ----------
add(
  h1('8. Actividades'),
  h2('Actividad 1 - calcularNotaFinal'),
  parrafo('Se agregó la herramienta calcularNotaFinal(nota1, nota2, nota3, porcentaje1, porcentaje2, porcentaje3), que devuelve la nota definitiva ponderada: (nota1·p1 + nota2·p2 + nota3·p3) / 100. Los porcentajes deben sumar 100; si no, el método lanza IllegalArgumentException con un mensaje claro (lo cubre una prueba unitaria). Con 4.0, 3.0 y 5.0 al 30 %, 30 % y 40 % el resultado es 4.1.'),
  ...captura('12-a3-1-nota-final', 'calcularNotaFinal ejecutada con las notas 4.0, 3.0, 5.0 y los porcentajes 30, 30, 40; respuesta: 4.1.'),
  h2('Actividad 2 - estadoAsistencia'),
  parrafo('Se agregó estadoAsistencia(porcentajeAsistencia) con una regla definida en Java: 80 % o más devuelve CUMPLE_ASISTENCIA; entre 60 % y menos de 80 % devuelve EN_RIESGO; por debajo de 60 % devuelve REPRUEBA_POR_INASISTENCIA. Con 75 % el estado es EN_RIESGO.'),
  ...captura('13-a3-2-asistencia', 'estadoAsistencia ejecutada con 75.0 %; respuesta: EN_RIESGO.'),
  h2('Actividad 3 - Prueba de selección con cuatro herramientas'),
  parrafo('Con las cuatro herramientas disponibles se formularon cinco preguntas distintas (A3-1 a A3-5) y se registró, mediante los mensajes de consola, cuál herramienta seleccionó el modelo. Además se conservó un resultado inesperado y se añadió una variante (A3-3b).'),
  tabla(['Id', 'Pregunta', 'Se esperaba', 'Herramienta ejecutada (consola)', 'Respuesta'], [
    ['A3-1', 'Notas 4.0, 3.0 y 5.0 con 30 %, 30 % y 40 %: ¿nota definitiva?', 'calcularNotaFinal', 'calcularNotaFinal(4.0, 3.0, 5.0, 30, 30, 40)', '4.1'],
    ['A3-2', 'Asistió al 75 % de las clases: ¿estado por asistencia?', 'estadoAsistencia', 'estadoAsistencia(75.0)', 'EN_RIESGO'],
    ['A3-3', 'Necesito el promedio simple de 3.2, 4.1 y 2.9.', 'calcularPromedio', 'Ninguna: el modelo calculó por sí mismo', '3.4'],
    ['A3-3b', 'Calcula el promedio de las notas 3.2, 4.1 y 2.9.', 'calcularPromedio', 'calcularPromedio(3.2, 4.1, 2.9)', '3.4'],
    ['A3-4', 'Si mi promedio es 3.4, ¿estoy aprobado o reprobado?', 'estadoAcademico', 'estadoAcademico(3.4)', 'APROBADO'],
    ['A3-5', 'Notas 2.0, 3.0 y 2.5: calcula mi promedio y dime si aprobé.', 'Dos herramientas encadenadas', 'calcularPromedio(2.0, 3.0, 2.5) y luego estadoAcademico(2.5)', '2.5 y REPROBADO'],
  ], [850, 2500, 1500, 2650, 1526]),
  espacio(),
  parrafo([negrita('Resultado inesperado en A3-3. '), t('El modelo respondió con la fórmula y el valor 3.4 sin invocar calcularPromedio (la consola no imprime ninguna línea). Para tres números sencillos juzgó que no necesitaba la herramienta. Con la redacción explícita de A3-3b ("Calcula el promedio de las notas ...") sí la invocó. Ambos intentos se conservan tal como ocurrieron: muestra que la decisión es del modelo, que depende de cómo se redacte la solicitud y que no es determinista. Para operaciones que siempre deben ejecutarse en Java conviene una descripción precisa de la herramienta y, si hace falta, instrucciones de sistema.')]),
  ...captura('14-a3-3-sin-tool-intento', 'A3-3: respuesta correcta pero sin ejecución de herramienta (la consola no añade ninguna línea ">>> TOOL").'),
  ...captura('19-a3-3b-promedio-explicito', 'A3-3b: con una redacción explícita el modelo sí ejecuta calcularPromedio(3.2, 4.1, 2.9).'),
  ...captura('15-a3-4-estado-academico', 'A3-4: el modelo elige estadoAcademico(3.4) y responde APROBADO.'),
  ...captura('16-a3-5-encadenadas', 'A3-5: dos herramientas encadenadas, calcularPromedio(2.0, 3.0, 2.5) y estadoAcademico(2.5).'),
  parrafo('La consola completa, con las nueve ejecuciones de herramientas en el orden real, se muestra a continuación (transcripción en docs/evidencias/consola/consola-tools.txt).'),
  ...captura('20-consola-completa', 'Consola de depuración con todas las ejecuciones de herramientas de la sesión.'),
  h2('Actividad 4 - Preguntas sin herramienta'),
  tabla(['Id', 'Pregunta', 'Por qué no era necesario Tool Calling'], [
    ['A4-1', '¿Qué diferencia hay entre una nota definitiva y una nota parcial?', 'Es una pregunta conceptual: la respuesta es conocimiento general que el modelo ya tiene. No hay números que procesar ni una operación de la aplicación que ejecutar, y ninguna de las cuatro herramientas aplica.'],
    ['A4-2', 'Dame tres consejos breves para organizar mi tiempo de estudio.', 'Es una petición de consejo general y creativa. No depende de datos ni de reglas de la aplicación; ejecutar una herramienta no aportaría nada y solo añadiría una ida y vuelta innecesaria.'],
  ], [800, 3000, 5226]),
  espacio(),
  parrafo('En ambos casos la consola no registra ninguna línea nueva con ">>> TOOL", lo que confirma que el modelo respondió sin invocar herramientas.'),
  ...captura('17-a4-1-sin-tool', 'A4-1: explicación sobre nota definitiva y parcial; sin nuevas líneas de herramienta.'),
  ...captura('18-a4-2-sin-tool', 'A4-2: tres consejos de estudio; sin nuevas líneas de herramienta.'),
);

// ---------- 9. Preguntas de análisis ----------
add(
  h1('9. Preguntas de análisis'),
  parrafo('La guía numera estas preguntas del 5 al 11 por un error de formato; aquí se numeran del 1 al 7 en el mismo orden.'),
  parrafo([negrita('1. ¿Cuál es la función de la anotación @Tool?')]),
  parrafo('Marca un método Java como una capacidad que Spring AI puede presentar al modelo. A partir del método y de su descripción, Spring AI construye el nombre, la descripción y el esquema de parámetros que se envían al LLM, y, cuando el modelo pide la herramienta, invoca ese método.'),
  parrafo([negrita('2. ¿Por qué es importante escribir una descripción precisa para cada herramienta?')]),
  parrafo('Porque es lo único que el modelo conoce de la herramienta: con esa descripción decide si le sirve, cuándo pedirla y qué valores enviar. Una descripción vaga o ambigua puede llevar a no usarla, a elegir la equivocada o a enviar mal los argumentos. Por eso a las herramientas nuevas se les añadió además @ToolParam con la descripción de cada parámetro (por ejemplo, que los porcentajes van de 0 a 100). El caso A3-3 muestra que la forma de pedir también influye en que el modelo la use o no.'),
  parrafo([negrita('3. ¿Quién decide qué herramienta solicitar: el controlador o el LLM?')]),
  parrafo('El LLM. El controlador solo registra las herramientas disponibles con .tools(herramientas); nunca llama directamente a calcularPromedio() ni a estadoAcademico(). En la Prueba 3, por ejemplo, el modelo eligió estadoAcademico y no calcularPromedio.'),
  parrafo([negrita('4. ¿Quién ejecuta realmente el método Java?')]),
  parrafo('La aplicación Spring, dentro de la JVM. El modelo solo devuelve una solicitud con el nombre de la herramienta y los argumentos; Spring AI la recibe, invoca el método y devuelve el resultado al modelo para que redacte la respuesta final. La evidencia es el mensaje ">>> TOOL ..." que imprime el propio método en la consola de la aplicación; además, una prueba automatizada con un modelo simulado comprueba el ciclo completo.'),
  parrafo([negrita('5. ¿Qué diferencia existe entre exponer una Tool y llamar directamente un método desde el controlador?')]),
  parrafo('En la llamada directa el programador elige el método y escribe los argumentos en el código: el comportamiento es determinista y no interviene ningún modelo. Al exponer una Tool, el LLM interpreta la intención del usuario, elige entre las herramientas y extrae los argumentos del lenguaje natural; se gana flexibilidad, pero se pierde determinismo (A3-3) y se añaden latencia y costo por los viajes adicionales al modelo.'),
  parrafo([negrita('6. ¿Por qué no todas las preguntas deberían provocar una llamada a herramientas?')]),
  parrafo('Porque muchas se responden directamente (Prueba 4 y actividad 4) y llamar a una herramienta sería innecesario: añade una ida y vuelta más al modelo (más latencia y consumo de tokens), ejecuta código sin necesidad y amplía la superficie de riesgo. Registrar herramientas no obliga a usarlas; el modelo decide.'),
  parrafo([negrita('7. ¿Qué riesgos existirían si una aplicación expusiera como Tools operaciones sensibles sin controles adicionales?')]),
  punto('Ejecución no deseada: el modelo puede equivocarse, alucinar argumentos o ser manipulado con inyección de instrucciones (prompt injection) desde el texto del usuario o desde datos que lea.'),
  punto('Impacto real: borrar o modificar datos, transferir dinero o filtrar información privada, porque la herramienta se ejecuta con los permisos de la aplicación.'),
  punto('Falta de autorización: el modelo no distingue quién es el usuario ni a qué tiene derecho.'),
  punto('Mitigaciones: exponer lo mínimo y preferir herramientas de solo lectura; validar siempre los argumentos en Java; comprobar autenticación y autorización dentro de cada herramienta; pedir confirmación humana en acciones destructivas; limitar frecuencia y registrar cada ejecución para auditoría.'),
);

// ---------- 10. Verificación ----------
add(
  h1('10. Criterios de verificación'),
  tabla(['Criterio', 'Evidencia', 'Verificación'], [
    ['Proyecto ejecuta correctamente', 'Aplicación inicia sin errores (Figura 7) y 16 pruebas automatizadas en verde', 'Sí'],
    ['Tools correctamente declaradas', 'Uso de @Tool y descripciones (Figuras 2 y 3)', 'Sí'],
    ['LLM utiliza calcularPromedio', 'Mensaje de consola en la Prueba 1 (Figura 8)', 'Sí'],
    ['LLM selecciona otra Tool', 'Prueba 3: estadoAcademico (Figura 10)', 'Sí'],
    ['Existe una consulta sin Tool', 'Prueba 4 y actividad 4: respuesta sin ejecución en consola', 'Sí'],
    ['Extensión realizada', 'calcularNotaFinal y estadoAsistencia funcionales (A3-1 y A3-2)', 'Sí'],
  ], [3000, 4726, 1300], { centro: [2] }),
  espacio(),
  parrafo([negrita('Entregables de la guía: '), t('captura de la estructura (Figura 1); código de HerramientasAcademicas.java y ChatController.java (secciones 4 y 5); pruebas.http (sección 5.3); capturas de las respuestas (secciones 7 y 8); captura de consola (Figuras 8, 9, 10 y 18); herramientas adicionales (actividades 1 y 2); respuestas a las preguntas de análisis (sección 9).')]),
);

// ---------- 11. Pruebas automatizadas ----------
add(
  h1('11. Pruebas automatizadas'),
  parrafo('Además de la ejecución manual contra Groq, el repositorio incluye 16 pruebas que se ejecutan sin API ni clave (y por eso funcionan en el CI de GitHub Actions):'),
  punto('HerramientasAcademicasTest (13 pruebas): comprueba la lógica Java pura de cada herramienta, incluidos los límites 3.0 (aprobado) y 80 % / 60 % (asistencia) y el rechazo de porcentajes que no suman 100.'),
  punto('ChatControllerToolCallingTest (3 pruebas): un servidor HTTP local imita el endpoint de chat de un proveedor compatible con OpenAI y responde lo que respondería un LLM (primero pide una herramienta, luego redacta la respuesta final). Verifica que Spring envía las cuatro herramientas al modelo, que ejecuta el método Java, que devuelve el resultado en un mensaje de rol "tool" y que una pregunta conceptual no ejecuta ninguna.'),
);

// ---------- 12. Conclusión ----------
add(
  h1('12. Conclusión'),
  punto([negrita('Mecanismo. '), t('Con @Tool y .tools(herramientas) bastó para que el modelo eligiera entre cuatro métodos Java, extrajera sus argumentos del lenguaje natural y devolviera una respuesta apoyada en el resultado que calculó la aplicación.')]),
  punto([negrita('Control. '), t('El programador define qué puede hacerse (las herramientas y sus reglas, como el mínimo de 3.0); el modelo solo decide cuándo pedirlo. Por eso las validaciones deben vivir en Java y no en el texto del prompt.')]),
  punto([negrita('Límite. '), t('La selección no es obligatoria ni determinista: la Prueba 4 y la actividad 4 no usaron herramientas, y A3-3 muestra un caso en que, aun pudiendo, el modelo no la usó. Esto hay que tenerlo en cuenta al diseñar sistemas donde una operación debe ejecutarse siempre.')]),
  punto([negrita('Siguiente paso. '), t('Las herramientas pueden evolucionar hacia @Service, repositorios JPA, bases de datos, APIs externas, RAG, memoria conversacional y MCP, aplicando siempre los controles de seguridad descritos en la pregunta 7.')]),
);

// ---------- 13. Repositorio ----------
add(
  h1('13. Repositorio y reproducción'),
  parrafo([t('Código, pruebas, evidencias y este informe: '), t(REPO_URL, { color: '2F5496', underline: {} })]),
  ...bloqueCodigo([
    '# 1) Definir la clave (nunca se escribe en el repositorio)',
    '#    PowerShell:  $env:GROQ_API_KEY = "<su-clave>"',
    '#    Bash:        export GROQ_API_KEY=<su-clave>',
    '',
    '# 2) Ejecutar',
    './mvnw spring-boot:run',
    '',
    '# 3) Probar',
    'curl -X POST http://localhost:8080/api/chat -H "Content-Type: text/plain" \\',
    '     -d "Calcula el promedio de las notas 4.0, 3.5 y 4.5"',
    '',
    '# Pruebas automatizadas (no necesitan clave real)',
    './mvnw verify',
  ], 'Ejecución rápida'),
);

const doc = new Document({
  creator: AUTOR,
  title: 'Taller práctico: Tool Calling con Spring AI',
  styles: {
    default: { document: { run: { font: 'Calibri', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 34, bold: true, color: AZUL, font: 'Calibri' }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, color: '2F5496', font: 'Calibri' }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: {
    config: [
      { reference: 'vinetas', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [t('Taller Tool Calling con Spring AI  |  Página ', { size: 18, color: '7F7F7F' }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '7F7F7F' })],
        })],
      }),
    },
    children: hijos,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(SALIDA, buf);
  console.log(`Informe generado: ${SALIDA} (${(buf.length / 1024).toFixed(0)} KB)`);
  if (faltantes.length) console.log(`Capturas pendientes (${faltantes.length}): ${faltantes.join(', ')}`);
});
