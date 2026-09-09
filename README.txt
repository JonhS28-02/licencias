Gestión de personal — Hospital de la Mujer

Ejecución local
  python3 servir.py
  Abrir http://127.0.0.1:8000

No requiere compilación ni instalación de paquetes. El servidor local solo
escucha en este equipo y únicamente entrega los archivos de la aplicación;
no publica el directorio, .git ni los demás Excel operativos.
Se necesita conexión para descargar las bibliotecas de los CDN.

Archivos
  index.html: estructura y dependencias fijadas con integridad SHA-384.
  styles.css: interfaz, adaptación móvil y estilos de impresión.
  cs.js: consulta, análisis de faltas, vales y reportes.
  eventos.js: acciones explícitas de los botones y campos.
  seguridad.js: validación, lectura de Excel y protección de exportaciones.
  constancia.js: generación de constancias con ExcelJS y JSZip.
  data.json: base de personal; debe permanecer junto a index.html.
  TEMPLATE_GLOBALES.xlsx: plantilla original de constancias.

Comprobaciones
  node tests/security.test.cjs
  node --check cs.js
  node --check seguridad.js
  node --check eventos.js
  node --check constancia.js
  git diff --check

Datos y operación
  Los Excel cargados se procesan en el navegador y no se envían a terceros.
  Faltas: cargar .xlsx manualmente; comprobar el mes/año antes de exportar.
  Cambiar el periodo requiere volver a cargar el listado correspondiente.
  Vales: cargar BASE VALES y FALTAS del periodo antes de calcular.
  Las tarjetas ambiguas no se usan para justificar ni para decidir vales.
  Las diferencias de nombre conservan las faltas para revisión manual.
  Los archivos admiten hasta 20 MB; hojas de hasta 50.001 filas, 256 columnas
  y 2 millones de celdas declaradas. Contenido ZIP: máximo 100 MB declarados.

Publicación
  No publicar este directorio con datos reales en alojamiento estático público.
  data.json contiene información de personal y es descargable para cualquiera
  que pueda acceder a su URL. Una instalación compartida requiere servidor
  autenticado, autorización y protección también de JSON y plantillas.
  Ver REVISION.md para alcance, cambios y riesgos pendientes.
