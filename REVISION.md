# Revisión del proyecto · 9 de septiembre de 2026

## Correcciones aplicadas

- Se separaron datos y código: RFC y tarjetas viajan en atributos de datos;
  los eventos HTML se sustituyeron por acciones explícitas en `eventos.js`.
  La CSP impide JavaScript en atributos y scripts incrustados, limita conexiones
  al mismo origen y bloquea objetos, formularios y cambios de URL base.
- Se escaparon nombres de fuentes y detalles que aún llegaban sin protección
  al HTML. Los errores de carga no se insertan como HTML.
- Se neutralizan fórmulas en las celdas y encabezados CSV, incluso con espacios
  o controles antes del operador. Se retrasa la liberación de descargas Blob.
- Dependencias PDF actualizadas a jsPDF 4.2.1 y AutoTable 5.0.8; todas las
  bibliotecas remotas tienen versión fija, SHA-384 y CORS anónimo.
- Validación estructural de `data.json` y diccionario sin prototipo. Una base
  inválida no deja los módulos activos con información parcial.
- Validación común de archivos XLSX, tamaño comprimido, tamaños declarados del
  ZIP y dimensiones de hojas. No se importan fórmulas ni HTML de celdas.
- Fechas DMY/ISO con validación de calendario, días de faltas y rangos invertidos;
  las horas Excel conservan su hora literal sin conversión de zona horaria.
- Búsqueda sin distinción de acentos y con prioridad de coincidencias exactas.
  Las tarjetas duplicadas conservan todas sus coincidencias de búsqueda y se
  excluyen del cruce automático. Un nombre distinto no justifica faltas.
- Cálculos de vales incompletos bloqueados; falta de coincidencia en BASE VALES
  requiere revisión. Se rechazan duplicados de las bases importadas y periodos
  de faltas incompatibles cuando el encabezado indica mes/año.
- Cambiar el periodo invalida resultados y evita exportar un cálculo anterior.
  Las importaciones obsoletas, canceladas o de vistas cerradas no aplican cambios.
- Facilidades en ficha y calendario reconocen el año de sus columnas.
- Se retiró la conexión automática que requería compartir públicamente un
  Google Sheet con información de personal; se conserva carga local de XLSX.
- Rediseño con tonos vino/papel, portada editorial, tipografía local, navegación
  móvil con etiquetas, foco visible y uso de teclado en búsqueda y detalles.
- Servidor local de lista permitida, sin índice del directorio, sin caché y
  ligado a 127.0.0.1. Documentación corregida: el script es `cs.js`, no `app.js`.

## Validación

Pruebas de regresión reproducibles: `node tests/security.test.cjs`.
Se comprobó la compatibilidad estructural de la base real sin imprimir datos.
Pruebas de navegador sobre una copia temporal con una persona completamente
ficticia: búsqueda por tarjeta, navegación, PDF individual, tres PDF de reportes,
cruce de faltas con facilidades/licencias/LCGS, Excel corregido, PDF de faltas,
constancia y exportación de vales. Se revisaron escritorio (1440 px) y móvil
(390 px). No se sustituyeron datos ni la plantilla original.

## Límites y pendientes de despliegue

1. **Control de acceso:** una aplicación estática no protege `data.json`.
   Para uso compartido se necesita autenticación y autorización en el servidor,
   HTTPS y protección de cada recurso. El servidor local no es un servidor de
   producción. Si datos o Google Sheets ya se publicaron, restringirlos y revisar
   el acceso desde sus servicios; esta revisión no cambió permisos remotos.
2. **Formatos administrativos:** BASE VALES conserva su distribución de columnas
   y años 2025/2026. Si cambia la plantilla, debe adaptarse su lector. En FALTAS
   las columnas adicionales de omisiones/retardos conservan el orden esperado.
   Si el encabezado no indica periodo, se usa el elegido; verificarlo al cargar.
3. **Reglas de negocio:** se mantuvo la convención existente `10/12 = 10–12`.
   Las discrepancias de nombre y tarjetas requieren revisión humana. El sistema
   no acredita por sí solo la integridad o vigencia de los archivos aportados.
4. **Archivos hostiles:** los límites reducen bloqueos accidentales; no sustituyen
   aislamiento en un Worker ni garantizan protección contra todo ZIP malicioso.
   JSZip usa metadatos internos para inspeccionar tamaños; verificar este control
   al actualizar la dependencia. El lector todavía se ejecuta en el hilo principal.
5. **Dependencias:** SRI detecta cambios en bytes, pero no audita todo el código
   del proveedor. Queda dependencia de disponibilidad de CDN. La CSP conserva
   estilos inline porque el renderizado actual los necesita.
6. Esta revisión es de código y pruebas funcionales dirigidas, no una prueba de
   penetración ni certificación de ausencia total de vulnerabilidades. No incluye
   auditoría del alojamiento, permisos de usuarios ni todos los listados reales.

Fuentes primarias consultadas:
- [jsPDF 4.2.1: correcciones de seguridad](https://github.com/parallax/jsPDF/releases/tag/v4.2.1)
- [AutoTable 5.0.8](https://github.com/simonbengtsson/jsPDF-AutoTable/releases/tag/v5.0.8)
- [Distribución oficial de SheetJS](https://docs.sheetjs.com/docs/getting-started/installation/standalone/)
