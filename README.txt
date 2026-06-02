Estructura refactorizada

- index.html: estructura HTML y carga de librerías jsPDF/autoTable.
- styles.css: todo el diseño visual.
- app.js: lógica de carga, búsqueda, renderizado y generación de PDF.

Cambios aplicados:
1) Se corrigió el choque de nombre PDF.text vs color text, usando PDF.ink.
2) Se aumentó el margen inferior de autoTable para que las tablas no invadan el pie de página.
3) Se ajustaron posiciones del dashboard avanzado en horizontal: antes usaba y=213 en una hoja de 210 mm de alto.
4) Se bajó el riesgo de corte/encimado dejando el contenido por encima del footer.
5) Se añadieron validaciones si jsPDF no carga por bloqueo/CDN.

Mantén data.json junto a index.html, styles.css y app.js.
