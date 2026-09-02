'use strict';

/* ═══════════════════════════════════════════════════════════
   CONSTANCIA GLOBAL DE FALTAS — generador en el navegador
   Puerto de Integration/generar_constancia.py a JS (ExcelJS + JSZip).
   Genera el mismo documento oficial (hoja "QNA N GLOBAL" + "CARÁTULA")
   a partir del análisis de faltas ya corregido (Análisis de Faltas).

   Estrategia: se carga TEMPLATE_GLOBALES.xlsx (la misma plantilla que
   usaba Python) con ExcelJS para heredar automáticamente fuentes,
   bordes, colores y anchos — no se reinventa el estilo a mano. El
   membrete/logo (guardado como grupo de imágenes, algo que ni ExcelJS
   ni openpyxl saben leer/escribir correctamente) se reinyecta aparte,
   a nivel de archivo .xlsx (zip), exactamente como hacía Python.
═══════════════════════════════════════════════════════════ */

const TEMPLATE_CONSTANCIA_PATH = 'TEMPLATE_GLOBALES.xlsx';
const RENGLONES_POR_HOJA = 25;

const NUMEROS_EN_LETRA = {
  1:'UNO', 2:'DOS', 3:'TRES', 4:'CUATRO', 5:'CINCO', 6:'SEIS', 7:'SIETE', 8:'OCHO', 9:'NUEVE', 10:'DIEZ',
  11:'ONCE', 12:'DOCE', 13:'TRECE', 14:'CATORCE', 15:'QUINCE'
};
function numeroALetra(n) { return NUMEROS_EN_LETRA[n] || String(n); }

const FIRMAS_GLOBAL = [
  ['JEFE DE DEPTO. EN ÁREA MÉDICA "A" RECURSOS HUMANOS', 'MTRA. MARTHA B. AGUILAR BLANCAS'],
  ['COORDINADOR ADMINISTRATIVO  DEL HOSPITAL DE LA MUJER', 'MTRO. LUIS STEEB MORENO SÁNCHEZ.'],
  ['COORDINADORA', 'C. IRMA GUERRERO MARÍN'],
  ['SUBDIRECTORA DE ÁREA EN LA DIRECCIÓN DE PERSONAL', 'LIC. MARIANA MORALES NAVA'],
];
const FIRMAS_CARATULA = [
  ['ELABORÓ HOSPITAL DE LA MUJER', 'MTRA. MARTHA B. AGUILAR BLANCAS', 'JEFA DE DEPTO. EN ÁREA MÉDICA "A" RECURSOS HUMANOS'],
  ['AUTORIZÓ HOSPITAL DE LA MUJER', 'MTRO. LUIS STEEB MORENO SÁNCHEZ', 'COORDINADOR ADMINISTRATIVO  DEL HOSPITAL DE LA MUJER'],
];

/* ═══════════════════════════════════════════════════════════
   BASE DE CÓDIGO DE PUESTO (RFC → código) — opcional, se sube cada
   quincena porque data.json de servicio no trae el código de puesto.
═══════════════════════════════════════════════════════════ */
let _constanciaBaseCodigo   = null; // Map RFC -> {codigo, desc}
let _constanciaBaseAmbiguos = null; // Set de RFC con datos distintos repetidos

async function handleConstanciaBaseCodigoFile(input) {
  const file = input.files[0];
  const estado = document.getElementById('constanciaBaseCodigoEstado');
  if (!file || !estado) return;
  if (!window.XLSX) { showToast('SheetJS no cargó', 'err'); return; }
  try {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames.find(n => n.toUpperCase() === 'HORARIOS') || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });

    let headerRowIdx = -1, col = {};
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const row = rows[i]; if (!row) continue;
      const c = {};
      row.forEach((cell, idx) => { const s = String(cell ?? '').trim().toUpperCase(); if (s && c[s] == null) c[s] = idx; });
      if (c['R.F.C.'] != null && c['CODIGO ACTUAL'] != null) { headerRowIdx = i; col = c; break; }
    }
    if (headerRowIdx < 0) {
      estado.textContent = 'No encontré columnas R.F.C. / CODIGO ACTUAL en el archivo';
      estado.style.color = 'var(--danger)';
      return;
    }

    const crudo = new Map(); // RFC -> [{codigo,desc}, ...]
    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const rfc = String(row[col['R.F.C.']] ?? '').trim().toUpperCase();
      if (!rfc) continue;
      const codigo = row[col['CODIGO ACTUAL']];
      const desc   = col['DESCRIPCION DE CODIGO'] != null ? row[col['DESCRIPCION DE CODIGO']] : null;
      if (!crudo.has(rfc)) crudo.set(rfc, []);
      crudo.get(rfc).push({ codigo: codigo ?? '', desc: desc ?? '' });
    }

    _constanciaBaseCodigo   = new Map();
    _constanciaBaseAmbiguos = new Set();
    for (const [rfc, vistos] of crudo) {
      const uniq = new Set(vistos.map(v => `${v.codigo}|${v.desc}`));
      if (uniq.size === 1) _constanciaBaseCodigo.set(rfc, vistos[0]);
      else _constanciaBaseAmbiguos.add(rfc);
    }
    estado.textContent = `Cargada · ${_constanciaBaseCodigo.size} RFC con código`;
    estado.style.color = 'var(--acc2)';
    showToast('Base de código de puesto cargada');
  } catch (e) {
    estado.textContent = 'Error al leer el archivo';
    estado.style.color = 'var(--danger)';
    console.error(e);
  }
}

/* ═══════════════════════════════════════════════════════════
   HELPERS DE COLUMNAS (A, B, C… ↔ 1, 2, 3…)
═══════════════════════════════════════════════════════════ */
function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function colToIdx(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function copyCellStyle(src, dest) {
  if (src.font)       dest.font       = JSON.parse(JSON.stringify(src.font));
  if (src.border)     dest.border     = JSON.parse(JSON.stringify(src.border));
  if (src.fill)       dest.fill       = JSON.parse(JSON.stringify(src.fill));
  if (src.alignment)  dest.alignment  = JSON.parse(JSON.stringify(src.alignment));
  if (src.protection) dest.protection = JSON.parse(JSON.stringify(src.protection));
  if (src.numFmt)      dest.numFmt    = src.numFmt;
}

/* Clona una hoja (valores + estilos + merges + anchos) de un workbook
   fuente hacia un workbook destino — ExcelJS no trae un copy_worksheet
   como openpyxl, así que se hace celda por celda. */
function cloneWorksheetInto(destWb, srcWs, name, maxCol, maxRow) {
  const newWs = destWb.addWorksheet(name.slice(0, 31));
  if (srcWs.properties && srcWs.properties.defaultRowHeight != null) {
    newWs.properties.defaultRowHeight = srcWs.properties.defaultRowHeight;
  }
  for (let c = 1; c <= maxCol; c++) {
    const w = srcWs.getColumn(c).width;
    if (w != null) newWs.getColumn(c).width = w;
  }
  for (let r = 1; r <= maxRow; r++) {
    const srcRow = srcWs.getRow(r);
    const newRow = newWs.getRow(r);
    if (srcRow.height != null) newRow.height = srcRow.height;
    if (srcRow.hidden) newRow.hidden = true;
    for (let c = 1; c <= maxCol; c++) {
      const srcCell = srcRow.getCell(c);
      const newCell = newRow.getCell(c);
      if (srcCell.value != null) newCell.value = srcCell.value;
      copyCellStyle(srcCell, newCell);
    }
  }
  const merges = (srcWs.model && srcWs.model.merges) || [];
  merges.forEach(range => { try { newWs.mergeCells(range); } catch (e) { /* rango ya cubierto */ } });
  newWs.views = [{ showGridLines: false }];
  return newWs;
}

/* ═══════════════════════════════════════════════════════════
   EXTENSIÓN DE COLUMNAS DE DÍA (si alguien tuvo >4 faltas en la qna)
═══════════════════════════════════════════════════════════ */
function extenderColumnasDia(ws, nDiasNecesarios, colDiaInicio = 8) {
  const colDiaFinActual = colDiaInicio + 3; // H..K = 4 columnas
  if (nDiasNecesarios <= 4) return colDiaFinActual;

  const nuevas = nDiasNecesarios - 4;
  const colEstilo = colLetter(colDiaFinActual); // 'K'

  const merges = (ws.model && ws.model.merges) || [];
  if (merges.includes('H10:K10')) { try { ws.unMergeCells('H10:K10'); } catch (e) {} }

  for (let i = 1; i <= nuevas; i++) {
    const nuevaColIdx = colDiaFinActual + i;
    const nuevaCol = colLetter(nuevaColIdx);
    const wEstilo = ws.getColumn(colDiaFinActual).width;
    if (wEstilo != null) ws.getColumn(nuevaColIdx).width = wEstilo;
    copyCellStyle(ws.getCell(`${colEstilo}10`), ws.getCell(`${nuevaCol}10`));
    const c11 = ws.getCell(`${nuevaCol}11`);
    copyCellStyle(ws.getCell(`${colEstilo}11`), c11);
    c11.value = 'DÍA';
  }
  const nuevoColFin = colDiaFinActual + nuevas;
  ws.mergeCells(`H10:${colLetter(nuevoColFin)}10`);
  return nuevoColFin;
}

/* ═══════════════════════════════════════════════════════════
   BLOQUE DE FIRMAS
═══════════════════════════════════════════════════════════ */
function escribirFirmas(ws, colFinDia) {
  const filaTitulo = 37, filaLinea = 39, filaNombre = 40;
  const colIni = 2; // B
  const totalCols = Math.max(colFinDia - colIni + 1, 8);
  const n = FIRMAS_GLOBAL.length;
  const base = Math.floor(totalCols / n);
  const sobra = totalCols % n;

  const fuente = { name: 'Calibri', size: 11, bold: true };
  const centrado = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const centradoSinWrap = { horizontal: 'center', vertical: 'middle', wrapText: false };

  const merges = (ws.model && ws.model.merges) || [];
  merges.forEach(range => {
    const [a, b] = range.split(':');
    const ra = parseInt(a.match(/\d+/)[0], 10), rb = parseInt((b || a).match(/\d+/)[0], 10);
    if (ra <= filaNombre && rb >= filaTitulo) { try { ws.unMergeCells(range); } catch (e) {} }
  });

  for (let fila = filaTitulo; fila <= filaNombre; fila++) {
    for (let col = colIni; col <= colFinDia; col++) ws.getCell(fila, col).border = {};
  }
  ws.getRow(filaTitulo).height = 28;
  ws.getRow(38).height = 10;
  ws.getRow(filaLinea).height = 10;
  ws.getRow(filaNombre).height = 16;

  let c0 = colIni;
  const anchos = [];
  for (let i = 0; i < n; i++) anchos.push(base + (i === n - 1 ? sobra : 0));

  FIRMAS_GLOBAL.forEach(([puesto, nombre], i) => {
    const anchoBloque = anchos[i];
    const c1 = c0 + anchoBloque - 1;
    const col0 = colLetter(c0), col1 = colLetter(c1);

    ws.mergeCells(`${col0}${filaTitulo}:${col1}${filaTitulo}`);
    const ct = ws.getCell(`${col0}${filaTitulo}`);
    ct.value = puesto; ct.font = fuente; ct.alignment = centrado;

    const largoLinea = Math.max(anchoBloque * 6, 10);
    ws.mergeCells(`${col0}${filaLinea}:${col1}${filaLinea}`);
    const cl = ws.getCell(`${col0}${filaLinea}`);
    cl.value = '_'.repeat(largoLinea); cl.font = fuente; cl.alignment = centradoSinWrap;

    ws.mergeCells(`${col0}${filaNombre}:${col1}${filaNombre}`);
    const cn = ws.getCell(`${col0}${filaNombre}`);
    cn.value = nombre; cn.font = fuente; cn.alignment = centrado;

    c0 = c1 + 1;
  });
}

function escribirFirmasCaratula(wsC) {
  const fuenteEnc = { name: 'Calibri', size: 16, bold: false };
  const fuenteDetalle = { name: 'Calibri', size: 14, bold: false };
  const centrado = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const centradoSinWrap = { horizontal: 'center', vertical: 'middle', wrapText: false };

  const bloques = [[2, 3], [4, 5]];
  const FILA_TITULO = 29, FILA_LINEA = 31, FILA_NOMBRE = 32, FILA_PUESTO = 33;
  const filas = [29, 30, 31, 32, 33, 34];
  const minF = Math.min(...filas), maxF = Math.max(...filas);

  const merges = (wsC.model && wsC.model.merges) || [];
  merges.forEach(range => {
    const [a, b] = range.split(':');
    const ca = a.match(/[A-Z]+/)[0], cb = (b || a).match(/[A-Z]+/)[0];
    const ra = parseInt(a.match(/\d+/)[0], 10), rb = parseInt((b || a).match(/\d+/)[0], 10);
    if (ra <= maxF && rb >= minF && colToIdx(ca) <= 5) { try { wsC.unMergeCells(range); } catch (e) {} }
  });

  for (let fila = minF; fila <= maxF; fila++) {
    for (let col = 2; col <= 5; col++) wsC.getCell(fila, col).border = {};
  }
  wsC.getCell('B31').value = null;
  const rowT = wsC.getRow(FILA_TITULO);
  rowT.height = Math.max(rowT.height || 21, 34);

  bloques.forEach(([c0, c1], i) => {
    const [encabezado, nombre, puesto] = FIRMAS_CARATULA[i];
    const col0 = colLetter(c0), col1 = colLetter(c1);
    const anchoBloque = c1 - c0 + 1;

    wsC.mergeCells(`${col0}${FILA_TITULO}:${col1}${FILA_TITULO}`);
    const ce = wsC.getCell(`${col0}${FILA_TITULO}`);
    ce.value = encabezado; ce.font = fuenteEnc; ce.alignment = centrado;

    const largoLinea = Math.max(anchoBloque * 16, 16);
    wsC.mergeCells(`${col0}${FILA_LINEA}:${col1}${FILA_LINEA}`);
    const cl = wsC.getCell(`${col0}${FILA_LINEA}`);
    cl.value = '_'.repeat(largoLinea); cl.font = fuenteEnc; cl.alignment = centradoSinWrap;

    wsC.mergeCells(`${col0}${FILA_NOMBRE}:${col1}${FILA_NOMBRE}`);
    const cn = wsC.getCell(`${col0}${FILA_NOMBRE}`);
    cn.value = nombre; cn.font = fuenteDetalle; cn.alignment = centrado;

    wsC.mergeCells(`${col0}${FILA_PUESTO}:${col1}34`);
    const cp = wsC.getCell(`${col0}${FILA_PUESTO}`);
    cp.value = puesto; cp.font = fuenteDetalle; cp.alignment = centrado;
  });
}

/* ═══════════════════════════════════════════════════════════
   LLENADO DE UNA HOJA FÍSICA "QNA N GLOBAL Hx"
═══════════════════════════════════════════════════════════ */
function llenarHojaGlobal(ws, empleadosHoja, folioInicialHoja, numHoja, nHojas, esUltimaHoja, config, noEncontrados) {
  ws.getCell('B3').value = `      ${config.unidadNombre}`;
  ws.getCell('B5').value = `UNIDAD RESPONSABLE: ${config.unidadCod}  ${config.unidadNombre}.`;
  ws.getCell('K5').value = config.unidadCod;
  ws.getCell('K2').value = config.noDocumento;
  ws.getCell('H6').value = `           QNA: ${config.qnaNum}`;
  ws.getCell('I7').value = config.mes;
  ws.getCell('I8').value = String(config.anio);
  ws.getCell('B9').value = `                   DEL PERSONAL QUE A CONTINUACIÓN SE DETALLA: ${config.quincenaDesc}`;
  ws.getCell('I9').value = numHoja;
  ws.getCell('K9').value = nHojas;

  const maxDias = empleadosHoja.reduce((m, e) => Math.max(m, e.dias.length), 0);
  const colFinDia = extenderColumnasDia(ws, maxDias);

  const FILA_INICIO_DETALLE = 12, FILA_ULTIMA_PLANTILLA = 36;
  let filaFinalUsada = FILA_INICIO_DETALLE - 1;

  empleadosHoja.forEach((emp, idx) => {
    const fila = FILA_INICIO_DETALLE + idx;
    const folio = folioInicialHoja + idx;
    ws.getCell(`B${fila}`).value = folio;

    const rfc = (emp.rfc || '').trim().toUpperCase();
    let codigo = '';
    if (rfc && _constanciaBaseAmbiguos && _constanciaBaseAmbiguos.has(rfc)) {
      codigo = '(VERIFICAR)';
      noEncontrados.push([emp.nombreDisplay, 'RFC duplicado en la base de código con datos distintos']);
    } else if (rfc && _constanciaBaseCodigo && _constanciaBaseCodigo.has(rfc)) {
      codigo = _constanciaBaseCodigo.get(rfc).codigo || '';
    } else if (_constanciaBaseCodigo) {
      noEncontrados.push([emp.nombreDisplay, 'RFC no está en la base de código de puesto']);
    }

    ws.getCell(`C${fila}`).value = rfc;
    ws.getCell(`D${fila}`).value = codigo;
    ws.getCell(`E${fila}`).value = emp.nombreDisplay;
    const nDias = emp.dias.length;
    ws.getCell(`F${fila}`).value = nDias;
    ws.getCell(`G${fila}`).value = numeroALetra(nDias);
    for (let c = 8; c <= colFinDia; c++) ws.getCell(fila, c).value = null;
    emp.dias.forEach((dia, j) => { ws.getCell(fila, 8 + j).value = dia; });
    filaFinalUsada = fila;
  });

  for (let fila = filaFinalUsada + 1; fila <= FILA_ULTIMA_PLANTILLA; fila++) {
    for (let c = 2; c <= colFinDia; c++) ws.getCell(fila, c).value = null;
  }

  if (esUltimaHoja) escribirFirmas(ws, colFinDia);

  ws.pageSetup = Object.assign({}, ws.pageSetup, { printArea: `B1:${colLetter(colFinDia)}${esUltimaHoja ? 41 : 44}` });
}

/* ═══════════════════════════════════════════════════════════
   REINYECCIÓN DEL MEMBRETE (logo) — puerto de
   restaurar_membrete_original() de generar_constancia.py
═══════════════════════════════════════════════════════════ */
function quitarCuadrosDeTexto(drawingXmlText) {
  const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const doc = new DOMParser().parseFromString(drawingXmlText, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) {
    throw new Error('No se pudo interpretar el XML del membrete de la plantilla');
  }
  const root = doc.documentElement;
  Array.from(root.children).forEach(child => {
    const tieneImagen = child.getElementsByTagNameNS(NS_XDR, 'pic').length > 0;
    if (!tieneImagen) root.removeChild(child);
  });
  // Serializa el ELEMENTO raíz (no el Document) — así nunca se duplica el
  // prólogo "<?xml ...?>" que ya agregamos a mano abajo.
  const serialized = new XMLSerializer().serializeToString(root);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serialized;
}

async function restaurarMembreteOriginal(outputBuffer, templateBuffer) {
  const enc = new TextEncoder(), dec = new TextDecoder('utf-8');
  const zt = await JSZip.loadAsync(templateBuffer);

  const mediaFiles = {};
  for (const name of Object.keys(zt.files)) {
    if (name.startsWith('xl/media/') && !zt.files[name].dir) mediaFiles[name] = await zt.file(name).async('uint8array');
  }
  const drawingCaratula = quitarCuadrosDeTexto(await zt.file('xl/drawings/drawing2.xml').async('string'));
  const relsCaratula    = await zt.file('xl/drawings/_rels/drawing2.xml.rels').async('string');
  const drawingGlobal   = quitarCuadrosDeTexto(await zt.file('xl/drawings/drawing3.xml').async('string'));
  const relsGlobal      = await zt.file('xl/drawings/_rels/drawing3.xml.rels').async('string');

  const zin = await JSZip.loadAsync(outputBuffer);
  const nombresOriginales = new Set(Object.keys(zin.files).filter(n => !zin.files[n].dir));
  const contenidos = {};
  for (const n of nombresOriginales) contenidos[n] = await zin.file(n).async('uint8array');

  const wbXml = dec.decode(contenidos['xl/workbook.xml']);
  const sheets = [...wbXml.matchAll(/<sheet [^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"/g)].map(m => [m[1], m[2]]);
  const wbRelsXml = dec.decode(contenidos['xl/_rels/workbook.xml.rels']);
  const ridToTarget = {};
  for (const m of wbRelsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const idM = m[0].match(/Id="(rId\d+)"/), tgtM = m[0].match(/Target="([^"]+)"/);
    if (idM && tgtM) ridToTarget[idM[1]] = tgtM[1];
  }

  let siguienteDrawing = 1;
  for (const n of nombresOriginales) {
    const m = n.match(/^xl\/drawings\/drawing(\d+)\.xml$/);
    if (m) siguienteDrawing = Math.max(siguienteDrawing, parseInt(m[1], 10) + 1);
  }

  const drawingPorParte = {}, relsPorParte = {};
  const cambiosSheetXml = {}, cambiosSheetRels = {}, cambiosContentTypes = [];

  for (const [nombreHoja, rid] of sheets) {
    const target = ridToTarget[rid] || '';
    const sheetPart = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    const sheetFile = sheetPart.split('/').pop();
    const esCaratula = nombreHoja.toUpperCase().includes('CARATULA') || nombreHoja.includes('CARÁTULA');
    const drawingXml = esCaratula ? drawingCaratula : drawingGlobal;
    const relsXml     = esCaratula ? relsCaratula    : relsGlobal;

    const relsPath = `xl/worksheets/_rels/${sheetFile}.rels`;
    let drawingFile = null;
    if (nombresOriginales.has(relsPath)) {
      const sheetRels = dec.decode(contenidos[relsPath]);
      const mm = sheetRels.match(/Target="[^"]*drawings\/(drawing\d+\.xml)"/);
      if (mm) drawingFile = mm[1];
    }

    if (drawingFile) {
      drawingPorParte[`xl/drawings/${drawingFile}`] = drawingXml;
      relsPorParte[`xl/drawings/_rels/${drawingFile}.rels`] = relsXml;
      continue;
    }

    drawingFile = `drawing${siguienteDrawing}.xml`;
    siguienteDrawing++;
    drawingPorParte[`xl/drawings/${drawingFile}`] = drawingXml;
    relsPorParte[`xl/drawings/_rels/${drawingFile}.rels`] = relsXml;

    let relsActual, nuevoRid;
    if (nombresOriginales.has(relsPath)) {
      relsActual = dec.decode(contenidos[relsPath]);
      const ids = [...relsActual.matchAll(/Id="rId(\d+)"/g)].map(m2 => parseInt(m2[1], 10));
      nuevoRid = `rId${ids.length ? Math.max(...ids) + 1 : 1}`;
      const nuevaRel = `<Relationship Id="${nuevoRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFile}"/>`;
      relsActual = relsActual.replace('</Relationships>', nuevaRel + '</Relationships>');
    } else {
      nuevoRid = 'rId1';
      relsActual = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${nuevoRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingFile}"/></Relationships>`;
    }
    cambiosSheetRels[relsPath] = relsActual;

    let sheetXml = dec.decode(contenidos[sheetPart]);
    if (!sheetXml.includes('<drawing ')) {
      const wsTagStart = sheetXml.indexOf('<worksheet');
      const wsTagEnd = sheetXml.indexOf('>', wsTagStart);
      const wsTagFull = sheetXml.slice(wsTagStart, wsTagEnd + 1);
      if (!wsTagFull.includes('xmlns:r=')) {
        sheetXml = sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
      }
      sheetXml = sheetXml.replace('</worksheet>', `<drawing r:id="${nuevoRid}"/></worksheet>`);
      cambiosSheetXml[sheetPart] = sheetXml;
    }

    cambiosContentTypes.push(`<Override PartName="/xl/drawings/${drawingFile}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
  }

  if (cambiosContentTypes.length) {
    let ct = dec.decode(contenidos['[Content_Types].xml']);
    ct = ct.replace('</Types>', cambiosContentTypes.join('') + '</Types>');
    contenidos['[Content_Types].xml'] = enc.encode(ct);
  }
  for (const [k, v] of Object.entries(cambiosSheetXml))  contenidos[k] = enc.encode(v);
  for (const [k, v] of Object.entries(cambiosSheetRels)) contenidos[k] = enc.encode(v);

  const zout = new JSZip();
  for (const nombre of nombresOriginales) {
    if (nombre.startsWith('xl/media/')) continue;
    if (drawingPorParte[nombre] || relsPorParte[nombre]) continue;
    zout.file(nombre, contenidos[nombre]);
  }
  for (const nombre of Object.keys(cambiosSheetRels)) {
    if (!nombresOriginales.has(nombre)) zout.file(nombre, contenidos[nombre]);
  }
  for (const [nombre, bytes] of Object.entries(mediaFiles)) zout.file(nombre, bytes);
  for (const [parte, xml] of Object.entries(drawingPorParte)) zout.file(parte, typeof xml === 'string' ? enc.encode(xml) : xml);
  for (const [parte, xml] of Object.entries(relsPorParte))    zout.file(parte, typeof xml === 'string' ? enc.encode(xml) : xml);

  return await zout.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

/* ═══════════════════════════════════════════════════════════
   CONSTRUCCIÓN DEL WORKBOOK COMPLETO
═══════════════════════════════════════════════════════════ */
async function construirWorkbookConstancia(empleados, folioInicial, config) {
  const tplResp = await fetch(TEMPLATE_CONSTANCIA_PATH);
  if (!tplResp.ok) throw new Error(`No se pudo cargar ${TEMPLATE_CONSTANCIA_PATH}`);
  const tplBuf = await tplResp.arrayBuffer();

  const tpl = new ExcelJS.Workbook();
  await tpl.xlsx.load(tplBuf);
  const wsBaseTpl = tpl.getWorksheet('QNA 13 GLOBAL');
  const wsCTpl    = tpl.getWorksheet('CARÁTULA  (3)');
  if (!wsBaseTpl || !wsCTpl) throw new Error('La plantilla no tiene las hojas esperadas ("QNA 13 GLOBAL" / "CARÁTULA  (3)")');

  const outWb = new ExcelJS.Workbook();

  const bloques = [];
  for (let i = 0; i < empleados.length; i += RENGLONES_POR_HOJA) bloques.push(empleados.slice(i, i + RENGLONES_POR_HOJA));
  if (!bloques.length) bloques.push([]);
  const nHojas = bloques.length;

  // CARÁTULA primero, para que quede al frente del libro.
  const wsC = cloneWorksheetInto(outWb, wsCTpl, 'CARÁTULA', 12, 49);

  const hojas = bloques.map((_, i) =>
    cloneWorksheetInto(outWb, wsBaseTpl, `QNA ${config.qnaNum} GLOBAL H${i + 1}`, 12, 41)
  );

  const noEncontrados = [];
  let folioCursor = folioInicial;
  hojas.forEach((ws, i) => {
    const esUltima = i === hojas.length - 1;
    llenarHojaGlobal(ws, bloques[i], folioCursor, i + 1, nHojas, esUltima, config, noEncontrados);
    folioCursor += bloques[i].length;
  });

  const folioFinal = empleados.length ? folioInicial + empleados.length - 1 : folioInicial;

  wsC.getCell('K2').value = new Date();
  wsC.getCell('K2').numFmt = 'dd/mm/yyyy';
  wsC.getCell('K6').value = `QNA-${config.qnaNum}`;
  wsC.getCell('B12').value = folioInicial;
  wsC.getCell('B21').value = folioFinal;
  wsC.getCell('K15').value = empleados.length;
  wsC.getCell('K21').value = nHojas;
  wsC.getCell('J1').value = `HOJA NO.   1         DE        ${nHojas}   `;
  escribirFirmasCaratula(wsC);

  const buf = await outWb.xlsx.writeBuffer();
  const conMembrete = await restaurarMembreteOriginal(buf, tplBuf);
  return { buffer: conMembrete, folioFinal, nHojas, noEncontrados };
}

/* ═══════════════════════════════════════════════════════════
   PUNTO DE ENTRADA — botón "Generar Constancia Global"
═══════════════════════════════════════════════════════════ */
function buildConstanciaEmpleados() {
  if (!_faltasAnalysis) return [];
  return _faltasAnalysis.results
    .filter(r => r.diasNoJustificados && r.diasNoJustificados.length)
    .map(r => ({
      nombreDisplay: (r.nombreDB && r.nombreDB !== '—') ? r.nombreDB : r.nombre,
      rfc: r.rfc,
      dias: r.diasNoJustificados.slice().sort((a, b) => a - b)
    }))
    .sort((a, b) => a.nombreDisplay.localeCompare(b.nombreDisplay, 'es'));
}

async function generarConstanciaGlobalXLSX() {
  if (!_faltasAnalysis) { showToast('Primero sube y analiza un archivo de faltas', 'err'); return; }
  if (!window.ExcelJS) { showToast('ExcelJS no cargó. Revisa la conexión a internet.', 'err'); return; }
  if (!window.JSZip)   { showToast('JSZip no cargó. Revisa la conexión a internet.', 'err'); return; }

  const qnaNum   = parseInt(document.getElementById('cQnaNum')?.value, 10);
  const folioIni = parseInt(document.getElementById('cFolioInicial')?.value, 10);
  if (!qnaNum || !folioIni) { showToast('Falta el número de quincena o el folio inicial', 'err'); return; }

  const quincena     = parseInt(document.getElementById('cQuincena')?.value, 10) || 2;
  const noDocumento  = parseInt(document.getElementById('cNoDocumento')?.value, 10) || 8001;
  const unidadCod    = parseInt(document.getElementById('cUnidadCod')?.value, 10) || 160;
  const unidadNombre = (document.getElementById('cUnidadNombre')?.value || 'HOSPITAL DE LA MUJER').trim();
  const mes  = MESES_FAC[_faltasMes - 1];
  const anio = _faltasAnio;
  const quincenaDesc = `${quincena === 1 ? '1A' : '2A'} QNA. DE ${mes} ${anio}`;

  const empleados = buildConstanciaEmpleados();
  if (!empleados.length) { showToast('No hay faltas sin justificar para incluir en la constancia', 'err'); return; }

  showToast('Generando Constancia Global…');
  try {
    const { buffer, folioFinal, nHojas, noEncontrados } = await construirWorkbookConstancia(
      empleados, folioIni, { qnaNum, mes, anio, quincenaDesc, noDocumento, unidadCod, unidadNombre }
    );
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = `CONSTANCIA_GLOBAL_QNA${qnaNum}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);

    showToast(`Constancia generada · ${empleados.length} personas · folios ${folioIni}–${folioFinal} · ${nHojas} hoja(s)`);
    if (noEncontrados.length) {
      setTimeout(() => showToast(`${noEncontrados.length} caso(s) sin código de puesto — revisar a mano`, 'warn'), 400);
    }
  } catch (e) {
    console.error(e);
    showToast('Error al generar la Constancia Global: ' + e.message, 'err');
  }
}
