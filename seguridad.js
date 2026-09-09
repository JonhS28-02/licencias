'use strict';

// Validación en el límite de entrada; no modifica el archivo de origen.
function validateDatabase(value) {
  const isObject = v => v !== null && typeof v === 'object' && !Array.isArray(v);
  if (!isObject(value)) throw new Error('La base debe contener un objeto de personas.');
  const clean = Object.create(null);
  for (const [rfc, person] of Object.entries(value)) {
    if (!rfc.trim() || !isObject(person) || typeof person.nombre !== 'string' || !isObject(person.fuentes)) {
      throw new Error('La base contiene una persona con estructura inválida.');
    }
    for (const sheets of Object.values(person.fuentes)) {
      if (!isObject(sheets)) throw new Error('La base contiene una fuente inválida.');
      for (const rows of Object.values(sheets)) {
        if (!Array.isArray(rows) || rows.some(row => !isObject(row) || Object.values(row).some(v => v !== null && !['string', 'number', 'boolean'].includes(typeof v)))) {
          throw new Error('La base contiene registros inválidos.');
        }
      }
    }
    clean[rfc] = { ...person, rfc };
  }
  return clean;
}

function validDate(year, month, day) {
  const y = Number(year), m = Number(month), d = Number(day);
  if (![y, m, d].every(Number.isInteger) || y < 1900 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const date = new Date(y, m - 1, d);
  return date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d ? date : null;
}

function csvCell(value) {
  let text = String(value ?? '');
  if (/^[\s\u0000-\u001f]*[=+@-]/.test(text) || /^[\t\r\n]/.test(text)) text = "'" + text;
  return '"' + text.replace(/"/g, '""') + '"';
}

async function readExcelFile(file) {
  if (!/\.xlsx$/i.test(file.name)) throw new Error('Selecciona un archivo .xlsx.');
  if (!file.size || file.size > 20 * 1024 * 1024) throw new Error('El Excel debe pesar entre 1 byte y 20 MB.');
  return file.arrayBuffer();
}

async function readWorkbook(data, options = {}) {
  if (!window.XLSX || !window.JSZip) throw new Error('No se cargaron las bibliotecas de Excel. Recarga la página.');
  if (!data.byteLength || data.byteLength > 20 * 1024 * 1024) throw new Error('El Excel excede el límite de 20 MB o está vacío.');
  const zip = await JSZip.loadAsync(data);
  const entries = Object.values(zip.files);
  // Inspeccionar tamaños declarados antes de descomprimir las hojas.
  if (entries.length > 2000 || entries.reduce((sum, entry) => sum + (entry._data?.uncompressedSize || 0), 0) > 100 * 1024 * 1024) {
    throw new Error('El contenido descomprimido del Excel excede los límites permitidos.');
  }
  const wb = XLSX.read(data, { ...options, type: 'array', cellFormula: false, cellHTML: false });
  for (const sheet of Object.values(wb.Sheets)) {
    if (!sheet['!ref']) continue;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (range.e.r > 50000 || range.e.c > 255 || (range.e.r + 1) * (range.e.c + 1) > 2000000) {
      throw new Error('La hoja excede el límite de filas o columnas permitido.');
    }
  }
  return wb;
}

function setDataStatus(message, ready = false) {
  document.querySelectorAll('[data-load-status]').forEach(el => { el.textContent = message; });
  document.body.classList.toggle('data-ready', ready);
  document.querySelectorAll('.module-nav, .capability-grid button, #si').forEach(el => { el.disabled = !ready; });
}

// Los identificadores se transportan como datos, nunca como código ejecutable.
document.addEventListener('click', async event => {
  const target = event.target.closest('[data-person], [data-copy-rfc], [data-vale]');
  if (!target) return;
  if (target.hasAttribute('data-person')) pick(target.dataset.person);
  if (target.hasAttribute('data-vale')) selectValesPerson(target.dataset.vale);
  if (target.hasAttribute('data-copy-rfc')) {
    try {
      await navigator.clipboard.writeText(target.dataset.copyRfc);
      showToast('RFC copiado');
    } catch {
      showToast('No se pudo copiar. Selecciona el RFC y cópialo manualmente.', 'warn');
    }
  }
});
document.addEventListener('keydown', event => {
  const target = event.target.closest('[data-vale], .faltas-drop');
  if (target && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); target.click(); }
});

// Cambiar el periodo invalida resultados; nunca exportar un cálculo del mes anterior.
document.addEventListener('change', event => {
  if (['faltasMesSel', 'faltasAnioInput'].includes(event.target.id)) {
    cancelImport(document.getElementById('faltasResult'));
    _faltasAnalysis = null;
    _sinNadaCache = null;
    document.getElementById('faltasResult')?.replaceChildren();
    showToast('Periodo actualizado. Vuelve a cargar el listado para analizarlo.');
  }
  if (['vEvalMes', 'vEvalAnio'].includes(event.target.id)) {
    cancelImport(document.getElementById('vFaltasSt'));
    _valesResults = null;
    _valesFaltasMap = null;
    document.getElementById('valesResult')?.replaceChildren();
    const status = document.getElementById('vFaltasSt');
    if (status) status.textContent = 'Carga las faltas del periodo seleccionado';
  }
});

const _importVersions = new WeakMap();
function beginImport(element) {
  const version = Symbol('import');
  _importVersions.set(element, version);
  return () => element.isConnected && _importVersions.get(element) === version;
}
function cancelImport(element) {
  if (element) _importVersions.delete(element);
}
