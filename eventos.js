'use strict';

// Acciones explícitas: sin eval, Function ni eventos ejecutables dentro del HTML.
const UI_ACTIONS = Object.freeze({
  resetWorkspace: function (event) { resetWorkspace(); },
  openFaltasPanel: function (event) { openFaltasPanel(); },
  openReporteFacilidadesAvanzado: function (event) { openReporteFacilidadesAvanzado('facilidades'); },
  openValesPanel: function (event) { openValesPanel(); },
  focusSearch: function (event) { document.getElementById('si').focus(); },
  genPDF: function (event) { genPDF(); },
  toggleF: function (event) { toggleF(this); },
  genPDFAvanzado: function (event) { genPDFAvanzado(); },
  exportCSV: function (event) { exportCSV(); },
  switchAdvTab: function (event) { switchAdvTab('facilidades'); },
  showLicenciasTab: function (event) { switchAdvTab('licencias'); },
  showLCGSTab: function (event) { switchAdvTab('lcgs'); },
  chooseFaltasFile: function (event) { document.getElementById('faltasFile').click(); },
  handleFaltasFile: function (event) { handleFaltasFile(this); },
  changeFaltasMonth: function (event) { _faltasMes=parseInt(this.value);_sinNadaCache=null;document.getElementById('faltasPanelTitulo').textContent='Análisis de Faltas — '+MESES_FAC[_faltasMes-1]+' '+_faltasAnio; },
  changeFaltasYear: function (event) { _faltasAnio=parseInt(this.value)||_faltasAnio;_sinNadaCache=null;document.getElementById('faltasPanelTitulo').textContent='Análisis de Faltas — '+MESES_FAC[_faltasMes-1]+' '+_faltasAnio; },
  downloadCorregidas: function (event) { downloadCorregidas(); },
  downloadReporteJustificaciones: function (event) { downloadReporteJustificaciones(); },
  downloadReporteSinNada: function (event) { downloadReporteSinNada(); },
  genPDFFaltas: function (event) { genPDFFaltas(); },
  chooseCodigoFile: function (event) { document.getElementById('constanciaBaseCodigoFile').click(); },
  handleConstanciaBaseCodigoFile: function (event) { handleConstanciaBaseCodigoFile(this); },
  generarConstanciaGlobalXLSX: function (event) { generarConstanciaGlobalXLSX(); },
  downloadConstanciaPack: function (event) { downloadConstanciaPack(); },
  changeValesMonth: function (event) { _valesEvalM=parseInt(this.value); },
  changeValesYear: function (event) { _valesEvalY=parseInt(this.value)||_valesEvalY; },
  chooseValesBase: function (event) { document.getElementById('valesBaseFile').click(); },
  handleValesBaseUpload: function (event) { handleValesBaseUpload(this); },
  chooseValesFaltas: function (event) { document.getElementById('valesFaltasFile').click(); },
  handleValesFaltasUpload: function (event) { handleValesFaltasUpload(this); },
  chooseValesNombres: function (event) { document.getElementById('valesNombresFile').click(); },
  clearNamesFilter: function (event) { _valesNombresFilter=null;_valesNombresNoMatch=[];document.getElementById('vNombresSt').textContent='— no cargada';document.getElementById('vNombresSt').style.color='';if(_valesResults)renderValesResultados(); },
  handleValesNombresUpload: function (event) { handleValesNombresUpload(this); },
  runValesAnalysis: function (event) { runValesAnalysis(); },
  clearValesFiltros: function (event) { clearValesFiltros(); },
  downloadValesElegibles: function (event) { downloadValesElegibles(); },
  downloadValesCompleto: function (event) { downloadValesCompleto(); },
  downloadValesPDFEncargado: function (event) { downloadValesPDFEncargado(); },
  closeValesDetalle: function (event) { closeValesDetalle(); },
});

for (const type of ['click', 'change']) {
  document.addEventListener(type, event => {
    const target = event.target.closest(`[data-${type}]`);
    if (!target || target.disabled || typeof target['on' + type] === 'function') return;
    const key = target.getAttribute(`data-${type}`);
    if (!Object.hasOwn(UI_ACTIONS, key)) return;
    const action = UI_ACTIONS[key];
    if (typeof action === 'function') action.call(target, event);
  });
}
