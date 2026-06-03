
'use strict';

/* ═══════════════════════════════════════════════════════════
   ESTADO GLOBAL
═══════════════════════════════════════════════════════════ */
let DB = null, hist = [], cur = null, acItems = [], acIdx = -1;
let _eventosCache = null;
let _sinNadaCache = null;

/* Toast notifications */
function showToast(msg, type = 'ok') {
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  document.body.appendChild(el);
  requestAnimationFrame(() => el.classList.add('visible'));
  setTimeout(() => { el.classList.remove('visible'); setTimeout(() => el.remove(), 380); }, 2800);
}

/* ═══════════════════════════════════════════════════════════
   UTILIDADES GENERALES
═══════════════════════════════════════════════════════════ */
function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function pad(d) { return String(parseInt(d, 10)).padStart(2, '0'); }
function isExcelTime(s) { return typeof s === 'string' && s.startsWith('1899-12-30T'); }
function isISODate(s)   { return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(s) && !s.startsWith('1899'); }

function fmtTime(s) {
  const d = new Date(s);
  if (isNaN(d)) return s;
  return String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0');
}
function fmtDate(s) {
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
}
function fmtCell(key, val) {
  if (val === null || val === undefined || val === '') return '—';
  const s = String(val).trim();
  if (!s) return '—';
  if (isExcelTime(s)) return fmtTime(s);
  if (isISODate(s))   return fmtDate(s);
  return s;
}
function cellClass(key, val) {
  const s  = String(val ?? '').trim();
  const kl = key.toLowerCase();
  if (isExcelTime(s) || kl === 'entrada' || kl === 'salida') return 'tc';
  if (isISODate(s) || /^\d{2}\/\d{2}\/\d{4}$/.test(s) || kl.includes('fecha')) return 'dc';
  if (s && !isNaN(s) && /^\d+(\.\d+)?$/.test(s)) return 'nc';
  return '';
}
function getIni(n) {
  if (!n) return '?';
  const p = n.replace(/[,\/]/g,' ').split(/\s+/).filter(Boolean);
  return p.length >= 2 ? (p[0][0] + p[1][0]).toUpperCase() : n[0].toUpperCase();
}

/* ── FACILIDADES ── */
function fmtFacilidad(v) {
  if (!v) return '';
  let s = String(v).trim();
  if (!s || s === '.' || s === '}' || /^\s+$/.test(s)) return '';
  if (isISODate(s)) return `<span class="day-chip">${esc(fmtDate(s))}</span>`;
  s = s.replace(/\\/g, '/').replace(/,\s*$/, '').trim();
  const yrMatch = s.match(/\/(\d{4})$/);
  if (!yrMatch) return `<span class="day-chip">${esc(s)}</span>`;
  const yr = yrMatch[1];
  const rest = s.slice(0, s.length - 5);
  const mmMatch = rest.match(/\/(\d{2})$/);
  if (!mmMatch) return `<span class="day-chip">${esc(s)}</span>`;
  const mm = mmMatch[1];
  const daysPart = rest.slice(0, rest.length - 3);
  return parseDayPart(daysPart, mm, yr);
}
function parseDayPart(part, mm, yr) {
  if (!part) return `<span class="day-chip">${mm}/${yr}</span>`;
  const tokens = part.split('/').map(t => t.trim()).filter(Boolean);
  let chips = '';
  for (const t of tokens) {
    if (t.includes('-')) {
      const [a, b] = t.split('-').map(x => x.trim()).filter(Boolean);
      if (a && b && parseInt(b) > parseInt(a)) chips += `<span class="day-range">${pad(a)}–${pad(b)}/${mm}/${yr}</span>`;
      else { if(a) chips += `<span class="day-chip">${pad(a)}/${mm}/${yr}</span>`; if(b) chips += `<span class="day-chip">${pad(b)}/${mm}/${yr}</span>`; }
    } else if (t.includes(',')) {
      t.split(',').forEach(d => chips += `<span class="day-chip">${pad(d.trim())}/${mm}/${yr}</span>`);
    } else {
      chips += `<span class="day-chip">${pad(t)}/${mm}/${yr}</span>`;
    }
  }
  return chips || `<span class="day-chip">${esc(part)}/${mm}/${yr}</span>`;
}

/* ═══════════════════════════════════════════════════════════
   METADATOS DE FUENTE
═══════════════════════════════════════════════════════════ */
function srcDisplayName(n) {
  if (n.includes('Facilidad'))  return 'Facilidades Administrativas 2026';
  if (n.includes('LCGS') || n.includes('2025')) return 'Licencias con Goce de Sueldo';
  if (n.includes('2023') || n.includes('Control') || n.includes('Licencia')) return 'Licencias Médicas';
  return n;
}
function srcColor(n) {
  if (n.includes('2023') || n.includes('Control')) return 'var(--acc)';
  if (n.includes('LCGS') || n.includes('2025'))   return 'var(--acc2)';
  return 'var(--amber)';
}

/* ═══════════════════════════════════════════════════════════
   CARGA DE DATOS
═══════════════════════════════════════════════════════════ */
async function loadData() {
  try {
    const r = await fetch('./data.json?t=' + Date.now());
    if (!r.ok) throw new Error('No se pudo cargar data.json');
    DB = await r.json();
    _eventosCache = null;
    let tr = 0;
    Object.values(DB).forEach(p => Object.values(p.fuentes).forEach(s => Object.values(s).forEach(rr => tr += rr.length)));
    document.getElementById('sp').textContent = Object.keys(DB).length.toLocaleString('es-MX');
    document.getElementById('sr').textContent = tr.toLocaleString('es-MX');
    document.getElementById('ld').style.display = 'none';
    document.getElementById('em').style.display = 'flex';
    setupSearch();
  } catch(e) {
    document.getElementById('ld').innerHTML = `<span style="color:#e05252;font-size:12px">Error: ${e.message}</span>`;
    console.error(e);
  }
}

/* ═══════════════════════════════════════════════════════════
   BÚSQUEDA
═══════════════════════════════════════════════════════════ */
function setupSearch() {
  const inp = document.getElementById('si');
  const ac  = document.getElementById('ac');
  const sc  = document.getElementById('sc');

  inp.addEventListener('input', () => {
    const q = inp.value.trim().toUpperCase();
    sc.style.display = q ? 'block' : 'none';
    if (q.length < 2) { hideAC(); return; }
    const res = searchDB(q, 14);
    if (!res.length) { hideAC(); return; }
    acItems = res; acIdx = -1;
    ac.innerHTML = res.map((p, i) => `
      <div class="aci" data-rfc="${esc(p.rfc)}" data-i="${i}">
        <div class="aci-rfc">${hl(p.rfc, q)}</div>
        <div class="aci-nom">${hl(p.nombre || '(sin nombre)', q)}</div>
        <div class="aci-tags">${Object.keys(p.fuentes).map(s => `<span class="tag">${srcDisplayName(s).split(' ')[0]}</span>`).join('')}</div>
      </div>`).join('');
    ac.querySelectorAll('.aci').forEach(el => el.addEventListener('mousedown', e => {
      e.preventDefault(); pick(el.dataset.rfc);
      inp.value = ''; sc.style.display = 'none';
    }));
    // Pie con total de resultados
    const total = searchDB(q, 9999).length;
    if (total > res.length) {
      ac.innerHTML += `<div style="padding:6px 13px;font-size:9px;color:var(--tx3);font-family:'IBM Plex Mono',monospace;border-top:1px solid var(--brd);text-align:center">+${total - res.length} más — escribe más para filtrar</div>`;
    }
    ac.style.display = 'block';
  });

  inp.addEventListener('keydown', e => {
    if (!acItems.length) return;
    const els = ac.querySelectorAll('.aci');
    if (e.key === 'ArrowDown') { e.preventDefault(); acIdx = Math.min(acIdx + 1, els.length - 1); els.forEach((el, i) => el.classList.toggle('act', i === acIdx)); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); acIdx = Math.max(acIdx - 1, -1); els.forEach((el, i) => el.classList.toggle('act', i === acIdx)); }
    else if (e.key === 'Enter') { e.preventDefault(); const t = acIdx >= 0 ? acItems[acIdx] : acItems[0]; if(t) { pick(t.rfc); inp.value = ''; sc.style.display = 'none'; } }
    else if (e.key === 'Escape') hideAC();
  });

  sc.addEventListener('click', () => { inp.value = ''; sc.style.display = 'none'; hideAC(); inp.focus(); });
  document.addEventListener('click', e => { if (!e.target.closest('#sw')) hideAC(); });

  // Ctrl+K / Cmd+K → enfocar búsqueda
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') { e.preventDefault(); inp.focus(); inp.select(); }
    // Escape → colapsar paneles avanzados y volver al estado vacío
    if (e.key === 'Escape' && !acItems.length) {
      const rp = document.getElementById('rp');
      if (rp && rp.classList.contains('on') && !cur) {
        rp.classList.remove('on'); rp.innerHTML = '';
        document.getElementById('em').style.display = 'flex';
      }
      hideAC();
    }
  });
}

function hideAC() { document.getElementById('ac').style.display = 'none'; acItems = []; acIdx = -1; }

function searchDB(q, lim) {
  const res = [];
  for (const [rfc, d] of Object.entries(DB)) {
    if (res.length >= lim) break;
    if (rfc.includes(q) || (d.nombre && d.nombre.toUpperCase().includes(q))) res.push(d);
  }
  res.sort((a, b) => (a.rfc.startsWith(q) ? 0 : 1) - (b.rfc.startsWith(q) ? 0 : 1));
  return res.slice(0, lim);
}

function hl(t, q) {
  if (!t) return t;
  const i = t.toUpperCase().indexOf(q); if (i < 0) return esc(t);
  return esc(t.slice(0,i)) + `<mark style="background:rgba(31,95,139,.15);color:var(--acc);border-radius:2px">${esc(t.slice(i, i + q.length))}</mark>` + esc(t.slice(i + q.length));
}

/* ═══════════════════════════════════════════════════════════
   SELECCIÓN
═══════════════════════════════════════════════════════════ */
function pick(rfc) {
  hideAC();
  const p = DB[rfc]; if (!p) return;
  cur = rfc;
  const ei = hist.indexOf(rfc); if (ei !== -1) hist.splice(ei, 1);
  hist.unshift(rfc); if (hist.length > 25) hist.pop();
  renderSide();
  renderPerson(p);
}

function renderSide() {
  document.getElementById('rl').innerHTML = hist.map(r => {
    const p = DB[r]; if (!p) return '';
    const dots = Object.keys(p.fuentes).map(s => `<div class="rdot" style="background:${srcColor(s)}"></div>`).join('');
    return `<div class="ri ${r === cur ? 'act' : ''}" onclick="pick('${esc(r)}')">
      <span class="ri-rfc">${esc(r)}</span>
      <span class="ri-nom">${esc(p.nombre || '—')}</span>
      <div class="ri-dots">${dots}</div>
    </div>`;
  }).join('');
}

/* ═══════════════════════════════════════════════════════════
   RENDER PERSONA — HELPERS
═══════════════════════════════════════════════════════════ */
function computePersonStats(p) {
  let licDias=0, licCount=0, lcgsDias=0, lcgsCount=0, facCount=0;
  let tarjeta='—', servicio='—', turno='—';
  for (const [src, sheets] of Object.entries(p.fuentes||{})) {
    for (const recs of Object.values(sheets||{})) {
      for (const rec of (recs||[])) {
        if (isSrcLicMed(src)) {
          licCount++;
          const d = parseFloat(guessDias(rec)); if (!isNaN(d) && d>0) licDias += d;
        } else if (isSrcLCGS(src)) {
          const d = parseFloat(guessDias(rec));
          if (!isNaN(d) && d>0) { lcgsDias+=d; lcgsCount++; }
          else if (hasData([rec['Fecha de Inicio']||''])) lcgsCount++;
        } else if (isSrcFac(src)) {
          const t=guessTarjeta(rec); if(t&&t!=='—') tarjeta=t;
          const sv=String(rec['SERVICIO']||'').trim(); if(sv) servicio=sv;
          const tu=normTurno(String(rec['TURNO']||'').trim()); if(tu&&tu!=='—') turno=tu;
          for (const mes of MESES_FAC) {
            const v=rec[`FACILIDADES ADMINISTRATIVAS ${mes} 2026`];
            if(v&&String(v).trim()&&String(v).trim()!=='.') facCount+=expandFac(v).length;
          }
        }
      }
    }
  }
  return {licDias,licCount,lcgsDias,lcgsCount,facCount,tarjeta,servicio,turno};
}

function renderMiniCal(facValue, mesIdx) {
  const year=2026, daysInMonth=new Date(year,mesIdx+1,0).getDate();
  const firstDow=(new Date(year,mesIdx,1).getDay()+6)%7; // Mon=0
  const highlighted=new Set(), rangeSet=new Set();
  for(const d of expandFac(facValue)){
    const parts=d.split('/');
    if(parts.length<3||parts[1]!==pad(mesIdx+1)||parts[2]!==String(year)) continue;
    if(parts[0].includes('-')){
      const[s,e]=parts[0].split('-').map(x=>parseInt(x,10));
      for(let i=s;i<=e;i++){highlighted.add(i);rangeSet.add(i);}
    } else { highlighted.add(parseInt(parts[0],10)); }
  }
  if(!highlighted.size) return null;
  const DOWS='LMXJVSD';
  let g=`<div class="mini-cal"><div class="mini-cal-row mcr-head">${[...DOWS].map(d=>`<div>${d}</div>`).join('')}</div><div class="mini-cal-row">`;
  for(let i=0;i<firstDow;i++) g+=`<div class="mc-empty"></div>`;
  let col=firstDow;
  for(let d=1;d<=daysInMonth;d++){
    const hi=highlighted.has(d), range=rangeSet.has(d);
    g+=`<div class="mc-day${hi?' mc-hi':''}${range?' mc-range':''}">${d}</div>`;
    col++;
    if(col%7===0&&d<daysInMonth) g+=`</div><div class="mini-cal-row">`;
  }
  return g+`</div></div>`;
}

/* ═══════════════════════════════════════════════════════════
   RENDER PERSONA
═══════════════════════════════════════════════════════════ */
function renderPerson(p) {
  const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  let totalR = 0;
  Object.values(p.fuentes).forEach(s => Object.values(s).forEach(rr => totalR += rr.length));
  const stats = computePersonStats(p);

  let h = `<div class="ph">
    <div class="av">${esc(getIni(p.nombre))}</div>
    <div class="pi">
      <div class="pi-rfc">RFC: ${esc(p.rfc)}${stats.tarjeta!=='—'?` · Tarjeta: <b>${esc(stats.tarjeta)}</b>`:''}</div>
      <div class="pi-nom">${esc(p.nombre || '(Sin nombre registrado)')}</div>
      <div class="pi-chips">${Object.keys(p.fuentes).map(s => `<span class="chip" style="border-color:${srcColor(s)};color:${srcColor(s)}">${esc(srcDisplayName(s))}</span>`).join('')}</div>
      ${stats.servicio!=='—'?`<div style="margin-top:5px;font-size:11px;color:var(--tx2)"><b>${esc(stats.servicio)}</b>${stats.turno!=='—'?' · '+esc(stats.turno):''}</div>`:''}
    </div>
    <div class="tot">
      <div class="tot-n">${totalR}</div>
      <div class="tot-l">registros totales</div>
      <div style="display:flex;gap:6px;margin-top:6px;flex-wrap:wrap;justify-content:flex-end">
        <button class="btn" onclick="genPDF()">⬇ PDF</button>
        <button class="btn sec" onclick="navigator.clipboard.writeText('${esc(p.rfc)}').then(()=>showToast('RFC copiado'))" title="Copiar RFC">⎘ RFC</button>
      </div>
    </div>
  </div>
  <div class="stats-bar">
    <div class="stat-pill stat-navy"><span class="stat-num">${stats.licCount}</span><span class="stat-lbl">Licencias</span>${stats.licDias>0?`<span class="stat-sub">${stats.licDias} días</span>`:''}</div>
    <div class="stat-pill stat-teal"><span class="stat-num">${stats.lcgsCount}</span><span class="stat-lbl">LCGS</span>${stats.lcgsDias>0?`<span class="stat-sub">${stats.lcgsDias} días</span>`:''}</div>
    <div class="stat-pill stat-amber"><span class="stat-num">${stats.facCount}</span><span class="stat-lbl">Facilidades 2026</span><span class="stat-sub">días individuales</span></div>
  </div>`;

  for (const [sname, sheets] of Object.entries(p.fuentes)) {
    const st = Object.values(sheets).reduce((a, rr) => a + rr.length, 0);
    const displayName = srcDisplayName(sname);
    const isFac = sname.includes('Facilidad');

    h += `<div class="fsec">
      <div class="fhdr" onclick="toggleF(this)">
        <div class="fdot" style="background:${srcColor(sname)}"></div>
        <div class="fnam">${esc(displayName)}</div>
        <div class="fcnt">${st} registro${st !== 1 ? 's' : ''}</div>
        <span class="chev">▾</span>
      </div>
      <div class="fbody">`;

    for (const [shName, recs] of Object.entries(sheets)) {
      if (!recs.length) continue;
      h += `<div class="ssec"><div class="sttl">${esc(shName)}<span class="sbdg">${recs.length} reg.</span></div>`;

      if (isFac) {
        const infoKeys = ['TARJETA','TURNO','ENTRADA','SALIDA','SERVICIO'];
        recs.forEach((rec, ri) => {
          if (recs.length > 1) h += `<div style="padding:4px 12px;font-size:9px;color:var(--tx3);font-family:'IBM Plex Mono',monospace;background:var(--soft);border:1px solid var(--brd);border-bottom:none;border-top:${ri>0?'1px solid var(--brd)':'none'}">REGISTRO ${ri + 1}</div>`;
          const infoFields = infoKeys.filter(k => rec[k] !== undefined && rec[k] !== null && rec[k] !== '');
          if (infoFields.length) {
            h += `<div class="fac-info">`;
            infoFields.forEach(k => {
              const vs = String(rec[k] ?? '').trim();
              const isTime = (k === 'ENTRADA' || k === 'SALIDA');
              const display = isTime ? (isExcelTime(vs) ? fmtTime(vs) : vs) : (isISODate(vs) ? fmtDate(vs) : vs);
              h += `<div class="fac-field"><span class="fac-key">${esc(k)}</span><span class="fac-val${isTime?' time':''}">${esc(display)}</span></div>`;
            });
            h += `</div>`;
          }
          h += `<div class="fac-grid">`;
          MESES.forEach((mes, mesIdx) => {
            const key  = `FACILIDADES ADMINISTRATIVAS ${mes} 2026`;
            const val  = rec[key];
            const vs   = String(val ?? '').trim();
            const isEmpty = !vs || vs === '.' || vs === '}' || /^\s+$/.test(vs);
            const cal  = isEmpty ? null : renderMiniCal(vs, mesIdx);
            const chips = isEmpty ? '' : fmtFacilidad(vs);
            h += `<div class="mes-card ${isEmpty ? 'empty-mes' : ''}">
              <div class="mes-nom">${esc(mes)}</div>
              ${isEmpty ? '<div class="mes-val">—</div>' : (cal ? cal : `<div class="day-chips">${chips}</div>`)}
            </div>`;
          });
          h += `</div>`;
        });
      } else {
        const allKeys = [];
        recs.forEach(r => Object.keys(r).forEach(k => { if (!allKeys.includes(k)) allKeys.push(k); }));
        // Excluir claves vacías/ruido
        const displayKeys = allKeys.filter(k => !k.startsWith('EMPTY_') && recs.some(r => { const v=r[k]; return v!==null&&v!==undefined&&String(v).trim()!==''; }));
        // Limpiar nombres de columna con caracteres corruptos para mostrar
        const cleanKey = k => k.replace(/\?/g,'í').replace(/_2$/, ' (2)').replace(/_/g,' ');
        h += `<div class="tw"><table><thead><tr><th>#</th>${displayKeys.map(k => `<th>${esc(cleanKey(k))}</th>`).join('')}</tr></thead><tbody>`;
        recs.forEach((rec, ri) => {
          h += `<tr><td class="rn">${ri + 1}</td>${displayKeys.map(k => {
            const raw = rec[k];
            const v   = fmtCell(k, raw);
            const cls = cellClass(k, String(raw ?? ''));
            return `<td class="${cls}" title="${esc(v)}">${esc(v)}</td>`;
          }).join('')}</tr>`;
        });
        h += `</tbody></table></div>`;
      }
      h += `</div>`; // ssec
    }
    h += `</div></div>`; // fbody + fsec
  }

  const rp = document.getElementById('rp');
  rp.innerHTML = h;
  rp.classList.add('on');
  rp.scrollTop = 0;
  document.getElementById('em').style.display = 'none';
}

function toggleF(hdr) {
  hdr.classList.toggle('col');
  hdr.nextElementSibling.classList.toggle('hid');
}

/* ═══════════════════════════════════════════════════════════
   PDF — PALETA Y HELPERS CENTRALIZADOS
═══════════════════════════════════════════════════════════ */
const PDF = {
  navy:    [31,  95,  139],
  navyDk:  [21,  67,  101],
  teal:    [47,  111,  99],
  amber:   [160, 108,  32],
  slate:   [88,  100, 116],
  ink:     [31,   41,  51],
  muted:   [91,  103, 118],
  light:   [130, 141, 155],
  border:  [216, 224, 234],
  soft:    [246, 248, 251],
  white:   [255, 255, 255],

  fill(doc, c)    { doc.setFillColor(c[0],c[1],c[2]); },
  draw(doc, c)    { doc.setDrawColor(c[0],c[1],c[2]); },
  color(doc, c)   { doc.setTextColor(c[0],c[1],c[2]); },

  text(doc, t, x, y, { size=8, bold=false, color, align, maxWidth } = {}) {
    doc.setFont('helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    this.color(doc, color || this.ink);
    const opt = {};
    if (align)    opt.align    = align;
    if (maxWidth) opt.maxWidth = maxWidth;
    doc.text(String(t ?? ''), x, y, opt);
  },

  wrap(doc, t, x, y, w, { size=7, color, lineH=3.8 } = {}) {
    doc.setFont('helvetica','normal');
    doc.setFontSize(size);
    this.color(doc, color || this.muted);
    const lines = doc.splitTextToSize(String(t ?? ''), w);
    doc.text(lines, x, y);
    return y + lines.length * lineH;
  },

  trunc(doc, t, w) {
    let s = String(t ?? '—');
    if (doc.getTextWidth(s) <= w) return s;
    while (s.length > 1 && doc.getTextWidth(s + '…') > w) s = s.slice(0,-1);
    return s + '…';
  },

  size(doc) { return { w: doc.internal.pageSize.getWidth(), h: doc.internal.pageSize.getHeight() }; },

  header(doc, title, subtitle = '', meta = '') {
    const { w } = this.size(doc);
    this.fill(doc, this.navy);    doc.rect(0,0,w,30,'F');
    this.fill(doc, this.navyDk); doc.rect(0,0,6,30,'F');
    this.text(doc, title,    14, 13, { size:14, bold:true, color:this.white });
    if (subtitle) this.text(doc, subtitle, 14, 20, { size:7.5, color:[220,230,240] });
    if (meta)     this.text(doc, meta, w-12, 13, { size:7, color:[200,215,230], align:'right' });
    this.draw(doc, this.border); doc.setLineWidth(.2); doc.line(14, 34, w-14, 34);
  },

  footer(doc, generated) {
    const n = doc.internal.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i);
      const { w, h } = this.size(doc);
      this.draw(doc, this.border); doc.setLineWidth(.18);
      doc.line(14, h-12, w-14, h-12);
      this.text(doc, `Hospital de la Mujer · Generado: ${generated}`, 14, h-6, { size:6.5, color:this.light });
      this.text(doc, `Página ${i} de ${n}`, w-14, h-6, { size:6.5, color:this.light, align:'right' });
    }
  },

  section(doc, title, y, color = null) {
    const { w } = this.size(doc);
    const c = color || this.navy;
    this.fill(doc, this.soft); this.draw(doc, this.border);
    doc.roundedRect(14, y, w-28, 11, 3, 3, 'FD');
    this.fill(doc, c); doc.roundedRect(14, y, 5, 11, 3, 3, 'F');
    this.text(doc, title, 22, y+7.5, { size:9.5, bold:true });
    return y + 16;
  },

  kpi(doc, x, y, w, label, value, sub, color = null) {
    const c = color || this.navy;
    this.fill(doc, this.white); this.draw(doc, this.border);
    doc.roundedRect(x, y, w, 26, 4, 4, 'FD');
    this.fill(doc, c); doc.roundedRect(x, y, 5, 26, 4, 4, 'F');
    this.text(doc, label, x+9, y+8,  { size:6.5, bold:true, color:this.muted });
    this.text(doc, String(value), x+9, y+19, { size:13, bold:true, color:c });
    this.text(doc, sub, x+w-5, y+19, { size:6, color:this.light, align:'right' });
  },

  ensure(doc, y, need, generated, title, subtitle) {
    const { h, w } = this.size(doc);
    if (y + need <= h - 28) return y;
    doc.addPage(w > h ? 'landscape' : 'portrait');
    this.header(doc, title, subtitle, generated);
    return 40;
  },

  noData(doc, y, text = 'No hay registros disponibles') {
    const { w } = this.size(doc);
    this.fill(doc, this.soft); this.draw(doc, this.border);
    doc.roundedRect(14, y, w-28, 28, 4, 4, 'FD');
    this.text(doc, text, w/2, y+12, { size:11, bold:true, color:this.muted, align:'center' });
    this.text(doc, 'Revisa los datos o cambia los filtros.', w/2, y+21, { size:7, color:this.light, align:'center' });
  },

  safeName(v) {
    return String(v||'REPORTE').replace(/[^\w\sáéíóúÁÉÍÓÚñÑ]/gi,'').replace(/\s+/g,'_').slice(0,80);
  },

  table(doc, startY, head, body, opt = {}) {
    if (!body || !body.length) return startY;
    const generated = opt.generated || '';
    const pageTitle = opt.pageTitle || '';
    const pageSub   = opt.pageSub   || '';
    const margin    = Object.assign({ left:14, right:14, bottom:26, top:42 }, opt.margin || {});
    const firstPage = doc.internal.getCurrentPageInfo().pageNumber;
    const self = this;
    doc.autoTable({
      startY, head:[head], body, theme:'grid',
      margin, tableWidth: opt.tableWidth || 'auto',
      styles: { font:'helvetica', fontSize:opt.fontSize||6.8, cellPadding:opt.cellPadding||2.2, overflow:'linebreak', valign:'middle', textColor:self.ink, lineColor:self.border, lineWidth:.12, minCellHeight:5 },
      headStyles: { fillColor: opt.headColor || self.navy, textColor:[255,255,255], fontStyle:'bold', halign:'center', valign:'middle' },
      alternateRowStyles: { fillColor:[249,251,253] },
      columnStyles: opt.columnStyles || {},
      didDrawPage() {
        const cur = doc.internal.getCurrentPageInfo().pageNumber;
        if (cur !== firstPage && pageTitle) self.header(doc, pageTitle, pageSub, generated);
      }
    });
    return doc.lastAutoTable.finalY + 10;
  },

  personCard(doc, y, p, tarjeta) {
    this.fill(doc, this.white); this.draw(doc, this.border);
    doc.roundedRect(14, y, 182, 38, 4, 4, 'FD');
    this.fill(doc, this.navy); doc.roundedRect(18, y+9, 20, 20, 3, 3, 'F');
    this.text(doc, getIni(p.nombre||p.rfc), 28, y+21, { size:9, bold:true, color:this.white, align:'center' });
    this.text(doc, this.trunc(doc, p.nombre||'SIN NOMBRE REGISTRADO', 136), 44, y+14, { size:11.5, bold:true });
    this.text(doc, `RFC: ${p.rfc}`, 44, y+21, { size:7.8, color:this.muted });
    this.text(doc, `Tarjeta / Nómina: ${tarjeta||'—'}`, 44, y+28, { size:7.8, bold:true, color:this.navy });
    this.text(doc, 'Reporte individual · Licencias médicas, LCGS y facilidades administrativas', 44, y+34, { size:6.5, color:this.light });
    return y + 46;
  },

  hBar(doc, { x, y, w, h, title, data, total, color, limit=5, labelW=55 }) {
    const c = color || this.navy;
    this.fill(doc, this.white); this.draw(doc, this.border);
    doc.roundedRect(x, y, w, h, 4, 4, 'FD');
    this.text(doc, title, x+5, y+8, { size:8.5, bold:true });
    const items = (data||[]).slice(0, limit);
    if (!items.length) { this.text(doc, 'Sin datos', x+5, y+18, { size:7, color:this.light }); return; }
    const rowH = Math.min(10, (h-18) / items.length);
    const max  = Math.max(...items.map(i=>i[1]), 1);
    const bX = x + labelW, valW = 26, bW = Math.max(10, w - labelW - valW - 6);
    let cy = y + 15;
    items.forEach(([name, val]) => {
      const pctVal = total ? (val/total*100) : 0;
      this.text(doc, this.trunc(doc, name, labelW-3), x+4, cy+3.5, { size:6.3, color:this.muted });
      this.fill(doc, [235,240,246]); doc.roundedRect(bX, cy+.5, bW, 4, 2, 2, 'F');
      this.fill(doc, c);             doc.roundedRect(bX, cy+.5, Math.max(1.5,(val/max)*bW), 4, 2, 2, 'F');
      this.text(doc, `${val} · ${pctVal.toFixed(0)}%`, x+w-2, cy+3.5, { size:6, color:this.muted, align:'right' });
      cy += rowH;
    });
  },

  vBar(doc, { x, y, w, h, title, data, total, color, limit=10 }) {
    const c = color || this.amber;
    this.fill(doc, this.white); this.draw(doc, this.border);
    doc.roundedRect(x, y, w, h, 4, 4, 'FD');
    this.text(doc, title, x+5, y+8, { size:8.5, bold:true });
    const items = (data||[]).slice(0, limit);
    if (!items.length) { this.text(doc, 'Sin datos', x+5, y+18, { size:7, color:this.light }); return; }
    const max    = Math.max(...items.map(i=>i[1]), 1);
    const baseY  = y + h - 12;
    const chartH = h - 25;
    const gap = 3;
    const bw  = Math.max(4, (w - 12 - gap*(items.length-1)) / items.length);
    items.forEach(([name, val], idx) => {
      const bx = x + 6 + idx*(bw+gap);
      const bh = Math.max(2, (val/max)*chartH);
      this.fill(doc, c); doc.roundedRect(bx, baseY-bh, bw, bh, 1, 1, 'F');
      this.text(doc, String(val), bx+bw/2, baseY-bh-1.5, { size:5.5, color:this.muted, align:'center' });
      this.text(doc, String(name).split(' ')[0].slice(0,5), bx+bw/2, baseY+5, { size:5.3, color:this.light, align:'center' });
    });
  }
};

/* ═══════════════════════════════════════════════════════════
   HELPERS DE EXTRACCIÓN DE DATOS
═══════════════════════════════════════════════════════════ */
function norm(k) {
  return String(k||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().replace(/\?/g,'I').replace(/\s+/g,' ').trim();
}
/* Normaliza valores de turno — maneja corrupción de encoding y typos del Excel */
function normTurno(raw) {
  if (!raw) return '—';
  let s = String(raw).trim().toUpperCase();
  if (!s || s === '—' || s === 'INSABI') return '—';
  s = s.replace(/VE\?ADA/g,'VELADA').replace(/VELDADA/g,'VELADA').replace(/VEKLADA/g,'VELADA')
       .replace(/^VELDA\b/,'VELADA').replace(/^VELAD\s/,'VELADA ');
  s = s.replace(/^MICTO\b/,'MIXTO').replace(/^MATUITINO\b/,'MATUTINO')
       .replace(/^MATURTINO\b/,'MATUTINO').replace(/^VESPERTIONO\b/,'VESPERTINO');
  s = s.replace(/VELADA\s*["']([ABC])["']?/g,'VELADA $1');
  const bases = ['MATUTINO','VESPERTINO','MIXTO','VELADA FIJA','VELADA A','VELADA B','VELADA C','DISCONTINUO'];
  for (const b of bases) { if (s.startsWith(b)) return b; }
  return s.split(/\s+/)[0] || '—';
}
function valNames(row, names) {
  const wanted = names.map(norm);
  for (const [k,v] of Object.entries(row||{})) {
    if (v===null||v===undefined||String(v).trim()==='') continue;
    if (wanted.includes(norm(k))) return v;
  }
  return '';
}
function valContains(row, words) {
  const wanted = words.map(norm);
  for (const [k,v] of Object.entries(row||{})) {
    if (v===null||v===undefined||String(v).trim()==='') continue;
    const nk = norm(k);
    if (wanted.every(w => nk.includes(w))) return v;
  }
  return '';
}
function firstVal(row, options) {
  for (const op of options) {
    const v = Array.isArray(op) ? valNames(row, op) : valContains(row, [op]);
    if (v!==null && v!==undefined && String(v).trim()!=='') return v;
  }
  return '—';
}
function outDate(v) {
  if (!v || String(v).trim()==='') return '—';
  const s = String(v).trim();
  return isISODate(s) ? fmtDate(s) : s;
}
function outTime(v) {
  if (!v || String(v).trim()==='') return '—';
  const s = String(v).trim();
  return isExcelTime(s) ? fmtTime(s) : s;
}
function hasData(values) {
  return values.some(v => { const s = String(v??'').trim(); return s && s!=='—' && !/^-{2,}/.test(s); });
}
function guessDias(row) {
  if (!row) return '—';
  const posibles = ['# DIAS','# DÍAS','NO DIAS','NO DÍAS','NO. DIAS','NO. DÍAS','N° DIAS','N° DÍAS','NUM DIAS','NUM DÍAS','NUM. DIAS','NUM. DÍAS','NUMERO DE DIAS','NÚMERO DE DÍAS','DIAS','DÍAS','DIA','DÍA','TOTAL DIAS','TOTAL DÍAS','DURACION','DURACIÓN','DÍAS OTORGADOS','DIAS OTORGADOS'];
  const exact = valNames(row, posibles);
  if (exact !== null && exact !== undefined && String(exact).trim() !== '') return String(exact).trim();
  for (const [k, v] of Object.entries(row)) {
    if (v === null || v === undefined || String(v).trim() === '') continue;
    const nk = norm(k);
    if (nk.includes('DIAGNOST')||nk.includes('FECHA')||nk.includes('RFC')||nk.includes('NOMBRE')||nk.includes('TARJETA')) continue;
    if (nk.includes('DIAS')||nk.includes('DIA')||nk.includes('DURACION')) return String(v).trim();
  }
  return '—';
}
function guessTarjeta(row) {
  return valNames(row, ['TARJETA','NO TARJETA','NO. TARJETA','N° TARJETA','NUMERO DE TARJETA','NÚMERO DE TARJETA','TARJETA DE ASISTENCIA','NUM TARJETA','NOMINA','NÓMINA','NO NOMINA','NO. NOMINA']) || '—';
}
function findPersonTarjeta(p) {
  for (const sheets of Object.values(p.fuentes||{}))
    for (const recs of Object.values(sheets||{}))
      for (const row of (recs||[])) { const t = guessTarjeta(row); if (t && t!=='—') return t; }
  return '—';
}

const MESES_FAC = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

function cleanFacVal(v) {
  if (!v) return '';
  let s = String(v).trim();
  if (!s || s==='.' || s==='}' || /^\s+$/.test(s)) return '';
  if (isISODate(s)) return fmtDate(s);
  return s.replace(/\\/g,'/').replace(/,\s*$/,'').trim();
}
function expandFac(v) {
  v = cleanFacVal(v); if (!v) return [];
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v)) { const[d,m,y]=v.split('/'); return [`${pad(d)}/${pad(m)}/${y}`]; }
  const match = v.match(/^(.+)\/(\d{1,2})\/(\d{4})$/);
  if (!match) return [v];
  const mm = pad(match[2]), yr = match[3], out = [];
  match[1].split('/').forEach(tok => {
    tok = tok.trim(); if (!tok) return;
    if (tok.includes('-')) {
      const [a,b] = tok.split('-').map(x=>x.trim()).filter(Boolean);
      out.push(a&&b ? `${pad(a)}-${pad(b)}/${mm}/${yr}` : `${tok}/${mm}/${yr}`);
    } else if (tok.includes(',')) {
      tok.split(',').map(x=>x.trim()).filter(Boolean).forEach(d => out.push(`${pad(d)}/${mm}/${yr}`));
    } else {
      out.push(`${pad(tok)}/${mm}/${yr}`);
    }
  });
  return out.length ? out : [v];
}
function facDias(row) {
  const dias = [];
  MESES_FAC.forEach(mes => {
    const key = Object.keys(row||{}).find(k => { const nk=norm(k); return nk.includes('FACILIDADES ADMINISTRATIVAS') && nk.includes(mes) && nk.includes('2026'); });
    if (key) expandFac(row[key]).forEach(d => dias.push(`${mes}: ${d}`));
  });
  return dias;
}
function isSrcLicMed(src) { const n=norm(src); return (n.includes('CONTROL')||n.includes('LICENCIA'))&&!n.includes('LCGS')&&!n.includes('GOCE')&&!n.includes('FACILIDAD'); }
function isSrcLCGS(src)   { const n=norm(src); return n.includes('LCGS')||n.includes('GOCE'); }
function isSrcFac(src)    { return norm(src).includes('FACILIDAD'); }

/* ═══════════════════════════════════════════════════════════
   PDF INDIVIDUAL
═══════════════════════════════════════════════════════════ */
async function genPDF() {
  if (!cur || !DB[cur]) return;
  if (!window.jspdf || !window.jspdf.jsPDF) { alert('No cargó jsPDF. Revisa conexión/CDN o usa archivos locales.'); return; }
  const p = DB[cur], { jsPDF } = window.jspdf;
  const doc = new jsPDF('portrait','mm','a4');
  const generated = new Date().toLocaleString('es-MX');
  const tarjeta   = findPersonTarjeta(p);

  const licencias=[], lcgs=[], facilidades=[];

  Object.entries(p.fuentes||{}).forEach(([src,sheets]) => Object.values(sheets||{}).forEach(recs => (recs||[]).forEach(r => {
    if (isSrcLicMed(src)) {
      const diagnostico = firstVal(r, [['DIAGNOSTICO','DIAGNÓSTICO','DIAGNOSTICO_2','DIAGNÓSTICO_2'],'DIAGNOSTICO']);
      const dias = guessDias(r);
      const d=valNames(r,['D']), m=valNames(r,['M']), a=valNames(r,['A']);
      const d2=valNames(r,['D_2']), m2=valNames(r,['M_2']), a2=valNames(r,['A_2'])||a;
      const inicio   = (d||m||a)  ? `${d||'--'}/${m||'--'}/${a||'--'}`   : outDate(firstVal(r,[['FECHA'],['FECHA DE INICIO'],['INICIO']]));
      const termino  = (d2||m2||a2) ? `${d2||'--'}/${m2||'--'}/${a2||'--'}` : outDate(firstVal(r,[['FECHA_2'],['FECHA DE TERMINO'],['FECHA TERMINO'],['TERMINO'],['TÉRMINO']]));
      if (hasData([diagnostico,dias]) || /\d/.test(String(inicio)) || /\d/.test(String(termino)))
        licencias.push([diagnostico, inicio, termino, dias]);

    } else if (isSrcLCGS(src)) {
      const row = [
        firstVal(r,[['CONSECUTIVO','NO.','NUMERO','NÚMERO','FOLIO']]),
        guessDias(r),
        outDate(firstVal(r,[['FECHA DE INICIO','FECHA INICIO','INICIO']])),
        outDate(firstVal(r,[['FECHA DE TERMINO','FECHA DE TÉRMINO','FECHA TERMINO','FECHA TÉRMINO','TERMINO','TÉRMINO']])),
        firstVal(r,[['TURNO','TURNO ','JORNADA'],'TURNO'])
      ];
      if (hasData(row)) lcgs.push(row);

    } else if (isSrcFac(src)) {
      const tarj    = guessTarjeta(r);
      const servicio= firstVal(r,[['SERVICIO','ÁREA','AREA']]);
      const turno   = firstVal(r,[['TURNO','TURNO ','JORNADA'],'TURNO']);
      const entrada = outTime(firstVal(r,[['ENTRADA','HORA ENTRADA']]));
      const salida  = outTime(firstVal(r,[['SALIDA','HORA SALIDA']]));
      const dias    = facDias(r);
      if (dias.length) dias.forEach(dia => facilidades.push([tarj, servicio, turno, entrada, salida, dia]));
      else if (hasData([servicio,turno,entrada,salida])) facilidades.push([tarj, servicio, turno, entrada, salida, '—']);
    }
  })));

  const total = licencias.length + lcgs.length + facilidades.length;
  PDF.header(doc, 'Reporte Integral de Licencias', 'Hospital de la Mujer · Control administrativo individual', generated);
  let y = 38;
  y = PDF.personCard(doc, y, p, tarjeta);

  // KPIs
  PDF.kpi(doc, 14, y, 42, 'LICENCIAS',   licencias.length,  'Médicas',            PDF.navy);
  PDF.kpi(doc, 60, y, 42, 'LCGS',        lcgs.length,       'Goce de sueldo',     PDF.teal);
  PDF.kpi(doc,106, y, 42, 'FACILIDADES', facilidades.length,'Administrativas',    PDF.amber);
  PDF.kpi(doc,152, y, 44, 'TOTAL',       total,             'Registros',          PDF.slate);
  y += 36;

  PDF.fill(doc, PDF.soft); PDF.draw(doc, PDF.border);
  doc.roundedRect(14, y, 182, 18, 3, 3, 'FD');
  PDF.text(doc, 'Los registros se distinguen por color. Las tablas largas repiten encabezado al paginar.', 18, y+7, { size:7, color:PDF.muted });
  PDF.text(doc, `Tarjeta/Nómina: ${tarjeta}`, 18, y+14, { size:7.5, bold:true, color:PDF.navy });
  y += 26;

  const opts = { generated, pageTitle:'Reporte Integral de Licencias' };

  if (licencias.length) {
    y = PDF.ensure(doc, y, 50, generated, 'Reporte Integral de Licencias', 'Continuación · Licencias Médicas');
    y = PDF.section(doc, 'Licencias Médicas', y, PDF.navy);
    y = PDF.table(doc, y, ['Diagnóstico','Inicio','Término','# Días'], licencias, { ...opts, pageSub:'Continuación · Licencias Médicas', headColor:PDF.navy, columnStyles:{ 0:{cellWidth:88}, 1:{cellWidth:32,halign:'center'}, 2:{cellWidth:32,halign:'center'}, 3:{cellWidth:22,halign:'center'} } });
  }
  if (lcgs.length) {
    y = PDF.ensure(doc, y, 50, generated, 'Reporte Integral de Licencias', 'Continuación · LCGS');
    y = PDF.section(doc, 'Licencias con Goce de Sueldo', y, PDF.teal);
    y = PDF.table(doc, y, ['Consecutivo','# Días','Inicio','Término','Turno'], lcgs, { ...opts, pageSub:'Continuación · LCGS', headColor:PDF.teal, columnStyles:{ 0:{cellWidth:30,halign:'center'}, 1:{cellWidth:20,halign:'center'}, 2:{cellWidth:36,halign:'center'}, 3:{cellWidth:36,halign:'center'}, 4:{cellWidth:54} } });
  }
  if (facilidades.length) {
    y = PDF.ensure(doc, y, 50, generated, 'Reporte Integral de Licencias', 'Continuación · Facilidades');
    y = PDF.section(doc, 'Facilidades Administrativas', y, PDF.amber);
    y = PDF.table(doc, y, ['Tarjeta','Servicio','Turno','Entrada','Salida','Día de facilidad'], facilidades, { ...opts, pageSub:'Continuación · Facilidades', headColor:PDF.amber, columnStyles:{ 0:{cellWidth:20,halign:'center'}, 1:{cellWidth:38}, 2:{cellWidth:24,halign:'center'}, 3:{cellWidth:18,halign:'center'}, 4:{cellWidth:18,halign:'center'}, 5:{cellWidth:58} } });
  }
  if (!total) PDF.noData(doc, y, 'No se encontraron registros para esta persona');

  PDF.footer(doc, generated);
  doc.save(`REPORTE_${PDF.safeName(p.nombre||p.rfc)}.pdf`);
  showToast(`PDF descargado · ${p.nombre||p.rfc}`);
}

/* ═══════════════════════════════════════════════════════════
   HELPERS DE EXTRACCIÓN — LICENCIAS Y LCGS
═══════════════════════════════════════════════════════════ */
function buildLicencias() {
  const rows = [];
  if (!DB) return rows;
  Object.values(DB).forEach(persona => {
    Object.entries(persona.fuentes||{}).forEach(([src, sheets]) => {
      if (!isSrcLicMed(src)) return;
      Object.values(sheets||{}).forEach(recs => {
        (recs||[]).forEach(r => {
          const diagnostico = String(firstVal(r,[['DIAGNOSTICO','DIAGNÓSTICO','DIAGNOSTICO_2','DIAGNÓSTICO_2'],'DIAGNOSTICO'])||'—').trim();
          const dias        = guessDias(r);
          // Turno viene del campo 'Horario' en licencias médicas (no TURNO)
          const turnoRaw    = valNames(r,['Horario','HORARIO','TURNO','JORNADA']) || '';
          const turno       = normTurno(String(turnoRaw).trim());
          const d=String(valNames(r,['D'])||'').trim(), m=String(valNames(r,['M'])||'').trim();
          const aRaw=String(valNames(r,['A'])||'').trim();
          // Convertir año 2 dígitos → 4 dígitos
          const a = aRaw.length <= 2 && aRaw ? '20' + aRaw.padStart(2,'0') : aRaw;
          const d2=String(valNames(r,['D_2'])||'').trim(), m2=String(valNames(r,['M_2'])||'').trim();
          const a2Raw=String(valNames(r,['A_2'])||aRaw).trim();
          const a2 = a2Raw.length <= 2 && a2Raw ? '20' + a2Raw.padStart(2,'0') : a2Raw;
          const inicio  = (d||m||a)   ? `${d||'--'}/${m||'--'}/${a||'--'}`   : outDate(firstVal(r,[['FECHA'],['FECHA DE INICIO'],['INICIO']]));
          const termino = (d2||m2||a2) ? `${d2||'--'}/${m2||'--'}/${a2||'--'}` : outDate(firstVal(r,[['FECHA_2'],['FECHA DE TERMINO'],['FECHA TERMINO'],['TERMINO'],['TÉRMINO']]));
          if (!hasData([diagnostico]) && !/\d/.test(String(inicio)) && !/\d/.test(String(termino))) return;
          const anio = String(inicio).match(/\d{4}/)?.[0] || (a.length===4 ? a : '—');
          rows.push({ rfc:persona.rfc||'—', nombre:persona.nombre||'SIN NOMBRE', diagnostico, inicio, termino, dias, turno, anio });
        });
      });
    });
  });
  return rows;
}

function buildLCGS() {
  const rows = [];
  if (!DB) return rows;
  Object.values(DB).forEach(persona => {
    Object.entries(persona.fuentes||{}).forEach(([src, sheets]) => {
      if (!isSrcLCGS(src)) return;
      Object.values(sheets||{}).forEach(recs => {
        (recs||[]).forEach(r => {
          const consec  = String(firstVal(r,[['CONSECUTIVO','NO.','NUMERO','NÚMERO','FOLIO']])||'—').trim();
          const dias    = guessDias(r);
          const inicio  = outDate(firstVal(r,[['FECHA DE INICIO','FECHA INICIO','INICIO']]));
          const termino = outDate(firstVal(r,[['FECHA DE TERMINO','FECHA DE TÉRMINO','FECHA TERMINO','FECHA TÉRMINO','TERMINO','TÉRMINO']]));
          const turno   = normTurno(String(firstVal(r,[['TURNO','TURNO ','JORNADA'],'TURNO'])||'').trim());
          if (!hasData([consec, dias, inicio, termino])) return;
          const anio = String(inicio).match(/\d{4}/)?.[0] || '—';
          rows.push({ rfc:persona.rfc||'—', nombre:persona.nombre||'SIN NOMBRE', consec, dias, inicio, termino, turno, anio });
        });
      });
    });
  });
  return rows;
}

/* ═══════════════════════════════════════════════════════════
   HELPERS DE FACILIDADES (con caché)
═══════════════════════════════════════════════════════════ */
function advParseMonthField(key) {
  const nk = norm(key);
  if (!nk.includes('FACILIDADES ADMINISTRATIVAS')) return null;
  const mes  = MESES_FAC.find(m => nk.includes(m));
  const anio = (nk.match(/(20\d{2})/)||[])[1];
  if (!mes || !anio) return null;
  return { mes, mesIndex: MESES_FAC.indexOf(mes), anio };
}
function buildEventos() {
  if (_eventosCache) return _eventosCache;
  const eventos = [];
  if (!DB) return eventos;
  Object.values(DB).forEach(persona => {
    Object.entries(persona.fuentes||{}).forEach(([src,sheets]) => {
      if (!isSrcFac(src)) return;
      Object.entries(sheets||{}).forEach(([,recs]) => {
        (recs||[]).forEach(row => {
          const servicio = String(valNames(row,['SERVICIO','AREA','ÁREA','DEPARTAMENTO'])||'—').trim()||'—';
          const turno    = String(valNames(row,['TURNO','JORNADA'])||'—').trim()||'—';
          const entrada  = outTime(valNames(row,['ENTRADA','HORA ENTRADA']));
          const salida   = outTime(valNames(row,['SALIDA','HORA SALIDA']));
          Object.keys(row||{}).forEach(k => {
            const meta = advParseMonthField(k); if (!meta) return;
            expandFac(row[k]).forEach(dia => eventos.push({
              rfc:persona.rfc||'—', nombre:persona.nombre||'SIN NOMBRE',
              servicio, turno, entrada, salida,
              dia, mes:meta.mes, mesIndex:meta.mesIndex, anio:String(meta.anio)
            }));
          });
        });
      });
    });
  });
  _eventosCache = eventos;
  return eventos;
}
function countMap(rows, keyFn) {
  const m = new Map();
  rows.forEach(r => { const k=keyFn(r)||'—'; m.set(k,(m.get(k)||0)+1); });
  return Array.from(m.entries()).sort((a,b)=>b[1]-a[1]||String(a[0]).localeCompare(String(b[0]),'es'));
}
function pct(n, total) { return total ? (n/total*100) : 0; }

/* ═══════════════════════════════════════════════════════════
   FILTROS FACILIDADES
═══════════════════════════════════════════════════════════ */
function applyFilters(eventos) {
  const g = id => document.getElementById(id)?.value||'TODOS';
  const persona=g('advPersona'), mes=g('advMes'), anio=g('advAnio'), servicio=g('advServicio'), turno=g('advTurno');
  const q = (document.getElementById('advQ')?.value||'').trim().toUpperCase();
  return eventos.filter(e => {
    if (persona!=='TODOS' && e.rfc!==persona) return false;
    if (mes!=='TODOS' && String(e.mesIndex)!==mes) return false;
    if (anio!=='TODOS' && e.anio!==anio) return false;
    if (servicio!=='TODOS' && e.servicio!==servicio) return false;
    if (turno!=='TODOS' && e.turno!==turno) return false;
    if (q && !(e.rfc.toUpperCase().includes(q)||e.nombre.toUpperCase().includes(q)||e.servicio.toUpperCase().includes(q))) return false;
    return true;
  });
}

/* ═══════════════════════════════════════════════════════════
   FILTROS LICENCIAS
═══════════════════════════════════════════════════════════ */
function applyFiltersLic(rows) {
  const g = id => document.getElementById(id)?.value||'TODOS';
  const anio=g('licAnio'), turno=g('licTurno');
  const q = (document.getElementById('licQ')?.value||'').trim().toUpperCase();
  return rows.filter(e => {
    if (anio!=='TODOS' && e.anio!==anio) return false;
    if (turno!=='TODOS' && e.turno!==turno) return false;
    if (q && !(e.rfc.toUpperCase().includes(q)||e.nombre.toUpperCase().includes(q)||e.diagnostico.toUpperCase().includes(q))) return false;
    return true;
  });
}

/* ═══════════════════════════════════════════════════════════
   FILTROS LCGS
═══════════════════════════════════════════════════════════ */
function applyFiltersLCGS(rows) {
  const g = id => document.getElementById(id)?.value||'TODOS';
  const anio=g('lcgsAnio'), turno=g('lcgsTurno');
  const q = (document.getElementById('lcgsQ')?.value||'').trim().toUpperCase();
  return rows.filter(e => {
    if (anio!=='TODOS' && e.anio!==anio) return false;
    if (turno!=='TODOS' && e.turno!==turno) return false;
    if (q && !(e.rfc.toUpperCase().includes(q)||e.nombre.toUpperCase().includes(q))) return false;
    return true;
  });
}

function fillSelect(id, values, labelAll = 'Todos') {
  const el = document.getElementById(id); if (!el) return;
  const current = el.value || 'TODOS';
  el.innerHTML = `<option value="TODOS">${labelAll}</option>` + values.map(v=>`<option value="${esc(v.value)}">${esc(v.label)}</option>`).join('');
  if (Array.from(el.options).some(o=>o.value===current)) el.value = current;
}

/* ═══════════════════════════════════════════════════════════
   SVG HELPERS
═══════════════════════════════════════════════════════════ */
const CHART_COLORS = ['#1f5f8b','#2f6f63','#a06f20','#7a8694','#3a7ab5','#3a8f7c','#c8901a','#5a6874'];

function svgDonut(data, total, size=120) {
  if (!data || !data.length || !total) return '<div style="color:var(--tx3);font-size:11px;padding:16px">Sin datos</div>';
  const r=46, cx=size/2, cy=size/2, stroke=18;
  let angle = -Math.PI/2, paths='', legend='';
  data.slice(0,8).forEach(([name,val],i) => {
    const sweep = (val/total)*2*Math.PI;
    const x1=cx+r*Math.cos(angle), y1=cy+r*Math.sin(angle);
    const x2=cx+r*Math.cos(angle+sweep), y2=cy+r*Math.sin(angle+sweep);
    const large = sweep>Math.PI?1:0;
    const col = CHART_COLORS[i%CHART_COLORS.length];
    paths += `<path d="M${x1.toFixed(2)},${y1.toFixed(2)} A${r},${r} 0 ${large},1 ${x2.toFixed(2)},${y2.toFixed(2)}" fill="none" stroke="${col}" stroke-width="${stroke}" stroke-linecap="butt"><title>${esc(name)}: ${val} (${pct(val,total).toFixed(1)}%)</title></path>`;
    legend += `<div class="donut-leg"><span class="donut-dot" style="background:${col}"></span><span class="donut-lname">${esc(String(name).length>18?String(name).slice(0,17)+'…':name)}</span><span class="donut-lval">${val}</span></div>`;
    angle += sweep;
  });
  return `<div class="donut-wrap">
    <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" style="flex-shrink:0">
      <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#edf1f5" stroke-width="${stroke}"/>
      ${paths}
      <text x="${cx}" y="${cy-4}" text-anchor="middle" font-size="18" font-weight="600" fill="#1f2933">${total}</text>
      <text x="${cx}" y="${cy+10}" text-anchor="middle" font-size="7" fill="#9aa4b2">TOTAL</text>
    </svg>
    <div class="donut-legend">${legend}</div>
  </div>`;
}

function svgHBars(data, total, limit=10) {
  if (!data || !data.length) return '<div class="empty-adv">Sin datos</div>';
  const shown = data.slice(0,limit);
  const max = shown[0]?.[1]||1;
  return '<div class="hbar-list">'+shown.map(([name,val],i)=>{
    const p = pct(val,total);
    const col = CHART_COLORS[i%CHART_COLORS.length];
    return `<div class="hbar-row">
      <div class="hbar-name" title="${esc(name)}">${esc(name)}</div>
      <div class="hbar-track"><div class="hbar-fill" style="width:${Math.max(p,1).toFixed(1)}%;background:${col}"></div></div>
      <div class="hbar-val">${val}<span class="hbar-pct">${p.toFixed(1)}%</span></div>
    </div>`;
  }).join('')+'</div>';
}

/* ═══════════════════════════════════════════════════════════
   PANEL REPORTES AVANZADOS — NAVEGACIÓN POR TABS
═══════════════════════════════════════════════════════════ */
let _advTab = 'facilidades';

function openReporteFacilidadesAvanzado(tabInit = 'facilidades') {
  if (!DB) { alert('Espera a que cargue la base de datos.'); return; }
  _advTab = tabInit;
  const rp = document.getElementById('rp');
  document.getElementById('em').style.display = 'none';
  rp.classList.add('on');

  const eventos   = buildEventos();
  const personas  = Array.from(new Map(eventos.map(e=>[e.rfc,{value:e.rfc,label:`${e.rfc} — ${e.nombre}`}])).values()).sort((a,b)=>a.label.localeCompare(b.label,'es'));
  const aniosFac  = [...new Set(eventos.map(e=>e.anio))].sort().map(y=>({value:y,label:y}));
  const servicios = [...new Set(eventos.map(e=>e.servicio))].sort((a,b)=>a.localeCompare(b,'es')).map(x=>({value:x,label:x}));
  const turnosFac = [...new Set(eventos.map(e=>e.turno))].sort((a,b)=>a.localeCompare(b,'es')).map(x=>({value:x,label:x}));

  const licAll   = buildLicencias();
  const aniosLic = [...new Set(licAll.map(e=>e.anio))].filter(a=>a!=='—').sort().map(y=>({value:y,label:y}));
  const turnosLic= [...new Set(licAll.map(e=>e.turno))].filter(t=>t!=='—').sort().map(t=>({value:t,label:t}));

  const lcgsAll  = buildLCGS();
  const aniosLcgs= [...new Set(lcgsAll.map(e=>e.anio))].filter(a=>a!=='—').sort().map(y=>({value:y,label:y}));
  const turnosLcgs=[...new Set(lcgsAll.map(e=>e.turno))].filter(t=>t!=='—').sort().map(t=>({value:t,label:t}));

  const backBtn = cur ? `<button class="btn sec" onclick="pick('${esc(cur)}')">← Volver a persona</button>` : '';

  rp.innerHTML = `<div class="adv-wrap">
    <div class="adv-head">
      <div class="adv-title">
        <h2>Reportes avanzados</h2>
        <p>Estadísticas globales de facilidades, licencias médicas y LCGS con filtros y exportación.</p>
      </div>
      <div class="adv-actions">
        ${backBtn}
        <button class="btn" id="btnPdfAdv" onclick="genPDFAvanzado()">⬇ PDF filtrado</button>
        <button class="btn sec" id="btnCsvAdv" onclick="exportCSV()">CSV</button>
      </div>
    </div>

    <div class="adv-tabs">
      <button class="adv-tab ${_advTab==='facilidades'?'act':''}" onclick="switchAdvTab('facilidades')">
        <span class="adv-tab-dot" style="background:var(--amber)"></span>Facilidades Administrativas
      </button>
      <button class="adv-tab ${_advTab==='licencias'?'act':''}" onclick="switchAdvTab('licencias')">
        <span class="adv-tab-dot" style="background:var(--acc)"></span>Licencias Médicas
      </button>
      <button class="adv-tab ${_advTab==='lcgs'?'act':''}" onclick="switchAdvTab('lcgs')">
        <span class="adv-tab-dot" style="background:var(--acc2)"></span>Licencias con Goce de Sueldo
      </button>
    </div>

    <!-- TAB: FACILIDADES -->
    <div id="tab-facilidades" class="adv-tab-pane ${_advTab==='facilidades'?'':'hidden'}">
      <div class="adv-filters">
        <div class="adv-field"><label>Persona</label><select id="advPersona"></select></div>
        <div class="adv-field"><label>Mes</label><select id="advMes"><option value="TODOS">Todos</option>${MESES_FAC.map((m,i)=>`<option value="${i}">${m}</option>`).join('')}</select></div>
        <div class="adv-field"><label>Año</label><select id="advAnio"></select></div>
        <div class="adv-field"><label>Área / servicio</label><select id="advServicio"></select></div>
        <div class="adv-field"><label>Turno</label><select id="advTurno"></select></div>
        <div class="adv-field"><label>Buscar</label><input id="advQ" placeholder="RFC, nombre o área…"></div>
      </div>
      <div id="advResult"></div>
    </div>

    <!-- TAB: LICENCIAS MÉDICAS -->
    <div id="tab-licencias" class="adv-tab-pane ${_advTab==='licencias'?'':'hidden'}">
      <div class="adv-filters">
        <div class="adv-field"><label>Año</label><select id="licAnio"></select></div>
        <div class="adv-field"><label>Turno</label><select id="licTurno"></select></div>
        <div class="adv-field adv-field-wide"><label>Buscar RFC, nombre o diagnóstico</label><input id="licQ" placeholder="escribe para filtrar…"></div>
      </div>
      <div id="licResult"></div>
    </div>

    <!-- TAB: LCGS -->
    <div id="tab-lcgs" class="adv-tab-pane ${_advTab==='lcgs'?'':'hidden'}">
      <div class="adv-filters">
        <div class="adv-field"><label>Año</label><select id="lcgsAnio"></select></div>
        <div class="adv-field"><label>Turno</label><select id="lcgsTurno"></select></div>
        <div class="adv-field adv-field-wide"><label>Buscar RFC o nombre</label><input id="lcgsQ" placeholder="escribe para filtrar…"></div>
      </div>
      <div id="lcgsResult"></div>
    </div>
  </div>`;

  // Facilidades
  fillSelect('advPersona', personas, 'Todas las personas');
  fillSelect('advAnio',    aniosFac, 'Todos los años');
  fillSelect('advServicio',servicios,'Todas las áreas');
  fillSelect('advTurno',   turnosFac,'Todos los turnos');
  const now = new Date();
  if (eventos.some(e=>e.mesIndex===now.getMonth())) document.getElementById('advMes').value = String(now.getMonth());
  if (eventos.some(e=>e.anio===String(now.getFullYear()))) document.getElementById('advAnio').value = String(now.getFullYear());
  ['advPersona','advMes','advAnio','advServicio','advTurno'].forEach(id => document.getElementById(id)?.addEventListener('change', renderAdvanzado));
  document.getElementById('advQ')?.addEventListener('input', renderAdvanzado);

  // Licencias — auto-select año más reciente
  fillSelect('licAnio',   aniosLic,  'Todos los años');
  fillSelect('licTurno',  turnosLic, 'Todos los turnos');
  if (aniosLic.length) document.getElementById('licAnio').value = aniosLic[aniosLic.length-1].value;
  document.getElementById('licAnio')?.addEventListener('change', renderLicencias);
  document.getElementById('licTurno')?.addEventListener('change', renderLicencias);
  document.getElementById('licQ')?.addEventListener('input', renderLicencias);

  // LCGS — auto-select año más reciente
  fillSelect('lcgsAnio',  aniosLcgs, 'Todos los años');
  fillSelect('lcgsTurno', turnosLcgs,'Todos los turnos');
  if (aniosLcgs.length) document.getElementById('lcgsAnio').value = aniosLcgs[aniosLcgs.length-1].value;
  document.getElementById('lcgsAnio')?.addEventListener('change', renderLCGS);
  document.getElementById('lcgsTurno')?.addEventListener('change', renderLCGS);
  document.getElementById('lcgsQ')?.addEventListener('input', renderLCGS);

  renderAdvanzado();
  renderLicencias();
  renderLCGS();
}

function switchAdvTab(tab) {
  _advTab = tab;
  document.querySelectorAll('.adv-tab').forEach(b => b.classList.remove('act'));
  document.querySelectorAll('.adv-tab-pane').forEach(p => p.classList.add('hidden'));
  const btn = document.querySelector(`.adv-tab[onclick="switchAdvTab('${tab}')"]`);
  if (btn) btn.classList.add('act');
  const pane = document.getElementById(`tab-${tab}`);
  if (pane) pane.classList.remove('hidden');
  // update PDF/CSV button targets
  const pdfBtn = document.getElementById('btnPdfAdv');
  const csvBtn = document.getElementById('btnCsvAdv');
  if (pdfBtn) pdfBtn.onclick = () => { if(tab==='facilidades') genPDFAvanzado(); else if(tab==='licencias') genPDFLicencias(); else genPDFLCGS(); };
  if (csvBtn) csvBtn.onclick = () => { if(tab==='facilidades') exportCSV(); else if(tab==='licencias') exportCSVLicencias(); else exportCSVLCGS(); };
}

/* ═══════════════════════════════════════════════════════════
   RENDER — FACILIDADES
═══════════════════════════════════════════════════════════ */
function renderAdvanzado() {
  const rows     = applyFilters(buildEventos());
  const total    = rows.length;
  const personas = new Set(rows.map(r=>r.rfc)).size;
  const servicios= new Set(rows.map(r=>r.servicio)).size;
  const porServicio = countMap(rows, r=>r.servicio);
  const porTurno    = countMap(rows, r=>r.turno);
  const porMes      = countMap(rows, r=>`${r.mes} ${r.anio}`);
  const porPersona  = countMap(rows, r=>`${r.rfc} — ${r.nombre}`);
  const topS = porServicio[0]||['—',0];

  const result = document.getElementById('advResult'); if (!result) return;
  result.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi kpi-amber"><div class="kpi-num">${total}</div><div class="kpi-lbl">Facilidades</div><div class="kpi-sub">Eventos filtrados</div></div>
      <div class="kpi kpi-navy"><div class="kpi-num">${personas}</div><div class="kpi-lbl">Personas</div><div class="kpi-sub">RFC únicos</div></div>
      <div class="kpi kpi-teal"><div class="kpi-num">${servicios}</div><div class="kpi-lbl">Áreas</div><div class="kpi-sub">Servicios distintos</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${personas?(total/personas).toFixed(1):'0'}</div><div class="kpi-lbl">Promedio</div><div class="kpi-sub">Facilidades / persona</div></div>
      <div class="kpi kpi-amber"><div class="kpi-num">${pct(topS[1],total).toFixed(1)}%</div><div class="kpi-lbl">Área mayor carga</div><div class="kpi-sub">${esc(topS[0])}</div></div>
    </div>
    <div class="chart-grid-4">
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por área / servicio</h3><span>${porServicio.length} áreas</span></div>
        <div class="adv-card-body">${svgHBars(porServicio, total, 12)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por turno</h3><span>${porTurno.length} turnos</span></div>
        <div class="adv-card-body">${svgDonut(porTurno, total)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por mes</h3><span>${porMes.length} meses</span></div>
        <div class="adv-card-body">${svgHBars(porMes, total, 12)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Personas con más facilidades</h3><span>Top 12</span></div>
        <div class="adv-card-body">${svgHBars(porPersona.map(([k,v])=>[k.split(' — ').slice(1).join(' — ')||k,v]), total, 12)}</div>
      </div>
    </div>
    <div class="adv-card">
      <div class="adv-card-h"><h3>Detalle filtrado</h3><span>${total} filas</span></div>
      ${total ? `<div class="adv-table-wrap"><table><thead><tr><th>#</th><th>RFC</th><th>Nombre</th><th>Área</th><th>Turno</th><th>Entrada</th><th>Salida</th><th>Día</th><th>Mes</th><th>Año</th></tr></thead><tbody>`+
        rows.sort((a,b)=>a.nombre.localeCompare(b.nombre,'es')||a.anio.localeCompare(b.anio)||a.mesIndex-b.mesIndex||a.dia.localeCompare(b.dia,'es')).map((r,i)=>`
          <tr><td class="rn">${i+1}</td><td class="mono acc">${esc(r.rfc)}</td><td>${esc(r.nombre)}</td><td>${esc(r.servicio)}</td><td>${esc(r.turno)}</td><td class="tc">${esc(r.entrada)}</td><td class="tc">${esc(r.salida)}</td><td class="dc">${esc(r.dia)}</td><td>${esc(r.mes)}</td><td class="mono">${esc(r.anio)}</td></tr>`).join('')+
        `</tbody></table></div>` : `<div class="empty-adv">No hay resultados con los filtros seleccionados.</div>`}
      <div class="adv-note">Conteo: <b>12/15/05/2026 = 2 facilidades</b> (días individuales) · <b>12-15/05/2026 = 1 facilidad</b> (rango). Para año completo deja el mes en "Todos".</div>
    </div>`;
}

/* ═══════════════════════════════════════════════════════════
   RENDER — LICENCIAS MÉDICAS
═══════════════════════════════════════════════════════════ */
function renderLicencias() {
  const all  = buildLicencias();
  const rows = applyFiltersLic(all);
  const total = rows.length;
  const personas = new Set(rows.map(r=>r.rfc)).size;
  const porTurno = countMap(rows, r=>r.turno);
  const porAnio  = countMap(rows, r=>r.anio);
  const porDiag  = countMap(rows, r=>r.diagnostico.length>50?r.diagnostico.slice(0,50)+'…':r.diagnostico);
  const diasNum  = rows.map(r=>parseFloat(r.dias)).filter(d=>!isNaN(d));
  const totalDias= diasNum.reduce((a,b)=>a+b,0);

  const result = document.getElementById('licResult'); if (!result) return;
  result.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi kpi-navy"><div class="kpi-num">${total}</div><div class="kpi-lbl">Licencias</div><div class="kpi-sub">Registros filtrados</div></div>
      <div class="kpi kpi-navy"><div class="kpi-num">${personas}</div><div class="kpi-lbl">Personas</div><div class="kpi-sub">RFC únicos</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${totalDias}</div><div class="kpi-lbl">Días totales</div><div class="kpi-sub">Suma de días</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${personas && diasNum.length ? (totalDias/personas).toFixed(1):'0'}</div><div class="kpi-lbl">Días promedio</div><div class="kpi-sub">Por persona</div></div>
    </div>
    <div class="chart-grid-3">
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por turno</h3><span>${porTurno.length} turnos</span></div>
        <div class="adv-card-body">${svgDonut(porTurno, total)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por año</h3><span>${porAnio.length} años</span></div>
        <div class="adv-card-body">${svgHBars(porAnio, total, 8)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Diagnósticos más frecuentes</h3><span>Top 10</span></div>
        <div class="adv-card-body">${svgHBars(porDiag, total, 10)}</div>
      </div>
    </div>
    <div class="adv-card">
      <div class="adv-card-h"><h3>Detalle filtrado</h3><span>${total} registros</span></div>
      ${total ? `<div class="adv-table-wrap"><table><thead><tr><th>#</th><th>RFC</th><th>Nombre</th><th>Diagnóstico</th><th>Inicio</th><th>Término</th><th># Días</th><th>Turno</th><th>Año</th></tr></thead><tbody>`+
        rows.map((r,i)=>`<tr><td class="rn">${i+1}</td><td class="mono acc">${esc(r.rfc)}</td><td>${esc(r.nombre)}</td><td style="max-width:220px;white-space:normal">${esc(r.diagnostico)}</td><td class="dc">${esc(r.inicio)}</td><td class="dc">${esc(r.termino)}</td><td class="nc">${esc(r.dias)}</td><td>${esc(r.turno)}</td><td class="mono">${esc(r.anio)}</td></tr>`).join('')+
        `</tbody></table></div>` : `<div class="empty-adv">No hay resultados con los filtros seleccionados.</div>`}
    </div>`;
}

/* ═══════════════════════════════════════════════════════════
   RENDER — LCGS
═══════════════════════════════════════════════════════════ */
function renderLCGS() {
  const all  = buildLCGS();
  const rows = applyFiltersLCGS(all);
  const total = rows.length;
  const personas = new Set(rows.map(r=>r.rfc)).size;
  const porTurno = countMap(rows, r=>r.turno);
  const porAnio  = countMap(rows, r=>r.anio);
  const diasNum  = rows.map(r=>parseFloat(r.dias)).filter(d=>!isNaN(d));
  const totalDias= diasNum.reduce((a,b)=>a+b,0);

  const result = document.getElementById('lcgsResult'); if (!result) return;
  result.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi kpi-teal"><div class="kpi-num">${total}</div><div class="kpi-lbl">Registros LCGS</div><div class="kpi-sub">Filtrados</div></div>
      <div class="kpi kpi-teal"><div class="kpi-num">${personas}</div><div class="kpi-lbl">Personas</div><div class="kpi-sub">RFC únicos</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${totalDias}</div><div class="kpi-lbl">Días totales</div><div class="kpi-sub">Suma de días</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${personas && diasNum.length ? (totalDias/personas).toFixed(1):'0'}</div><div class="kpi-lbl">Días promedio</div><div class="kpi-sub">Por persona</div></div>
    </div>
    <div class="chart-grid-2">
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por turno</h3><span>${porTurno.length} turnos</span></div>
        <div class="adv-card-body">${svgDonut(porTurno, total)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Por año</h3><span>${porAnio.length} años</span></div>
        <div class="adv-card-body">${svgHBars(porAnio, total, 8)}</div>
      </div>
    </div>
    <div class="adv-card">
      <div class="adv-card-h"><h3>Detalle filtrado</h3><span>${total} registros</span></div>
      ${total ? `<div class="adv-table-wrap"><table><thead><tr><th>#</th><th>RFC</th><th>Nombre</th><th>Consecutivo</th><th># Días</th><th>Inicio</th><th>Término</th><th>Turno</th><th>Año</th></tr></thead><tbody>`+
        rows.map((r,i)=>`<tr><td class="rn">${i+1}</td><td class="mono acc">${esc(r.rfc)}</td><td>${esc(r.nombre)}</td><td class="nc">${esc(r.consec)}</td><td class="nc">${esc(r.dias)}</td><td class="dc">${esc(r.inicio)}</td><td class="dc">${esc(r.termino)}</td><td>${esc(r.turno)}</td><td class="mono">${esc(r.anio)}</td></tr>`).join('')+
        `</tbody></table></div>` : `<div class="empty-adv">No hay resultados con los filtros seleccionados.</div>`}
    </div>`;
}

/* ── HELPERS FILTRO TEXTO ── */
function filterText() {
  const txt = [];
  [['Persona','advPersona'],['Mes','advMes'],['Año','advAnio'],['Área','advServicio'],['Turno','advTurno']].forEach(([lbl,id]) => {
    const el = document.getElementById(id);
    if (el && el.value !== 'TODOS') txt.push(`${lbl}: ${el.options[el.selectedIndex]?.text}`);
  });
  const q = document.getElementById('advQ')?.value.trim(); if (q) txt.push(`Búsqueda: ${q}`);
  return txt.length ? txt.join(' · ') : 'Todos los registros';
}

/* ═══════════════════════════════════════════════════════════
   PDF AVANZADO — FACILIDADES
═══════════════════════════════════════════════════════════ */
function genPDFAvanzado() {
  if (!DB) { alert('Espera a que cargue la base de datos.'); return; }
  if (!window.jspdf || !window.jspdf.jsPDF) { alert('No cargó jsPDF.'); return; }
  const rows = applyFilters(buildEventos()).slice().sort((a,b)=>a.nombre.localeCompare(b.nombre,'es')||a.anio.localeCompare(b.anio)||a.mesIndex-b.mesIndex||a.dia.localeCompare(b.dia,'es'));
  const { jsPDF } = window.jspdf;
  const doc       = new jsPDF('landscape','mm','a4');
  const generated = new Date().toLocaleString('es-MX');

  const total    = rows.length;
  const personas = new Set(rows.map(r=>r.rfc)).size;
  const servicios= new Set(rows.map(r=>r.servicio)).size;
  const turnos   = new Set(rows.map(r=>r.turno)).size;
  const promedio = personas ? (total/personas).toFixed(1) : '0';
  const porServicio= countMap(rows,r=>r.servicio);
  const porTurno   = countMap(rows,r=>r.turno);
  const porMes     = countMap(rows,r=>`${r.mes} ${r.anio}`);
  const porPersona = countMap(rows,r=>`${r.rfc} — ${r.nombre}`);
  const topS = porServicio[0]||['—',0];
  const topP = porPersona[0]||['—',0];

  /* PÁG 1: DASHBOARD */
  PDF.header(doc,'Reporte Avanzado · Facilidades Administrativas','Hospital de la Mujer · Dashboard visual filtrado',generated);
  PDF.wrap(doc,`Filtros: ${filterText()}`,14,36,268,{size:7});

  PDF.kpi(doc, 14, 42, 54, 'FACILIDADES', String(total),    'Eventos reales',  PDF.amber);
  PDF.kpi(doc, 72, 42, 54, 'PERSONAS',    String(personas), 'RFC únicos',      PDF.navy);
  PDF.kpi(doc,130, 42, 54, 'ÁREAS',       String(servicios),'Servicios',       PDF.teal);
  PDF.kpi(doc,188, 42, 54, 'TURNOS',      String(turnos),   'Jornadas',        PDF.slate);
  PDF.kpi(doc,246, 42, 40, 'PROM.',       String(promedio), 'Fac./persona',    PDF.navyDk);

  PDF.section(doc, 'Resumen visual', 74, PDF.navy);
  PDF.hBar(doc,{x:14,  y:88, w:136, h:50, title:'Top áreas / servicios',            data:porServicio, total, color:PDF.teal,  limit:5, labelW:60});
  PDF.hBar(doc,{x:158, y:88, w:129, h:50, title:'Top personas con más facilidades', data:porPersona,  total, color:PDF.navy,  limit:5, labelW:76});
  PDF.hBar(doc,{x:14, y:144, w:130, h:36, title:'Distribución por turno',           data:porTurno,    total, color:PDF.slate, limit:4, labelW:46});
  PDF.vBar(doc,{x:158,y:144, w:129, h:36, title:'Comportamiento por mes',           data:porMes,      total, color:PDF.amber, limit:10});

  PDF.fill(doc, PDF.soft); PDF.draw(doc, PDF.border);
  doc.roundedRect(14,184,268,15,3,3,'FD');
  PDF.text(doc,'Lectura rápida',18,190,{size:8,bold:true});
  PDF.text(doc,`Área con mayor carga: ${topS[0]} (${pct(topS[1],total).toFixed(1)}%).  Persona con más eventos: ${topP[0]} (${topP[1]}).`,18,196,{size:7,color:PDF.muted,maxWidth:258});

  /* PÁG 2: RESUMEN TABULAR */
  doc.addPage('landscape');
  PDF.header(doc,'Resumen tabular · Facilidades',`Filtros: ${filterText()}`,generated);

  if (!rows.length) {
    PDF.noData(doc,45,'No hay resultados con los filtros seleccionados');
  } else {
    let y2 = 42;
    y2 = PDF.table(doc,y2,['Área / servicio','Total','%'],porServicio.map(([k,v])=>[k,v,`${pct(v,total).toFixed(1)}%`]),{generated,tableWidth:130,margin:{left:14,right:155,bottom:26,top:42},fontSize:7,headColor:PDF.teal,columnStyles:{0:{cellWidth:90},1:{cellWidth:16,halign:'center'},2:{cellWidth:24,halign:'center'}}});
    PDF.table(doc,42,['Turno','Total','%'],porTurno.map(([k,v])=>[k,v,`${pct(v,total).toFixed(1)}%`]),{generated,tableWidth:60,margin:{left:156,right:81,bottom:26,top:42},fontSize:7,headColor:PDF.slate,columnStyles:{0:{cellWidth:30},1:{cellWidth:12,halign:'center'},2:{cellWidth:18,halign:'center'}}});
    PDF.table(doc,42,['Mes','Total','%'],porMes.slice(0,12).map(([k,v])=>[k,v,`${pct(v,total).toFixed(1)}%`]),{generated,tableWidth:60,margin:{left:228,right:14,bottom:26,top:42},fontSize:7,headColor:PDF.amber,columnStyles:{0:{cellWidth:30},1:{cellWidth:12,halign:'center'},2:{cellWidth:18,halign:'center'}}});

    const bottomY = Math.max(doc.lastAutoTable?.finalY + 14 || 100, 100);
    const { h: pageH } = PDF.size(doc);
    if (bottomY + 60 < pageH - 26) {
      PDF.section(doc,'Top personas con más facilidades', bottomY, PDF.navy);
      PDF.table(doc, bottomY+15, ['#','RFC','Nombre','Total','%'],
        porPersona.slice(0,15).map(([k,v],i) => { const parts=k.split(' — '); return [i+1,parts[0]||'—',parts.slice(1).join(' — ')||'—',v,`${pct(v,total).toFixed(1)}%`]; }),
        {generated,tableWidth:268,margin:{left:14,right:14,bottom:26,top:42},fontSize:7,headColor:PDF.navy,columnStyles:{0:{cellWidth:10,halign:'center'},1:{cellWidth:34},2:{cellWidth:170},3:{cellWidth:22,halign:'center'},4:{cellWidth:22,halign:'center'}}});
    } else {
      doc.addPage('landscape');
      PDF.header(doc,'Top personas · Facilidades',`Filtros: ${filterText()}`,generated);
      PDF.table(doc, 42, ['#','RFC','Nombre','Total','%'],
        porPersona.slice(0,15).map(([k,v],i) => { const parts=k.split(' — '); return [i+1,parts[0]||'—',parts.slice(1).join(' — ')||'—',v,`${pct(v,total).toFixed(1)}%`]; }),
        {generated,tableWidth:268,margin:{left:14,right:14,bottom:26,top:42},fontSize:7,headColor:PDF.navy,columnStyles:{0:{cellWidth:10,halign:'center'},1:{cellWidth:34},2:{cellWidth:170},3:{cellWidth:22,halign:'center'},4:{cellWidth:22,halign:'center'}}});
    }
  }

  /* PÁG: DETALLE */
  if (rows.length) {
    doc.addPage('landscape');
    PDF.header(doc,'Detalle de eventos filtrados',`Total de filas: ${rows.length}`,generated);
    PDF.table(doc,42,['#','RFC','Nombre','Área','Turno','Entrada','Salida','Día','Mes','Año'],
      rows.map((r,i)=>[i+1,r.rfc,r.nombre,r.servicio,r.turno,r.entrada,r.salida,r.dia,r.mes,r.anio]),
      {generated,pageTitle:'Detalle de eventos filtrados',pageSub:'Continuación',fontSize:5.8,cellPadding:1.6,headColor:PDF.navy,
       columnStyles:{0:{cellWidth:9,halign:'center'},1:{cellWidth:24},2:{cellWidth:50},3:{cellWidth:48},4:{cellWidth:20,halign:'center'},5:{cellWidth:17,halign:'center'},6:{cellWidth:17,halign:'center'},7:{cellWidth:28},8:{cellWidth:24},9:{cellWidth:14,halign:'center'}},
       margin:{left:14,right:14,bottom:26,top:42}});
  }

  PDF.footer(doc, generated);
  doc.save(`REPORTE_FACILIDADES_${new Date().toISOString().slice(0,10)}.pdf`);
}

/* ═══════════════════════════════════════════════════════════
   PDF LICENCIAS MÉDICAS
═══════════════════════════════════════════════════════════ */
function genPDFLicencias() {
  if (!window.jspdf?.jsPDF) { alert('No cargó jsPDF.'); return; }
  const rows = applyFiltersLic(buildLicencias()).sort((a,b)=>a.nombre.localeCompare(b.nombre,'es'));
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('landscape','mm','a4');
  const generated = new Date().toLocaleString('es-MX');

  const total   = rows.length;
  const personas= new Set(rows.map(r=>r.rfc)).size;
  const diasNum = rows.map(r=>parseFloat(r.dias)).filter(d=>!isNaN(d));
  const totalDias = diasNum.reduce((a,b)=>a+b,0);
  const porTurno= countMap(rows,r=>r.turno);
  const porAnio = countMap(rows,r=>r.anio);
  const porDiag = countMap(rows,r=>r.diagnostico.length>60?r.diagnostico.slice(0,60)+'…':r.diagnostico);

  PDF.header(doc,'Reporte · Licencias Médicas','Hospital de la Mujer · Datos filtrados',generated);
  PDF.kpi(doc, 14, 40, 58, 'LICENCIAS',   String(total),     'Registros',     PDF.navy);
  PDF.kpi(doc, 76, 40, 58, 'PERSONAS',    String(personas),  'RFC únicos',    PDF.navy);
  PDF.kpi(doc,138, 40, 58, 'DÍAS TOTAL',  String(totalDias), 'Acumulado',     PDF.slate);
  PDF.kpi(doc,200, 40, 58, 'DÍAS PROM.',  personas&&diasNum.length?(totalDias/personas).toFixed(1):'0','Por persona', PDF.slate);

  PDF.section(doc, 'Distribución por turno y año', 72, PDF.navy);
  PDF.hBar(doc,{x:14,  y:86, w:130, h:44, title:'Por turno', data:porTurno, total, color:PDF.navy, limit:5, labelW:50});
  PDF.hBar(doc,{x:152, y:86, w:135, h:44, title:'Por año',   data:porAnio,  total, color:PDF.slate,limit:8, labelW:30});

  PDF.section(doc, 'Diagnósticos más frecuentes', 136, PDF.navy);
  PDF.hBar(doc,{x:14, y:150, w:273, h:48, title:'Top diagnósticos', data:porDiag, total, color:PDF.navy, limit:8, labelW:120});

  if (rows.length) {
    doc.addPage('landscape');
    PDF.header(doc,'Detalle · Licencias Médicas',`Total: ${rows.length} registros`,generated);
    PDF.table(doc,42,['#','RFC','Nombre','Diagnóstico','Inicio','Término','# Días','Turno','Año'],
      rows.map((r,i)=>[i+1,r.rfc,r.nombre,r.diagnostico,r.inicio,r.termino,r.dias,r.turno,r.anio]),
      {generated,pageTitle:'Detalle · Licencias Médicas',pageSub:'Continuación',fontSize:5.8,cellPadding:1.6,headColor:PDF.navy,
       columnStyles:{0:{cellWidth:9,halign:'center'},1:{cellWidth:24},2:{cellWidth:50},3:{cellWidth:68},4:{cellWidth:22,halign:'center'},5:{cellWidth:22,halign:'center'},6:{cellWidth:14,halign:'center'},7:{cellWidth:22},8:{cellWidth:14,halign:'center'}},
       margin:{left:14,right:14,bottom:26,top:42}});
  }
  PDF.footer(doc, generated);
  doc.save(`REPORTE_LICENCIAS_${new Date().toISOString().slice(0,10)}.pdf`);
}

/* ═══════════════════════════════════════════════════════════
   PDF LCGS
═══════════════════════════════════════════════════════════ */
function genPDFLCGS() {
  if (!window.jspdf?.jsPDF) { alert('No cargó jsPDF.'); return; }
  const rows = applyFiltersLCGS(buildLCGS()).sort((a,b)=>a.nombre.localeCompare(b.nombre,'es'));
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('landscape','mm','a4');
  const generated = new Date().toLocaleString('es-MX');

  const total    = rows.length;
  const personas = new Set(rows.map(r=>r.rfc)).size;
  const diasNum  = rows.map(r=>parseFloat(r.dias)).filter(d=>!isNaN(d));
  const totalDias= diasNum.reduce((a,b)=>a+b,0);
  const porTurno = countMap(rows,r=>r.turno);
  const porAnio  = countMap(rows,r=>r.anio);

  PDF.header(doc,'Reporte · Licencias con Goce de Sueldo','Hospital de la Mujer · Datos filtrados',generated);
  PDF.kpi(doc, 14, 40, 58, 'LCGS',        String(total),     'Registros',     PDF.teal);
  PDF.kpi(doc, 76, 40, 58, 'PERSONAS',    String(personas),  'RFC únicos',    PDF.teal);
  PDF.kpi(doc,138, 40, 58, 'DÍAS TOTAL',  String(totalDias), 'Acumulado',     PDF.slate);
  PDF.kpi(doc,200, 40, 58, 'DÍAS PROM.',  personas&&diasNum.length?(totalDias/personas).toFixed(1):'0','Por persona', PDF.slate);

  PDF.section(doc, 'Distribución por turno y año', 72, PDF.teal);
  PDF.hBar(doc,{x:14,  y:86, w:130, h:44, title:'Por turno', data:porTurno, total, color:PDF.teal,  limit:5, labelW:50});
  PDF.hBar(doc,{x:152, y:86, w:135, h:44, title:'Por año',   data:porAnio,  total, color:PDF.slate, limit:8, labelW:30});

  if (rows.length) {
    doc.addPage('landscape');
    PDF.header(doc,'Detalle · LCGS',`Total: ${rows.length} registros`,generated);
    PDF.table(doc,42,['#','RFC','Nombre','Consecutivo','# Días','Inicio','Término','Turno','Año'],
      rows.map((r,i)=>[i+1,r.rfc,r.nombre,r.consec,r.dias,r.inicio,r.termino,r.turno,r.anio]),
      {generated,pageTitle:'Detalle · LCGS',pageSub:'Continuación',fontSize:5.8,cellPadding:1.6,headColor:PDF.teal,
       columnStyles:{0:{cellWidth:9,halign:'center'},1:{cellWidth:24},2:{cellWidth:60},3:{cellWidth:22,halign:'center'},4:{cellWidth:14,halign:'center'},5:{cellWidth:26,halign:'center'},6:{cellWidth:26,halign:'center'},7:{cellWidth:28},8:{cellWidth:14,halign:'center'}},
       margin:{left:14,right:14,bottom:26,top:42}});
  }
  PDF.footer(doc, generated);
  doc.save(`REPORTE_LCGS_${new Date().toISOString().slice(0,10)}.pdf`);
}

/* ═══════════════════════════════════════════════════════════
   EXPORT CSV
═══════════════════════════════════════════════════════════ */
function csvBlob(headers, rows) {
  const csv = [headers.join(',')].concat(rows.map(r => r.map(v => '"'+String(v??'').replace(/"/g,'""')+'"').join(','))).join('\n');
  return new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8;'});
}
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportCSV() {
  const rows = applyFilters(buildEventos());
  download(csvBlob(['RFC','Nombre','AreaServicio','Turno','Entrada','Salida','Dia','Mes','Anio'],
    rows.map(r=>[r.rfc,r.nombre,r.servicio,r.turno,r.entrada,r.salida,r.dia,r.mes,r.anio])),
    `FACILIDADES_${new Date().toISOString().slice(0,10)}.csv`);
  showToast(`CSV descargado · ${rows.length} filas`);
}
function exportCSVLicencias() {
  const rows = applyFiltersLic(buildLicencias());
  download(csvBlob(['RFC','Nombre','Diagnostico','Inicio','Termino','Dias','Turno','Anio'],
    rows.map(r=>[r.rfc,r.nombre,r.diagnostico,r.inicio,r.termino,r.dias,r.turno,r.anio])),
    `LICENCIAS_${new Date().toISOString().slice(0,10)}.csv`);
  showToast(`CSV licencias · ${rows.length} filas`);
}
function exportCSVLCGS() {
  const rows = applyFiltersLCGS(buildLCGS());
  download(csvBlob(['RFC','Nombre','Consecutivo','Dias','Inicio','Termino','Turno','Anio'],
    rows.map(r=>[r.rfc,r.nombre,r.consec,r.dias,r.inicio,r.termino,r.turno,r.anio])),
    `LCGS_${new Date().toISOString().slice(0,10)}.csv`);
  showToast(`CSV LCGS · ${rows.length} filas`);
}

/* ═══════════════════════════════════════════════════════════
   ANÁLISIS DE FALTAS — HELPERS DE FECHAS
═══════════════════════════════════════════════════════════ */
function parseDateDMY(s) {
  if (!s) return null;
  const str = String(s).trim();
  const parts = str.split('/');
  if (parts.length !== 3) return null;
  const d = parseInt(parts[0], 10), m = parseInt(parts[1], 10), y = parseInt(parts[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y) || y < 1900) return null;
  return new Date(y, m - 1, d);
}

function isMayoDayInFacilidad(facValue, dayNum) {
  if (!facValue || !String(facValue).trim()) return false;
  for (const d of expandFac(facValue)) {
    const parts = d.split('/');
    if (parts.length < 3 || parts[1] !== '05' || parts[2] !== '2026') continue;
    const dp = parts[0];
    if (dp.includes('-')) {
      const [s, e] = dp.split('-').map(x => parseInt(x, 10));
      if (!isNaN(s) && !isNaN(e) && dayNum >= s && dayNum <= e) return true;
    } else {
      if (parseInt(dp, 10) === dayNum) return true;
    }
  }
  return false;
}

function lcgsCoversMayoDay(rec, dayNum) {
  const fi = String(rec['Fecha de Inicio'] || '').trim();
  const ft = String(rec['Fecha de Termino'] || rec['Fecha de Término'] || '').trim();
  const start = parseDateDMY(fi);
  if (!start) return false;
  const end = parseDateDMY(ft) || start;
  const target = new Date(2026, 4, dayNum);
  return start <= target && target <= end;
}

function licMedCoversMayoDay(rec, dayNum) {
  const d  = String(rec['D']  || '').trim();
  const m  = String(rec['M']  || '').trim();
  const aR = String(rec['A']  || '').trim();
  const d2 = String(rec['D_2'] || d).trim();
  const m2 = String(rec['M_2'] || m).trim();
  const a2R= String(rec['A_2'] || aR).trim();
  if (!d || !m || !aR) return false;
  const yr  = aR.length  <= 2 ? 2000 + parseInt(aR,  10) : parseInt(aR,  10);
  const yr2 = a2R.length <= 2 ? 2000 + parseInt(a2R, 10) : parseInt(a2R, 10);
  const start  = new Date(yr,  parseInt(m,  10) - 1, parseInt(d,  10));
  const end    = new Date(yr2, parseInt(m2, 10) - 1, parseInt(d2, 10));
  if (isNaN(start.getTime())) return false;
  const target = new Date(2026, 4, dayNum);
  return start <= target && target <= end;
}

/* ═══════════════════════════════════════════════════════════
   ANÁLISIS DE FALTAS — MAPA TARJETA → PERSONA
═══════════════════════════════════════════════════════════ */
let _tarjetaMap = null;
function buildTarjetaMap() {
  if (_tarjetaMap) return _tarjetaMap;
  _tarjetaMap = new Map();
  for (const [rfc, persona] of Object.entries(DB || {})) {
    for (const sheets of Object.values(persona.fuentes || {})) {
      for (const recs of Object.values(sheets || {})) {
        for (const rec of (recs || [])) {
          const t = guessTarjeta(rec);
          if (t && t !== '—') {
            const key = String(t).trim().replace(/^0+/, '') || '0';
            if (!_tarjetaMap.has(key)) _tarjetaMap.set(key, {rfc, persona});
          }
        }
      }
    }
  }
  return _tarjetaMap;
}

/* ═══════════════════════════════════════════════════════════
   ANÁLISIS DE FALTAS — CORE
═══════════════════════════════════════════════════════════ */
function parseFaltasDias(raw) {
  const s = String(raw ?? '').trim();
  if (!s) return [];
  return [...new Set(s.split(/[,\s]+/).map(x => parseInt(x.trim(), 10)).filter(n => !isNaN(n) && n >= 1 && n <= 31))].sort((a, b) => a - b);
}

function checkCoberturaFalta(persona, diaNum) {
  for (const [src, sheets] of Object.entries(persona.fuentes || {})) {
    for (const recs of Object.values(sheets || {})) {
      for (const rec of (recs || [])) {
        if (isSrcFac(src)) {
          const facVal = rec['FACILIDADES ADMINISTRATIVAS MAYO 2026'];
          if (facVal && isMayoDayInFacilidad(facVal, diaNum))
            return { tipo: 'FACILIDAD', detalle: `Facilidad administrativa día ${diaNum}/05/2026` };
        } else if (isSrcLCGS(src)) {
          if (lcgsCoversMayoDay(rec, diaNum)) {
            const fi = String(rec['Fecha de Inicio'] || '').trim();
            const ft = String(rec['Fecha de Termino'] || rec['Fecha de Término'] || '').trim();
            const diasN = String(rec['No. Días'] || rec['No. D?as'] || '').trim();
            return { tipo: 'LCGS', detalle: `Lic. goce de sueldo ${fi}–${ft}${diasN?' ('+diasN+' días)':''}` };
          }
        } else if (isSrcLicMed(src)) {
          if (licMedCoversMayoDay(rec, diaNum)) {
            const d=rec['D'],m=rec['M'],aR=String(rec['A']||'').trim();
            const d2=rec['D_2']||d, m2=rec['M_2']||m, a2R=String(rec['A_2']||aR).trim();
            const a  = aR.length<=2  ? '20'+aR.padStart(2,'0')  : aR;
            const a2 = a2R.length<=2 ? '20'+a2R.padStart(2,'0') : a2R;
            const diag = String(rec['Diagnostico'] || '').slice(0, 45);
            return { tipo: 'LIC.MED.', detalle: `Licencia médica ${d}/${m}/${a}–${d2}/${m2}/${a2}${diag?' ('+diag+')':''}` };
          }
        }
      }
    }
  }
  return null;
}

function getPersonaServicioTurno(persona) {
  for (const [src, sheets] of Object.entries(persona.fuentes || {})) {
    if (!isSrcFac(src)) continue;
    for (const recs of Object.values(sheets || {})) {
      for (const rec of (recs || [])) {
        const s = String(rec['SERVICIO'] || '').trim();
        const t = normTurno(String(rec['TURNO'] || '').trim());
        if (s) return { servicio: s, turno: t };
      }
    }
  }
  return { servicio: '—', turno: '—' };
}

function normNombreSimple(n) {
  return String(n || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function analyzeFaltas(rows) {
  const tarjetaMap = buildTarjetaMap();
  const results    = [];
  const notFound   = [];
  // rows[0] = day-number headers, rows[1] = column headers, rows[2+] = data
  const dataRows   = rows.slice(2).filter(r => r && r[1] != null && String(r[1]).trim());

  for (const row of dataRows) {
    const tarjetaRaw = String(row[1] ?? '').trim();
    const nombreRaw  = String(row[2] ?? '').trim();
    const faltasRaw  = row[3];
    const diasOrig   = parseFaltasDias(faltasRaw);
    if (!diasOrig.length) continue;

    const tarjetaKey = tarjetaRaw.replace(/^0+/, '') || '0';
    const dbEntry    = tarjetaMap.get(tarjetaKey);

    if (!dbEntry) {
      notFound.push({ tarjeta: tarjetaRaw, nombre: nombreRaw, dias: diasOrig });
      continue;
    }

    const { rfc, persona } = dbEntry;
    const nombreDB   = normNombreSimple(persona.nombre || '');
    const nombreFalt = normNombreSimple(nombreRaw);
    const nameOk     = nombreDB.length < 4 || nombreFalt.length < 4 || nombreDB.includes(nombreFalt.slice(0, 6)) || nombreFalt.includes(nombreDB.slice(0, 6));
    const { servicio, turno } = getPersonaServicioTurno(persona);

    const justificaciones = [], noJustificados = [];
    for (const diaNum of diasOrig) {
      const cob = checkCoberturaFalta(persona, diaNum);
      if (cob) justificaciones.push({ dia: diaNum, ...cob });
      else     noJustificados.push(diaNum);
    }

    results.push({
      tarjeta: tarjetaKey, tarjetaRaw, nombre: nombreRaw, nombreDB: persona.nombre || '—',
      rfc, servicio, turno, nameMatch: nameOk,
      diasOriginales: diasOrig, justificaciones,
      diasJustificados: justificaciones.map(j => j.dia),
      diasNoJustificados: noJustificados
    });
  }

  return { results, notFound };
}

function buildPersonasSinNada(faltasResults) {
  if (_sinNadaCache) return _sinNadaCache;
  const tarjetasConFaltas = new Set(faltasResults.map(r => r.tarjeta));
  const sinNada = [];
  for (const [rfc, persona] of Object.entries(DB || {})) {
    let tieneFacMayo = false, tieneLicMayo = false, tieneLcgsMayo = false, tarjeta = '—', servicio = '—', turno = '—';
    for (const [src, sheets] of Object.entries(persona.fuentes || {})) {
      for (const recs of Object.values(sheets || {})) {
        for (const rec of (recs || [])) {
          if (isSrcFac(src)) {
            const t = guessTarjeta(rec);
            if (t && t !== '—') tarjeta = String(t).trim().replace(/^0+/, '') || '0';
            const sv = String(rec['SERVICIO'] || '').trim(); if (sv) servicio = sv;
            const tu = String(rec['TURNO'] || '').trim();   if (tu) turno = normTurno(tu);
            const mayo = rec['FACILIDADES ADMINISTRATIVAS MAYO 2026'];
            if (mayo && String(mayo).trim() && String(mayo).trim() !== '.' && expandFac(mayo).length)
              tieneFacMayo = true;
          } else if (isSrcLCGS(src)) {
            for (let d = 1; d <= 31 && !tieneLcgsMayo; d++)
              if (lcgsCoversMayoDay(rec, d)) tieneLcgsMayo = true;
          } else if (isSrcLicMed(src)) {
            for (let d = 1; d <= 31 && !tieneLicMayo; d++)
              if (licMedCoversMayoDay(rec, d)) tieneLicMayo = true;
          }
        }
      }
    }
    if (!tieneFacMayo && !tieneLicMayo && !tieneLcgsMayo && !tarjetasConFaltas.has(tarjeta))
      sinNada.push({ rfc, nombre: persona.nombre || '—', tarjeta, servicio, turno });
  }
  _sinNadaCache = sinNada;
  return sinNada;
}

/* ═══════════════════════════════════════════════════════════
   ANÁLISIS DE FALTAS — PANEL UI
═══════════════════════════════════════════════════════════ */
let _faltasAnalysis = null;

function openFaltasPanel() {
  if (!DB) { alert('Espera a que cargue la base de datos.'); return; }
  const rp = document.getElementById('rp');
  document.getElementById('em').style.display = 'none';
  rp.classList.add('on');

  const backBtn = cur ? `<button class="btn sec" onclick="pick('${esc(cur)}')">← Volver a persona</button>` : '';
  rp.innerHTML = `<div class="faltas-wrap">
    <div class="adv-head">
      <div class="adv-title">
        <h2>Análisis de Faltas — Mayo 2026</h2>
        <p>Sube el Excel con la hoja <b>FALTAS MAYO</b>. El sistema cruzará cada falta con facilidades, LCGS y licencias médicas activas ese día.</p>
      </div>
      <div class="adv-actions">
        ${backBtn}
        <button class="btn" onclick="document.getElementById('faltasFile').click()">📂 Subir Excel</button>
        <input type="file" id="faltasFile" accept=".xlsx" style="display:none" onchange="handleFaltasFile(this)">
      </div>
    </div>
    <div id="faltasDropzone" class="faltas-drop" onclick="document.getElementById('faltasFile').click()">
      <div class="faltas-drop-ico">📋</div>
      <div class="faltas-drop-txt">Haz clic o arrastra aquí el archivo Excel</div>
      <div class="faltas-drop-sub">Debe contener la hoja "FALTAS MAYO"</div>
    </div>
    <div id="faltasResult"></div>
  </div>`;

  // Drag & drop
  const dz = document.getElementById('faltasDropzone');
  dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) processFaltasFile(file);
  });
}

async function handleFaltasFile(input) {
  const file = input.files[0];
  if (file) await processFaltasFile(file);
  input.value = '';
}

async function processFaltasFile(file) {
  const res = document.getElementById('faltasResult');
  if (!res) return;
  if (!window.XLSX) { res.innerHTML = `<div class="empty-adv" style="color:#e05252">SheetJS no cargó. Revisa la conexión a internet.</div>`; return; }
  res.innerHTML = `<div class="empty-adv"><div class="spinner" style="margin:0 auto 8px"></div>Procesando ${esc(file.name)}…</div>`;
  _sinNadaCache = null; // reset cache for new file
  try {
    const data = await file.arrayBuffer();
    const wb   = XLSX.read(data, { type: 'array' });
    if (!wb.Sheets['FALTAS MAYO']) {
      res.innerHTML = `<div class="empty-adv" style="color:#e05252">No se encontró la hoja "FALTAS MAYO" en el archivo.</div>`;
      return;
    }
    const rows = XLSX.utils.sheet_to_json(wb.Sheets['FALTAS MAYO'], { header: 1, defval: null });
    const analysis = analyzeFaltas(rows);
    _faltasAnalysis = { ...analysis, rows, fileName: file.name };
    renderFaltasResultado(analysis);
    showToast(`Análisis completado · ${analysis.results.length} personas`);
  } catch(e) {
    res.innerHTML = `<div class="empty-adv" style="color:#e05252">Error al leer el archivo: ${esc(e.message)}</div>`;
    console.error(e);
  }
}

function renderFaltasResultado(analysis) {
  const { results, notFound } = analysis;
  const res = document.getElementById('faltasResult'); if (!res) return;

  const totalPersonas = results.length;
  const totalFaltas   = results.reduce((a, r) => a + r.diasOriginales.length, 0);
  const totalJust     = results.reduce((a, r) => a + r.diasJustificados.length, 0);
  const totalNoJust   = results.reduce((a, r) => a + r.diasNoJustificados.length, 0);
  const pctJust       = totalFaltas ? (totalJust/totalFaltas*100).toFixed(0) : 0;
  const servicios     = [...new Set(results.map(r => r.servicio))].filter(s => s && s !== '—').sort();
  const sinNada       = buildPersonasSinNada(results);

  // Distribución por servicio
  const porServFaltas = countMap(results, r => r.servicio);

  res.innerHTML = `
    <div class="kpi-grid" style="margin-top:4px">
      <div class="kpi kpi-navy"><div class="kpi-num">${totalPersonas}</div><div class="kpi-lbl">Personas con faltas</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${totalFaltas}</div><div class="kpi-lbl">Días falta originales</div></div>
      <div class="kpi kpi-teal"><div class="kpi-num">${totalJust}<span style="font-size:14px;margin-left:4px">(${pctJust}%)</span></div><div class="kpi-lbl">Justificados</div><div class="kpi-sub">Con cobertura activa</div></div>
      <div class="kpi kpi-amber"><div class="kpi-num">${totalNoJust}</div><div class="kpi-lbl">Sin justificar</div><div class="kpi-sub">Permanecen en reporte</div></div>
      <div class="kpi kpi-slate"><div class="kpi-num">${sinNada.length}</div><div class="kpi-lbl">Sin nada en Mayo</div><div class="kpi-sub">Sin faltas ni cobertura</div></div>
    </div>

    <div class="chart-grid-2">
      <div class="adv-card">
        <div class="adv-card-h"><h3>Faltas por servicio</h3><span>${porServFaltas.length} áreas</span></div>
        <div class="adv-card-body">${svgHBars(porServFaltas, totalPersonas, 10)}</div>
      </div>
      <div class="adv-card">
        <div class="adv-card-h"><h3>Resumen por persona</h3><span>${totalPersonas} personas</span></div>
        <div class="adv-card-body">
          <div class="adv-table-wrap" style="max-height:260px">
          <table><thead><tr><th>Tarjeta</th><th>Nombre</th><th>Orig.</th><th>Just.</th><th>Resto</th><th>Servicio</th></tr></thead><tbody>
          ${results.slice().sort((a,b)=>b.diasOriginales.length-a.diasOriginales.length).map(r=>`
            <tr>
              <td class="mono acc">${esc(r.tarjeta)}</td>
              <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(r.nombreDB)}</td>
              <td class="nc">${r.diasOriginales.length}</td>
              <td style="color:var(--acc2);font-family:'IBM Plex Mono',monospace;font-size:10px">${r.diasJustificados.length}</td>
              <td class="nc" style="color:${r.diasNoJustificados.length?'var(--amber)':'var(--acc2)'}">${r.diasNoJustificados.length}</td>
              <td style="font-size:10px;color:var(--tx2)">${esc(r.servicio)}</td>
            </tr>`).join('')}
          </tbody></table>
          </div>
        </div>
      </div>
    </div>

    <div class="adv-card">
      <div class="adv-card-h">
        <h3>Filtrar y exportar</h3>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
          <select id="fServicio" style="font-size:11px;padding:6px 10px;border:1px solid var(--brd);border-radius:7px;background:var(--soft);color:var(--tx)">
            <option value="">Todos los servicios</option>
            ${servicios.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('')}
          </select>
          <button class="btn" onclick="downloadCorregidas()">⬇ Faltas corregidas</button>
          <button class="btn sec" onclick="downloadReporteJustificaciones()">📄 Justificaciones</button>
          <button class="btn sec" onclick="downloadReporteSinNada()">📄 Sin cobertura</button>
          <button class="btn sec" onclick="genPDFFaltas()">🖨 PDF reporte</button>
        </div>
      </div>
    </div>

    <div id="faltasTablas"></div>`;

  document.getElementById('fServicio').addEventListener('change', () => renderFaltasTablas(analysis, sinNada));
  renderFaltasTablas(analysis, sinNada);
}

function renderFaltasTablas(analysis, sinNada) {
  const { results, notFound } = analysis;
  const filtroServ = document.getElementById('fServicio')?.value || '';
  const filtered   = filtroServ ? results.filter(r => r.servicio === filtroServ) : results;
  const filtSinNada = filtroServ ? sinNada.filter(r => r.servicio === filtroServ) : sinNada;

  const el = document.getElementById('faltasTablas'); if (!el) return;

  let h = '';

  // Tabla de justificaciones
  const justRows = [];
  filtered.forEach(r => r.justificaciones.forEach(j => justRows.push({
    tarjeta: r.tarjeta, nombre: r.nombreDB, rfc: r.rfc, servicio: r.servicio, turno: r.turno, dia: j.dia, tipo: j.tipo, detalle: j.detalle
  })));

  h += `<div class="adv-card">
    <div class="adv-card-h"><h3>Faltas justificadas (removidas)</h3><span>${justRows.length} días</span></div>
    <div class="adv-table-wrap">
    ${justRows.length ? `<table><thead><tr><th>Tarjeta</th><th>Nombre</th><th>Servicio</th><th>Turno</th><th>Día</th><th>Tipo</th><th>Cobertura</th></tr></thead><tbody>` +
      justRows.map(r => `<tr>
        <td class="mono acc">${esc(r.tarjeta)}</td><td>${esc(r.nombre)}</td><td>${esc(r.servicio)}</td><td>${esc(r.turno)}</td>
        <td class="nc">${r.dia}/05/2026</td>
        <td><span class="tipo-badge tipo-${r.tipo.replace('.','').toLowerCase().split('.')[0]}">${esc(r.tipo)}</span></td>
        <td style="font-size:10px;color:var(--tx2)">${esc(r.detalle)}</td>
      </tr>`).join('') + '</tbody></table>'
    : '<div class="empty-adv">No hay faltas justificadas.</div>'}
    </div></div>`;

  // Tabla de faltas restantes (no justificadas)
  const noJustRows = filtered.filter(r => r.diasNoJustificados.length);
  h += `<div class="adv-card">
    <div class="adv-card-h"><h3>Faltas sin justificar (permanecen)</h3><span>${noJustRows.reduce((a,r)=>a+r.diasNoJustificados.length,0)} días · ${noJustRows.length} personas</span></div>
    <div class="adv-table-wrap">
    ${noJustRows.length ? `<table><thead><tr><th>Tarjeta</th><th>Nombre</th><th>Servicio</th><th>Turno</th><th>Días sin justificar</th></tr></thead><tbody>` +
      noJustRows.map(r => `<tr>
        <td class="mono acc">${esc(r.tarjeta)}</td><td>${esc(r.nombreDB)}</td><td>${esc(r.servicio)}</td><td>${esc(r.turno)}</td>
        <td class="dc">${r.diasNoJustificados.map(d=>`<span class="day-chip">${d}/05</span>`).join(' ')}</td>
      </tr>`).join('') + '</tbody></table>'
    : '<div class="empty-adv">Todas las faltas fueron justificadas.</div>'}
    </div></div>`;

  // Personas sin nada
  h += `<div class="adv-card">
    <div class="adv-card-h"><h3>Personas sin faltas, sin facilidades ni licencias en Mayo 2026</h3><span>${filtSinNada.length} personas</span></div>
    <div class="adv-table-wrap">
    ${filtSinNada.length ? `<table><thead><tr><th>Tarjeta</th><th>RFC</th><th>Nombre</th><th>Servicio</th><th>Turno</th></tr></thead><tbody>` +
      filtSinNada.map(r => `<tr>
        <td class="mono acc">${esc(r.tarjeta)}</td><td class="mono">${esc(r.rfc)}</td><td>${esc(r.nombre)}</td><td>${esc(r.servicio)}</td><td>${esc(r.turno)}</td>
      </tr>`).join('') + '</tbody></table>'
    : '<div class="empty-adv">No hay personas en esta categoría.</div>'}
    </div></div>`;

  if (notFound.length) {
    h += `<div class="adv-card">
      <div class="adv-card-h" style="background:#fff8f0"><h3 style="color:#a06f20">No encontradas en la base de datos</h3><span>${notFound.length}</span></div>
      <div class="adv-table-wrap">
      <table><thead><tr><th>Tarjeta</th><th>Nombre en faltas</th><th>Días</th></tr></thead><tbody>` +
      notFound.map(r => `<tr><td class="mono">${esc(r.tarjeta)}</td><td>${esc(r.nombre)}</td><td class="mono">${r.dias.join(', ')}</td></tr>`).join('') +
      `</tbody></table></div></div>`;
  }

  el.innerHTML = h;
}

/* ═══════════════════════════════════════════════════════════
   ANÁLISIS DE FALTAS — DESCARGAS XLSX
═══════════════════════════════════════════════════════════ */
function downloadCorregidas() {
  if (!_faltasAnalysis || !window.XLSX) return;
  const { results, rows } = _faltasAnalysis;
  const justMap = new Map(results.map(r => [r.tarjeta, new Set(r.diasJustificados)]));

  // Rebuild rows: keep structure, remove justified dias from FALTAS column
  const outRows = rows.map((row, idx) => {
    if (idx < 2 || !row || row[1] == null) return row ? [...row] : [];
    const tarjeta = String(row[1] ?? '').trim().replace(/^0+/, '') || '0';
    const justified = justMap.get(tarjeta);
    if (!justified || !justified.size) return [...row];
    const diasOrig = parseFaltasDias(row[3]);
    const diasLeft = diasOrig.filter(d => !justified.has(d));
    const newRow   = [...row];
    newRow[3]      = diasLeft.length ? diasLeft.join(', ') : null;
    return newRow;
  });

  const ws = XLSX.utils.aoa_to_sheet(outRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'FALTAS MAYO');
  XLSX.writeFile(wb, `FALTAS_CORREGIDAS_MAYO_2026.xlsx`);
  showToast('Faltas corregidas descargadas');
}

function downloadReporteJustificaciones() {
  if (!_faltasAnalysis || !window.XLSX) return;
  const { results } = _faltasAnalysis;
  const filtroServ = document.getElementById('fServicio')?.value || '';
  const filtered = filtroServ ? results.filter(r => r.servicio === filtroServ) : results;

  const head = ['Tarjeta','RFC','Nombre','Servicio','Turno','Día justificado','Tipo de cobertura','Detalle'];
  const data = [head];
  filtered.forEach(r => {
    r.justificaciones.forEach(j =>
      data.push([r.tarjeta, r.rfc, r.nombreDB, r.servicio, r.turno, `${j.dia}/05/2026`, j.tipo, j.detalle]));
    if (!r.justificaciones.length)
      data.push([r.tarjeta, r.rfc, r.nombreDB, r.servicio, r.turno, '—', 'SIN COBERTURA', `Días sin justificar: ${r.diasNoJustificados.join(', ')}`]);
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), 'Justificaciones');
  XLSX.writeFile(wb, `REPORTE_JUSTIFICACIONES_MAYO_2026.xlsx`);
  showToast(`Reporte justificaciones · ${data.length - 1} filas`);
}

function downloadReporteSinNada() {
  if (!_faltasAnalysis || !window.XLSX) return;
  const { results } = _faltasAnalysis;
  const filtroServ = document.getElementById('fServicio')?.value || '';
  const sinNada    = buildPersonasSinNada(results);
  const filtered   = filtroServ ? sinNada.filter(r => r.servicio === filtroServ) : sinNada;

  const head = ['Tarjeta','RFC','Nombre','Servicio','Turno'];
  const data  = [head, ...filtered.map(r => [r.tarjeta, r.rfc, r.nombre, r.servicio, r.turno])];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), 'Sin cobertura Mayo');
  XLSX.writeFile(wb, `REPORTE_SIN_COBERTURA_MAYO_2026.xlsx`);
  showToast(`Reporte sin cobertura · ${filtered.length} personas`);
}

/* ═══════════════════════════════════════════════════════════
   PDF ANÁLISIS DE FALTAS
═══════════════════════════════════════════════════════════ */
function genPDFFaltas() {
  if (!_faltasAnalysis || !window.jspdf?.jsPDF) { showToast('No hay análisis activo o jsPDF no cargó', 'err'); return; }
  const { results, notFound } = _faltasAnalysis;
  const filtroServ  = document.getElementById('fServicio')?.value || '';
  const filtered    = filtroServ ? results.filter(r => r.servicio === filtroServ) : results;
  const sinNada     = buildPersonasSinNada(results);
  const filtSinNada = filtroServ ? sinNada.filter(r => r.servicio === filtroServ) : sinNada;

  const { jsPDF }  = window.jspdf;
  const doc        = new jsPDF('landscape','mm','a4');
  const generated  = new Date().toLocaleString('es-MX');
  const totalFalt  = filtered.reduce((a,r)=>a+r.diasOriginales.length,0);
  const totalJust  = filtered.reduce((a,r)=>a+r.diasJustificados.length,0);
  const totalNoJust= filtered.reduce((a,r)=>a+r.diasNoJustificados.length,0);
  const pct        = totalFalt ? (totalJust/totalFalt*100).toFixed(0) : 0;

  /* PÁG 1: RESUMEN */
  PDF.header(doc,'Análisis de Faltas · Mayo 2026',`Hospital de la Mujer${filtroServ?' · Servicio: '+filtroServ:''}`,generated);
  PDF.kpi(doc, 14, 40, 56,'PERSONAS',     String(filtered.length),  'Con faltas',     PDF.navy);
  PDF.kpi(doc, 74, 40, 56,'FALTAS ORIG.', String(totalFalt),        'Días',           PDF.slate);
  PDF.kpi(doc,134, 40, 56,'JUSTIFICADAS', `${totalJust} (${pct}%)`, 'Con cobertura',  PDF.teal);
  PDF.kpi(doc,194, 40, 56,'SIN JUSTIF.',  String(totalNoJust),      'Permanecen',     PDF.amber);
  PDF.kpi(doc,254, 40, 32,'SIN NADA',     String(filtSinNada.length),'Sin cobertura', PDF.slate);

  PDF.section(doc,'Resumen por persona',72,PDF.navy);
  PDF.table(doc,86,['Tarjeta','Nombre','Servicio','Turno','Orig.','Just.','Resta'],
    filtered.slice().sort((a,b)=>b.diasOriginales.length-a.diasOriginales.length).map(r=>[
      r.tarjeta, r.nombreDB, r.servicio, r.turno,
      r.diasOriginales.length, r.diasJustificados.length, r.diasNoJustificados.length
    ]),{generated,pageTitle:'Análisis de Faltas · Mayo 2026',pageSub:'Resumen por persona',fontSize:6.8,headColor:PDF.navy,
      columnStyles:{0:{cellWidth:18,halign:'center'},1:{cellWidth:62},2:{cellWidth:58},3:{cellWidth:28},4:{cellWidth:14,halign:'center'},5:{cellWidth:14,halign:'center'},6:{cellWidth:14,halign:'center'}}});

  /* PÁG 2: JUSTIFICACIONES */
  doc.addPage('landscape');
  PDF.header(doc,'Faltas justificadas','Días con cobertura activa en Mayo 2026',generated);
  const justRows=[];
  filtered.forEach(r=>r.justificaciones.forEach(j=>justRows.push([r.tarjeta,r.nombreDB,r.servicio,r.turno,`${j.dia}/05/2026`,j.tipo,j.detalle])));
  if(justRows.length){
    PDF.table(doc,42,['Tarjeta','Nombre','Servicio','Turno','Día','Tipo','Cobertura detectada'],justRows,
      {generated,pageTitle:'Faltas justificadas',pageSub:'Continuación',fontSize:6,cellPadding:1.8,headColor:PDF.teal,
       columnStyles:{0:{cellWidth:18,halign:'center'},1:{cellWidth:52},2:{cellWidth:48},3:{cellWidth:26},4:{cellWidth:22,halign:'center'},5:{cellWidth:22,halign:'center'},6:{cellWidth:99}}});
  } else PDF.noData(doc,42,'No se justificó ninguna falta');

  /* PÁG 3: SIN JUSTIFICAR */
  if(filtered.some(r=>r.diasNoJustificados.length)){
    doc.addPage('landscape');
    PDF.header(doc,'Faltas sin justificar','Días que permanecen en el reporte',generated);
    const noJRows=[];
    filtered.filter(r=>r.diasNoJustificados.length).forEach(r=>
      noJRows.push([r.tarjeta,r.nombreDB,r.servicio,r.turno,r.diasNoJustificados.join(', ')]));
    PDF.table(doc,42,['Tarjeta','Nombre','Servicio','Turno','Días sin justificar'],noJRows,
      {generated,pageTitle:'Sin justificar',pageSub:'Continuación',fontSize:6.8,headColor:PDF.amber,
       columnStyles:{0:{cellWidth:18,halign:'center'},1:{cellWidth:62},2:{cellWidth:60},3:{cellWidth:28},4:{cellWidth:119}}});
  }

  /* PÁG 4: SIN NADA */
  if(filtSinNada.length){
    doc.addPage('landscape');
    PDF.header(doc,'Personas sin cobertura en Mayo 2026','Sin faltas, facilidades, LCGS ni licencias activas',generated);
    PDF.table(doc,42,['Tarjeta','RFC','Nombre','Servicio','Turno'],
      filtSinNada.map(r=>[r.tarjeta,r.rfc,r.nombre,r.servicio,r.turno]),
      {generated,pageTitle:'Sin cobertura Mayo',pageSub:'Continuación',fontSize:6.8,headColor:PDF.slate,
       columnStyles:{0:{cellWidth:18,halign:'center'},1:{cellWidth:36},2:{cellWidth:72},3:{cellWidth:72},4:{cellWidth:28}}});
  }

  PDF.footer(doc,generated);
  doc.save(`ANALISIS_FALTAS_MAYO_2026${filtroServ?'_'+PDF.safeName(filtroServ):''}.pdf`);
  showToast('PDF análisis de faltas descargado');
}

/* ── INICIO ── */
loadData();
