/* Dashboard Construccion — La Pampina | PowerChina | SheetJS */
'use strict';

// ── State ────────────────────────────────────────────────────────────────────
let D = null;
let charts = {};

// ── Boot ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setupTheme();
  setupTabs();
  setupFileUpload();
  setupScenarios();
  setupTableFilters();
  setupTabFilters();
});

// ── Theme toggle ─────────────────────────────────────────────────────────────
function setupTheme() {
  const btn   = document.getElementById('themeToggle');
  const saved = localStorage.getItem('theme') || 'light';
  applyTheme(saved);
  btn?.addEventListener('click', () => {
    const next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    localStorage.setItem('theme', next);
  });
}
function applyTheme(t) {
  document.documentElement.setAttribute('data-theme', t);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = t === 'dark' ? '☀️' : '🌙';
}

// ── File loading (SheetJS) ────────────────────────────────────────────────────
function setupFileUpload() {
  on('fileInput', 'change', handleFile);
}

async function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  showToast('Leyendo archivo…', false, 60000);
  try {
    const buf = await file.arrayBuffer();
    // Yield ao browser para mostrar o toast antes do parsing pesado
    await new Promise(r => setTimeout(r, 30));
    showToast('Procesando datos…', false, 60000);
    await new Promise(r => setTimeout(r, 30));
    const wb = XLSX.read(buf, {
      type: 'array',
      cellDates: false,         // datas como serial numérico — mais rápido
      bookVBA: false,           // ignora código VBA (maior ganho em .xlsm)
      sheets: 'LPA CONSTR.',    // só lê a aba necessária
    });
    D = parseConstr(wb);
    document.getElementById('welcomeCard').style.display    = 'none';
    document.getElementById('resumenContent').style.display = 'block';
    populateAreaDropdowns();
    render();
    showToast('Cargado: ' + D.meta.dataWeek + ' · ' + D.meta.actTotal + ' actividades');
  } catch(err) {
    showToast('Error: ' + err.message, true);
    console.error(err);
  }
  e.target.value = '';
}

// ── Excel Parsing ─────────────────────────────────────────────────────────────
function parseConstr(wb) {
  const ws = wb.Sheets['LPA CONSTR.'];
  if (!ws) throw new Error('Aba "LPA CONSTR." não encontrada no arquivo.');

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1, raw: true, cellDates: true, defval: null
  });

  const PLAN_START = 22, PLAN_END = 90;
  const REAL_START = 92, REAL_END = 160;

  const dateRow   = rows[8] || [];
  const headerRow = rows[9] || [];
  const planWeeks = headerRow.slice(PLAN_START, PLAN_END);
  const planDates = dateRow.slice(PLAN_START, PLAN_END).map(xlsxDateToIso);

  const C = {
    nivel:1, resumen:2, edt:6, tarea:7, duracion:8,
    inicio:9, fin:10, hh:11, incidencia:13,
    pct_plan:15, pct_real:16, pct_comp_plan:17, pct_comp_real:18,
    desv_pond:19, desviacion:20,
  };

  const records = [];
  let totalRow = null;

  for (const row of rows.slice(10)) {
    if (!row || (row[0] == null && row[5] == null)) continue;
    const edt   = String(row[C.edt]   ?? '').trim();
    const tarea = String(row[C.tarea] ?? '').trim();
    if (!tarea) continue;

    const rec = {
      edt, tarea,
      nivel:        +(row[C.nivel]         ?? 0),
      resumen:      isResumenVal(row[C.resumen]),
      duracion:     String(row[C.duracion] ?? ''),
      inicio:       xlsxDateToIso(row[C.inicio]),
      fin:          xlsxDateToIso(row[C.fin]),
      hh:           +(row[C.hh]            ?? 0),
      incidencia:   +(row[C.incidencia]    ?? 0),
      pctPlan:      +(row[C.pct_plan]      ?? 0),
      pctReal:      +(row[C.pct_real]      ?? 0),
      pctCompPlan:  +(row[C.pct_comp_plan] ?? 0),
      pctCompReal:  +(row[C.pct_comp_real] ?? 0),
      desvPond:     +(row[C.desv_pond]     ?? 0),
      desviacion:   +(row[C.desviacion]    ?? 0),
      planSeries:   row.slice(PLAN_START, PLAN_END).map(v => +v || 0),
      realSeries:   row.slice(REAL_START, REAL_END).map(v => +v || 0),
    };
    rec.status = calcStatus(rec);
    records.push(rec);
    if (edt === '4.5') totalRow = rec;
  }

  // Auto-detect semana actual: último índice com real > 0 na linha 4.5
  let currIdx = 0;
  if (totalRow) {
    for (let i = totalRow.realSeries.length - 1; i >= 0; i--) {
      if (totalRow.realSeries[i] > 0) { currIdx = i; break; }
    }
  }
  const dataWeek = planWeeks[currIdx] || `W${currIdx}`;
  const dataDate = planDates[currIdx] || null;

  const scurve = planWeeks.map((wk, i) => ({
    week: wk || `W${i}`,
    date: planDates[i],
    plan: totalRow ? (totalRow.planSeries[i] || 0) : 0,
    real: (totalRow && totalRow.realSeries[i] > 0) ? totalRow.realSeries[i] : null,
    isCurrent: i === currIdx,
  }));

  const areas     = records.filter(r => r.resumen && (r.nivel === 3 || r.nivel === 4));
  const leaves    = records.filter(r => !r.resumen);

  const topDesvios = leaves.filter(r => r.incidencia > 0.0001)
    .sort((a, b) => a.desviacion - b.desviacion).slice(0, 50);
  const critical  = leaves.filter(r => r.incidencia > 0.003 && r.desviacion < -0.05)
    .sort((a, b) => b.incidencia - a.incidencia);
  const sinAvance = leaves.filter(r => r.pctCompPlan > 0 && r.pctCompReal === 0)
    .sort((a, b) => b.incidencia - a.incidencia);
  const ranking   = leaves.filter(r => r.incidencia > 0)
    .sort((a, b) => Math.abs(b.desvPond) - Math.abs(a.desvPond)).slice(0, 50);
  const future    = leaves
    .filter(r => r.pctCompReal < 1 && r.fin && r.fin > dataDate && r.incidencia > 0.0005)
    .sort((a, b) => (b.incidencia*(1-b.pctCompReal)) - (a.incidencia*(1-a.pctCompReal)))
    .slice(0, 30);

  const leavesW = leaves.filter(r => r.incidencia > 0);
  const avgIncidencia = leavesW.length
    ? leavesW.reduce((s, r) => s + r.incidencia, 0) / leavesW.length : 0;

  const tr = totalRow || {};
  return {
    meta: {
      dataDate, dataWeek,
      startLB:      tr.inicio      ?? null,
      endLB:        tr.fin         ?? null,
      totalHH:      tr.hh          || 0,
      pctPlan:      tr.pctPlan     || 0,
      pctReal:      tr.pctReal     || 0,
      desvio:       tr.desvPond    || 0,
      actTotal:     leaves.length,
      actSinAvance: sinAvance.length,
      avgIncidencia,
    },
    scurve, areas, topDesvios, critical, sinAvance, ranking, future, allLeaves: leaves,
  };
}

function xlsxDateToIso(v) {
  if (v == null) return null;
  // Serial numérico do Excel (ex: 45794 = 17/05/2026)
  if (typeof v === 'number' && v > 1) {
    const d = new Date((v - 25569) * 86400000);
    return d.toISOString().split('T')[0];
  }
  if (v instanceof Date && !isNaN(v)) return v.toISOString().split('T')[0];
  return null;
}
function isResumenVal(v) {
  if (v == null) return false;
  const s = String(v).toLowerCase().trim();
  return s === 'si' || s === 'sí' || s === 'yes' || s === 'true' || v === true || v === 1;
}
function calcStatus(r) {
  if (r.pctCompReal >= 0.995) return 'completed';
  if (r.pctCompReal > 0)      return 'inProgress';
  if (r.pctCompPlan > 0)      return 'late';
  return 'notStarted';
}

// ── Render ────────────────────────────────────────────────────────────────────
function render() {
  renderKPIs();
  renderResumen();
  renderAreaChart();
  renderStatusDonut();
  renderWeeklyBar();
  renderTop5();
  renderScurve();
  renderAreasBar();
  renderAreasTable();
  renderDesviosBar(D.topDesvios);
  renderDesviosTable(D.topDesvios);
  renderCriticasBar(D.critical);
  renderCriticasTable(D.critical);
  renderSinAvanceCharts(D.sinAvance);
  renderSinAvanceTable(D.sinAvance);
  renderRankingBar(D.ranking.slice(0, 20));
  renderRankingTable(D.ranking);
  renderFuture();
  renderTabla();
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs() {
  const m = D.meta;
  set('kPlan',    pct(m.pctPlan));
  set('kReal',    pct(m.pctReal));
  set('kDev',     (m.desvio >= 0 ? '+' : '') + pct(m.desvio));
  set('kHH',      (m.totalHH / 1000).toFixed(0) + 'K h');
  set('kAct',     m.actTotal.toLocaleString());
  set('kSin',     m.actSinAvance);
  set('metaDate', 'Semana ' + m.dataWeek + '  |  ' + fmtDate(m.dataDate));

  const devEl = document.getElementById('kpi-dev');
  devEl.classList.toggle('kpi-alert', m.desvio < 0);
  devEl.classList.toggle('positive',  m.desvio >= 0);

  const realPts = D.scurve.filter(s => s.real != null && s.real > 0);
  let forecast = '—';
  if (realPts.length >= 2) {
    const last = realPts[realPts.length - 1];
    const prev = realPts[realPts.length - 2];
    const rate = last.real - prev.real;
    if (rate > 0) {
      const weeks = Math.round((1 - last.real) / rate);
      const d = new Date(last.date);
      d.setDate(d.getDate() + weeks * 7);
      forecast = d.toLocaleDateString('es-CL', { month: 'short', year: 'numeric' });
    }
  }
  set('kForecast', forecast);
}

// ── Resumen ejecutivo ─────────────────────────────────────────────────────────
function renderResumen() {
  const m = D.meta;
  const dev = m.desvio;
  const devStr = (dev >= 0 ? '+' : '') + pct(dev);
  const status = dev < -0.02 ? '<span class="alert">PROYECTO EN RETRASO CRÍTICO</span>'
               : dev < 0    ? '<span class="alert">Proyecto con retraso</span>'
               :               '<span class="ok">Proyecto sin atraso</span>';

  const areas = D.areas.filter(a => a.nivel === 3 && a.edt !== '4.5.1' && a.edt !== '4.5.9');
  const areaRows = areas.map(a => {
    const d = a.desvPond;
    const cls = d < -0.005 ? 'alert' : d < 0 ? '' : 'ok';
    return `<span class="${cls}">• ${a.tarea.trim()}: Plan ${pct(a.pctPlan)} / Real ${pct(a.pctReal)} / Desvío ${(d >= 0 ? '+' : '') + pct(d)}</span>`;
  }).join('<br>');

  document.getElementById('resumenText').innerHTML = `
    <strong>Proyecto:</strong> LA PAMPINA &nbsp;|&nbsp; <strong>Disciplina:</strong> CONSTRUCCIÓN<br>
    <strong>Fecha de control:</strong> ${fmtDate(m.dataDate)} (${m.dataWeek})<br>
    <strong>Período LB:</strong> ${fmtDate(m.startLB)} → ${fmtDate(m.endLB)}<br>
    <strong>H-H Totales:</strong> ${Math.round(m.totalHH).toLocaleString()} horas-hombre<br>
    <br>
    <strong>Situación:</strong> ${status}<br>
    Avance real: <strong>${pct(m.pctReal)}</strong> vs planeado: <strong>${pct(m.pctPlan)}</strong><br>
    Desvío acumulado ponderado: <strong>${devStr}</strong><br>
    <br>
    <strong>Avance por área:</strong><br>
    ${areaRows}<br>
    <br>
    <strong>Actividades sin avance real:</strong> ${m.actSinAvance}<br>
    <strong>Top desvío:</strong> ${D.topDesvios[0] ? D.topDesvios[0].tarea.trim() + ' (' + pct(D.topDesvios[0].desviacion) + ')' : '—'}
  `;
}

// ── Status Donut ──────────────────────────────────────────────────────────────
function renderStatusDonut() {
  const counts = { completed:0, inProgress:0, late:0, notStarted:0 };
  D.allLeaves.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
  const total = D.allLeaves.length;

  destroyChart('statusDonut');
  charts['statusDonut'] = new Chart(document.getElementById('statusDonut'), {
    type: 'doughnut',
    data: {
      labels: ['Completada', 'En progreso', 'Atrasada', 'No iniciada'],
      datasets: [{
        data: [counts.completed, counts.inProgress, counts.late, counts.notStarted],
        backgroundColor: ['#00844a','#f0a500','#c00000','#a0a8c0'],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font:{ size:11 }, padding:8 } },
        tooltip: { callbacks: { label: ctx =>
          `${ctx.label}: ${ctx.parsed} (${((ctx.parsed/total)*100).toFixed(1)}%)`
        }}
      }
    }
  });
}

// ── Weekly Bar ────────────────────────────────────────────────────────────────
function renderWeeklyBar() {
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  const slice = sc.slice(Math.max(0, currIdx - 9), currIdx + 1);

  destroyChart('weeklyBar');
  charts['weeklyBar'] = new Chart(document.getElementById('weeklyBar'), {
    type: 'bar',
    data: {
      labels: slice.map(s => s.week),
      datasets: [
        { label:'% Plan', data: slice.map(s => +(s.plan*100).toFixed(2)), backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label:'% Real', data: slice.map(s => s.real != null ? +(s.real*100).toFixed(2) : null), backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top', labels:{ font:{size:11} } } },
      scales: { y: { ticks:{ callback: v => v+'%', font:{size:10} } } }
    }
  });
}

// ── Top 5 ─────────────────────────────────────────────────────────────────────
function renderTop5() {
  document.getElementById('top5List').innerHTML = D.topDesvios.slice(0, 5).map((r, i) => `
    <div class="top5-item">
      <span class="top5-rank">${i+1}</span>
      <div class="top5-info">
        <div class="top5-name">${r.tarea.trim().length > 44 ? r.tarea.trim().slice(0,44)+'…' : r.tarea.trim()}</div>
        <div class="top5-meta">${r.edt} · incid. ${pct(r.incidencia,3)}</div>
      </div>
      <span class="top5-dev dev-neg">${signPct(r.desviacion)}</span>
    </div>
  `).join('');
}

// ── Area bar chart ────────────────────────────────────────────────────────────
function renderAreaChart() {
  const areas = D.areas.filter(a => a.nivel===3 && a.incidencia > 0.001);
  const plan  = areas.map(a => +(a.pctPlan*100).toFixed(2));
  const real  = areas.map(a => +(a.pctReal*100).toFixed(2));

  destroyChart('areaChart');
  charts['areaChart'] = new Chart(document.getElementById('areaChart'), {
    type: 'bar',
    data: {
      labels: areas.map(a => a.tarea.trim()),
      datasets: [
        { label:'% Plan', data:plan, backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label:'% Real', data:real, backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top' } },
      scales: { y: { ticks:{ callback: v => v+'%' }, max: Math.max(...plan,...real,5)*1.2 } }
    }
  });
}

// ── S-Curve ───────────────────────────────────────────────────────────────────
function renderScurve() {
  const period = document.getElementById('scurvePeriod')?.value || 'around';
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);

  let slice;
  if      (period === 'around') slice = sc.slice(Math.max(0,currIdx-12), Math.min(sc.length,currIdx+13));
  else if (period === 'past20') slice = sc.slice(Math.max(0,currIdx-19), currIdx+1);
  else if (period === 'next30') slice = sc.slice(currIdx, Math.min(sc.length,currIdx+31));
  else                          slice = sc;

  destroyChart('scurveChart');
  charts['scurveChart'] = new Chart(document.getElementById('scurveChart'), {
    type: 'line',
    data: {
      labels: slice.map(s => s.week),
      datasets: [
        { label:'% Planeado', data: slice.map(s => +(s.plan*100).toFixed(3)),
          borderColor:'#0054a6', backgroundColor:'rgba(0,84,166,.08)',
          pointRadius:0, fill:true, tension:0.3, borderWidth:2 },
        { label:'% Real', data: slice.map(s => s.real!=null ? +(s.real*100).toFixed(3) : null),
          borderColor:'#00a36c', backgroundColor:'rgba(0,163,108,.08)',
          pointRadius:1, fill:false, tension:0.3, borderWidth:2.5, spanGaps:false },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      interaction:{ mode:'index', intersect:false },
      plugins: {
        legend:{ position:'top' },
        tooltip:{ callbacks:{ label: ctx => ctx.dataset.label+': '+(ctx.parsed.y||0).toFixed(2)+'%' } },
      },
      scales: { y:{ ticks:{ callback: v => v+'%' }, min:0 } }
    }
  });
}

// ── Areas Bar + Table ─────────────────────────────────────────────────────────
function renderAreasBar() {
  const n = parseInt(document.getElementById('areaLvlBox')?.value || '3');
  const rows = D.areas.filter(a => a.nivel === n);

  destroyChart('areasBarChart');
  charts['areasBarChart'] = new Chart(document.getElementById('areasBarChart'), {
    type: 'bar',
    data: {
      labels: rows.map(a => a.tarea.trim().slice(0,35)),
      datasets: [
        { label:'% Plan Pond.', data: rows.map(a => +(a.pctPlan*100).toFixed(2)), backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label:'% Real Pond.', data: rows.map(a => +(a.pctReal*100).toFixed(2)), backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top' } },
      scales: { x:{ ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderAreasTable() {
  const n = parseInt(document.getElementById('areaLvlBox')?.value || '3');
  const rows = D.areas.filter(a => a.nivel === n);

  document.getElementById('areasTable').innerHTML = tableWrap(
    `<tr>
      <th class="left">Área / Grupo</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
      <th>H-H</th><th>Incidencia</th><th>% Plan Pond.</th><th>% Real Pond.</th>
      <th>Desvío Pond.</th><th>% Comp. Plan</th><th>% Comp. Real</th><th>Status</th>
    </tr>`,
    rows.map(r => `<tr>
      <td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td>
      <td>${pct(r.incidencia,3)}</td><td>${pct(r.pctPlan)}</td><td>${pct(r.pctReal)}</td>
      <td class="${devClass(r.desvPond)}">${signPct(r.desvPond)}</td>
      <td>${pbarDuo(r.pctCompPlan,r.pctCompReal)}</td>
      <td>${pct(r.pctCompReal)}</td><td>${statusBadge(r)}</td>
    </tr>`).join('')
  );
}

// ── Top Desvios Bar + Table ───────────────────────────────────────────────────
function renderDesviosBar(rows) {
  const top  = rows.slice(0,15);
  const vals = top.map(r => +(r.desviacion*100).toFixed(2));

  destroyChart('desviosBarChart');
  charts['desviosBarChart'] = new Chart(document.getElementById('desviosBarChart'), {
    type:'bar',
    data: {
      labels: top.map(r => r.edt),
      datasets:[{ label:'Desvío (%)', data:vals,
        backgroundColor: vals.map(v => v<0 ? 'rgba(192,0,0,0.7)' : 'rgba(0,132,74,0.7)'),
        borderRadius:3 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label: ctx => {
          const r = top[ctx.dataIndex];
          return [`Desvío: ${ctx.parsed.x.toFixed(2)}%`, r.tarea.trim().slice(0,50)];
        }}}
      },
      scales:{ x:{ ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderDesviosTable(rows) {
  document.getElementById('desviosTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
     <th>H-H</th><th>Incidencia</th><th>% Plan</th><th>% Real</th><th>Desvío</th></tr>`,
    rows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
      <td class="${devClass(r.desviacion)}">${signPct(r.desviacion)}</td>
    </tr>`).join('')
  );
}

// ── Críticas Bar + Table ──────────────────────────────────────────────────────
function renderCriticasBar(rows) {
  const top = rows.slice(0,15);
  destroyChart('criticasBarChart');
  charts['criticasBarChart'] = new Chart(document.getElementById('criticasBarChart'), {
    type:'bar',
    data: {
      labels: top.map(r => r.edt),
      datasets:[
        { label:'% Plan', data: top.map(r => +(r.pctCompPlan*100).toFixed(1)), backgroundColor:'rgba(0,84,166,0.6)', borderRadius:3 },
        { label:'% Real', data: top.map(r => +(r.pctCompReal*100).toFixed(1)), backgroundColor:'rgba(192,0,0,0.65)', borderRadius:3 },
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'top'},
        tooltip:{ callbacks:{ afterTitle: items => top[items[0].dataIndex]?.tarea.trim().slice(0,55) }}
      },
      scales:{ x:{ ticks:{ callback: v => v+'%' }, max:100 } }
    }
  });
}

function renderCriticasTable(rows) {
  document.getElementById('criticasTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
     <th>H-H</th><th>Incidencia</th><th>% Plan</th><th>% Real</th><th>Desvío</th><th>Impacto Pond.</th></tr>`,
    rows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
      <td class="dev-neg">${signPct(r.desviacion)}</td><td class="dev-neg">${signPct(r.desvPond)}</td>
    </tr>`).join('')
  );
}

// ── Sin Avance Charts + Table ─────────────────────────────────────────────────
function renderSinAvanceCharts(rows) {
  const areaNames = {};
  D.areas.filter(a => a.nivel===3).forEach(a => { areaNames[a.edt] = a.tarea.trim().slice(0,22); });

  const byCount = {}, byHH = {};
  rows.forEach(r => {
    const k = r.edt.split('.').slice(0,3).join('.');
    byCount[k] = (byCount[k]||0) + 1;
    byHH[k]    = (byHH[k]   ||0) + r.hh;
  });

  const keys   = Object.keys(byCount).sort();
  const labels = keys.map(k => areaNames[k] || k);
  const colors = ['#c00000','#e68a00','#0054a6','#6a0dad','#00844a','#008080','#888'];

  destroyChart('sinAvanceDonut');
  charts['sinAvanceDonut'] = new Chart(document.getElementById('sinAvanceDonut'), {
    type:'doughnut',
    data:{ labels, datasets:[{ data: keys.map(k=>byCount[k]),
      backgroundColor: colors.slice(0,keys.length), borderWidth:2, borderColor:'#fff' }] },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'right', labels:{ font:{size:11}, padding:8 } } } }
  });

  const hhEntries = Object.entries(byHH).sort((a,b) => b[1]-a[1]);
  destroyChart('sinAvanceHHBar');
  charts['sinAvanceHHBar'] = new Chart(document.getElementById('sinAvanceHHBar'), {
    type:'bar',
    data:{ labels: hhEntries.map(([k]) => areaNames[k]||k),
      datasets:[{ label:'H-H sin avance', data: hhEntries.map(([,v]) => Math.round(v)),
        backgroundColor:'rgba(192,0,0,0.65)', borderRadius:3 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{ x:{ ticks:{ font:{size:11} } } } }
  });
}

function renderSinAvanceTable(rows) {
  document.getElementById('sinAvanceTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
     <th>H-H</th><th>Incidencia</th><th>% Plan</th><th>% Real</th></tr>`,
    rows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td class="dev-neg">0.0%</td>
    </tr>`).join('')
  );
}

// ── Ranking Bar + Table ───────────────────────────────────────────────────────
function renderRankingBar(rows) {
  const vals = rows.map(r => +(Math.abs(r.desvPond)*100).toFixed(3));
  destroyChart('rankingBarChart');
  charts['rankingBarChart'] = new Chart(document.getElementById('rankingBarChart'), {
    type:'bar',
    data:{ labels: rows.map(r => r.edt),
      datasets:[{ label:'Impacto Ponderado (%)', data:vals,
        backgroundColor: vals.map(v => v>0.5?'rgba(192,0,0,0.7)':v>0.2?'rgba(230,138,0,0.7)':'rgba(0,84,166,0.6)'),
        borderRadius:3 }]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label: ctx => {
          const r = rows[ctx.dataIndex];
          return [`Impacto: ${ctx.parsed.x.toFixed(3)}%`, r.tarea.trim().slice(0,50)];
        }}}
      },
      scales:{ x:{ ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderRankingTable(rows) {
  document.getElementById('rankingTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>H-H</th>
     <th>Incidencia</th><th>% Plan Pond.</th><th>% Real Pond.</th><th>Impacto Pond.</th>
     <th>% Plan</th><th>% Real</th><th>Clasif.</th></tr>`,
    rows.map((r,i) => {
      const imp = Math.abs(r.desvPond);
      const cls = imp>0.005?'badge badge-crit':imp>0.002?'badge badge-late':imp>0.0005?'badge badge-warn':'badge badge-ok';
      const lbl = imp>0.005?'CRITICO':imp>0.002?'ALTO':imp>0.0005?'MEDIO':'BAJO';
      return `<tr>
        <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
        <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
        <td>${pct(r.pctPlan)}</td><td>${pct(r.pctReal)}</td>
        <td class="${devClass(r.desvPond)}">${signPct(r.desvPond)}</td>
        <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
        <td><span class="${cls}">${lbl}</span></td>
      </tr>`;
    }).join('')
  );
}

// ── Future ────────────────────────────────────────────────────────────────────
function renderFuture() {
  document.getElementById('futureTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th>
     <th>H-H</th><th>Incidencia</th><th>% Real actual</th><th>% Recuperable</th><th>Fin LB</th></tr>`,
    D.future.map((r,i) => {
      const rec = r.incidencia*(1-r.pctCompReal);
      return `<tr>
        <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
        <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
        <td>${pct(r.pctCompReal)}</td><td class="dev-pos">${pct(rec,3)}</td>
        <td>${fmtDate(r.fin)}</td>
      </tr>`;
    }).join('')
  );
}

// ── Tabla ─────────────────────────────────────────────────────────────────────
let tablaRows = [];
function renderTabla(filtered) {
  tablaRows = filtered || D.allLeaves;

  const areaBox = document.getElementById('areaBox');
  if (areaBox.options.length === 1) {
    [...new Set(D.allLeaves.map(r => r.edt.split('.').slice(0,2).join('.')))].sort()
      .forEach(g => areaBox.appendChild(new Option(g, g)));
  }

  document.getElementById('tablaWrap').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Dur.</th>
     <th>Inicio LB</th><th>Fin LB</th><th>H-H</th><th>Incidencia</th>
     <th>% Plan</th><th>% Real</th><th>Desvío</th><th>Status</th></tr>`,
    tablaRows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td>
      <td>${r.edt}</td><td>${r.duracion}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
      <td class="${devClass(r.desviacion)}">${signPct(r.desviacion)}</td>
      <td>${statusBadge(r)}</td>
    </tr>`).join('')
  );
}

// ── Populate area dropdowns ───────────────────────────────────────────────────
function populateAreaDropdowns() {
  const areaNames = {};
  D.areas.filter(a => a.nivel===3).forEach(a => { areaNames[a.edt] = a.tarea.trim().slice(0,28); });
  const opts = Object.entries(areaNames).sort()
    .map(([k,v]) => `<option value="${k}">${k} — ${v}</option>`).join('');
  ['desvAreaBox','critAreaBox','sinAreaBox','rankAreaBox'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = '<option value="">— Todas las áreas —</option>' + opts;
  });
}

// ── Tab filters ───────────────────────────────────────────────────────────────
function setupTabFilters() {
  on('scurvePeriod', 'change', renderScurve);
  on('areaLvlBox',  'change', () => { renderAreasBar(); renderAreasTable(); });

  function updateDesv() {
    if (!D) return;
    const q = document.getElementById('desvSearch').value.toLowerCase();
    const area = document.getElementById('desvAreaBox').value;
    const rows = D.topDesvios.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderDesviosBar(rows); renderDesviosTable(rows);
  }
  on('desvSearch',  'input',  updateDesv);
  on('desvAreaBox', 'change', updateDesv);

  function updateCrit() {
    if (!D) return;
    const q = document.getElementById('critSearch').value.toLowerCase();
    const area = document.getElementById('critAreaBox').value;
    const rows = D.critical.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderCriticasBar(rows); renderCriticasTable(rows);
  }
  on('critSearch',  'input',  updateCrit);
  on('critAreaBox', 'change', updateCrit);

  function updateSin() {
    if (!D) return;
    const q = document.getElementById('sinSearch').value.toLowerCase();
    const area = document.getElementById('sinAreaBox').value;
    const rows = D.sinAvance.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderSinAvanceCharts(rows); renderSinAvanceTable(rows);
  }
  on('sinSearch',  'input',  updateSin);
  on('sinAreaBox', 'change', updateSin);

  function updateRank() {
    if (!D) return;
    const q      = document.getElementById('rankSearch').value.toLowerCase();
    const area   = document.getElementById('rankAreaBox').value;
    const clasif = document.getElementById('rankClasif').value;
    const rows = D.ranking.filter(r => {
      const imp = Math.abs(r.desvPond);
      const c = imp>0.005?'CRITICO':imp>0.002?'ALTO':imp>0.0005?'MEDIO':'BAJO';
      return (!q      || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
             (!area   || r.edt.startsWith(area)) &&
             (!clasif || c === clasif);
    });
    renderRankingBar(rows.slice(0,20)); renderRankingTable(rows);
  }
  on('rankSearch',  'input',  updateRank);
  on('rankAreaBox', 'change', updateRank);
  on('rankClasif',  'change', updateRank);
}

// ── Table filters (Tabla tab) ─────────────────────────────────────────────────
function setupTableFilters() {
  on('searchBox', 'input',  applyFilters);
  on('statusBox', 'change', applyFilters);
  on('areaBox',   'change', applyFilters);
  on('exportBtn', 'click',  exportCsv);
}

function applyFilters() {
  if (!D) return;
  const q      = document.getElementById('searchBox').value.toLowerCase();
  const status = document.getElementById('statusBox').value;
  const area   = document.getElementById('areaBox').value;
  renderTabla(D.allLeaves.filter(r =>
    (!q      || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
    (!status || r.status === status) &&
    (!area   || r.edt.startsWith(area))
  ));
}

function exportCsv() {
  const hdr = ['EDT','Actividad','Duracion','Inicio LB','Fin LB','HH','Incidencia','Plan%','Real%','Desvio','Status'];
  const lines = [hdr.join(','), ...tablaRows.map(r =>
    [r.edt, `"${r.tarea.trim()}"`, r.duracion, r.inicio, r.fin,
     r.hh.toFixed(0), r.incidencia.toFixed(6),
     (r.pctCompPlan*100).toFixed(2), (r.pctCompReal*100).toFixed(2),
     (r.desviacion*100).toFixed(2), r.status].join(',')
  )];
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'}));
  a.download = 'actividades_construccion.csv';
  a.click();
}

// ── Scenarios ─────────────────────────────────────────────────────────────────
function setupScenarios() {
  on('scCalcA', 'click', () => {
    const edt     = document.getElementById('scEdtA').value.trim();
    const newReal = parseFloat(document.getElementById('scRealA').value) / 100;
    const act     = D?.allLeaves?.find(r => r.edt === edt);
    const el      = document.getElementById('scResultA');
    if (!act) { el.innerHTML = 'Actividad no encontrada. Verifique el EDT.'; el.classList.remove('hidden'); return; }
    const gain     = act.incidencia*(newReal - act.pctCompReal);
    const newTotal = D.meta.pctReal + gain;
    el.innerHTML = `
      <strong>Actividad:</strong> ${act.tarea.trim()}<br>
      <strong>Incidencia:</strong> ${pct(act.incidencia,4)}<br>
      <strong>% Real actual:</strong> ${pct(act.pctCompReal)}<br>
      <strong>% Real propuesto:</strong> ${pct(newReal)}<br>
      <strong>Ganancia en % ponderado:</strong> ${signPct(gain)}<br>
      <strong>Nuevo % Real total estimado:</strong> ${pct(newTotal)}
      ${gain >= 0 ? ' <span class="ok">▲ Mejora</span>' : ' <span class="alert">▼ Sin mejora</span>'}
    `;
    el.classList.remove('hidden');
  });

  on('scCalcB', 'click', () => {
    const target = parseFloat(document.getElementById('scDevB').value) / 100;
    const avg    = D.meta.avgIncidencia;
    const needed = Math.ceil(target / avg);
    const el     = document.getElementById('scResultB');
    el.innerHTML = `
      <strong>Desvío a recuperar:</strong> ${pct(target)}<br>
      <strong>Incidencia media por actividad:</strong> ${pct(avg,4)}<br>
      <strong>Actividades necesarias (estimado):</strong>
      <span style="font-size:20px;color:#002f6c;font-weight:700">${needed}</span><br>
      <small>Asumiendo completar cada actividad de 0% → 100%</small>
    `;
    el.classList.remove('hidden');
  });
}

// ── Tabs ──────────────────────────────────────────────────────────────────────
function setupTabs() {
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab)?.classList.add('active');
    });
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function on(id, event, fn) { document.getElementById(id)?.addEventListener(event, fn); }
function destroyChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }
function set(id, v) { const el = document.getElementById(id); if (el) el.textContent = v; }
function pct(v, d=2)     { return v == null ? '—' : (v*100).toFixed(d)+'%'; }
function signPct(v, d=2) { return v == null ? '—' : (v>=0?'+':'')+(v*100).toFixed(d)+'%'; }
function fmtDate(iso) {
  if (!iso) return '—';
  const [y,m,d] = iso.split('-');
  return `${d}/${m}/${y}`;
}
function devClass(v) { return v==null?'dev-neutral':v<-0.001?'dev-neg':v>0.001?'dev-pos':'dev-neutral'; }
function statusBadge(r) {
  const map = { completed:['badge-completed','Completada'], inProgress:['badge-inProgress','En progreso'],
                late:['badge-late','Atrasada'], notStarted:['badge-notStarted','No iniciada'] };
  const [cls,lbl] = map[r.status] || ['badge-notStarted','—'];
  return `<span class="badge ${cls}">${lbl}</span>`;
}
function pbarDuo(plan, real) {
  return `<div class="pbar-wrap">
    <div class="pbar"><div class="pbar-fill pbar-plan" style="width:${(plan*100).toFixed(1)}%"></div></div>
    <div class="pbar"><div class="pbar-fill pbar-real" style="width:${(real*100).toFixed(1)}%"></div></div>
  </div>`;
}
function tableWrap(thead, tbody) {
  return `<div class="grid-wrap"><table><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
}

let toastTimer;
function showToast(msg, isError = false, ms = 3500) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.background = isError ? 'var(--danger)' : 'var(--text)';
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  if (ms < 60000) toastTimer = setTimeout(() => el.classList.add('hidden'), ms);
}
