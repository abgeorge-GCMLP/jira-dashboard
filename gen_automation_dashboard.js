'use strict';
const fs   = require('fs');
const path = require('path');
const stats = require('./automation_stats.json');

// ── Derive sorted month list ──────────────────────────────────────────────────
const months = [...new Set(stats.monthly.map(r => r.Month))].sort();
const monthLabels = months.map(m => {
  const [y, mo] = m.split('-');
  return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][+mo-1] + " '" + y.slice(2);
});

// ── Index monthly data by DocType+Month ──────────────────────────────────────
const byTypeMonth = {};
stats.monthly.forEach(r => {
  byTypeMonth[r.DocType + '|' + r.Month] = {
    total: +r.Total, completed: +r.Completed, automatable: +r.Automatable
  };
});

const caData   = months.map(m => byTypeMonth['CapitalActivity|'   + m] || {total:0,completed:0,automatable:0});
const pcapData = months.map(m => byTypeMonth['Pcap|'              + m] || {total:0,completed:0,automatable:0});
const fsData   = months.map(m => byTypeMonth['FinancialStatement|'+ m] || {total:0,completed:0,automatable:0});

// ── KPIs ─────────────────────────────────────────────────────────────────────
const caTotal   = caData.reduce((s,r)=>s+r.total,0);
const pcapTotal = pcapData.reduce((s,r)=>s+r.total,0);
const fsTotal   = fsData.reduce((s,r)=>s+r.total,0);
const grandTotal= caTotal + pcapTotal + fsTotal;

const caComp   = caData.reduce((s,r)=>s+r.completed,0);
const pcapComp = pcapData.reduce((s,r)=>s+r.completed,0);
const fsComp   = fsData.reduce((s,r)=>s+r.completed,0);
const grandComp= caComp + pcapComp + fsComp;

const caAuto   = caData.reduce((s,r)=>s+r.automatable,0);
const pcapAuto = pcapData.reduce((s,r)=>s+r.automatable,0);

const compRate = (grandComp / grandTotal * 100).toFixed(1);
const caAutoRate   = (caAuto   / caTotal   * 100).toFixed(1);
const pcapAutoRate = (pcapAuto / pcapTotal * 100).toFixed(1);

// ── CA Types aggregated by type ───────────────────────────────────────────────
const caTypesByMonth = {};
stats.caTypes.forEach(r => {
  if (!caTypesByMonth[r.CapitalActivityType]) caTypesByMonth[r.CapitalActivityType] = {};
  caTypesByMonth[r.CapitalActivityType][r.Month] = +r.Cnt;
});
const capitalCallData    = months.map(m => caTypesByMonth['CapitalCall']?.[m] || 0);
const distributionData   = months.map(m => caTypesByMonth['Distribution']?.[m] || 0);
const dikData            = months.map(m => caTypesByMonth['DistributionInKind']?.[m] || 0);

const totalCC   = capitalCallData.reduce((s,v)=>s+v,0);
const totalDist = distributionData.reduce((s,v)=>s+v,0);
const totalDIK  = dikData.reduce((s,v)=>s+v,0);

// ── Logo (base64 from file) ───────────────────────────────────────────────────
let logoSrc = '';
const logoPath = path.join(__dirname, 'jira-dashboard', 'GCM-Logo.png');
if (fs.existsSync(logoPath)) {
  const b64 = fs.readFileSync(logoPath).toString('base64');
  logoSrc = `data:image/png;base64,${b64}`;
}

// ── Deals data ────────────────────────────────────────────────────────────────
const dealNames     = stats.deals.map(d => d.DealName);
const dealCA        = stats.deals.map(d => +d.CA);
const dealPCAP      = stats.deals.map(d => +d.PCAP);
const dealFS        = stats.deals.map(d => +d.FS);
const dealTotal     = stats.deals.map(d => +d.Total);
const dealCompleted = stats.deals.map(d => +d.Completed);

// ── Completion rates per month per doc type ────────────────────────────────────
const caCompRate   = caData.map(r=>r.total>0?(r.completed/r.total*100).toFixed(1):null);
const pcapCompRate = pcapData.map(r=>r.total>0?(r.completed/r.total*100).toFixed(1):null);
const fsCompRate   = fsData.map(r=>r.total>0?(r.completed/r.total*100).toFixed(1):null);

const now = new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});

// ─────────────────────────────────────────────────────────────────────────────
// HTML
// ─────────────────────────────────────────────────────────────────────────────
const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Canoe Doc Automation Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Outfit:wght@400;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #0d0f14; --surface: #161a22; --surface2: #1e2330; --border: #2a3040;
  --text: #e2e8f0; --muted: #8b97b0;
  --cyan: #22d3ee; --green: #4ade80; --red: #f87171;
  --amber: #fbbf24; --blue: #60a5fa; --purple: #a78bfa;
  --orange: #fb923c; --teal: #2dd4bf; --pink: #f472b6;
}
body { font-family:'Inter',sans-serif; background:var(--bg); color:var(--text); min-height:100vh; }

/* Header */
.header {
  background:linear-gradient(135deg,#0d0f14,#1a1f2e); border-bottom:1px solid var(--border);
  padding:20px 32px; display:flex; align-items:center; justify-content:space-between;
  position:sticky; top:0; z-index:100; backdrop-filter:blur(10px);
}
.header-left { display:flex; align-items:center; gap:20px; }
.header-logo { height:36px;width:auto;display:block;filter:brightness(0) invert(1); }
.header-divider { width:1px;height:36px;background:var(--border); }
.header-title { font-family:'Outfit',sans-serif;font-size:1.4rem;font-weight:700; }
.header-sub { font-size:0.78rem;color:var(--muted);margin-top:2px; }
.period-pill {
  padding:5px 14px;border-radius:20px;font-size:0.75rem;font-weight:600;
  background:rgba(34,211,238,.12);color:var(--cyan);border:1px solid rgba(34,211,238,.3);
}

/* Main */
.main { padding:28px 32px;max-width:1700px;margin:0 auto; }
.section-title { font-family:'Outfit',sans-serif;font-size:1rem;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:14px;margin-top:28px; }

/* KPI */
.kpi-grid { display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:28px; }
.kpi-card {
  background:var(--surface);border:1px solid var(--border);border-radius:14px;
  padding:20px 18px;position:relative;overflow:hidden;transition:transform .15s;
}
.kpi-card:hover { transform:translateY(-2px); }
.kpi-card::before { content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--accent,var(--cyan)); }
.kpi-icon { font-size:1.2rem;margin-bottom:8px; }
.kpi-val { font-family:'Outfit',sans-serif;font-size:1.9rem;font-weight:700;line-height:1;margin-bottom:5px;color:var(--accent,var(--cyan)); }
.kpi-label { font-size:0.73rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-weight:500; }
.kpi-sub { font-size:0.72rem;color:var(--muted);margin-top:4px; }

/* Charts */
.charts-row { display:grid;gap:20px;margin-bottom:20px; }
.charts-row-2 { grid-template-columns:1fr 1fr; }
.charts-row-3 { grid-template-columns:2fr 1fr; }
.chart-card { background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:22px; }
.chart-title { font-family:'Outfit',sans-serif;font-size:1rem;font-weight:600;margin-bottom:6px; }
.chart-subtitle { font-size:0.76rem;color:var(--muted);margin-bottom:16px; }
.chart-wrap { position:relative; }
.chart-wrap-md  { height:220px; }
.chart-wrap-sm  { height:200px; }
.chart-wrap-lg  { height:260px; }
.chart-wrap-deal { height:460px; }

/* Legend row */
.legend-row { display:flex;gap:14px;flex-wrap:wrap;margin-top:12px; }
.legend-item { display:flex;align-items:center;gap:5px;font-size:0.73rem;color:var(--muted); }
.legend-dot { width:10px;height:10px;border-radius:2px;flex-shrink:0; }

/* Summary table */
.summary-table { width:100%;border-collapse:collapse;font-size:0.82rem; }
.summary-table th { text-align:left;padding:9px 14px;font-weight:600;font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--border); }
.summary-table td { padding:10px 14px;border-bottom:1px solid rgba(42,48,64,.6); }
.summary-table tr:last-child td { border-bottom:none; }
.summary-table tr:hover td { background:var(--surface2); }
.badge { display:inline-block;padding:2px 8px;border-radius:8px;font-size:0.71rem;font-weight:600; }
.badge-ca   { background:rgba(96,165,250,.15);color:var(--blue); }
.badge-pcap { background:rgba(167,139,250,.15);color:var(--purple); }
.badge-fs   { background:rgba(45,212,191,.15);color:var(--teal); }
.progress-bar { background:var(--surface2);border-radius:4px;height:6px;overflow:hidden;width:80px;display:inline-block;vertical-align:middle;margin-left:6px; }
.progress-fill { height:100%;border-radius:4px;background:var(--green); }
.progress-fill.amber { background:var(--amber); }
.progress-fill.blue  { background:var(--blue); }
.text-right { text-align:right; }
.text-green { color:var(--green); }
.text-amber { color:var(--amber); }
.text-blue  { color:var(--blue); }
.text-muted { color:var(--muted); }

/* Footer */
.footer { text-align:center;padding:24px 32px;color:var(--muted);font-size:0.75rem;border-top:1px solid var(--border);margin-top:32px; }
</style>
</head>
<body>

<!-- HEADER -->
<header class="header">
  <div class="header-left">
    ${logoSrc ? `<img src="${logoSrc}" class="header-logo" alt="GCM Grosvenor">
    <div class="header-divider"></div>` : ''}
    <div>
      <div class="header-title">Canoe Doc Automation Dashboard</div>
      <div class="header-sub">Capital Activity · PCAP · Financial Statements — document processing stats</div>
    </div>
  </div>
  <div class="period-pill">Mar 2025 – Mar 2026</div>
</header>

<div class="main">

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi-card" style="--accent:var(--cyan)">
      <div class="kpi-icon">📄</div>
      <div class="kpi-val">${grandTotal.toLocaleString()}</div>
      <div class="kpi-label">Total Docs</div>
      <div class="kpi-sub">Last 13 months</div>
    </div>
    <div class="kpi-card" style="--accent:var(--green)">
      <div class="kpi-icon">✅</div>
      <div class="kpi-val">${compRate}%</div>
      <div class="kpi-label">Completion Rate</div>
      <div class="kpi-sub">${grandComp.toLocaleString()} completed</div>
    </div>
    <div class="kpi-card" style="--accent:var(--blue)">
      <div class="kpi-icon">📊</div>
      <div class="kpi-val">${caTotal.toLocaleString()}</div>
      <div class="kpi-label">Capital Activity</div>
      <div class="kpi-sub">${caAutoRate}% automatable</div>
    </div>
    <div class="kpi-card" style="--accent:var(--purple)">
      <div class="kpi-icon">📋</div>
      <div class="kpi-val">${pcapTotal.toLocaleString()}</div>
      <div class="kpi-label">PCAP Docs</div>
      <div class="kpi-sub">${pcapAutoRate}% automatable</div>
    </div>
    <div class="kpi-card" style="--accent:var(--teal)">
      <div class="kpi-icon">📑</div>
      <div class="kpi-val">${fsTotal.toLocaleString()}</div>
      <div class="kpi-label">Financial Statements</div>
      <div class="kpi-sub">${(fsComp/fsTotal*100).toFixed(1)}% completed</div>
    </div>
    <div class="kpi-card" style="--accent:var(--amber)">
      <div class="kpi-icon">⚡</div>
      <div class="kpi-val">${((caAuto+pcapAuto)/(caTotal+pcapTotal)*100).toFixed(1)}%</div>
      <div class="kpi-label">Automatable (CA+PCAP)</div>
      <div class="kpi-sub">${(caAuto+pcapAuto).toLocaleString()} of ${(caTotal+pcapTotal).toLocaleString()}</div>
    </div>
  </div>

  <!-- ROW 1: Monthly Volume + CA Types Donut -->
  <div class="charts-row charts-row-3">
    <div class="chart-card">
      <div class="chart-title">Monthly Document Volume</div>
      <div class="chart-subtitle">Total docs received per month by type</div>
      <div class="chart-wrap chart-wrap-lg">
        <canvas id="volumeChart"></canvas>
      </div>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-dot" style="background:#60a5fa"></div>Capital Activity</div>
        <div class="legend-item"><div class="legend-dot" style="background:#a78bfa"></div>PCAP</div>
        <div class="legend-item"><div class="legend-dot" style="background:#2dd4bf"></div>Financial Statements</div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Capital Activity Types</div>
      <div class="chart-subtitle">Distribution across all 13 months</div>
      <div class="chart-wrap chart-wrap-lg" style="display:flex;align-items:center;justify-content:center;">
        <canvas id="caTypeDonut"></canvas>
      </div>
    </div>
  </div>

  <!-- ROW 2: Completion Rates + Monthly CA Types Breakdown -->
  <div class="charts-row charts-row-2">
    <div class="chart-card">
      <div class="chart-title">Monthly Completion Rate</div>
      <div class="chart-subtitle">% of docs completed per month by type</div>
      <div class="chart-wrap chart-wrap-md">
        <canvas id="compRateChart"></canvas>
      </div>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-dot" style="background:#60a5fa"></div>Capital Activity</div>
        <div class="legend-item"><div class="legend-dot" style="background:#a78bfa"></div>PCAP</div>
        <div class="legend-item"><div class="legend-dot" style="background:#2dd4bf"></div>Financial Statements</div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">CA Types by Month</div>
      <div class="chart-subtitle">Capital Call · Distribution · Distribution in Kind</div>
      <div class="chart-wrap chart-wrap-md">
        <canvas id="caTypesChart"></canvas>
      </div>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-dot" style="background:#22d3ee"></div>Capital Call</div>
        <div class="legend-item"><div class="legend-dot" style="background:#4ade80"></div>Distribution</div>
        <div class="legend-item"><div class="legend-dot" style="background:#fbbf24"></div>Dist. in Kind</div>
      </div>
    </div>
  </div>

  <!-- Top 20 Deals -->
  <div class="chart-card" style="margin-bottom:20px">
    <div class="chart-title">Top 20 Deals by Document Volume</div>
    <div class="chart-subtitle">Capital Activity · PCAP · Financial Statements — sorted by total</div>
    <div class="chart-wrap chart-wrap-deal">
      <canvas id="dealsChart"></canvas>
    </div>
    <div class="legend-row" style="margin-top:14px">
      <div class="legend-item"><div class="legend-dot" style="background:#60a5fa"></div>Capital Activity</div>
      <div class="legend-item"><div class="legend-dot" style="background:#a78bfa"></div>PCAP</div>
      <div class="legend-item"><div class="legend-dot" style="background:#2dd4bf"></div>Financial Statements</div>
    </div>
  </div>

  <!-- Summary Table -->
  <div class="chart-card" style="margin-bottom:20px">
    <div class="chart-title" style="margin-bottom:16px">Document Type Summary</div>
    <table class="summary-table">
      <thead>
        <tr>
          <th>Doc Type</th>
          <th class="text-right">Total Docs</th>
          <th class="text-right">Completed</th>
          <th class="text-right">Completion Rate</th>
          <th class="text-right">Automatable</th>
          <th class="text-right">Auto Rate</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><span class="badge badge-ca">Capital Activity</span></td>
          <td class="text-right">${caTotal.toLocaleString()}</td>
          <td class="text-right text-green">${caComp.toLocaleString()}</td>
          <td class="text-right">
            ${(caComp/caTotal*100).toFixed(1)}%
            <span class="progress-bar"><span class="progress-fill blue" style="width:${(caComp/caTotal*100).toFixed(1)}%"></span></span>
          </td>
          <td class="text-right text-blue">${caAuto.toLocaleString()}</td>
          <td class="text-right">${caAutoRate}%
            <span class="progress-bar"><span class="progress-fill" style="width:${caAutoRate}%"></span></span>
          </td>
        </tr>
        <tr>
          <td><span class="badge badge-pcap">PCAP</span></td>
          <td class="text-right">${pcapTotal.toLocaleString()}</td>
          <td class="text-right text-green">${pcapComp.toLocaleString()}</td>
          <td class="text-right">
            ${(pcapComp/pcapTotal*100).toFixed(1)}%
            <span class="progress-bar"><span class="progress-fill blue" style="width:${(pcapComp/pcapTotal*100).toFixed(1)}%"></span></span>
          </td>
          <td class="text-right text-blue">${pcapAuto.toLocaleString()}</td>
          <td class="text-right">${pcapAutoRate}%
            <span class="progress-bar"><span class="progress-fill amber" style="width:${pcapAutoRate}%"></span></span>
          </td>
        </tr>
        <tr>
          <td><span class="badge badge-fs">Financial Statements</span></td>
          <td class="text-right">${fsTotal.toLocaleString()}</td>
          <td class="text-right text-green">${fsComp.toLocaleString()}</td>
          <td class="text-right">
            ${(fsComp/fsTotal*100).toFixed(1)}%
            <span class="progress-bar"><span class="progress-fill blue" style="width:${(fsComp/fsTotal*100).toFixed(1)}%"></span></span>
          </td>
          <td class="text-right text-muted">N/A</td>
          <td class="text-right text-muted">—</td>
        </tr>
      </tbody>
    </table>
  </div>

</div><!-- .main -->

<footer class="footer">
  Generated ${now} · Canoe Doc Automation · Source: documentautomation-prd · Mar 2025 – Mar 2026
</footer>

<script>
Chart.defaults.color = '#8b97b0';
Chart.defaults.borderColor = '#2a3040';
Chart.defaults.font.family = "'Inter', sans-serif";

const MONTHS = ${JSON.stringify(monthLabels)};

// ── Monthly Volume (stacked bar) ──────────────────────────────────────────────
new Chart(document.getElementById('volumeChart'), {
  type: 'bar',
  data: {
    labels: MONTHS,
    datasets: [
      { label:'Capital Activity', data:${JSON.stringify(caData.map(r=>r.total))},    backgroundColor:'rgba(96,165,250,0.75)',  stack:'s' },
      { label:'PCAP',             data:${JSON.stringify(pcapData.map(r=>r.total))},  backgroundColor:'rgba(167,139,250,0.75)', stack:'s' },
      { label:'Financial Stmts',  data:${JSON.stringify(fsData.map(r=>r.total))},    backgroundColor:'rgba(45,212,191,0.65)',  stack:'s' }
    ]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{display:false}, tooltip:{mode:'index',intersect:false} },
    scales:{
      x:{ stacked:true, grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} },
      y:{ stacked:true, grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} }
    }
  }
});

// ── CA Type Donut ─────────────────────────────────────────────────────────────
new Chart(document.getElementById('caTypeDonut'), {
  type: 'doughnut',
  data: {
    labels: ['Capital Call','Distribution','Dist. in Kind'],
    datasets: [{
      data: [${totalCC}, ${totalDist}, ${totalDIK}],
      backgroundColor: ['rgba(34,211,238,0.8)','rgba(74,222,128,0.8)','rgba(251,191,36,0.8)'],
      borderColor: ['#22d3ee','#4ade80','#fbbf24'],
      borderWidth: 2,
      hoverOffset: 8
    }]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    cutout:'65%',
    plugins:{
      legend:{ position:'bottom', labels:{ padding:14, font:{size:12}, color:'#8b97b0' } },
      tooltip:{
        callbacks:{
          label: ctx => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            return ' ' + ctx.label + ': ' + ctx.parsed.toLocaleString() + ' (' + (ctx.parsed/total*100).toFixed(1) + '%)';
          }
        }
      }
    }
  }
});

// ── Completion Rate Line Chart ─────────────────────────────────────────────────
new Chart(document.getElementById('compRateChart'), {
  type: 'line',
  data: {
    labels: MONTHS,
    datasets: [
      {
        label:'Capital Activity', data:${JSON.stringify(caCompRate)},
        borderColor:'#60a5fa', backgroundColor:'rgba(96,165,250,0.1)',
        borderWidth:2, pointRadius:3, tension:0.35, fill:false
      },
      {
        label:'PCAP', data:${JSON.stringify(pcapCompRate)},
        borderColor:'#a78bfa', backgroundColor:'rgba(167,139,250,0.1)',
        borderWidth:2, pointRadius:3, tension:0.35, fill:false
      },
      {
        label:'Financial Stmts', data:${JSON.stringify(fsCompRate)},
        borderColor:'#2dd4bf', backgroundColor:'rgba(45,212,191,0.1)',
        borderWidth:2, pointRadius:3, tension:0.35, fill:false
      }
    ]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{display:false}, tooltip:{mode:'index',intersect:false} },
    scales:{
      x:{ grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} },
      y:{
        grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11},callback:v=>v+'%'},
        min:80, max:100
      }
    }
  }
});

// ── CA Types Monthly (stacked bar) ────────────────────────────────────────────
new Chart(document.getElementById('caTypesChart'), {
  type: 'bar',
  data: {
    labels: MONTHS,
    datasets: [
      { label:'Capital Call',      data:${JSON.stringify(capitalCallData)},  backgroundColor:'rgba(34,211,238,0.75)', stack:'s' },
      { label:'Distribution',      data:${JSON.stringify(distributionData)}, backgroundColor:'rgba(74,222,128,0.75)', stack:'s' },
      { label:'Dist. in Kind',     data:${JSON.stringify(dikData)},          backgroundColor:'rgba(251,191,36,0.75)', stack:'s' }
    ]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{display:false}, tooltip:{mode:'index',intersect:false} },
    scales:{
      x:{ stacked:true, grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} },
      y:{ stacked:true, grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} }
    }
  }
});

// ── Top 20 Deals (horizontal stacked bar) ─────────────────────────────────────
new Chart(document.getElementById('dealsChart'), {
  type: 'bar',
  data: {
    labels: ${JSON.stringify(dealNames.map(n => n.length > 45 ? n.slice(0,44)+'…' : n))},
    datasets: [
      { label:'Capital Activity', data:${JSON.stringify(dealCA)},   backgroundColor:'rgba(96,165,250,0.8)',  stack:'s' },
      { label:'PCAP',             data:${JSON.stringify(dealPCAP)}, backgroundColor:'rgba(167,139,250,0.8)', stack:'s' },
      { label:'Financial Stmts',  data:${JSON.stringify(dealFS)},   backgroundColor:'rgba(45,212,191,0.75)', stack:'s' }
    ]
  },
  options: {
    indexAxis:'y',
    responsive:true, maintainAspectRatio:false,
    plugins:{
      legend:{display:false},
      tooltip:{
        mode:'index', intersect:false,
        callbacks:{
          afterBody: items => {
            const total = items.reduce((s,i)=>s+i.parsed.x,0);
            const completed = ${JSON.stringify(dealCompleted)}[items[0].dataIndex];
            return ['──────────────', 'Total: '+total, 'Completed: '+completed+' ('+(completed/total*100).toFixed(0)+'%)'];
          }
        }
      }
    },
    scales:{
      x:{ stacked:true, grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} },
      y:{ stacked:true, grid:{color:'rgba(42,48,64,.5)'}, ticks:{font:{size:11}} }
    }
  }
});
</script>
</body>
</html>`;

const outPath = path.join(__dirname, 'canoe-automation-dashboard.html');
fs.writeFileSync(outPath, html);
console.log('Written:', outPath);
console.log(`KPIs: Total=${grandTotal.toLocaleString()}, Completion=${compRate}%, CA Automatable=${caAutoRate}%, PCAP Automatable=${pcapAutoRate}%`);
