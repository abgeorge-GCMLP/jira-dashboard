const fs   = require('fs');
const XLSX = require('./node_modules/xlsx');
const data = require('./pcap_data.json');

// ── PEREI contacts: key = "deal|le" ──────────────────────────────────────────
const pereibyDealLe = {};
try {
  const csv = fs.readFileSync('C:/Users/ABGEORGE/PEREI Funds - Investment Contacts _Working file(PEREI Funds).csv', 'utf8');
  csv.split('\n').map(l => l.replace(/\r/g,'').trim()).filter(l => l).slice(1).forEach(line => {
    const cols = []; let cur = '', inQ = false;
    for (const c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { cols.push(cur.trim()); cur = ''; }
      else { cur += c; }
    }
    cols.push(cur.trim());
    if (cols.length >= 3) {
      const deal = cols[0].replace(/^"|"$/g,'').trim();
      const le   = cols[1].replace(/^"|"$/g,'').trim();
      const email= cols[2].replace(/^"|"$/g,'').trim();
      if (deal && email) pereibyDealLe[deal + '|' + le] = email;
    }
  });
  console.log('PEREI contacts loaded:', Object.keys(pereibyDealLe).length);
} catch(e) { console.log('PEREI contacts not found'); }

// ── CIC contacts: key = "Fund Name" ──────────────────────────────────────────
const cicByDeal = {};
try {
  const wb = XLSX.readFile('C:/Users/ABGEORGE/CIC_Epic List.xlsx');
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''});
  rows.forEach(r => {
    const deal  = (r['Fund Name'] || '').trim();
    const email = (r['GP Contact Email'] || '').trim();
    if (deal && email) cicByDeal[deal] = email;
  });
  console.log('CIC contacts loaded:', Object.keys(cicByDeal).length);
} catch(e) { console.log('CIC Epic List not found'); }

// ── NPS contacts: key = "Manager" (all 3 sheets) ─────────────────────────────
const npsByDeal = {};
try {
  const wb = XLSX.readFile('C:/Users/ABGEORGE/GP contacts NPS Feb 2026.xlsx');
  wb.SheetNames.forEach(sheet => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {defval:''});
    rows.forEach(r => {
      const deal  = (r['Manager'] || '').trim();
      const email = (r['Template/ Accounting/Document contact'] || r['Mail'] || '').trim();
      if (deal && email) npsByDeal[deal] = email;
    });
  });
  console.log('NPS contacts loaded:', Object.keys(npsByDeal).length);
} catch(e) { console.log('NPS contacts not found'); }

// ── Build merged contacts map (PEREI priority, then CIC, then NPS) ────────────
// Store all sources per deal for the browser to pick at runtime
// Key = "deal|le" → { perei, cic, nps } (whichever are available)
const contactsMap = {};
const dealLeSet = new Set(data.map(r => r.deal + '|' + r.le));
dealLeSet.forEach(key => {
  const [deal, le] = key.split('|');
  const perei = pereibyDealLe[key] || '';
  const cic   = cicByDeal[deal]   || '';
  const nps   = npsByDeal[deal]   || '';
  if (perei || cic || nps) contactsMap[key] = { perei, cic, nps };
});
console.log('Merged contacts (deal|le keys with any match):', Object.keys(contactsMap).length);

const LP_ADMIN_IDS = [418, 511, 534, 546, 616, 646, 647, 648, 649, 650];

const q1Received = data.filter(r => r.q1 === 'Received').length;
const q1Missing  = data.filter(r => r.q1 === 'Not Received').length;
const totalQ4    = data.length;

const acctReceived = data.filter(r => r.docType === 'Account Statement' && r.q1 === 'Received').length;
const acctMissing  = data.filter(r => r.docType === 'Account Statement' && r.q1 === 'Not Received').length;
const finReceived  = data.filter(r => r.docType === 'Financials' && r.q1 === 'Received').length;
const finMissing   = data.filter(r => r.docType === 'Financials' && r.q1 === 'Not Received').length;
const acctTotal    = acctReceived + acctMissing;
const finTotal     = finReceived  + finMissing;

const dataJson        = JSON.stringify(data);
const lpAdminIdsJson  = JSON.stringify(LP_ADMIN_IDS);
const contactsMapJson = JSON.stringify(contactsMap);

const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Canoe Doc Receipt Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Outfit:wght@400;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js"></script>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg: #0d0f14; --surface: #161a22; --surface2: #1e2330; --border: #2a3040;
    --text: #e2e8f0; --muted: #8b97b0;
    --cyan: #22d3ee; --green: #4ade80; --red: #f87171;
    --amber: #fbbf24; --blue: #60a5fa; --purple: #a78bfa;
    --orange: #fb923c; --teal: #2dd4bf;
  }
  body { font-family: 'Inter', sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }

  /* Header */
  .header {
    background: linear-gradient(135deg,#0d0f14,#1a1f2e); border-bottom: 1px solid var(--border);
    padding: 20px 32px; display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 100; backdrop-filter: blur(10px);
  }
  .header-left { display: flex; align-items: center; gap: 20px; }
  .header-logo { height:36px;width:auto;display:block;filter:brightness(0) invert(1); }
  .header-divider { width:1px;height:36px;background:var(--border); }
  .header-title { font-family:'Outfit',sans-serif;font-size:1.4rem;font-weight:700; }
  .header-sub { font-size:0.78rem;color:var(--muted);margin-top:2px; }
  .period-pills { display:flex;gap:8px;align-items:center; }
  .period-pill { padding:5px 12px;border-radius:20px;font-size:0.75rem;font-weight:600;border:1px solid; }
  .pill-q4 { background:rgba(250,191,36,.12);color:var(--amber);border-color:rgba(250,191,36,.3); }
  .pill-q1 { background:rgba(34,211,238,.12);color:var(--cyan);border-color:rgba(34,211,238,.3); }
  .pill-sep { color:var(--muted);font-size:.85rem; }

  /* Main */
  .main { padding:28px 32px;max-width:1700px;margin:0 auto; }

  /* KPI - 3 tiles */
  .kpi-grid { display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:28px;max-width:700px; }
  .kpi-card {
    background:var(--surface);border:1px solid var(--border);border-radius:14px;
    padding:20px 22px;position:relative;overflow:hidden;transition:transform .15s;
  }
  .kpi-card:hover { transform:translateY(-2px); }
  .kpi-card::before { content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--accent,var(--cyan)); }
  .kpi-icon { font-size:1.4rem;margin-bottom:10px; }
  .kpi-val { font-family:'Outfit',sans-serif;font-size:2rem;font-weight:700;line-height:1;margin-bottom:6px;color:var(--accent,var(--cyan)); }
  .kpi-label { font-size:0.78rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-weight:500; }

  /* Charts */
  .charts-row { display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:28px; }
  .chart-card { background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:22px; }
  .chart-title { font-family:'Outfit',sans-serif;font-size:1rem;font-weight:600;margin-bottom:18px; }
  .chart-wrap { position:relative;height:200px; }

  /* Filters bar */
  .filters-bar {
    background:var(--surface);border:1px solid var(--border);border-radius:14px;
    padding:14px 22px;margin-bottom:12px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;
  }
  .filter-divider { width:1px;height:28px;background:var(--border);flex-shrink:0; }
  .filter-group { display:flex;align-items:center;gap:8px; }
  .filter-label { font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;font-weight:600;white-space:nowrap; }
  .search-input {
    background:var(--surface2);border:1px solid var(--border);border-radius:8px;
    padding:7px 12px;color:var(--text);font-size:0.83rem;font-family:'Inter',sans-serif;
    width:190px;outline:none;transition:border-color .15s;
  }
  .search-input:focus { border-color:var(--cyan); }
  .search-input::placeholder { color:var(--muted); }
  .date-input {
    background:var(--surface2);border:1px solid var(--border);border-radius:8px;
    padding:6px 10px;color:var(--text);font-size:0.83rem;font-family:'Inter',sans-serif;
    outline:none;transition:border-color .15s;color-scheme:dark;cursor:pointer;
  }
  .date-input:focus { border-color:var(--cyan); }
  .date-range { display:flex;align-items:center;gap:6px; }
  .date-range-sep { font-size:0.75rem;color:var(--muted); }
  .btn-group { display:flex;gap:4px; }
  .btn-filter {
    padding:6px 12px;border-radius:8px;font-size:0.79rem;font-weight:500;cursor:pointer;
    border:1px solid var(--border);background:var(--surface2);color:var(--muted);
    transition:all .15s;font-family:'Inter',sans-serif;
  }
  .btn-filter:hover { border-color:var(--cyan);color:var(--text); }
  .btn-filter.active { background:rgba(34,211,238,.15);border-color:var(--cyan);color:var(--cyan); }

  /* Admin segmented */
  .admin-seg { display:flex;background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:3px;gap:2px; }
  .admin-seg-btn {
    display:flex;align-items:center;gap:6px;padding:6px 14px;border-radius:8px;
    font-size:0.81rem;font-weight:600;cursor:pointer;border:none;background:transparent;
    color:var(--muted);transition:all .18s;font-family:'Inter',sans-serif;white-space:nowrap;
  }
  .admin-seg-btn:hover:not(.seg-active) { color:var(--text);background:rgba(255,255,255,.04); }
  .admin-seg-btn.seg-active.seg-all { background:rgba(34,211,238,.15);color:var(--cyan);box-shadow:0 0 0 1px rgba(34,211,238,.35); }
  .admin-seg-btn.seg-active.seg-lp  { background:rgba(251,146,60,.18);color:var(--orange);box-shadow:0 0 0 1px rgba(251,146,60,.45); }
  .admin-seg-btn.seg-active.seg-gp  { background:rgba(45,212,191,.18);color:var(--teal);box-shadow:0 0 0 1px rgba(45,212,191,.45); }
  .seg-dot { width:7px;height:7px;border-radius:50%;flex-shrink:0; }
  .seg-all .seg-dot { background:var(--muted); }  .seg-active.seg-all .seg-dot { background:var(--cyan); }
  .seg-lp  .seg-dot { background:var(--orange); }
  .seg-gp  .seg-dot { background:var(--teal); }
  .seg-count { font-size:0.7rem;font-weight:700;padding:1px 6px;border-radius:20px;background:rgba(255,255,255,.07);color:var(--muted);transition:all .18s; }
  .seg-active.seg-all .seg-count { background:rgba(34,211,238,.2);color:var(--cyan); }
  .seg-active.seg-lp  .seg-count { background:rgba(251,146,60,.2);color:var(--orange); }
  .seg-active.seg-gp  .seg-count { background:rgba(45,212,191,.2);color:var(--teal); }

  /* Export button */
  .export-btn {
    display:flex;align-items:center;gap:7px;padding:7px 16px;border-radius:9px;
    background:rgba(74,222,128,.1);border:1px solid rgba(74,222,128,.35);color:var(--green);
    font-size:0.81rem;font-weight:600;cursor:pointer;font-family:'Inter',sans-serif;
    transition:all .18s;white-space:nowrap;margin-left:auto;
  }
  .export-btn:hover { background:rgba(74,222,128,.2);border-color:rgba(74,222,128,.6); }
  .export-btn svg { flex-shrink:0; }
  .clear-btn {
    display:flex;align-items:center;gap:6px;padding:7px 14px;border-radius:9px;
    background:rgba(251,113,133,.1);border:1px solid rgba(251,113,133,.35);color:#fb7185;
    font-size:0.81rem;font-weight:600;cursor:pointer;font-family:'Inter',sans-serif;
    transition:all .18s;white-space:nowrap;
  }
  .clear-btn:hover { background:rgba(251,113,133,.2);border-color:rgba(251,113,133,.6); }
  .clear-btn svg { flex-shrink:0; }
  .draft-cell { white-space:nowrap; }
  .draft-btn {
    display:inline-flex;align-items:center;gap:5px;padding:4px 11px;border-radius:7px;
    background:rgba(139,92,246,.12);border:1px solid rgba(139,92,246,.35);color:#a78bfa;
    font-size:0.78rem;font-weight:600;cursor:pointer;font-family:'Inter',sans-serif;
    transition:all .15s;white-space:nowrap;
  }
  .draft-btn:hover { background:rgba(139,92,246,.25);border-color:rgba(139,92,246,.65); }

  .results-count { font-size:0.8rem;color:var(--muted); }

  /* Table */
  .table-card { background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden; }
  .table-header {
    padding:16px 22px;border-bottom:1px solid var(--border);
    display:flex;align-items:center;justify-content:space-between;
  }
  .table-title { font-family:'Outfit',sans-serif;font-size:1rem;font-weight:600; }
  .table-wrap { overflow-x:auto; }
  table { width:100%;border-collapse:collapse;font-size:0.83rem; }
  thead { background:var(--surface2); }
  thead tr:first-child th {
    padding:10px 14px;text-align:left;font-size:0.71rem;font-weight:600;
    color:var(--muted);text-transform:uppercase;letter-spacing:.07em;
    border-bottom:1px solid rgba(42,48,64,.5);white-space:nowrap;cursor:pointer;user-select:none;
  }
  thead tr:first-child th:hover { color:var(--text); }
  /* Column search row */
  thead tr.col-search th {
    padding:5px 8px;border-bottom:1px solid var(--border);background:var(--surface);
  }
  .col-search-input {
    width:100%;background:var(--surface2);border:1px solid var(--border);border-radius:6px;
    padding:5px 8px;color:var(--text);font-size:0.78rem;font-family:'Inter',sans-serif;outline:none;
    transition:border-color .15s;
  }
  .col-search-input:focus { border-color:var(--cyan); }
  .col-search-input::placeholder { color:rgba(139,151,176,.5); }
  td { padding:10px 14px;border-bottom:1px solid rgba(42,48,64,.55);vertical-align:middle; }
  tr:last-child td { border-bottom:none; }
  tr:hover td { background:rgba(255,255,255,.02); }
  tr.row-lp td:nth-child(2) { border-left:2px solid rgba(251,146,60,.5);padding-left:12px; }
  tr.row-gp td:nth-child(2) { border-left:2px solid rgba(45,212,191,.4);padding-left:12px; }
  .tick-cell { text-align:center;width:36px;padding:10px 4px !important; }
  .row-tick { font-size:1rem;color:var(--green);opacity:0;transition:opacity .15s; }
  .row-tick.visible { opacity:1; }
  .deal-name { font-weight:500;color:var(--text); }
  .le-lp { color:var(--orange)!important;font-weight:500; }
  .le-gp { color:var(--teal)!important;font-weight:500; }
  .le-def { color:var(--muted);font-size:0.8rem; }
  .doc-badge { display:inline-block;padding:3px 9px;border-radius:6px;font-size:0.72rem;font-weight:600;background:rgba(167,139,250,.15);color:var(--purple);border:1px solid rgba(167,139,250,.3); }
  .status-pill { display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:600;white-space:nowrap; }
  .pill-received { background:rgba(74,222,128,.12);color:var(--green);border:1px solid rgba(74,222,128,.25); }
  .pill-missing  { background:rgba(248,113,113,.12);color:var(--red);  border:1px solid rgba(248,113,113,.25); }
  /* FM Contacted display badge */
  .fm-cell { text-align:center;width:60px; }
  .fm-badge {
    display:inline-flex;align-items:center;justify-content:center;
    width:28px;height:28px;border-radius:50%;
    border:2px solid rgba(74,222,128,.25);color:transparent;
    font-size:14px;transition:all .2s;
  }
  .fm-badge.contacted {
    background:rgba(74,222,128,.15);border-color:var(--green);
    color:var(--green);
  }

  /* Pagination */
  .pagination { padding:14px 22px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px; }
  .page-info { font-size:0.8rem;color:var(--muted); }
  .page-btns { display:flex;gap:5px;align-items:center; }
  .page-btn {
    padding:5px 11px;border-radius:7px;border:1px solid var(--border);background:var(--surface2);
    color:var(--muted);font-size:0.79rem;cursor:pointer;transition:all .15s;font-family:'Inter',sans-serif;
  }
  .page-btn:hover:not(:disabled) { border-color:var(--cyan);color:var(--cyan); }
  .page-btn:disabled { opacity:.35;cursor:default; }
  .page-btn.current { background:rgba(34,211,238,.15);border-color:var(--cyan);color:var(--cyan); }
  .page-size-select { background:var(--surface2);border:1px solid var(--border);border-radius:7px;color:var(--text);font-size:0.79rem;padding:5px 8px;font-family:'Inter',sans-serif;outline:none; }
  .empty-row td { text-align:center;padding:48px;color:var(--muted);font-size:.9rem; }
  .canoe-link {
    display:inline-flex;align-items:center;gap:4px;
    color:var(--blue);font-size:0.75rem;font-weight:500;font-family:'Inter',monospace;
    text-decoration:none;padding:3px 8px;border-radius:6px;
    background:rgba(96,165,250,.1);border:1px solid rgba(96,165,250,.25);
    transition:all .15s;white-space:nowrap;
  }
  .canoe-link:hover { background:rgba(96,165,250,.2);border-color:rgba(96,165,250,.5); }
  .canoe-link::after { content:'↗';font-size:0.65rem;opacity:.7; }
  .dt-text { font-size:0.79rem;color:var(--text);white-space:nowrap; }
  .dt-na   { font-size:0.78rem;color:var(--muted);font-style:italic; }
  .na-text { color:var(--muted);font-size:0.78rem;font-style:italic; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <img src="GCM-Logo.png" alt="GCM" class="header-logo">
    <div class="header-divider"></div>
    <div>
      <div class="header-title">Canoe Doc Receipt Dashboard</div>
      <div class="header-sub">Document receipt tracking · Private Markets</div>
    </div>
  </div>
  <div class="period-pills">
    <span class="period-pill pill-q4">Q4 2025 Baseline</span>
    <span class="pill-sep">→</span>
    <span class="period-pill pill-q1">Q1 2026 Current</span>
  </div>
</div>

<div class="main">

  <!-- KPIs (3 tiles) -->
  <div class="kpi-grid">
    <div class="kpi-card" style="--accent:var(--amber)">
      <div class="kpi-icon">✅</div>
      <div class="kpi-val">${totalQ4.toLocaleString()}</div>
      <div class="kpi-label">Q4 2025 Received</div>
    </div>
    <div class="kpi-card" style="--accent:var(--green)">
      <div class="kpi-icon">📬</div>
      <div class="kpi-val">${q1Received.toLocaleString()}</div>
      <div class="kpi-label">Q1 2026 Received</div>
    </div>
    <div class="kpi-card" style="--accent:var(--red)">
      <div class="kpi-icon">⚠️</div>
      <div class="kpi-val">${q1Missing.toLocaleString()}</div>
      <div class="kpi-label">Q1 2026 Missing</div>
    </div>
  </div>

  <!-- Charts -->
  <div class="charts-row">
    <div class="chart-card">
      <div class="chart-title">Q1 2026 Receipt Status by Document Type</div>
      <div class="chart-wrap"><canvas id="statusChart"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Q4 vs Q1 Comparison by Document Type</div>
      <div class="chart-wrap"><canvas id="compareChart"></canvas></div>
    </div>
  </div>

  <!-- Filters -->
  <div class="filters-bar">
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Doc Type</span>
      <div class="btn-group" id="docTypeGroup">
        <button class="btn-filter active" onclick="setFilter('docType','all',this)">All</button>
        <button class="btn-filter" onclick="setFilter('docType','Account Statement',this)">Account Statement</button>
        <button class="btn-filter" onclick="setFilter('docType','Financials',this)">Financials</button>
      </div>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Q1 Status</span>
      <div class="btn-group" id="statusGroup">
        <button class="btn-filter active" onclick="setFilter('status','all',this)">All</button>
        <button class="btn-filter" onclick="setFilter('status','Received',this)">Received</button>
        <button class="btn-filter" onclick="setFilter('status','Not Received',this)">Not Received</button>
      </div>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">FM Contacted</span>
      <div class="btn-group" id="fmGroup">
        <button class="btn-filter active" onclick="setFilter('fm','all',this)">All</button>
        <button class="btn-filter" onclick="setFilter('fm','yes',this)">Contacted</button>
        <button class="btn-filter" onclick="setFilter('fm','no',this)">Not Contacted</button>
      </div>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Q4 Received</span>
      <div class="date-range">
        <input type="date" class="date-input" id="q4DateFrom" onchange="applyFilters()" title="From">
        <span class="date-range-sep">→</span>
        <input type="date" class="date-input" id="q4DateTo" onchange="applyFilters()" title="To">
      </div>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Q1 Received</span>
      <div class="date-range">
        <input type="date" class="date-input" id="q1DateFrom" onchange="applyFilters()" title="From">
        <span class="date-range-sep">→</span>
        <input type="date" class="date-input" id="q1DateTo" onchange="applyFilters()" title="To">
      </div>
    </div>
    <div class="filter-divider"></div>
    <div class="filter-group">
      <span class="filter-label">Admin View</span>
      <div class="admin-seg" id="adminSeg">
        <button class="admin-seg-btn seg-all seg-active" onclick="setAdmin('all',this)">
          <span class="seg-dot"></span>All<span class="seg-count" id="cnt-all"></span>
        </button>
        <button class="admin-seg-btn seg-lp" onclick="setAdmin('lp',this)">
          <span class="seg-dot"></span>LP Admin<span class="seg-count" id="cnt-lp"></span>
        </button>
        <button class="admin-seg-btn seg-gp" onclick="setAdmin('gp',this)">
          <span class="seg-dot"></span>GP Admin<span class="seg-count" id="cnt-gp"></span>
        </button>
      </div>
    </div>
    <button class="clear-btn" id="clearColBtn" onclick="clearColFilters()" style="display:none">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
      </svg>
      Clear Column Filters
    </button>
    <button class="export-btn" onclick="exportExcel()">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
      </svg>
      Export to Excel
    </button>
  </div>
  <div style="padding:0 4px 14px;display:flex;justify-content:flex-end;">
    <span class="results-count" id="resultsCount"></span>
  </div>

  <!-- Table -->
  <div class="table-card">
    <div class="table-header">
      <span class="table-title">Document Receipt Details</span>
      <div style="display:flex;align-items:center;gap:10px;">
        <span style="font-size:0.78rem;color:var(--muted);">Rows per page:</span>
        <select class="page-size-select" id="pageSizeSelect" onchange="changePageSize(this.value)">
          <option value="25">25</option><option value="50" selected>50</option><option value="100">100</option>
        </select>
      </div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th class="tick-cell" title="FM Contacted indicator"></th>
            <th onclick="sortTable('deal')">Deal <span id="sort-deal">↑</span></th>
            <th onclick="sortTable('le')">Legal Entity <span id="sort-le"></span></th>
            <th onclick="sortTable('docType')">Doc Type <span id="sort-docType"></span></th>
            <th onclick="sortTable('q4')">Q4 2025 <span id="sort-q4"></span></th>
            <th>Q4 CanoeID</th>
            <th onclick="sortTable('q4date')">Q4 Received (CT) <span id="sort-q4date"></span></th>
            <th onclick="sortTable('q1')">Q1 2026 <span id="sort-q1"></span></th>
            <th>Q1 CanoeID</th>
            <th onclick="sortTable('q1date')">Q1 Received (CT) <span id="sort-q1date"></span></th>
            <th>Draft Email</th>
            <th class="fm-cell" onclick="sortTable('fm')">FM Contacted <span id="sort-fm"></span></th>
          </tr>
          <tr class="col-search">
            <th></th>
            <th><input class="col-search-input" id="cs-deal" placeholder="Filter…" oninput="applyFilters()"></th>
            <th><input class="col-search-input" id="cs-le"   placeholder="Filter…" oninput="applyFilters()"></th>
            <th><input class="col-search-input" id="cs-doc"  placeholder="Filter…" oninput="applyFilters()"></th>
            <th></th>
            <th></th>
            <th></th>
            <th>
              <select class="col-search-input" id="cs-q1" onchange="applyFilters()" style="appearance:auto">
                <option value="">All</option>
                <option value="Received">Received</option>
                <option value="Not Received">Not Received</option>
              </select>
            </th>
            <th></th>
            <th></th>
            <th></th>
            <th>
              <select class="col-search-input" id="cs-fm" onchange="applyFilters()" style="appearance:auto">
                <option value="">All</option>
                <option value="yes">Contacted</option>
                <option value="no">Not Contacted</option>
              </select>
            </th>
          </tr>
        </thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
    <div class="pagination">
      <div class="page-info" id="pageInfo"></div>
      <div class="page-btns" id="pageBtns"></div>
    </div>
  </div>

</div>

<script>
const RAW = ${dataJson};
const LP_ADMIN_IDS = new Set(${lpAdminIdsJson});
const CONTACTS = ${contactsMapJson};

function draftEmail(deal, le, docType){
  const key = deal + '|' + le;
  const src  = CONTACTS[key] || {};
  const toAddr = src.perei || src.cic || src.nps || '';
  const cc = 'skamaraj@gcmlp.com;nkumar@gcmlp.com;psunkey@gcmlp.com;pgajarla@gcmlp.com;dfrancis@gcmlp.com;mvadapalli@gcmlp.com;rgattu@gcmlp.com;dmateam@gcmlp.com;abgeorge@gcmlp.com';
  const subject = 'Request for Q1 2026 Valuations';
  const body =
    'Dear Fund Manager,' +
    '\\n\\nAs a result of public company reporting obligations, GCM Grosvenor (\\'GCM\\') is required to gather valuations from our fund managers on an accelerated timeline. In order to assist us in meeting our reporting obligations, we are reaching out to you today to kindly ask you provide us with a Q1 2026 Partnership Capital Account Statement(s) (\\u2018PCAP\\u2019) for the following investment(s) no later than [PLEASE INSERT DATE]:' +
    '\\n\\nDeal: ' + deal +
    '\\nProduct: ' + le +
    '\\nDoc Type: ' + docType +
    '\\n\\nIf the Q1 2026 PCAPs are not yet available, we ask that you please provide us with a DRAFT or Estimated Q1 2026 capital account statement(s).' +
    '\\n\\nWe understand that you may not normally be required to provide year-end PCAPs within this time frame and appreciate your cooperation in helping us meet our reporting obligations.' +
    '\\n\\nPlease let us know if you have any questions.' +
    '\\n\\nRegards,' +
    '\\nGCM Grosvenor';
  const mailto = 'mailto:' + encodeURIComponent(toAddr)
    + '?cc=' + encodeURIComponent(cc)
    + '&subject=' + encodeURIComponent(subject)
    + '&body=' + encodeURIComponent(body);
  window.location.href = mailto;
}

// FM Contacted state — keyed by unique row id (deal|leId|docType), persisted to localStorage
const FM_STORAGE_KEY = 'pcap_fm_v1';
const FM = new Map(Object.entries(JSON.parse(localStorage.getItem(FM_STORAGE_KEY) || '{}')));
function fmKey(r) { return r.deal + '|' + r.leId + '|' + r.docType; }
function isFM(r)  { return FM.get(fmKey(r)) === true; }
function saveFM()  { localStorage.setItem(FM_STORAGE_KEY, JSON.stringify(Object.fromEntries(FM))); }

// Charts
const acctReceived = RAW.filter(r=>r.docType==='Account Statement'&&r.q1==='Received').length;
const acctMissing  = RAW.filter(r=>r.docType==='Account Statement'&&r.q1==='Not Received').length;
const finReceived  = RAW.filter(r=>r.docType==='Financials'&&r.q1==='Received').length;
const finMissing   = RAW.filter(r=>r.docType==='Financials'&&r.q1==='Not Received').length;

new Chart(document.getElementById('statusChart').getContext('2d'),{type:'bar',data:{labels:['Account Statement','Financials'],datasets:[
  {label:'Received',    data:[acctReceived,finReceived],backgroundColor:'rgba(74,222,128,.7)', borderColor:'#4ade80',borderWidth:1,borderRadius:6},
  {label:'Not Received',data:[acctMissing, finMissing], backgroundColor:'rgba(248,113,113,.7)',borderColor:'#f87171',borderWidth:1,borderRadius:6}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#8b97b0',font:{family:'Inter',size:12}}}},scales:{x:{stacked:true,grid:{color:'rgba(42,48,64,.6)'},ticks:{color:'#8b97b0'}},y:{stacked:true,grid:{color:'rgba(42,48,64,.6)'},ticks:{color:'#8b97b0'}}}}});

new Chart(document.getElementById('compareChart').getContext('2d'),{type:'bar',data:{labels:['Account Statement','Financials'],datasets:[
  {label:'Q4 2025 (Baseline)',data:[acctReceived+acctMissing,finReceived+finMissing],backgroundColor:'rgba(250,191,36,.7)',borderColor:'#fbbf24',borderWidth:1,borderRadius:6},
  {label:'Q1 2026 Received', data:[acctReceived,finReceived],                        backgroundColor:'rgba(34,211,238,.7)', borderColor:'#22d3ee',borderWidth:1,borderRadius:6}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#8b97b0',font:{family:'Inter',size:12}}}},scales:{x:{grid:{color:'rgba(42,48,64,.6)'},ticks:{color:'#8b97b0'}},y:{grid:{color:'rgba(42,48,64,.6)'},ticks:{color:'#8b97b0'},beginAtZero:true}}}});

// Segment counts
document.getElementById('cnt-all').textContent = RAW.length.toLocaleString();
document.getElementById('cnt-lp').textContent  = RAW.filter(r=>LP_ADMIN_IDS.has(r.leId)).length.toLocaleString();
document.getElementById('cnt-gp').textContent  = RAW.filter(r=>!LP_ADMIN_IDS.has(r.leId)).length.toLocaleString();

// State
let filtered=[...RAW], page=1, pageSize=50, sortKey='deal', sortDir=1, adminMode='all';
const activeFilters={docType:'all',status:'all',fm:'all'};

function setAdmin(mode,btn){
  adminMode=mode;
  document.querySelectorAll('#adminSeg .admin-seg-btn').forEach(b=>b.classList.remove('seg-active'));
  btn.classList.add('seg-active'); page=1; applyFilters();
}

function setFilter(type,val,btn){
  if(type==='status') activeFilters.status=val;
  else if(type==='fm') activeFilters.fm=val;
  else activeFilters[type]=val;
  const gMap={docType:'docTypeGroup',status:'statusGroup',fm:'fmGroup'};
  document.querySelectorAll('#'+gMap[type]+' .btn-filter').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); page=1; applyFilters();
}

function clearColFilters(){
  ['cs-deal','cs-le','cs-doc'].forEach(id=>{ document.getElementById(id).value=''; });
  ['cs-q1','cs-fm'].forEach(id=>{ document.getElementById(id).value=''; });
  ['q4DateFrom','q4DateTo','q1DateFrom','q1DateTo'].forEach(id=>{ document.getElementById(id).value=''; });
  document.getElementById('clearColBtn').style.display='none';
  applyFilters();
}

function applyFilters(){
  const csDeal    = document.getElementById('cs-deal').value.trim().toLowerCase();
  const csLE      = document.getElementById('cs-le').value.trim().toLowerCase();
  const csDoc     = document.getElementById('cs-doc').value.trim().toLowerCase();
  const csQ1      = document.getElementById('cs-q1').value;
  const csFm      = document.getElementById('cs-fm').value;
  const q4From    = document.getElementById('q4DateFrom').value;
  const q4To      = document.getElementById('q4DateTo').value;
  const q1From    = document.getElementById('q1DateFrom').value;
  const q1To      = document.getElementById('q1DateTo').value;
  const anyColFilter = csDeal || csLE || csDoc || csQ1 || csFm || q4From || q4To || q1From || q1To;
  document.getElementById('clearColBtn').style.display = anyColFilter ? '' : 'none';
  const dt=activeFilters.docType, st=activeFilters.status, fm=activeFilters.fm;

  filtered=RAW.filter(r=>{
    if(adminMode==='lp' && !LP_ADMIN_IDS.has(r.leId)) return false;
    if(adminMode==='gp' &&  LP_ADMIN_IDS.has(r.leId)) return false;
    if(csDeal && !r.deal.toLowerCase().includes(csDeal)) return false;
    if(csLE   && !r.le.toLowerCase().includes(csLE))     return false;
    if(csDoc  && !r.docType.toLowerCase().includes(csDoc)) return false;
    if(csQ1   && r.q1!==csQ1) return false;
    if(q4From && (!r.q4date || r.q4date < q4From)) return false;
    if(q4To   && (!r.q4date || r.q4date > q4To))   return false;
    if(q1From && (!r.q1date || r.q1date < q1From)) return false;
    if(q1To   && (!r.q1date || r.q1date > q1To))   return false;
    if(dt!=='all' && r.docType!==dt) return false;
    if(st!=='all' && r.q1!==st) return false;
    const fmVal = isFM(r)?'yes':'no';
    if(fm!=='all' && fmVal!==fm) return false;
    if(csFm!=='' && fmVal!==csFm) return false;
    return true;
  });

  filtered.sort((a,b)=>{
    let av,bv;
    if(sortKey==='fm'){av=isFM(a)?'1':'0';bv=isFM(b)?'1':'0';}
    else if(sortKey==='q4'){av='Received';bv='Received';}
    else{av=a[sortKey]||'';bv=b[sortKey]||'';}
    return av<bv?-sortDir:av>bv?sortDir:0;
  });

  page=1; render();
}

function sortTable(key){
  if(sortKey===key) sortDir*=-1; else{sortKey=key;sortDir=1;}
  ['deal','le','docType','q4','q4date','q1','q1date','fm'].forEach(k=>{ // q4date/q1date span id still used
    const el=document.getElementById('sort-'+k);
    if(el) el.textContent = sortKey===k ? (sortDir===1?' \\u2191':' \\u2193') : '';
  });
  applyFilters();
}

function changePageSize(val){pageSize=parseInt(val);page=1;render();}

function markFM(btn){
  const key = btn.dataset.key;
  FM.set(key, true);
  saveFM();
  // Update FM badge in same row
  const row = btn.closest('tr');
  const badge = row.querySelector('.fm-badge');
  if(badge) badge.classList.add('contacted');
  // Update row tick
  const tick = row.querySelector('.row-tick');
  if(tick) tick.classList.add('visible');
}

function render(){
  const tbody=document.getElementById('tableBody');
  const total=filtered.length;
  const totalPages=Math.max(1,Math.ceil(total/pageSize));
  if(page>totalPages) page=totalPages;
  const start=(page-1)*pageSize, end=Math.min(start+pageSize,total);
  const rows=filtered.slice(start,end);

  document.getElementById('resultsCount').textContent=total.toLocaleString()+' records';

  if(!rows.length){
    tbody.innerHTML='<tr class="empty-row"><td colspan="12">No records match the current filters</td></tr>';
  } else {
    tbody.innerHTML=rows.map(r=>{
      const isLp=LP_ADMIN_IDS.has(r.leId);
      const rowCls=isLp?'row-lp':'row-gp';
      const leCls =isLp?'le-lp':'le-gp';
      const q1Cls =r.q1==='Received'?'pill-received':'pill-missing';
      const key=fmKey(r);
      const checked=FM.get(key)===true;
      return '<tr class="'+rowCls+'">'
        +'<td class="tick-cell"><span class="row-tick'+(checked?' visible':'')+'">&#10003;</span></td>'
        +'<td class="deal-name">'+esc(r.deal)+'</td>'
        +'<td class="'+leCls+'">'+esc(r.le)+'</td>'
        +'<td><span class="doc-badge">'+docLabel(r.docType)+'</span></td>'
        +'<td><span class="status-pill pill-received">&#9679; Received</span></td>'
        +'<td>'+(r.q4url?'<a href="'+esc(r.q4url)+'" target="_blank" class="canoe-link">'+esc(r.q4id?r.q4id.substring(0,8)+'…':'')+'</a>':'<span class="na-text">N/A</span>')+'</td>'
        +'<td>'+(r.q4date?'<span class="dt-text">'+esc(r.q4date)+' '+esc((r.q4time||'').substring(0,5))+'</span>':'<span class="dt-na">—</span>')+'</td>'
        +'<td><span class="status-pill '+q1Cls+'">&#9679; '+esc(r.q1)+'</span></td>'
        +'<td>'+(r.q1url?'<a href="'+esc(r.q1url)+'" target="_blank" class="canoe-link">'+esc(r.q1id?r.q1id.substring(0,8)+'…':'')+'</a>':'<span class="na-text">N/A</span>')+'</td>'
        +'<td>'+(r.q1date?'<span class="dt-text">'+esc(r.q1date)+' '+esc(r.q1time||'')+'</span>':'<span class="dt-na">—</span>')+'</td>'
        +'<td class="draft-cell">'
          +(r.q1==='Not Received'?'<button class="draft-btn" onclick="markFM(this);draftEmail(this.dataset.deal,this.dataset.le,this.dataset.doc)" data-deal="'+esc(r.deal)+'" data-le="'+esc(r.le)+'" data-doc="'+esc(r.docType)+'" data-key="'+esc(key)+'">&#9993; Draft Email</button>':'')
        +'</td>'
        +'<td class="fm-cell"><span class="fm-badge'+(checked?' contacted':'')+'">&#10003;</span></td>'
        +'</tr>';
    }).join('');
  }

  document.getElementById('pageInfo').textContent=
    total===0?'No results':'Showing '+(start+1).toLocaleString()+'\\u2013'+end.toLocaleString()+' of '+total.toLocaleString();

  const btns=[];
  btns.push('<button class="page-btn" onclick="goPage('+(page-1)+')" '+(page<=1?'disabled':'')+'>&#8249; Prev</button>');
  getPageRange(page,totalPages).forEach(p=>{
    if(p==='...') btns.push('<span style="color:var(--muted);padding:0 4px">&#8230;</span>');
    else btns.push('<button class="page-btn '+(p===page?'current':'')+'" onclick="goPage('+p+')" >'+p+'</button>');
  });
  btns.push('<button class="page-btn" onclick="goPage('+(page+1)+')" '+(page>=totalPages?'disabled':'')+'>Next &#8250;</button>');
  document.getElementById('pageBtns').innerHTML=btns.join('');
}

function goPage(p){const tp=Math.max(1,Math.ceil(filtered.length/pageSize));if(p<1||p>tp)return;page=p;render();}
function getPageRange(c,t){
  if(t<=7)return Array.from({length:t},(_,i)=>i+1);
  if(c<=4)return[1,2,3,4,5,'...',t];
  if(c>=t-3)return[1,'...',t-4,t-3,t-2,t-1,t];
  return[1,'...',c-1,c,c+1,'...',t];
}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function docLabel(dt){ return dt==='Financials'?'Financial Statement':esc(dt); }

// ── Export to Excel ──
async function exportExcel(){
  const btn=document.querySelector('.export-btn');
  btn.textContent='Generating…'; btn.disabled=true;

  try{
    const wb = new ExcelJS.Workbook();
    wb.creator='PCAP Dashboard';
    wb.created=new Date();

    const ws = wb.addWorksheet('PCAP Receipt Details',{
      views:[{state:'frozen',ySplit:1}]
    });

    // Column definitions
    ws.columns=[
      {header:'Deal',          key:'deal',    width:42},
      {header:'Legal Entity',  key:'le',      width:42},
      {header:'Admin Type',    key:'admin',   width:14},
      {header:'Doc Type',      key:'docType', width:22},
      {header:'Q4 2025',       key:'q4',      width:16},
      {header:'Q4 CanoeID',       key:'q4url',    width:20},
      {header:'Q4 Received (CT)', key:'q4recvct', width:20},
      {header:'Q1 2026',       key:'q1',      width:16},
      {header:'Q1 CanoeID',       key:'q1url',    width:20},
      {header:'Q1 Received (CT)', key:'q1recvct', width:20},
      {header:'FM Contacted',  key:'fm',      width:16},
    ];

    // Header row styling
    const hdrRow=ws.getRow(1);
    hdrRow.height=28;
    hdrRow.eachCell(cell=>{
      cell.font={bold:true,color:{argb:'FFFFFFFF'},size:11,name:'Calibri'};
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1B2A4A'}};
      cell.alignment={vertical:'middle',horizontal:'center',wrapText:false};
      cell.border={
        top:{style:'thin',color:{argb:'FF2A3A5A'}},
        bottom:{style:'medium',color:{argb:'FF3B7DD8'}},
        left:{style:'thin',color:{argb:'FF2A3A5A'}},
        right:{style:'thin',color:{argb:'FF2A3A5A'}}
      };
    });

    // Export the full filtered dataset
    filtered.forEach((r,i)=>{
      const isLp=LP_ADMIN_IDS.has(r.leId);
      const fmVal=FM.get(fmKey(r))===true;
      const rowIdx=i+2;

      const row=ws.addRow({
        deal:   r.deal,
        le:     r.le,
        admin:  isLp?'LP Admin':'GP Admin',
        docType:r.docType==='Account Statement'?'PCAP':r.docType==='Financials'?'FS':r.docType,
        q4:     'Received',
        q4url:  r.q4url || '',
        q4recvct: r.q4date ? r.q4date + ' ' + (r.q4time||'') : '',
        q1:     r.q1,
        q1url:  r.q1url || 'N/A',
        q1recvct: r.q1date ? r.q1date + ' ' + (r.q1time||'') : 'N/A',
        fm:     fmVal?'Yes':'No',
      });
      row.height=20;

      // Base row fill — alternating
      const baseFill = i%2===0
        ? (isLp?'FFFFF4EC':'FFF0FAFA')   // LP=light orange, GP=light teal
        : (isLp?'FFFDECD8':'FFE6F7F5');  // slightly deeper alternating shade

      row.eachCell(cell=>{
        cell.font={size:10,name:'Calibri'};
        cell.alignment={vertical:'middle',wrapText:false};
        cell.border={
          top:{style:'thin',color:{argb:'FFE0E7EF'}},
          bottom:{style:'thin',color:{argb:'FFE0E7EF'}},
          left:{style:'thin',color:{argb:'FFE0E7EF'}},
          right:{style:'thin',color:{argb:'FFE0E7EF'}}
        };
        cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:baseFill}};
      });

      // Deal column — bold
      row.getCell('deal').font={size:10,name:'Calibri',bold:true};

      // Admin badge colour
      const adminCell=row.getCell('admin');
      if(isLp){
        adminCell.font={size:10,name:'Calibri',bold:true,color:{argb:'FFC05000'}};
        adminCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFDE8D0'}};
      } else {
        adminCell.font={size:10,name:'Calibri',bold:true,color:{argb:'FF0D7D6C'}};
        adminCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD0F5F0'}};
      }
      adminCell.alignment={vertical:'middle',horizontal:'center'};

      // Doc Type — purple tint
      const dtCell=row.getCell('docType');
      dtCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF3EFFE'}};
      dtCell.font={size:10,name:'Calibri',color:{argb:'FF6B48C8'}};

      // Q4 — always green
      const q4Cell=row.getCell('q4');
      q4Cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8F5E9'}};
      q4Cell.font={size:10,name:'Calibri',bold:true,color:{argb:'FF2E7D32'}};
      q4Cell.alignment={vertical:'middle',horizontal:'center'};

      // Q4 CanoeID — blue hyperlink
      const q4urlCell=row.getCell('q4url');
      if(r.q4url){
        q4urlCell.value={text:'Open in Canoe',hyperlink:r.q4url};
        q4urlCell.font={size:10,name:'Calibri',color:{argb:'FF1565C0'},underline:true};
        q4urlCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8F0FE'}};
      }
      q4urlCell.alignment={vertical:'middle',horizontal:'center'};

      // Q1 — green or red
      const q1Cell=row.getCell('q1');
      const isReceived=r.q1==='Received';
      q1Cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:isReceived?'FFE8F5E9':'FFFCE4E4'}};
      q1Cell.font={size:10,name:'Calibri',bold:true,color:{argb:isReceived?'FF2E7D32':'FFC62828'}};
      q1Cell.alignment={vertical:'middle',horizontal:'center'};

      // Q1 CanoeID — blue hyperlink or N/A
      const q1urlCell=row.getCell('q1url');
      if(r.q1url){
        q1urlCell.value={text:'Open in Canoe',hyperlink:r.q1url};
        q1urlCell.font={size:10,name:'Calibri',color:{argb:'FF1565C0'},underline:true};
        q1urlCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE8F0FE'}};
      } else {
        q1urlCell.font={size:10,name:'Calibri',color:{argb:'FF9AA3B2'},italic:true};
      }
      q1urlCell.alignment={vertical:'middle',horizontal:'center'};

      // FM Contacted
      const fmCell=row.getCell('fm');
      if(fmVal){
        fmCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE3F2FD'}};
        fmCell.font={size:10,name:'Calibri',bold:true,color:{argb:'FF1565C0'}};
      }
      fmCell.alignment={vertical:'middle',horizontal:'center'};
    });

    // Auto-filter on header row
    ws.autoFilter={from:'A1',to:'K1'};

    // Generate and download
    const buf=await wb.xlsx.writeBuffer();
    const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    const ts=new Date().toISOString().slice(0,10);
    a.download='PCAP_Receipt_'+ts+'.xlsx';
    a.click(); URL.revokeObjectURL(url);
  } finally {
    btn.innerHTML='<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Export to Excel';
    btn.disabled=false;
  }
}

applyFilters();
</script>
</body>
</html>`;

fs.writeFileSync('C:/Users/ABGEORGE/pcap-receipt-dashboard.html', html, 'utf8');
console.log('Done! Size:', Math.round(html.length/1024)+'KB');
