/**
 * CR Cost Saving & Material Analysis Dashboard
 * app.js – Frontend-only, data from Excel upload (SheetJS + Highcharts)
 * ─────────────────────────────────────────────────────────────────────
 */

'use strict';

// ═══════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════
let RAW_DATA   = [];   // full parsed dataset
let FILTERED   = [];   // after filters applied
let chartInstances = {}; // { id: HighchartsChart }

// ═══════════════════════════════════════════════════════
// HELPERS – FORMATTING
// ═══════════════════════════════════════════════════════

/** Format angka sebagai currency Indonesia singkat: Rp 1,25 M / Rp 500 JT */
function fmtCurrency(val) {
  if (val == null || isNaN(val)) return '-';
  const abs = Math.abs(val);
  let str;
  if (abs >= 1_000_000_000) {
    str = 'Rp ' + (val / 1_000_000_000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' B';
  } else if (abs >= 1_000_000) {
    str = 'Rp ' + (val / 1_000_000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' JT';
  } else if (abs >= 1_000) {
    str = 'Rp ' + (val / 1_000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' RB';
  } else {
    str = 'Rp ' + val.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }
  return str;
}

/** Format number dengan thousand separator */
function fmtNumber(val) {
  if (val == null || isNaN(val)) return '-';
  return Number(val).toLocaleString('id-ID');
}

/** Format persen dengan 2 desimal */
function fmtPct(val) {
  if (val == null || isNaN(val)) return '-';
  return val.toFixed(2) + '%';
}

// ═══════════════════════════════════════════════════════
// HELPERS – PARSING
// ═══════════════════════════════════════════════════════

/** Parse angka dari string Indonesia (koma desimal) atau number biasa */
function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  // Handle format Indonesia: "111,72" → 111.72
  const s = String(v).trim().replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/** Parse tanggal Excel (serial number atau string) */
function parseDate(v) {
  if (!v) return null;
  if (typeof v === 'number') {
    // Excel date serial
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d;
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

/** Hitung cost value untuk satu row */
function calcCost(row) {
  return (row.qty_g / 1000) * row.harga * row.qty_prod;
}

// ═══════════════════════════════════════════════════════
// EXCEL PARSING
// ═══════════════════════════════════════════════════════

function parseExcel(buffer) {
  const wb   = XLSX.read(buffer, { type: 'array', cellDates: false });
  const wsName = wb.SheetNames[0];
  const ws   = wb.Sheets[wsName];
  const json = XLSX.utils.sheet_to_json(ws, { raw: true, defval: '' });

  return json.map((row, idx) => {
    const r = {};
    // Normalise key names (lowercase, trim)
    for (const k in row) r[k.trim().toLowerCase().replace(/\s+/g,'_')] = row[k];

    return {
      _idx       : idx,
      no_sap     : String(r.no_sap   || '').trim(),
      part_name  : String(r.part_name || r.part || '').trim(),
      component  : String(r.component || '').trim(),
      material   : String(r.material  || '').trim().toUpperCase(),
      qty_g      : parseNum(r.qty_g),
      scenario   : String(r.scenario  || '').trim(),
      harga      : parseNum(r.harga),
      qty_prod   : parseNum(r.qty_prod),
      ok_prod    : parseNum(r.ok_prod),
      ng_prod    : parseNum(r.ng_prod),
      tanggal    : parseDate(r.tanggal),
      kategori   : String(r.kategori  || '').trim(),
      // computed
      cost_value : 0, // will be set below
    };
  }).map(r => {
    r.cost_value = calcCost(r);
    return r;
  });
}

// ═══════════════════════════════════════════════════════
// UPLOAD HANDLERS
// ═══════════════════════════════════════════════════════

function handleDragOver(e) {
  e.preventDefault();
  document.getElementById('upload-zone').classList.add('dragover');
}
function handleDragLeave(e) {
  document.getElementById('upload-zone').classList.remove('dragover');
}
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('upload-zone').classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) readFile(file);
}
function handleFile(e) {
  const file = e.target.files[0];
  if (file) readFile(file);
}

function readFile(file) {
  // Show loading skeleton on cards briefly
  showSkeleton();

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      RAW_DATA = parseExcel(new Uint8Array(e.target.result));
      FILTERED = [...RAW_DATA];

      // Update upload pill UI
      const pill = document.getElementById('upload-pill');
      pill.classList.add('success');
      document.getElementById('upload-icon').className = 'fa-solid fa-circle-check';
      // Truncate filename if too long
      const fname = file.name.length > 22 ? file.name.slice(0, 20) + '…' : file.name;
      document.getElementById('pill-main').textContent = fname;
      document.getElementById('pill-sub').textContent  = RAW_DATA.length + ' baris dimuat';

      // Populate filter dropdowns
      populateFilters();

      // Render everything
      hideSkeleton();
      renderAll();
    } catch (err) {
      console.error('Parse error:', err);
      alert('Gagal membaca file. Pastikan format kolom sesuai.');
      hideSkeleton();
    }
  };
  reader.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════════════════
// FILTER SYSTEM
// ═══════════════════════════════════════════════════════

function populateFilters() {
  // ── Tahun (derived from tanggal column)
  const years = [...new Set(
    RAW_DATA.map(r => r.tanggal ? r.tanggal.getFullYear() : null).filter(Boolean)
  )].sort((a, b) => a - b);
  const selTahun = document.getElementById('filter-tahun');
  selTahun.innerHTML = '<option value="">Semua Tahun</option>';
  years.forEach(y => {
    const o = document.createElement('option');
    o.value = y; o.textContent = y; selTahun.appendChild(o);
  });

  // ── Material
  const mats = [...new Set(RAW_DATA.map(r => r.material).filter(Boolean))].sort();
  const selMat = document.getElementById('filter-material');
  selMat.innerHTML = '<option value="">Semua</option>';
  mats.forEach(m => { const o = document.createElement('option'); o.value = m; o.textContent = m; selMat.appendChild(o); });

  // ── Part
  const parts = [...new Set(RAW_DATA.map(r => r.part_name).filter(Boolean))].sort();
  const selPart = document.getElementById('filter-part');
  selPart.innerHTML = '<option value="">Semua Part</option>';
  parts.forEach(p => { const o = document.createElement('option'); o.value = p; o.textContent = p; selPart.appendChild(o); });
}

function applyFilters() {
  const mat      = document.getElementById('filter-material').value;
  const part     = document.getElementById('filter-part').value;
  const scenario = document.getElementById('filter-scenario').value;
  const bulan    = document.getElementById('filter-bulan').value;   // "1"–"12" or ""
  const tahun    = document.getElementById('filter-tahun').value;   // "2024" or ""

  FILTERED = RAW_DATA.filter(r => {
    if (mat      && r.material  !== mat)      return false;
    if (part     && r.part_name !== part)     return false;
    if (scenario && r.scenario  !== scenario) return false;
    if (bulan && r.tanggal) {
      if ((r.tanggal.getMonth() + 1) !== parseInt(bulan)) return false;
    }
    if (tahun && r.tanggal) {
      if (r.tanggal.getFullYear() !== parseInt(tahun)) return false;
    }
    return true;
  });

  renderAll();
}

function resetFilters() {
  document.getElementById('filter-material').value = '';
  document.getElementById('filter-part').value     = '';
  document.getElementById('filter-scenario').value = '';
  document.getElementById('filter-bulan').value    = '';
  document.getElementById('filter-tahun').value    = '';
  FILTERED = [...RAW_DATA];
  if (RAW_DATA.length) renderAll();
}

// ── Attach filter events
document.addEventListener('DOMContentLoaded', () => {
  ['filter-material','filter-part','filter-scenario','filter-bulan','filter-tahun'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });

  // Init empty charts
  initEmptyCharts();
});

// ═══════════════════════════════════════════════════════
// NAVIGATION
// ═══════════════════════════════════════════════════════

function navigate(page) {
  document.querySelectorAll('.nav-item')
    .forEach(el => el.classList.remove('active'));

  document.querySelector(`[data-page="${page}"]`)
    .classList.add('active');

  const target = page === 'overview' ? 'row-overview'
               : page === 'saving'   ? 'row-overview'
               : page === 'ng'       ? 'page-ng'
               : 'page-material';

  const element = document.getElementById(target);

  if (element) {
    const headerHeight = 200; // tinggi header

    const elementPosition =
      element.getBoundingClientRect().top + window.pageYOffset;

    const offsetPosition = elementPosition - headerHeight;

    window.scrollTo({
      top: offsetPosition,
      behavior: 'smooth'
    });
  }

  closeSidebar();
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('overlay').classList.toggle('show');
}
function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('overlay').classList.remove('show');
}

// ═══════════════════════════════════════════════════════
// SKELETON
// ═══════════════════════════════════════════════════════

function showSkeleton() {
  ['card-standard','card-alt','card-saving','card-pct','card-ng'].forEach(id => {
    const el = document.getElementById(id);
    el.classList.add('skeleton');
    el.textContent = '';
    el.style.height = '32px';
  });
}
function hideSkeleton() {
  ['card-standard','card-alt','card-saving','card-pct','card-ng'].forEach(id => {
    const el = document.getElementById(id);
    el.classList.remove('skeleton');
    el.style.height = '';
  });
}

// ═══════════════════════════════════════════════════════
// RENDER ALL
// ═══════════════════════════════════════════════════════

function renderAll() {
  const data = FILTERED;

  // ── Calculations
  const stdRows  = data.filter(r => r.scenario === 'Standard');
  const altRows  = data.filter(r => r.scenario === 'Alternative');

  const totalStd = stdRows.reduce((s, r) => s + r.cost_value, 0);
  const totalAlt = altRows.reduce((s, r) => s + r.cost_value, 0);
  const saving   = totalStd - totalAlt;
  const pctSave  = totalStd ? (saving / totalStd) * 100 : 0;
  const totalNG  = data.reduce((s, r) => s + r.ng_prod, 0);

  // ── Cards with animated counter
  animateCounter('card-standard', totalStd, fmtCurrency);
  animateCounter('card-alt',      totalAlt,  fmtCurrency);
  animateCounter('card-saving',   saving,    fmtCurrency);
  document.getElementById('card-pct').textContent = fmtPct(pctSave);
  animateCounter('card-ng',       totalNG,   fmtNumber);

  // ── Charts
  renderComparisonChart(totalStd, totalAlt, saving);
  renderTop10SavingChart(data);
  renderMaterialCards(data);
  renderTop10NGChart(data);
  renderScatterChart(data);
  renderNGTable(data);
  renderDonutChart(data);
  renderByMaterialChart(data);
}

// ═══════════════════════════════════════════════════════
// ANIMATED COUNTER
// ═══════════════════════════════════════════════════════

function animateCounter(id, target, formatter, duration = 800) {
  const el = document.getElementById(id);
  if (!el) return;
  const start = performance.now();
  function step(now) {
    const p = Math.min((now - start) / duration, 1);
    const val = target * easeOut(p);
    el.textContent = formatter(val);
    if (p < 1) requestAnimationFrame(step);
    else el.textContent = formatter(target);
  }
  requestAnimationFrame(step);
}
function easeOut(t) { return 1 - Math.pow(1 - t, 3); }

// ═══════════════════════════════════════════════════════
// HIGHCHART DEFAULTS
// ═══════════════════════════════════════════════════════

const HC_COLORS = ['#2563EB','#16A34A','#7C3AED','#F97316','#DC2626','#0891B2','#D97706','#4F46E5','#059669','#E11D48'];

Highcharts.setOptions({
  chart: { backgroundColor: 'transparent', style: { fontFamily: 'DM Sans, sans-serif' } },
  credits: { enabled: false },
  exporting: { enabled: false },
  colors: HC_COLORS,
  tooltip: {
    backgroundColor: '#1e293b',
    borderColor: 'transparent',
    borderRadius: 10,
    style: { color: '#f8fafc', fontSize: '12px' },
    shadow: true
  },
  plotOptions: {
    series: { animation: { duration: 700 } },
    column: { borderRadius: 6, borderWidth: 0 },
    bar    : { borderRadius: 4, borderWidth: 0 },
  }
});

/** Destroy & re-create a chart. Returns chart instance. */
function renderChart(id, opts) {
  if (chartInstances[id]) {
    try { chartInstances[id].destroy(); } catch(e) {}
  }
  chartInstances[id] = Highcharts.chart(id, opts);
  return chartInstances[id];
}

// ═══════════════════════════════════════════════════════
// EMPTY CHART INIT
// ═══════════════════════════════════════════════════════

const CHART_IDS = [
  'chart-comparison','chart-top10-saving','chart-top10-ng',
  'chart-scatter','chart-donut','chart-by-material'
];

function initEmptyCharts() {
  const tpl = document.getElementById('tpl-empty').innerHTML;
  CHART_IDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = tpl;
  });
  document.getElementById('ng-table-body').innerHTML =
    '<tr><td colspan="5" class="text-center text-slate-400 py-8 text-sm">Belum ada data – upload file Excel terlebih dahulu</td></tr>';
}

// ═══════════════════════════════════════════════════════
// CHART 1 – COST COMPARISON COLUMN
// ═══════════════════════════════════════════════════════

function renderComparisonChart(std, alt, saving) {
  if (std === 0 && alt === 0) {
    const el = document.getElementById('chart-comparison');
    if (el) el.innerHTML = document.getElementById('tpl-empty').innerHTML;
    return;
  }
  renderChart('chart-comparison', {
    chart : { type: 'column', height: 330 },
    title : { text: '' },
    xAxis : {
      categories: ['Cost Standard', 'Cost Alternative', 'Total Saving'],
      labels: { style: { fontSize: '11px' } }
    },
    yAxis : {
      title: { text: '' },
      gridLineColor: '#f1f5f9',
      labels: { formatter() { return fmtCurrency(this.value); }, style: { fontSize: '10px' } }
    },
    legend: { enabled: false },
    series: [{
      name: 'Nilai',
      data: [
        { y: std,    color: '#2563EB' },
        { y: alt,    color: '#16A34A' },
        { y: saving, color: saving >= 0 ? '#7C3AED' : '#DC2626' }
      ],
      dataLabels: {
        enabled: true,
        formatter() { return fmtCurrency(this.y); },
        style: { fontSize: '10px', fontWeight: '600', textOutline: 'none' }
      }
    }]
  });
}

// ═══════════════════════════════════════════════════════
// CHART 2 – TOP 10 SAVING PART (horizontal bar)
// ═══════════════════════════════════════════════════════

function renderTop10SavingChart(data) {
  // Group by part_name, get std & alt cost per part
  const partMap = {};
  data.forEach(r => {
    if (!partMap[r.part_name]) partMap[r.part_name] = { std: 0, alt: 0 };
    if (r.scenario === 'Standard')    partMap[r.part_name].std += r.cost_value;
    if (r.scenario === 'Alternative') partMap[r.part_name].alt += r.cost_value;
  });

  const parts = Object.entries(partMap)
    .map(([k, v]) => ({ name: k, saving: v.std - v.alt }))
    .filter(p => p.saving > 0)
    .sort((a, b) => b.saving - a.saving)
    .slice(0, 10);

  if (!parts.length) {
    const el = document.getElementById('chart-top10-saving');
    if (el) el.innerHTML = document.getElementById('tpl-empty').innerHTML;
    return;
  }

  renderChart('chart-top10-saving', {
    chart : { type: 'bar', height: 330 },
    title : { text: '' },
    xAxis : { categories: parts.map(p => truncate(p.name, 20)), labels: { style: { fontSize: '10px' } } },
    yAxis : {
      title: { text: '' }, gridLineColor: '#f1f5f9',
      labels: { formatter() { return fmtCurrency(this.value); }, style: { fontSize: '10px' } }
    },
    legend: { enabled: false },
    series: [{
      name: 'Saving',
      data: parts.map((p, i) => ({ y: p.saving, color: HC_COLORS[i % HC_COLORS.length] })),
      dataLabels: {
        enabled: true, inside: false,
        formatter() { return fmtCurrency(this.y); },
        style: { fontSize: '10px', fontWeight: '600', textOutline: 'none' }
      }
    }]
  });
}

// ═══════════════════════════════════════════════════════
// SECTION 2 – MATERIAL CARDS (PP, PPGF, ABS)
// ═══════════════════════════════════════════════════════

// Dynamic color/icon palette for materials
const MAT_PALETTE = [
  { color: '#2563EB', bg: '#EFF6FF', icon: 'fa-circle-nodes' },
  { color: '#16A34A', bg: '#F0FDF4', icon: 'fa-layer-group'  },
  { color: '#F97316', bg: '#FFF7ED', icon: 'fa-cubes'         },
  { color: '#7C3AED', bg: '#F5F3FF', icon: 'fa-atom'          },
  { color: '#0891B2', bg: '#ECFEFF', icon: 'fa-droplet'       },
  { color: '#DC2626', bg: '#FEF2F2', icon: 'fa-fire'          },
  { color: '#D97706', bg: '#FFFBEB', icon: 'fa-star'          },
  { color: '#4F46E5', bg: '#EEF2FF', icon: 'fa-gem'           },
];

function renderMaterialCards(data) {
  const container = document.getElementById('material-cards');

  if (!data || data.length === 0) {
    container.innerHTML = `<div class="col-span-3 flex flex-col items-center justify-center py-10 text-center text-slate-400">
      <i class="fa-solid fa-cubes text-3xl mb-3 opacity-20"></i>
      <p class="text-sm font-medium">Belum ada data material</p>
      <p class="text-xs mt-1">Upload file Excel untuk menampilkan ringkasan material</p>
    </div>`;
    return;
  }

  // Build part-level map: part → { std, alt, kategori }
  const partMap = {};
  data.forEach(r => {
    const key = r.part_name + '|' + r.kategori;
    if (!partMap[key]) partMap[key] = { part: r.part_name, kategori: r.kategori, std: 0, alt: 0 };
    if (r.scenario === 'Standard')    partMap[key].std += r.cost_value;
    if (r.scenario === 'Alternative') partMap[key].alt += r.cost_value;
  });

  // Group by kategori
  const matGroups = {};
  Object.values(partMap).forEach(p => {
    const mat = p.kategori || '(Lainnya)';
    if (!matGroups[mat]) matGroups[mat] = [];
    matGroups[mat].push({ ...p, saving: p.std - p.alt });
  });

  // Get all unique kategoris from data, sorted
  const allkategoris = Object.keys(matGroups).sort();

  container.innerHTML = '';

  // Adjust grid cols dynamically
  const ncols = Math.min(allkategoris.length, 4);
  container.className = `grid gap-4 grid-cols-1 md:grid-cols-${ncols <= 2 ? ncols : '3'} ${ncols >= 4 ? 'lg:grid-cols-4' : ''}`;

  allkategoris.forEach((matKey, idx) => {
    const cfg = { key: matKey, ...MAT_PALETTE[idx % MAT_PALETTE.length] };
    const parts  = (matGroups[cfg.key] || []).sort((a,b) => b.saving - a.saving);
    const total  = parts.reduce((s, p) => s + p.saving, 0);
    const nParts = parts.length;
    const avg    = nParts ? total / nParts : 0;
    const top5   = parts.slice(0, 5);
    const maxSave= top5[0]?.saving || 1;

    const html = `
      <div class="mat-card fade-up">
        <div style="background:${cfg.bg}; border-bottom:3px solid ${cfg.color}" class="p-4 flex items-center gap-3">
          <div class="w-10 h-10 rounded-xl flex items-center justify-center text-white text-sm" style="background:${cfg.color}">
            <i class="fa-solid ${cfg.icon}"></i>
          </div>
          <div>
            <div class="font-display font-bold text-lg" style="color:${cfg.color}">${cfg.key}</div>
            <div class="text-xs text-slate-500">${nParts} Part</div>
          </div>
          <div class="ml-auto text-right">
            <div class="font-bold text-sm text-slate-700">${fmtCurrency(total)}</div>
            <div class="text-xs text-slate-400">Avg ${fmtCurrency(avg)}</div>
          </div>
        </div>
        <div class="p-4 space-y-3">
          ${top5.length ? top5.map(p => `
            <div>
              <div class="flex justify-between text-xs text-slate-600 mb-1">
                <span class="truncate max-w-[60%]" title="${p.part}">${truncate(p.part,22)}</span>
                <span class="font-semibold" style="color:${cfg.color}">${fmtCurrency(p.saving)}</span>
              </div>
              <div class="mini-bar-track">
                <div class="mini-bar-fill" style="background:${cfg.color};width:${Math.max(2,(p.saving/maxSave*100)).toFixed(1)}%"></div>
              </div>
            </div>`).join('')
          : '<p class="text-xs text-slate-400 text-center py-2">Tidak ada data</p>'}
        </div>
      </div>`;
    container.insertAdjacentHTML('beforeend', html);
  });
}

// ═══════════════════════════════════════════════════════
// SECTION 3 – NG CHARTS
// ═══════════════════════════════════════════════════════

function renderTop10NGChart(data) {
  // Group ng_prod by part
  const partNG = {};
  data.forEach(r => {
    partNG[r.part_name] = (partNG[r.part_name] || 0) + r.ng_prod;
  });

  const top10 = Object.entries(partNG)
    .map(([k, v]) => ({ name: k, ng: v }))
    .sort((a, b) => b.ng - a.ng).slice(0, 10);

  renderChart('chart-top10-ng', {
    chart : { type: 'bar', height: 310 },
    title : { text: '' },
    xAxis : { categories: top10.map(p => truncate(p.name, 20)), labels: { style: { fontSize: '10px' } } },
    yAxis : {
      title: { text: '' }, gridLineColor: '#f1f5f9',
      labels: { style: { fontSize: '10px' } }
    },
    legend: { enabled: false },
    series: [{
      name: 'NG',
      data: top10.map((p, i) => ({ y: p.ng, color: HC_COLORS[i % HC_COLORS.length] })),
      dataLabels: {
        enabled: true, inside: false,
        formatter() { return fmtNumber(this.y); },
        style: { fontSize: '10px', fontWeight: '600', textOutline: 'none' }
      }
    }]
  });
}

function renderScatterChart(data) {
  // Group: part → { totalNG, potSaving }
  const partMap = {};
  data.forEach(r => {
    if (!partMap[r.part_name]) partMap[r.part_name] = { ng: 0, std: 0, alt: 0 };
    partMap[r.part_name].ng += r.ng_prod;
    if (r.scenario === 'Standard')    partMap[r.part_name].std += r.cost_value;
    if (r.scenario === 'Alternative') partMap[r.part_name].alt += r.cost_value;
  });

  const points = Object.entries(partMap).map(([name, v]) => ({
    name, x: v.std - v.alt, y: v.ng
  })).filter(p => p.x !== 0 || p.y !== 0);

  renderChart('chart-scatter', {
    chart : { type: 'scatter', height: 310, zoomType: 'xy' },
    title : { text: '' },
    xAxis : {
      title: { text: 'Potensi Saving', style: { fontSize: '11px' } },
      labels: { formatter() { return fmtCurrency(this.value); }, style: { fontSize: '10px' } },
      gridLineColor: '#f1f5f9', gridLineWidth: 1
    },
    yAxis : {
      title: { text: 'Total NG', style: { fontSize: '11px' } },
      gridLineColor: '#f1f5f9',
      labels: { style: { fontSize: '10px' } }
    },
    legend: { enabled: false },
    tooltip: {
      formatter() {
        return `<b>${this.point.name}</b><br/>NG: ${fmtNumber(this.y)}<br/>Saving: ${fmtCurrency(this.x)}`;
      }
    },
    plotOptions: {
      scatter: {
        marker: { radius: 6, symbol: 'circle' },
        dataLabels: {
          enabled: true,
          formatter() { return truncate(this.point.name, 12); },
          style: { fontSize: '9px', textOutline: 'none', color: '#334155' }
        }
      }
    },
    series: [{
      name: 'Parts',
      data: points.map((p, i) => ({ ...p, color: HC_COLORS[i % HC_COLORS.length] }))
    }]
  });
}

function renderNGTable(data) {
  // Group part
  const partMap = {};
  data.forEach(r => {
    if (!partMap[r.part_name]) partMap[r.part_name] = { part: r.part_name, material: r.material, ng: 0, std: 0, alt: 0 };
    partMap[r.part_name].ng  += r.ng_prod;
    if (r.scenario === 'Standard')    partMap[r.part_name].std += r.cost_value;
    if (r.scenario === 'Alternative') partMap[r.part_name].alt += r.cost_value;
  });

  const rows = Object.values(partMap)
    .map(p => ({ ...p, saving: p.std - p.alt }))
    .sort((a, b) => b.ng - a.ng);

  // Assign status mock (based on NG thresholds)
  function getStatus(ng) {
    if (ng > 500) return { label: 'On Development', cls: 'badge-yellow' };
    if (ng > 100) return { label: 'Pending',        cls: 'badge-red'    };
    return             { label: 'Done',             cls: 'badge-green'  };
  }

  const tbody = document.getElementById('ng-table-body');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-slate-400 py-8 text-sm">Tidak ada data</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => {
    const s = getStatus(r.ng);
    return `<tr>
      <td class="font-medium text-slate-700">${truncate(r.part, 24)}</td>
      <td><span class="badge badge-blue">${r.material}</span></td>
      <td class="font-semibold text-slate-700">${fmtNumber(r.ng)}</td>
      <td class="font-semibold text-purple-700">${fmtCurrency(r.saving)}</td>
      <td><span class="badge ${s.cls}">${s.label}</span></td>
    </tr>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════
// SECTION 4 – MATERIAL USAGE
// ═══════════════════════════════════════════════════════

function renderDonutChart(data) {
  // Count parts per kategori
  const catMap = {};
  const seen = new Set();
  data.forEach(r => {
    const key = r.part_name + '|' + r.kategori;
    if (!seen.has(key)) {
      seen.add(key);
      catMap[r.kategori] = (catMap[r.kategori] || 0) + 1;
    }
  });

  const series = Object.entries(catMap).map(([k, v]) => ({ name: k || '(Tanpa Kategori)', y: v }));

  renderChart('chart-donut', {
    chart : { type: 'pie', height: 290 },
    title : { text: '' },
    tooltip: { pointFormat: '<b>{point.name}</b>: {point.y} part ({point.percentage:.1f}%)' },
    plotOptions: {
      pie: {
        innerSize: '52%',
        dataLabels: {
          enabled: true,
          format: '<b>{point.name}</b>: {point.y}',
          style: { fontSize: '11px', textOutline: 'none' }
        }
      }
    },
    series: [{ name: 'Part', colorByPoint: true, data: series }]
  });
}

function renderByMaterialChart(data) {
  // Count unique parts per material
  const matMap = {};
  const seen = new Set();
  data.forEach(r => {
    const key = r.part_name + '|' + r.material;
    if (!seen.has(key)) {
      seen.add(key);
      matMap[r.material] = (matMap[r.material] || 0) + 1;
    }
  });

  const cats = Object.keys(matMap).sort();
  const vals = cats.map(k => matMap[k]);

  renderChart('chart-by-material', {
    chart : { type: 'column', height: 290 },
    title : { text: '' },
    xAxis : { categories: cats, labels: { style: { fontSize: '11px' } } },
    yAxis : {
      title: { text: 'Jumlah Part' }, gridLineColor: '#f1f5f9',
      labels: { style: { fontSize: '10px' } }
    },
    legend: { enabled: false },
    series: [{
      name: 'Jumlah Part',
      data: vals.map((v, i) => ({ y: v, color: HC_COLORS[i % HC_COLORS.length] })),
      dataLabels: {
        enabled: true,
        formatter() { return this.y; },
        style: { fontSize: '11px', fontWeight: '700', textOutline: 'none' }
      }
    }]
  });
}

// ═══════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════

function truncate(str, n) {
  if (!str) return '';
  return str.length > n ? str.slice(0, n) + '…' : str;
}
