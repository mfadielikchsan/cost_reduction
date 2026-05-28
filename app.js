'use strict';

// ===== STATE =====
const APP = {
  rawData: [],
  filteredData: [],
  charts: {},
  ngTableSort: { col: 'ng_loss', dir: 'desc' },
  matTableSort: { col: 'usage_kg', dir: 'desc' },
  ngTablePage: 1,
  matTablePage: 1,
  PAGE_SIZE: 10,
  topChartMode: 'saving',
  comparisonMode: 'nominal',
  categoryViewMode: 'saving',
  debounceTimer: null,
  fileName: '',
};

const BULAN_NAMES = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
const CATEGORY_COLORS = ['#1d4ed8','#0891b2','#7c3aed','#059669','#d97706','#db2777','#dc2626','#84cc16'];

// ===== FORMAT HELPERS =====
function formatIDR(val, showRp = true) {
  if (val == null || isNaN(val)) return '—';
  const abs = Math.abs(val);
  const sign = val < 0 ? '-' : '';
  let str;
  if (abs >= 1_000_000_000) str = (abs / 1_000_000_000).toFixed(2).replace('.', ',') + ' M';
  else if (abs >= 1_000_000) str = (abs / 1_000_000).toFixed(2).replace('.', ',') + ' JT';
  else if (abs >= 1_000) str = (abs / 1_000).toFixed(1).replace('.', ',') + ' RB';
  else str = abs.toFixed(0);
  return (showRp ? 'Rp ' : '') + sign + str;
}

function formatNumber(val) {
  if (val == null || isNaN(val)) return '—';
  return new Intl.NumberFormat('id-ID').format(Math.round(val));
}

function formatPct(val) {
  if (val == null || isNaN(val)) return '—';
  return (val >= 0 ? '' : '') + val.toFixed(2).replace('.', ',') + '%';
}

function formatKg(val) {
  if (val == null || isNaN(val)) return '—';
  return formatNumber(val) + ' kg';
}

// ===== PARSE DATE =====
function parseDate(str) {
  if (!str) return null;
  if (typeof str === 'number') {
    // Excel serial date
    const d = new Date((str - 25569) * 86400 * 1000);
    return d;
  }
  if (typeof str === 'string') {
    const parts = str.split('/');
    if (parts.length === 3) {
      return new Date(+parts[2], +parts[1] - 1, +parts[0]);
    }
  }
  return null;
}

// ===== PARSE EXCEL =====
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: false });
        const sheetName = wb.SheetNames.find(s => s.toLowerCase().replace(/\s/g,'') === 'data_dashboard') || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

// ===== NORMALIZE ROW =====
function normalizeRow(row) {
  const d = parseDate(row.tanggal || row.Tanggal || row.TANGGAL);
  return {
    no_sap: String(row.no_sap || row.NO_SAP || ''),
    part_name: String(row.part_name || row.PART_NAME || ''),
    component: String(row.component || row.COMPONENT || ''),
    material: String(row.material || row.MATERIAL || ''),
    qty_g: parseFloat(row.qty_g || row.QTY_G || 0) || 0,
    scenario: String(row.scenario || row.SCENARIO || ''),
    harga: parseFloat(row.harga || row.HARGA || 0) || 0,
    qty_prod: parseFloat(row.qty_prod || row.QTY_PROD || 0) || 0,
    ok_prod: parseFloat(row.ok_prod || row.OK_PROD || 0) || 0,
    ng_prod: parseFloat(row.ng_prod || row.NG_PROD || 0) || 0,
    tanggal: d,
    kategori: String(row.kategori || row.KATEGORI || ''),
    bulan: d ? d.getMonth() + 1 : null,
    tahun: d ? d.getFullYear() : null,
    // Calculated
    get total_cost() { return (this.qty_g / 1000) * this.harga * this.qty_prod; },
    get ng_loss() { return (this.qty_g / 1000) * this.harga * this.ng_prod; },
    get usage_kg() { return (this.qty_g * this.qty_prod) / 1000 / 1000; },
  };
}

// ===== PROCESS DATA =====
function processData(rows) {
  return rows.map(normalizeRow);
}

// ===== APPLY FILTERS =====
function applyFilters() {
  const bulan = Array.from(document.getElementById('filter-bulan').selectedOptions).map(o => +o.value).filter(Boolean);
  const tahun = Array.from(document.getElementById('filter-tahun').selectedOptions).map(o => +o.value).filter(Boolean);
  const scenario = document.getElementById('filter-scenario').value;
  const kategoris = $('#filter-kategori').val() || [];
  const parts = $('#filter-part').val() || [];

  APP.filteredData = APP.rawData.filter(row => {
    if (bulan.length && !bulan.includes(row.bulan)) return false;
    if (tahun.length && !tahun.includes(row.tahun)) return false;
    if (scenario && row.scenario !== scenario) return false;
    if (kategoris.length && !kategoris.includes(row.kategori)) return false;
    if (parts.length && !parts.includes(row.part_name)) return false;
    return true;
  });

  document.getElementById('header-filtered-rows').textContent = formatNumber(APP.filteredData.length);
  renderDashboard();
}

// ===== CALCULATE OVERVIEW =====
function calculateOverview() {
  const std = APP.filteredData.filter(r => r.scenario === 'Standard');
  const alt = APP.filteredData.filter(r => r.scenario === 'Alternative');

  const totalStd = std.reduce((s, r) => s + r.total_cost, 0);
  const totalAlt = alt.reduce((s, r) => s + r.total_cost, 0);
  const totalSaving = totalStd - totalAlt;
  const pctSaving = totalStd !== 0 ? (totalSaving / totalStd) * 100 : 0;
  const totalNG = APP.filteredData.reduce((s, r) => s + r.ng_loss, 0);

  return { totalStd, totalAlt, totalSaving, pctSaving, totalNG };
}

// ===== RENDER KPI CARDS =====
function renderCards() {
  const { totalStd, totalAlt, totalSaving, pctSaving, totalNG } = calculateOverview();
  const isSaving = totalSaving >= 0;
  const savingClass = isSaving ? 'green' : 'red';
  const savingColor = isSaving ? '#10b981' : '#ef4444';

  const kpis = [
    {
      title: 'Total Cost Standard', val: formatIDR(totalStd), color: 'blue',
      icon: 'fas fa-circle-dollar-to-slot', iconColor: '#1d4ed8', bg: '#eff6ff',
      sub: 'Scenario Standard',
    },
    {
      title: 'Total Cost Alternative', val: formatIDR(totalAlt), color: 'cyan',
      icon: 'fas fa-arrows-rotate', iconColor: '#0891b2', bg: '#ecfeff',
      sub: 'Scenario Alternative',
    },
    {
      title: 'Total Saving', val: formatIDR(totalSaving), color: savingClass,
      icon: isSaving ? 'fas fa-piggy-bank' : 'fas fa-arrow-trend-down',
      iconColor: savingColor, bg: isSaving ? '#f0fdf4' : '#fef2f2',
      sub: isSaving ? 'Cost berhasil ditekan' : '🔴 Biaya meningkat',
    },
    {
      title: '% Saving', val: formatPct(pctSaving), color: savingClass,
      icon: 'fas fa-percent', iconColor: savingColor, bg: isSaving ? '#f0fdf4' : '#fef2f2',
      sub: 'Relative to Standard',
    },
    {
      title: 'Total NG Loss', val: formatIDR(totalNG), color: 'orange',
      icon: 'fas fa-triangle-exclamation', iconColor: '#f59e0b', bg: '#fffbeb',
      sub: 'Kerugian produk NG',
    },
  ];

  const container = document.getElementById('kpi-container');
  container.innerHTML = kpis.map((k, i) => `
    <div class="kpi-card ${k.color} fade-in fade-in-delay-${i+1}">
      <div class="flex items-start justify-between">
        <div class="w-10 h-10 rounded-xl flex items-center justify-center flex-shrink-0" style="background:${k.bg};">
          <i class="${k.icon}" style="color:${k.iconColor}; font-size:16px;"></i>
        </div>
        <div class="text-right flex-1 ml-2">
          <div class="text-xs font-semibold text-slate-500 mb-0.5">${k.title}</div>
          <div class="font-mono font-bold text-slate-800 leading-tight" style="font-size:16px;">${k.val}</div>
          <div class="text-xs text-slate-400 mt-0.5">${k.sub}</div>
        </div>
      </div>
    </div>
  `).join('');
}

// ===== RENDER COMPARISON CHART =====
function renderComparisonChart() {
  const std = APP.filteredData.filter(r => r.scenario === 'Standard');
  const alt = APP.filteredData.filter(r => r.scenario === 'Alternative');
  const totalStd = std.reduce((s, r) => s + r.total_cost, 0);
  const totalAlt = alt.reduce((s, r) => s + r.total_cost, 0);
  const saving = totalStd - totalAlt;

  destroyChart('comparison');

  const isNominal = APP.comparisonMode === 'nominal';
  let categories, seriesData;

  if (isNominal) {
    categories = ['Cost Standard', 'Cost Alternative', 'Total Saving'];
    seriesData = [
      { x: 'Cost Standard', y: Math.round(totalStd), fillColor: '#1d4ed8' },
      { x: 'Cost Alternative', y: Math.round(totalAlt), fillColor: '#0891b2' },
      { x: 'Total Saving', y: Math.round(saving), fillColor: saving >= 0 ? '#10b981' : '#ef4444' },
    ];
  } else {
    const pctAlt = totalStd > 0 ? (totalAlt / totalStd) * 100 : 0;
    const pctSaving = totalStd > 0 ? ((totalStd - totalAlt) / totalStd) * 100 : 0;
    categories = ['Cost Standard', 'Cost Alternative', 'Saving'];
    seriesData = [
      { x: 'Cost Standard', y: 100, fillColor: '#1d4ed8' },
      { x: 'Cost Alternative', y: +pctAlt.toFixed(2), fillColor: '#0891b2' },
      { x: 'Saving', y: +pctSaving.toFixed(2), fillColor: pctSaving >= 0 ? '#10b981' : '#ef4444' },
    ];
  }

  const opts = {
    series: [{ name: 'Value', data: seriesData }],
    chart: {
      type: 'bar',
      height: 360,
      toolbar: { show: false },
      animations: { enabled: true, easing: 'easeinout', speed: 600 },
      fontFamily: 'Plus Jakarta Sans, sans-serif',
    },
    plotOptions: {
      bar: {
        horizontal: true,
        distributed: true,
        borderRadius: 8,
        barHeight: '50%',
        dataLabels: { position: 'top' },
      }
    },
    dataLabels: {
      enabled: true,
      formatter: v => isNominal ? formatIDR(v) : v.toFixed(1) + '%',
      style: { fontSize: '12px', fontWeight: '700', colors: ['#1e293b'] },
      offsetX: 5,
    },
    legend: { show: false },
    tooltip: {
      y: { formatter: v => isNominal ? formatIDR(v, true) : v.toFixed(2) + '%' }
    },
    xaxis: {
      labels: { formatter: v => isNominal ? formatIDR(v) : v + '%', style: { fontSize: '11px', colors: '#94a3b8' } },
      axisBorder: { show: false },
      axisTicks: { show: false },
    },
    yaxis: { labels: { style: { fontSize: '12px', fontWeight: '600', colors: '#475569' } } },
    grid: { borderColor: '#f1f5f9', strokeDashArray: 3 },
    colors: seriesData.map(s => s.fillColor),
  };

  APP.charts.comparison = new ApexCharts(document.getElementById('chart-comparison'), opts);
  APP.charts.comparison.render();
}

// ===== RENDER TOP SAVING CHART =====
function renderTopSavingChart() {
  const mode = APP.topChartMode;

  // Group by part_name: std vs alt
  const byPart = {};
  APP.filteredData.forEach(r => {
    if (!byPart[r.part_name]) byPart[r.part_name] = { std: 0, alt: 0 };
    if (r.scenario === 'Standard') byPart[r.part_name].std += r.total_cost;
    if (r.scenario === 'Alternative') byPart[r.part_name].alt += r.total_cost;
  });

  let parts = Object.entries(byPart).map(([name, v]) => ({
    name, saving: v.std - v.alt
  }));

  parts.sort((a, b) => mode === 'saving' ? b.saving - a.saving : a.saving - b.saving);
  const top10 = parts.slice(0, 10);

  destroyChart('topSaving');

  const isSaving = mode === 'saving';
  const barColor = isSaving ? '#10b981' : '#ef4444';

  const opts = {
    series: [{ name: isSaving ? 'Saving' : 'Loss', data: top10.map(p => Math.round(Math.abs(p.saving))) }],
    chart: {
      type: 'bar',
      height: 360,
      toolbar: { show: false },
      animations: { enabled: true, easing: 'easeinout', speed: 500 },
      fontFamily: 'Plus Jakarta Sans, sans-serif',
    },
    plotOptions: {
      bar: {
        horizontal: true,
        borderRadius: 6,
        barHeight: '55%',
      }
    },
    colors: [barColor],
    dataLabels: {
      enabled: true,
      formatter: v => formatIDR(v),
      style: { fontSize: '11px', fontWeight: '600', colors: ['#fff'] },
    },
    tooltip: {
      y: { formatter: v => formatIDR(v, true) }
    },
    xaxis: {
      categories: top10.map(p => p.name.length > 22 ? p.name.substring(0,20)+'…' : p.name),
      labels: { formatter: v => formatIDR(v), style: { fontSize: '10px', colors: '#94a3b8' } },
      axisBorder: { show: false }, axisTicks: { show: false },
    },
    yaxis: { labels: { style: { fontSize: '11px', fontWeight: '600', colors: '#475569' } } },
    grid: { borderColor: '#f1f5f9', strokeDashArray: 3 },
  };

  APP.charts.topSaving = new ApexCharts(document.getElementById('chart-top-saving'), opts);
  APP.charts.topSaving.render();
}

// ===== RENDER CATEGORY CARDS =====
function renderCategoryCards() {
  const mode = APP.categoryViewMode;

  // Group by kategori
  const byKat = {};
  APP.filteredData.forEach(r => {
    if (!byKat[r.kategori]) byKat[r.kategori] = { parts: new Set(), std: 0, alt: 0, rows: [] };
    byKat[r.kategori].parts.add(r.part_name);
    if (r.scenario === 'Standard') byKat[r.kategori].std += r.total_cost;
    if (r.scenario === 'Alternative') byKat[r.kategori].alt += r.total_cost;
    byKat[r.kategori].rows.push(r);
  });

  const container = document.getElementById('category-cards');

  if (Object.keys(byKat).length === 0) {
    container.innerHTML = `<div class="col-span-3 empty-state"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg><div class="text-base font-semibold">Data belum tersedia</div><div class="text-sm mt-1">Upload file Excel terlebih dahulu</div></div>`;
    return;
  }

  container.innerHTML = Object.entries(byKat).map(([kat, val], idx) => {
    const saving = val.std - val.alt;
    const partCount = val.parts.size;
    const avgSaving = partCount > 0 ? saving / partCount : 0;
    const isSav = saving >= 0;
    const color = CATEGORY_COLORS[idx % CATEGORY_COLORS.length];

    // Top parts by saving or loss
    const partSaving = {};
    val.rows.forEach(r => {
      if (!partSaving[r.part_name]) partSaving[r.part_name] = { std: 0, alt: 0 };
      if (r.scenario === 'Standard') partSaving[r.part_name].std += r.total_cost;
      if (r.scenario === 'Alternative') partSaving[r.part_name].alt += r.total_cost;
    });
    let topParts = Object.entries(partSaving)
      .map(([name, v]) => ({ name, saving: v.std - v.alt }));

    if (mode === 'saving') topParts.sort((a, b) => b.saving - a.saving);
    else topParts.sort((a, b) => a.saving - b.saving);
    const top5 = topParts.slice(0, 5);

    const maxAbsSaving = Math.max(...top5.map(p => Math.abs(p.saving)), 1);

    return `
      <div class="category-card fade-in fade-in-delay-${(idx % 4) + 1}">
        <div class="h-1.5" style="background: linear-gradient(90deg, ${color}, ${color}88);"></div>
        <div class="p-4">
          <div class="flex items-start justify-between mb-3">
            <div>
              <div class="font-bold text-slate-800 text-base">${kat || 'Unknown'}</div>
              <div class="text-xs text-slate-500 mt-0.5">${partCount} Part</div>
            </div>
            <div class="text-right">
              <div class="font-mono font-bold text-base ${isSav ? 'text-emerald-600' : 'text-red-600'}">${formatIDR(saving)}</div>
              <div class="text-xs text-slate-400">Avg ${formatIDR(avgSaving)}</div>
            </div>
          </div>

          <div class="mb-3">
            <div class="text-xs font-semibold text-slate-500 mb-2">${mode === 'saving' ? 'Top 5 Saving' : 'Top 5 Loss'}</div>
            <div class="space-y-1.5">
              ${top5.map((p, i) => {
                const pct = Math.abs(p.saving) / maxAbsSaving * 100;
                const pSav = p.saving >= 0;
                const rankClass = i === 0 ? 'rank-1' : i === 1 ? 'rank-2' : i === 2 ? 'rank-3' : 'rank-n';
                return `
                  <div>
                    <div class="flex items-center justify-between mb-0.5">
                      <div class="flex items-center gap-1.5">
                        <span class="rank-badge ${rankClass}">${i+1}</span>
                        <span class="text-xs text-slate-700 font-medium truncate" style="max-width:140px;" title="${p.name}">${p.name}</span>
                      </div>
                      <span class="text-xs font-mono font-semibold ${pSav ? 'text-emerald-600' : 'text-red-600'}">${formatIDR(p.saving)}</span>
                    </div>
                    <div class="progress-bar">
                      <div class="progress-bar-fill" style="width:${pct}%; background: ${pSav ? '#10b981' : '#ef4444'};"></div>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        </div>
      </div>
    `;
  }).join('');
}

// ===== RENDER NG CHARTS =====
function renderNGCharts() {

  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(
    r => r.scenario === 'Alternative'
  );

  // ==============================
  // Group by part
  // ==============================
  const byPart = {};

  altData.forEach(r => {

    if (!byPart[r.part_name]) {

      byPart[r.part_name] = {
        ng_loss: 0,
        ng_prod: 0,
        qty_prod: 0,
        material: r.material,
        kategori: r.kategori
      };
    }

    byPart[r.part_name].ng_loss += r.ng_loss;
    byPart[r.part_name].ng_prod += r.ng_prod;
    byPart[r.part_name].qty_prod += r.qty_prod;
  });

  const parts = Object.entries(byPart)
    .map(([name, v]) => ({
      ...v,
      part_name: name,
      ng_rate:
        v.qty_prod > 0
          ? (v.ng_prod / v.qty_prod) * 100
          : 0
    }))
    .sort((a, b) => b.ng_loss - a.ng_loss)
    .slice(0, 10);

  // ==============================
  // Top 10 NG Part Bar
  // ==============================
  destroyChart('ngTop');

  APP.charts.ngTop = new ApexCharts(
    document.getElementById('chart-ng-top'),
    {
      series: [{
        name: 'NG Loss',
        data: parts.map(p => Math.round(p.ng_loss))
      }],

      chart: {
        type: 'bar',
        height: 320,
        toolbar: { show: false },
        fontFamily: 'Plus Jakarta Sans, sans-serif',
        animations: { speed: 500 }
      },

      plotOptions: {
        bar: {
          horizontal: true,
          borderRadius: 6,
          barHeight: '60%'
        }
      },

      colors: ['#ef4444'],

      dataLabels: {
        enabled: true,
        formatter: v => formatIDR(v),

        style: {
          fontSize: '11px',
          colors: ['#fff'],
          fontWeight: '600'
        }
      },

      tooltip: {
        y: {
          formatter: v => formatIDR(v, true)
        }
      },

      xaxis: {

        categories: parts.map(p =>
          p.part_name.length > 22
            ? p.part_name.substring(0, 20) + '…'
            : p.part_name
        ),

        labels: {
          formatter: v => formatIDR(v),

          style: {
            fontSize: '10px',
            colors: '#94a3b8'
          }
        },

        axisBorder: { show: false },
        axisTicks: { show: false },
      },

      yaxis: {
        labels: {
          style: {
            fontSize: '11px',
            fontWeight: '600',
            colors: '#475569'
          }
        }
      },

      grid: {
        borderColor: '#f1f5f9',
        strokeDashArray: 3
      },
    }
  );

  APP.charts.ngTop.render();

  // ==============================
  // NG Loss by Kategori Donut
  // ==============================
  const byKat = {};

  altData.forEach(r => {
    byKat[r.kategori] =
      (byKat[r.kategori] || 0) + r.ng_loss;
  });

  const katLabels = Object.keys(byKat);

  const katVals = Object.values(byKat)
    .map(v => Math.round(v));

  destroyChart('ngDonut');

  APP.charts.ngDonut = new ApexCharts(
    document.getElementById('chart-ng-donut'),
    {
      series: katVals,
      labels: katLabels,

      chart: {
        type: 'donut',
        height: 320,
        fontFamily: 'Plus Jakarta Sans, sans-serif',
        animations: { speed: 500 }
      },

      colors: CATEGORY_COLORS,

      dataLabels: {
        enabled: true,

        formatter: (v, o) =>
          o.w.globals.labels[o.seriesIndex] +
          '\n' +
          v.toFixed(1) +
          '%'
      },

      plotOptions: {
        pie: {
          donut: {
            size: '65%',

            labels: {
              show: true,

              total: {
                show: true,
                label: 'Total Loss',

                formatter: () =>
                  formatIDR(
                    katVals.reduce((a, b) => a + b, 0)
                  )
              }
            }
          }
        }
      },

      tooltip: {
        y: {
          formatter: v => formatIDR(v, true)
        }
      },

      legend: {
        position: 'bottom',
        fontSize: '11px'
      },
    }
  );

  APP.charts.ngDonut.render();

  // ==============================
  // NG Rate Gauge
  // ==============================
  const totalNG = altData.reduce(
    (s, r) => s + r.ng_prod,
    0
  );

  const totalQty = altData.reduce(
    (s, r) => s + r.qty_prod,
    0
  );

  const ngRate =
    totalQty > 0
      ? (totalNG / totalQty) * 100
      : 0;

  destroyChart('ngGauge');

  APP.charts.ngGauge = new ApexCharts(
    document.getElementById('chart-ng-gauge'),
    {
      series: [+ngRate.toFixed(2)],

      chart: {
        type: 'radialBar',
        height: 280,
        fontFamily: 'Plus Jakarta Sans, sans-serif'
      },

      plotOptions: {
        radialBar: {

          startAngle: -135,
          endAngle: 135,

          hollow: {
            size: '60%'
          },

          dataLabels: {

            name: {
              show: true,
              fontSize: '13px',
              color: '#64748b',
              offsetY: -6
            },

            value: {
              fontSize: '24px',
              fontWeight: '700',
              color: '#1e293b',
              offsetY: 8,
              formatter: v => v + '%'
            },
          },

          track: {
            background: '#f1f5f9',
            strokeWidth: '97%'
          },
        }
      },

      fill: {
        type: 'gradient',

        gradient: {
          shade: 'light',
          type: 'horizontal',

          gradientToColors:
            ngRate < 3
              ? ['#10b981']
              : ngRate < 8
                ? ['#f59e0b']
                : ['#ef4444'],

          stops: [0, 100]
        }
      },

      colors:
        ngRate < 3
          ? ['#34d399']
          : ngRate < 8
            ? ['#fbbf24']
            : ['#f87171'],

      labels: ['NG Rate'],
    }
  );

  APP.charts.ngGauge.render();

  document.getElementById('ng-rate-label')
    .textContent = ngRate.toFixed(2) + '%';

  // ==============================
  // Worst Material
  // ==============================
  const byMat = {};

  altData.forEach(r => {

    if (!byMat[r.material]) {

      byMat[r.material] = {
        ng: 0,
        loss: 0
      };
    }

    byMat[r.material].ng += r.ng_prod;
    byMat[r.material].loss += r.ng_loss;
  });

  const worstMats = Object.entries(byMat)
    .map(([mat, v]) => ({
      mat,
      ...v
    }))
    .sort((a, b) => b.loss - a.loss)
    .slice(0, 5);

  const wc =
    document.getElementById('worst-material-list');

  if (worstMats.length === 0) {

    wc.innerHTML = `
      <div class="empty-state"
           style="padding: 30px 20px;">
        <div class="text-sm">
          Data kosong
        </div>
      </div>
    `;

  } else {

    wc.innerHTML = worstMats.map((m, i) => `
      <div
        class="flex items-center justify-between p-3 rounded-xl"
        style="
          background: ${i === 0 ? '#fef2f2' : '#f8fafc'};
          border: 1px solid ${i === 0 ? '#fecaca' : '#e2e8f0'};
        ">

        <div class="flex items-center gap-2">

          <div class="rank-badge ${['rank-1','rank-2','rank-3','rank-n','rank-n'][i]}">
            ${i + 1}
          </div>

          <div>

            <div class="text-sm font-semibold text-slate-700">
              ${m.mat || '—'}
            </div>

            <div class="text-xs text-slate-400">
              NG: ${formatNumber(m.ng)}
            </div>

          </div>

        </div>

        <div class="text-right">

          <div class="text-sm font-bold font-mono text-red-600">
            ${formatIDR(m.loss)}
          </div>

          <div class="text-xs text-slate-400">
            NG Loss
          </div>

        </div>

      </div>
    `).join('');
  }
}

// ===== RENDER NG TABLE =====
let ngTableData = [];

function renderNGTable() {

  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(
    r => r.scenario === 'Alternative'
  );

  const byPart = {};

  altData.forEach(r => {

    const key =
      r.part_name + '|' + r.material;

    if (!byPart[key]) {

      byPart[key] = {
        part_name: r.part_name,
        material: r.material,
        qty_prod: 0,
        ng_prod: 0,
        ng_loss: 0
      };
    }

    byPart[key].qty_prod += r.qty_prod;
    byPart[key].ng_prod += r.ng_prod;
    byPart[key].ng_loss += r.ng_loss;
  });

  ngTableData = Object.values(byPart)
    .map(r => ({
      ...r,
      ng_rate:
        r.qty_prod > 0
          ? (r.ng_prod / r.qty_prod) * 100
          : 0
    }));

  renderNGTablePage(1);
}

function renderNGTablePage(page) {

  APP.ngTablePage = page;

  const search = (
    document.getElementById('ng-search')?.value || ''
  ).toLowerCase();

  const { col, dir } = APP.ngTableSort;

  let data = ngTableData.filter(r =>
    r.part_name.toLowerCase().includes(search) ||
    r.material.toLowerCase().includes(search)
  );

  data.sort((a, b) => {

    const av = a[col];
    const bv = b[col];

    return dir === 'asc'
      ? (av > bv ? 1 : -1)
      : (av < bv ? 1 : -1);
  });

  const total = data.length;

  const start =
    (page - 1) * APP.PAGE_SIZE;

  const slice = data.slice(
    start,
    start + APP.PAGE_SIZE
  );

  const tbody =
    document.getElementById('ng-table-body');

  if (slice.length === 0) {

    tbody.innerHTML = `
      <tr>
        <td colspan="6"
            class="text-center py-8 text-slate-400">
          Tidak ada data
        </td>
      </tr>
    `;

  } else {

    tbody.innerHTML = slice.map(r => {

      const ngClass =
        r.ng_rate > 8
          ? 'badge-red'
          : r.ng_rate > 3
            ? 'badge-yellow'
            : 'badge-green';

      return `
        <tr>

          <td class="font-medium text-slate-700">
            ${r.part_name}
          </td>

          <td class="text-slate-500">
            ${r.material}
          </td>

          <td class="font-mono text-slate-600">
            ${formatNumber(r.qty_prod)}
          </td>

          <td class="font-mono text-red-600 font-semibold">
            ${formatNumber(r.ng_prod)}
          </td>

          <td>
            <span class="badge ${ngClass}">
              ${r.ng_rate.toFixed(2)}%
            </span>
          </td>

          <td class="font-mono font-bold text-red-700">
            ${formatIDR(r.ng_loss)}
          </td>

        </tr>
      `;
    }).join('');
  }

  renderPagination(
    'ng-pagination',
    page,
    Math.ceil(total / APP.PAGE_SIZE),
    renderNGTablePage
  );
}

function sortNGTable(col) {

  if (APP.ngTableSort.col === col) {

    APP.ngTableSort.dir =
      APP.ngTableSort.dir === 'asc'
        ? 'desc'
        : 'asc';

  } else {

    APP.ngTableSort.col = col;
    APP.ngTableSort.dir = 'desc';
  }

  renderNGTablePage(1);
}

// ===== RENDER MATERIAL CHARTS =====
function renderMaterialAnalysis() {

  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(
    r => r.scenario === 'Alternative'
  );

  // ==============================
  // Composition Donut by usage_kg
  // ==============================
  const byKat = {};

  altData.forEach(r => {
    byKat[r.kategori] =
      (byKat[r.kategori] || 0) + r.usage_kg;
  });

  const katLabels = Object.keys(byKat);

  const katVals = Object.values(byKat)
    .map(v => +v.toFixed(2));

  destroyChart('matComposition');

  APP.charts.matComposition = new ApexCharts(
    document.getElementById('chart-mat-composition'),
    {
      series: katVals,
      labels: katLabels,

      chart: {
        type: 'donut',
        height: 300,
        fontFamily: 'Plus Jakarta Sans, sans-serif',
        animations: {
          speed: 500
        }
      },

      colors: CATEGORY_COLORS,

      dataLabels: {
        enabled: true,
        formatter: (v, o) =>
          o.w.globals.labels[o.seriesIndex] +
          '\n' +
          v.toFixed(1) +
          '%'
      },

      plotOptions: {
        pie: {
          donut: {
            size: '65%',

            labels: {
              show: true,

              total: {
                show: true,
                label: 'Total Usage',

                formatter: () =>
                  formatKg(
                    katVals.reduce((a, b) => a + b, 0)
                  )
              }
            }
          }
        }
      },

      tooltip: {
        y: {
          formatter: v => formatKg(v)
        }
      },

      legend: {
        position: 'bottom',
        fontSize: '11px'
      },
    }
  );

  APP.charts.matComposition.render();

  // ==============================
  // Top Material Cost Bar
  // ==============================
  const byMat = {};

  altData.forEach(r => {

    if (!byMat[r.material]) {
      byMat[r.material] = 0;
    }

    byMat[r.material] += r.total_cost;
  });

  const matSorted = Object.entries(byMat)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  destroyChart('matCost');

  APP.charts.matCost = new ApexCharts(
    document.getElementById('chart-mat-cost'),
    {
      series: [{
        name: 'Total Cost',
        data: matSorted.map(m => Math.round(m[1]))
      }],

      chart: {
        type: 'bar',
        height: 300,
        toolbar: {
          show: false
        },
        fontFamily: 'Plus Jakarta Sans, sans-serif',
        animations: {
          speed: 500
        }
      },

      plotOptions: {
        bar: {
          horizontal: true,
          borderRadius: 6,
          barHeight: '55%'
        }
      },

      colors: ['#1d4ed8'],

      dataLabels: {
        enabled: true,

        formatter: v => formatIDR(v),

        style: {
          fontSize: '11px',
          colors: ['#fff'],
          fontWeight: '600'
        }
      },

      tooltip: {
        y: {
          formatter: v => formatIDR(v, true)
        }
      },

      xaxis: {

        categories: matSorted.map(([m]) =>
          m.length > 22
            ? m.substring(0, 20) + '…'
            : m
        ),

        labels: {
          formatter: v => formatIDR(v),

          style: {
            fontSize: '10px',
            colors: '#94a3b8'
          }
        },

        axisBorder: {
          show: false
        },

        axisTicks: {
          show: false
        },
      },

      yaxis: {
        labels: {
          style: {
            fontSize: '11px',
            fontWeight: '600',
            colors: '#475569'
          }
        }
      },

      grid: {
        borderColor: '#f1f5f9',
        strokeDashArray: 3
      },
    }
  );

  APP.charts.matCost.render();

  // ==============================
  // Most Used Material Table
  // ==============================
  renderMatTable();
}

let matTableData = [];

function renderMatTable() {

  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(
    r => r.scenario === 'Alternative'
  );

  const byMat = {};

  altData.forEach(r => {

    if (!byMat[r.material]) {

      byMat[r.material] = {
        material: r.material,
        kategori: r.kategori,
        usage_kg: 0,
        total_cost: 0,
        parts: new Set()
      };
    }

    byMat[r.material].usage_kg += r.usage_kg;
    byMat[r.material].total_cost += r.total_cost;
    byMat[r.material].parts.add(r.part_name);
  });

  matTableData = Object.values(byMat)
    .map(r => ({
      ...r,
      total_part: r.parts.size
    }));

  renderMatTablePage(1);
}

function renderMatTablePage(page) {

  APP.matTablePage = page;

  const search = (
    document.getElementById('mat-search')?.value || ''
  ).toLowerCase();

  const { col, dir } = APP.matTableSort;

  let data = matTableData.filter(r =>
    r.material.toLowerCase().includes(search) ||
    r.kategori.toLowerCase().includes(search)
  );

  data.sort((a, b) => {

    const av = a[col];
    const bv = b[col];

    return dir === 'asc'
      ? (av > bv ? 1 : -1)
      : (av < bv ? 1 : -1);
  });

  const maxUsage = Math.max(
    ...data.map(r => r.usage_kg),
    1
  );

  const total = data.length;

  const start =
    (page - 1) * APP.PAGE_SIZE;

  const slice = data.slice(
    start,
    start + APP.PAGE_SIZE
  );

  const tbody =
    document.getElementById('mat-table-body');

  if (slice.length === 0) {

    tbody.innerHTML = `
      <tr>
        <td colspan="5"
            class="text-center py-8 text-slate-400">
          Tidak ada data
        </td>
      </tr>
    `;

  } else {

    tbody.innerHTML = slice.map(r => `
      <tr>

        <td class="font-medium text-slate-700">
          ${r.material}
        </td>

        <td>
          <span class="badge badge-blue">
            ${r.kategori}
          </span>
        </td>

        <td>

          <div class="flex items-center gap-2">

            <div style="width:60px;">

              <div class="progress-bar">

                <div
                  class="progress-bar-fill"
                  style="
                    width:${(r.usage_kg / maxUsage * 100).toFixed(1)}%;
                    background:#1d4ed8;
                  ">
                </div>

              </div>

            </div>

            <span class="font-mono text-xs text-slate-600">
              ${formatKg(r.usage_kg)}
            </span>

          </div>

        </td>

        <td class="font-mono font-semibold text-slate-700">
          ${formatIDR(r.total_cost)}
        </td>

        <td class="font-mono text-slate-600">
          ${r.total_part}
        </td>

      </tr>
    `).join('');
  }

  renderPagination(
    'mat-pagination',
    page,
    Math.ceil(total / APP.PAGE_SIZE),
    renderMatTablePage
  );
}

function sortMatTable(col) {

  if (APP.matTableSort.col === col) {

    APP.matTableSort.dir =
      APP.matTableSort.dir === 'asc'
        ? 'desc'
        : 'asc';

  } else {

    APP.matTableSort.col = col;
    APP.matTableSort.dir = 'desc';
  }

  renderMatTablePage(1);
}

// ===== PAGINATION =====
function renderPagination(containerId, currentPage, totalPages, callback) {
  const c = document.getElementById(containerId);
  if (!c) return;
  if (totalPages <= 1) { c.innerHTML = ''; return; }

  let html = `<div class="flex items-center gap-1 text-xs">`;
  html += `<button onclick="${callback.name}(${currentPage-1})" ${currentPage===1?'disabled':''} class="px-2 py-1 rounded border border-slate-200 ${currentPage===1?'opacity-40 cursor-not-allowed':'hover:bg-slate-50'}">‹</button>`;

  const pages = [];
  if (totalPages <= 5) {
    for (let i = 1; i <= totalPages; i++) pages.push(i);
  } else {
    pages.push(1);
    if (currentPage > 3) pages.push('...');
    for (let i = Math.max(2, currentPage-1); i <= Math.min(totalPages-1, currentPage+1); i++) pages.push(i);
    if (currentPage < totalPages - 2) pages.push('...');
    pages.push(totalPages);
  }

  pages.forEach(p => {
    if (p === '...') html += `<span class="px-2">…</span>`;
    else html += `<button onclick="${callback.name}(${p})" class="px-2 py-1 rounded border ${p===currentPage?'bg-blue-600 text-white border-blue-600':'border-slate-200 hover:bg-slate-50'}">${p}</button>`;
  });

  html += `<button onclick="${callback.name}(${currentPage+1})" ${currentPage===totalPages?'disabled':''} class="px-2 py-1 rounded border border-slate-200 ${currentPage===totalPages?'opacity-40 cursor-not-allowed':'hover:bg-slate-50'}">›</button>`;
  html += `<span class="ml-2 text-slate-400">Halaman ${currentPage} / ${totalPages}</span>`;
  html += `</div>`;
  c.innerHTML = html;
}

// ===== DESTROY CHART =====
function destroyChart(key) {
  if (APP.charts[key]) {
    try { APP.charts[key].destroy(); } catch (e) { /* ignore */ }
    delete APP.charts[key];
  }
}

// ===== RENDER DASHBOARD =====
function renderDashboard() {
  const activeTab = document.querySelector('.section-container.active')?.id?.replace('tab-','') || 'overview';

  if (activeTab === 'overview') {
    renderCards();
    renderComparisonChart();
    renderTopSavingChart();
  } else if (activeTab === 'category') {
    renderCategoryCards();
  } else if (activeTab === 'ng') {
    renderNGCharts();
    renderNGTable();
  } else if (activeTab === 'material') {
    renderMaterialAnalysis();
  }
}

// ===== SWITCH TAB =====
function switchTab(tab) {
  document.querySelectorAll('.section-container').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
  document.getElementById(`tab-${tab}`)?.classList.add('active');
  document.querySelectorAll('.nav-tab').forEach(el => {
    if (el.getAttribute('onclick')?.includes(tab)) el.classList.add('active');
  });
  renderDashboard();
  applyFilters();
}

// ===== TOGGLE BUTTONS =====
function toggleComparison(mode) {
  APP.comparisonMode = mode;
  document.getElementById('btn-comparison-nominal').classList.toggle('active', mode === 'nominal');
  document.getElementById('btn-comparison-pct').classList.toggle('active', mode === 'pct');
  renderComparisonChart();
}

function toggleTopChart(mode) {
  APP.topChartMode = mode;
  document.getElementById('btn-top-saving').classList.toggle('active', mode === 'saving');
  document.getElementById('btn-top-loss').classList.toggle('active', mode === 'loss');
  renderTopSavingChart();
}

function toggleCategoryView(mode) {
  APP.categoryViewMode = mode;
  document.getElementById('btn-cat-saving').classList.toggle('active', mode === 'saving');
  document.getElementById('btn-cat-loss').classList.toggle('active', mode === 'loss');
  renderCategoryCards();
}

// ==============================
// POPULATE FILTERS
// ==============================
function populateFilters() {

  const bulanSet = new Set();
  const tahunSet = new Set();
  const katSet = new Set();
  const partSet = new Set();

  APP.rawData.forEach(r => {

    if (r.bulan !== undefined && r.bulan !== null) {
      bulanSet.add(r.bulan);
    }

    if (r.tahun !== undefined && r.tahun !== null) {
      tahunSet.add(r.tahun);
    }

    if (r.kategori) {
      katSet.add(r.kategori);
    }

    if (r.part_name) {
      partSet.add(r.part_name);
    }
  });

  // =========================
  // SET OPTIONS
  // =========================

  $('#filter-bulan').html(
    [...bulanSet]
      .sort((a, b) => a - b)
      .map(m => `
        <option value="${m}">
          ${BULAN_NAMES[m - 1]}
        </option>
      `)
      .join('')
  );

  $('#filter-tahun').html(
    [...tahunSet]
      .sort()
      .map(y => `
        <option value="${y}">
          ${y}
        </option>
      `)
      .join('')
  );

  $('#filter-kategori').html(
    [...katSet]
      .sort()
      .map(k => `
        <option value="${k}">
          ${k}
        </option>
      `)
      .join('')
  );

  $('#filter-part').html(
    [...partSet]
      .sort()
      .map(p => `
        <option value="${p}">
          ${p}
        </option>
      `)
      .join('')
  );

  // =========================
  // DESTROY OLD SELECT2
  // =========================

  $('.filter-multi').each(function () {

    if ($(this).hasClass('select2-hidden-accessible')) {
      $(this).select2('destroy');
    }

  });

  // =========================
  // INIT SELECT2
  // =========================

  $('.filter-multi').select2({

    width: '100%',

    closeOnSelect: false,

    allowClear: true,

    placeholder: 'Pilih Data',

    language: {
      noResults: () => 'Data tidak ditemukan'
    }

  });

  // =========================
  // UPDATE COUNTER
  // =========================

  function updateCounter(id, emptyText) {

    const val = $(id).val();

    const total =
      Array.isArray(val)
        ? val.length
        : 0;

    const rendered = $(id)
      .next('.select2-container')
      .find('.select2-selection__rendered');

    // kosongkan default render
    rendered.find('li').hide();

    // custom text
    if (total === 0) {

      rendered.attr(
        'data-count',
        emptyText
      );

    } else {

      rendered.attr(
        'data-count',
        total + ' dipilih'
      );
    }
  }

  // =========================
  // INIT COUNTER
  // =========================

  updateCounter(
    '#filter-bulan',
    'Semua Bulan'
  );

  updateCounter(
    '#filter-tahun',
    'Semua Tahun'
  );

  updateCounter(
    '#filter-scenario',
    'Semua Scenario'
  );

  updateCounter(
    '#filter-kategori',
    'Semua Kategori'
  );

  updateCounter(
    '#filter-part',
    'Semua Part'
  );

  // =========================
  // REMOVE OLD EVENT
  // =========================

  $('.filter-multi').off('change');

  // =========================
  // CHANGE EVENT
  // =========================

  $('.filter-multi').on('change', function () {

    const id =
      '#' + $(this).attr('id');

    let emptyText = 'Pilih Data';

    switch (id) {

      case '#filter-bulan':
        emptyText = 'Semua Bulan';
        break;

      case '#filter-tahun':
        emptyText = 'Semua Tahun';
        break;

      case '#filter-scenario':
        emptyText = 'Semua Scenario';
        break;

      case '#filter-kategori':
        emptyText = 'Semua Kategori';
        break;

      case '#filter-part':
        emptyText = 'Semua Part';
        break;
    }

    updateCounter(id, emptyText);

    debounceFilter();
  });

  // =========================
  // FIX SEARCH INPUT
  // =========================

  $('.select2-search__field').css({
    width: '100%',

    minWidth: '120px',

    fontSize: '13px',

    fontFamily: 'Plus Jakarta Sans, sans-serif'
  });

}

// ===== DEBOUNCE FILTER =====
function debounceFilter() {
  clearTimeout(APP.debounceTimer);
  APP.debounceTimer = setTimeout(() => {
    applyFilters();
  }, 300);
}

// ===== RESET FILTERS =====
function resetFilters() {
  document.getElementById('filter-bulan').selectedIndex = -1;
  document.getElementById('filter-tahun').selectedIndex = -1;
  document.getElementById('filter-scenario').value = '';
  $('#filter-kategori').val(null).trigger('change');
  $('#filter-part').val(null).trigger('change');
  applyFilters();
}

// ===== EXPORT EXCEL =====
function exportExcel() {
  if (!APP.filteredData.length) return alert('Tidak ada data untuk di-export.');
  const ws = XLSX.utils.json_to_sheet(APP.filteredData.map(r => ({
    no_sap: r.no_sap,
    part_name: r.part_name,
    component: r.component,
    material: r.material,
    qty_g: r.qty_g,
    scenario: r.scenario,
    harga: r.harga,
    qty_prod: r.qty_prod,
    ok_prod: r.ok_prod,
    ng_prod: r.ng_prod,
    tanggal: r.tanggal ? `${String(r.tanggal.getDate()).padStart(2,'0')}/${String(r.tanggal.getMonth()+1).padStart(2,'0')}/${r.tanggal.getFullYear()}` : '',
    kategori: r.kategori,
    total_cost: +r.total_cost.toFixed(2),
    ng_loss: +r.ng_loss.toFixed(2),
    usage_kg: +r.usage_kg.toFixed(4),
  })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');
  XLSX.writeFile(wb, `astra_cost_saving_export_${Date.now()}.xlsx`);
}

function exportNGTableCSV() {
  const data = ngTableData;
  if (!data.length) return;
  const header = 'Part,Material,Qty Prod,NG,NG Rate,NG Loss';
  const rows = data.map(r => [r.part_name, r.material, r.qty_prod, r.ng_prod, r.ng_rate.toFixed(2) + '%', r.ng_loss.toFixed(2)].join(','));
  const csv = [header, ...rows].join('\n');
  const a = document.createElement('a');
  a.href = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);
  a.download = 'ng_analysis.csv';
  a.click();
}

function exportMatTableExcel() {
  if (!matTableData.length) return;
  const ws = XLSX.utils.json_to_sheet(matTableData.map(r => ({
    material: r.material,
    kategori: r.kategori,
    usage_kg: +r.usage_kg.toFixed(4),
    total_cost: +r.total_cost.toFixed(2),
    total_part: r.total_part,
  })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Material Usage');
  XLSX.writeFile(wb, `material_usage_${Date.now()}.xlsx`);
}

// ===== UPLOAD HANDLER =====
async function handleFile(file) {
  if (!file) return;
  APP.fileName = file.name;

  // Show loading
  const overlay = document.getElementById('loading-overlay');
  overlay.style.display = 'flex';
  document.getElementById('loading-info').textContent = 'Membaca file: ' + file.name;

  try {
    await new Promise(r => setTimeout(r, 50)); // let browser repaint
    const rows = await parseExcel(file);

    document.getElementById('loading-info').textContent = `Memproses ${rows.length} baris data…`;
    await new Promise(r => setTimeout(r, 30));

    APP.rawData = processData(rows);
    APP.filteredData = [...APP.rawData];

    document.getElementById('header-total-rows').textContent = formatNumber(APP.rawData.length);
    document.getElementById('header-filtered-rows').textContent = formatNumber(APP.filteredData.length);
    document.getElementById('header-date').textContent = new Date().toLocaleDateString('id-ID', { day:'2-digit', month:'short', year:'numeric' });

    document.getElementById('upload-text').textContent = file.name.length > 24 ? file.name.substring(0,22)+'…' : file.name;
    document.getElementById('upload-info').textContent = formatNumber(rows.length) + ' rows loaded';
    document.getElementById('upload-area').style.borderColor = '#10b981';
    document.getElementById('upload-area').style.background = '#f0fdf4';

    populateFilters();
    renderDashboard();
    applyFilters();
  } catch (err) {
    console.error(err);
    alert('Gagal membaca file. Pastikan file Excel valid dan sheet bernama "data_dashboard".');
  } finally {
    overlay.style.display = 'none';
  }
}

// ===== DRAG & DROP =====
function initDragDrop() {
  const overlay = document.getElementById('drag-overlay');
  let dragCounter = 0;

  document.addEventListener('dragenter', (e) => {
    if (e.dataTransfer.types.includes('Files')) {
      dragCounter++;
      overlay.style.display = 'flex';
    }
  });
  document.addEventListener('dragleave', () => {
    dragCounter--;
    if (dragCounter === 0) overlay.style.display = 'none';
  });
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    dragCounter = 0;
    overlay.style.display = 'none';
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  document.getElementById('file-input').addEventListener('change', (e) => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });

  // Upload area drag
  const ua = document.getElementById('upload-area');
  ua.addEventListener('dragover', e => { e.preventDefault(); ua.classList.add('dragover'); });
  ua.addEventListener('dragleave', () => ua.classList.remove('dragover'));
}

// ===== INIT =====
document.addEventListener('DOMContentLoaded', () => {
  // Set current date
  document.getElementById('header-date').textContent = new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });

  // Init Select2 (empty)
  $('#filter-kategori').select2({ placeholder: 'Pilih Kategori', allowClear: true });
  $('#filter-part').select2({ placeholder: 'Pilih Part', allowClear: true });

  // Init drag & drop
  initDragDrop();
  populateFilters()

  // Init Lucide icons
  if (window.lucide) lucide.createIcons();
});
