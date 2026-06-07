'use strict';

// ===== STATE =====
const APP = {
  rawData: [],
  filteredData: [],
  charts: {},
  planTarget: {
    plan: [],
    target: []
  },
  ngTableSort: { col: 'ng_loss', dir: 'desc' },
  matTableSort: { col: 'usage_kg', dir: 'desc' },
  ngTablePage: 1,
  matTablePage: 1,
  PAGE_SIZE: 10,
  comparisonMode: 'nominal',
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

        const wb =
          XLSX.read(
            e.target.result,
            {
              type: 'binary',
              cellDates: false
            }
          );

        // =====================================
        // PLAN TARGET
        // =====================================

        APP.planTarget = {

          plan: Array(12).fill(0),

          target: Array(12).fill(0)

        };

        const planSheetName =
          wb.SheetNames.find(
            s =>
              s.toLowerCase()
                .replace(/\s/g, '') ===
              'plan_target'
          );

        if (planSheetName) {

          const planWs =
            wb.Sheets[
              planSheetName
            ];

          const planRows =
            XLSX.utils.sheet_to_json(
              planWs,
              {
                header: 1,
                raw: true
              }
            );

          // Baris PLAN CR 2026
          for (
            let i = 1;
            i <= 12;
            i++
          ) {

            APP.planTarget.plan[
              i - 1
            ] =
              (
                Number(
                  planRows?.[1]?.[i]
                ) || 0
              ) * 1000;

          }

          // Baris TARGET
          for (
            let i = 1;
            i <= 12;
            i++
          ) {

            APP.planTarget.target[
              i - 1
            ] =
              (
                Number(
                  planRows?.[2]?.[i]
                ) || 0
              ) * 1000;

          }

        }

        // =====================================
        // DATA DASHBOARD
        // =====================================

        const sheetName =
          wb.SheetNames.find(
            s =>
              s.toLowerCase()
                .replace(/\s/g, '') ===
              'data_dashboard'
          ) ||
          wb.SheetNames[0];

        const ws =
          wb.Sheets[
            sheetName
          ];

        const rows =
          XLSX.utils.sheet_to_json(
            ws,
            {
              defval: null
            }
          );

        resolve(rows);

      } catch (err) {

        reject(err);

      }

    };

    reader.onerror = reject;

    reader.readAsBinaryString(
      file
    );

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
    // periode key: "Apr 2026"
    get periode() {
      if (!this.bulan || !this.tahun) return null;
      return BULAN_NAMES[this.bulan - 1] + ' ' + this.tahun;
    },
    // Calculated fields
    // total_cost = (qty_g/1000) * harga * ok_prod
    get total_cost() { return (this.qty_g / 1000) * this.harga * this.ok_prod; },
    // ng_loss = (qty_g/1000) * harga * ng_prod
    get ng_loss() { return (this.qty_g / 1000) * this.harga * this.ng_prod; },
    // usage_kg = (qty_g * ok_prod) / 1_000_000  (qty_g in grams, result in kg)
    get usage_kg() { return (this.qty_g * this.ok_prod) / 1_000_000; },
  };
}

// ===== PROCESS DATA =====
function processData(rows) {
  return rows.map(normalizeRow);
}

// ===== APPLY FILTERS =====
function applyFilters() {
  const periodes = $('#filter-periode').val() || [];
  const scenarios = $('#filter-scenario').val() || [];
  const kategoris = $('#filter-kategori').val() || [];
  const parts = $('#filter-part').val() || [];
  const materials = $('#filter-material').val() || [];

  APP.filteredData = APP.rawData.filter(row => {
    if (periodes.length && !periodes.includes(row.periode)) return false;
    if (scenarios.length && !scenarios.includes(row.scenario)) return false;
    if (kategoris.length && !kategoris.includes(row.kategori)) return false;
    if (parts.length && !parts.includes(row.part_name)) return false;
    if (materials.length && !materials.includes(row.material)) return false;
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
          <div class="font-mono font-bold text-slate-800 leading-tight" style="font-size:15px;">${k.val}</div>
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
  let seriesData;

  if (isNominal) {
    seriesData = [
      { x: 'Cost Standard', y: Math.round(totalStd), fillColor: '#1d4ed8' },
      { x: 'Cost Alternative', y: Math.round(totalAlt), fillColor: '#0891b2' },
      { x: 'Total Saving', y: Math.round(saving), fillColor: saving >= 0 ? '#10b981' : '#ef4444' },
    ];
  } else {
    const pctAlt = totalStd > 0 ? (totalAlt / totalStd) * 100 : 0;
    const pctSaving = totalStd > 0 ? ((totalStd - totalAlt) / totalStd) * 100 : 0;
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

// ===== RENDER TOP PART CHART (Bar with negative values, ALL parts, scrollable) =====
function renderTopSavingChart() {

  const byPart = {};

  APP.filteredData.forEach(r => {

    if (!byPart[r.part_name]) {
      byPart[r.part_name] = {
        std: 0,
        alt: 0
      };
    }

    if (r.scenario === 'Standard')
      byPart[r.part_name].std += r.total_cost;

    if (r.scenario === 'Alternative')
      byPart[r.part_name].alt += r.total_cost;
  });

  const parts = Object.entries(byPart)
    .map(([name, v]) => ({
      name,
      saving: v.std - v.alt
    }))
    .sort((a, b) => b.saving - a.saving);

  const container =
    document.getElementById('chart-top-saving');

  if (parts.length === 0) {
    container.innerHTML =
      '<div class="empty-state"><div class="text-sm">Data belum tersedia</div></div>';
    return;
  }

  const totalSaving =
    parts.reduce((sum, p) => sum + p.saving, 0);

  const avgSaving =
    totalSaving / parts.length;

  const maxAbsSaving = Math.max(
    ...parts.map(p => Math.abs(p.saving)),
    1
  );

  container.innerHTML = `

    <div>

      <div class="flex items-start justify-between mb-4">

        <div>
          <div class="font-bold text-slate-800 text-base">
            All Part - Saving vs Loss
          </div>

          <div class="text-xs text-slate-500 mt-1">
            ${parts.length} Part
          </div>
        </div>

        <div class="text-right">

          <div class="font-mono font-bold text-base ${
            totalSaving >= 0
              ? 'text-emerald-600'
              : 'text-red-600'
          }">
            ${formatIDR(totalSaving)}
          </div>

          <div class="text-xs text-slate-400">
            Avg ${formatIDR(avgSaving)}
          </div>

        </div>

      </div>

      <div class="text-xs font-semibold text-slate-500 mb-2">
        Semua Part (${parts.length})
      </div>

      <div class="cat-part-list space-y-2">

        ${parts.map((p, i) => {

          const pct =
            Math.abs(p.saving) /
            maxAbsSaving * 100;

          const isSaving =
            p.saving >= 0;

          const rankClass =
            i === 0 ? 'rank-1' :
            i === 1 ? 'rank-2' :
            i === 2 ? 'rank-3' :
            'rank-n';

          return `

          <div>

            <div class="flex items-center justify-between mb-1">

              <div
                class="flex items-center gap-1.5"
                style="
                  flex:1;
                  min-width:0;
                ">

                <span class="rank-badge ${rankClass}">
                  ${i + 1}
                </span>

                <span
                  class="text-xs text-slate-700 font-medium cursor-pointer hover:text-blue-600 hover:underline"
                  style="
                    flex:1;
                    min-width:0;
                  "
                  title="Klik untuk melihat detail"
                  onclick="showPartDetail('${encodeURIComponent(p.name)}')">

                  ${p.name}

                </span>

              </div>

              <span
                class="text-xs font-mono font-semibold ${
                  isSaving
                    ? 'text-emerald-600'
                    : 'text-red-600'
                }">

                ${formatIDR(p.saving)}

              </span>

            </div>

            <div class="progress-bar">

              <div
                class="progress-bar-fill"
                style="
                  width:${pct}%;
                  background:${
                    isSaving
                      ? '#10b981'
                      : '#ef4444'
                  };
                ">
              </div>

            </div>

          </div>

          `;

        }).join('')}

      </div>

    </div>

  `;
}

function closePartDetail() {

  document
    .getElementById('part-detail-modal')
    .classList
    .add('hidden');

}

function showPartDetail(encodedPartName) {

    const partName =
        decodeURIComponent(encodedPartName);

    const rows =
        APP.filteredData.filter(
            x => x.part_name === partName
        );

    if (!rows.length)
        return;

    // =====================================
    // SUMMARY
    // =====================================

    const totalStd =
        rows
        .filter(r => r.scenario === 'Standard')
        .reduce(
            (sum, r) =>
                sum + (r.total_cost || 0),
            0
        );

    const totalAlt =
        rows
        .filter(r => r.scenario === 'Alternative')
        .reduce(
            (sum, r) =>
                sum + (r.total_cost || 0),
            0
        );

    const saving =
        totalStd - totalAlt;

    // =====================================
    // MATERIAL SUMMARY BY CATEGORY
    // =====================================

    const materialSummary = {};

    rows.forEach(r => {

        const kategori =
            r.kategori || '-';

        if (!materialSummary[kategori]) {

            materialSummary[kategori] = {

                standardMaterial: null,
                standardCost: 0,

                alternativeMaterials: [],
                alternativeCost: 0

            };

        }

        const materialName =
            (r.material || '')
            .toUpperCase();

        const isRecycleMaterial =
            materialName.includes('PELETIZING') ||
            materialName.includes('REGRIND');

        if (r.scenario === 'Standard') {

            // simpan material standard pertama yg valid
            if (
                !isRecycleMaterial &&
                !materialSummary[kategori].standardMaterial
            ) {

                materialSummary[kategori].standardMaterial =
                    r.material;

            }

            // REGRIND & PELETIZING tidak dihitung
            if (!isRecycleMaterial) {

                materialSummary[kategori].standardCost +=
                    r.total_cost || 0;

            }

        }

        if (r.scenario === 'Alternative') {

            if (!isRecycleMaterial) {

                materialSummary[kategori]
                    .alternativeMaterials
                    .push(r.material);

                materialSummary[kategori].alternativeCost +=
                    r.total_cost || 0;

            }

        }

    });

    // =====================================
    // RANKING CATEGORY
    // =====================================

    const rankedCategories =
        Object.entries(materialSummary)
        .map(([kategori, v]) => ({

            kategori,

            standardMaterial:
                v.standardMaterial,

            standardCost:
                v.standardCost,

            alternativeCost:
                v.alternativeCost,

            saving:
                v.standardCost -
                v.alternativeCost

        }))
        .sort(
            (a, b) =>
                b.saving - a.saving
        );

    const bestCategory =
        rankedCategories.length
            ? rankedCategories.reduce((best, current) => {

                if (!best)
                    return current;

                return Math.abs(current.saving) >
                    Math.abs(best.saving)
                    ? current
                    : best;

            }, null)
            : null;

    let html = '';

    // =====================================
    // KPI
    // =====================================

    html += `

    <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">

        <div class="kpi-card blue">

            <div class="text-xs text-slate-500">
                Total Standard
            </div>

            <div class="mt-2 font-bold text-2xl text-blue-700">
                ${formatIDR(totalStd)}
            </div>

        </div>

        <div class="kpi-card cyan">

            <div class="text-xs text-slate-500">
                Total Alternative
            </div>

            <div class="mt-2 font-bold text-2xl text-cyan-700">
                ${formatIDR(totalAlt)}
            </div>

        </div>

        <div class="kpi-card ${saving >= 0 ? 'green' : 'red'}">

            <div class="text-xs text-slate-500">
                Total Saving
            </div>

            <div class="
                mt-2
                font-bold
                text-2xl
                ${
                    saving >= 0
                    ? 'text-emerald-700'
                    : 'text-red-700'
                }">

                ${formatIDR(saving)}

            </div>

        </div>

    </div>

    `;

    // =====================================
    // INSIGHT + FORMULA
    // =====================================

    const insightPositive =
        bestCategory &&
        bestCategory.saving >= 0;

    const insightText =
        !bestCategory
            ? '-'
            : insightPositive
                ? 'Saving Terbesar'
                : 'Kerugian Terbesar';

    const insightBadge =
        !bestCategory
            ? 'badge-blue'
            : insightPositive
                ? 'badge-green'
                : 'badge-red';

    const insightValueClass =
        !bestCategory
            ? 'text-slate-700'
            : insightPositive
                ? 'text-emerald-600'
                : 'text-red-600';

    html += `

    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">

        <div
            class="
                glass-card
                p-4
                h-full
                flex
                flex-col
            ">

            <div class="section-header">

                Insight

            </div>

            <div
                class="
                    flex-1
                    flex
                    flex-col
                    justify-center
                ">

                <div class="mb-3">

                    <span class="badge ${insightBadge}">

                        ${insightText}

                    </span>

                </div>

                <div class="text-sm leading-relaxed">

                    ${
                        bestCategory
                        ? `
                        Kategori

                        <span class="font-semibold">

                            ${bestCategory.kategori}

                        </span>

                        menghasilkan

                        <span class="
                            font-bold
                            ${insightValueClass}
                        ">

                            ${formatIDR(bestCategory.saving)}

                        </span>
                        `
                        : '-'
                    }

                </div>

            </div>

        </div>

        <div
            class="
                glass-card
                p-4
                h-full
                flex
                flex-col
            ">

            <div class="section-header">

                Formula Perhitungan

            </div>

            <div
                class="
                    flex-1
                    flex
                    items-center
                ">

                <div
                    class="
                        text-sm
                        text-slate-700
                        leading-relaxed
                    ">

                    <div class="font-semibold mb-2">

                        Total Cost

                    </div>

                    <div>

                        (Qty_g / 1000)

                        × Harga

                        × OK_Prod

                    </div>

                </div>

            </div>

        </div>

    </div>

    `;

    // =====================================
    // SAVING PER MATERIAL
    // =====================================

    html += `

    <div class="glass-card p-4 mb-6">

        <div class="section-header">

            Saving Per Material

        </div>

        <div class="overflow-auto">

            <table class="data-table">

                <thead>

                    <tr>

                        <th>Kategori</th>
                        <th>Material Standard</th>
                        <th>Standard Cost</th>
                        <th>Alternative Cost</th>
                        <th>Saving</th>

                    </tr>

                </thead>

                <tbody>

    `;

    rankedCategories.forEach(v => {

        html += `

        <tr>

            <td>

                <span class="badge badge-blue">

                    ${v.kategori}

                </span>

            </td>

            <td>

                ${v.standardMaterial || '-'}

            </td>

            <td class="text-right">

                ${formatIDR(v.standardCost)}

            </td>

            <td class="text-right">

                ${formatIDR(v.alternativeCost)}

            </td>

            <td class="text-right">

                <span class="
                    badge
                    ${
                        v.saving >= 0
                        ? 'badge-green'
                        : 'badge-red'
                    }
                ">

                    ${formatIDR(v.saving)}

                </span>

            </td>

        </tr>

        `;

    });

    html += `

                </tbody>

            </table>

        </div>

    </div>

    `;

    // =====================================
    // DETAIL SOURCE DATA
    // =====================================

    html += `

    <div class="glass-card p-4">

        <div class="section-header">

            Detail Source Data

        </div>

        <div class="overflow-auto">

            <table class="data-table">

                <thead>

                    <tr>

                        <th>Scenario</th>
                        <th>Kategori</th>
                        <th>Material</th>
                        <th>Component</th>
                        <th>Qty</th>
                        <th>Harga</th>
                        <th>OK Prod</th>
                        <th>Total Cost</th>

                    </tr>

                </thead>

                <tbody>

    `;

    rows.forEach(r => {

        html += `

        <tr>

            <td>

                <span class="
                    badge
                    ${
                        r.scenario === 'Standard'
                        ? 'badge-blue'
                        : 'badge-yellow'
                    }
                ">
                    ${r.scenario}
                </span>

            </td>

            <td>

                <span class="badge badge-blue">

                    ${r.kategori || '-'}

                </span>

            </td>

            <td>${r.material}</td>

            <td>${r.component}</td>

            <td class="text-right">
                ${formatNumber(r.qty_g)}
            </td>

            <td class="text-right">
                ${formatIDR(r.harga)}
            </td>

            <td class="text-right">
                ${formatNumber(r.ok_prod)}
            </td>

            <td class="text-right font-semibold">
                ${formatIDR(r.total_cost)}
            </td>

        </tr>

        `;

    });

    html += `

                </tbody>

            </table>

        </div>

    </div>

    `;

    document.getElementById(
        'part-detail-title'
    ).textContent = partName;

    document.getElementById(
        'part-detail-subtitle'
    ).textContent =
        `${rows.length} Record`;

    document.getElementById(
        'part-detail-body'
    ).innerHTML = html;

    document.getElementById(
        'part-detail-modal'
    ).classList.remove('hidden');

}

// ===== RENDER TOP MATERIAL CHART =====
function renderTopMaterialChart() {

    const materialSaving = {};

    // =====================================
    // GROUP PART + KATEGORI
    // =====================================

    const partKategoriMap = {};

    APP.filteredData.forEach(r => {

        const part =
            r.part_name || '-';

        const kategori =
            r.kategori || '-';

        if (!partKategoriMap[part]) {

            partKategoriMap[part] = {};

        }

        if (!partKategoriMap[part][kategori]) {

            partKategoriMap[part][kategori] = {

                standard: [],
                alternative: []

            };

        }

        if (r.scenario === 'Standard') {

            partKategoriMap[part][kategori]
                .standard.push(r);

        }

        if (r.scenario === 'Alternative') {

            partKategoriMap[part][kategori]
                .alternative.push(r);

        }

    });

    // =====================================
    // HITUNG PER PART + KATEGORI
    // =====================================

    Object.entries(partKategoriMap)
        .forEach(([partName, kategoriMap]) => {

        Object.entries(kategoriMap)
            .forEach(([kategori, data]) => {

            if (!data.standard.length) {

                return;

            }

            // ============================
            // MATERIAL STANDARD UTAMA
            // ============================

            const mainMaterialRow =
                data.standard.find(x => {

                    const name =
                        (x.material || '')
                        .toUpperCase();

                    return (
                        !name.includes('REGRIND') &&
                        !name.includes('PELETIZING')
                    );

                });

            if (!mainMaterialRow) {

                return;

            }

            const mainMaterial =
                mainMaterialRow.material;

            // ============================
            // COST STANDARD
            // ============================

            let standardCost = 0;

            data.standard.forEach(r => {

                const qty =
                    Number(r.qty_g || 0);

                const harga =
                    Number(r.harga || 0);

                const qtyProd =
                    Number(r.qty_prod || 0);

                standardCost +=
                    qty *
                    harga *
                    qtyProd /
                    1000;

            });

            // ============================
            // COST ALTERNATIVE
            // ============================

            let alternativeCost = 0;

            data.alternative.forEach(r => {

                const qty =
                    Number(r.qty_g || 0);

                const harga =
                    Number(r.harga || 0);

                const qtyProd =
                    Number(r.qty_prod || 0);

                alternativeCost +=
                    qty *
                    harga *
                    qtyProd /
                    1000;

            });

            // ============================
            // SAVING
            // ============================

            const saving =
                standardCost -
                alternativeCost;

            const key =
                `${kategori}|${mainMaterial}`;

            if (!materialSaving[key]) {

                materialSaving[key] = {

                    kategori,

                    material:
                        mainMaterial,

                    saving: 0,

                    partCount: 0

                };

            }

            materialSaving[key].saving +=
                saving;

            materialSaving[key].partCount++;

        });

    });

    // =====================================
    // SORT
    // =====================================

    const materials =
        Object.values(materialSaving)
        .sort(
            (a, b) =>
                b.saving - a.saving
        );

    const container =
        document.getElementById(
            'chart-top-material'
        );

    if (!materials.length) {

        container.innerHTML = `
            <div class="empty-state">
                <div class="text-sm">
                    Data belum tersedia
                </div>
            </div>
        `;

        return;

    }

    const totalSaving =
        materials.reduce(
            (sum, m) =>
                sum + m.saving,
            0
        );

    const avgSaving =
        totalSaving /
        materials.length;

    const maxAbsSaving =
        Math.max(
            ...materials.map(
                m => Math.abs(m.saving)
            ),
            1
        );

    // =====================================
    // RENDER
    // =====================================

    container.innerHTML = `

        <div>

            <div
                class="
                    flex
                    items-start
                    justify-between
                    mb-4
                ">

                <div>

                    <div
                        class="
                            font-bold
                            text-slate-800
                            text-base
                        ">

                        All Material - Saving vs Loss

                    </div>

                    <div
                        class="
                            text-xs
                            text-slate-500
                            mt-1
                        ">

                        ${materials.length} Material

                    </div>

                </div>

                <div class="text-right">

                    <div
                        class="
                            font-mono
                            font-bold
                            text-base
                            ${
                                totalSaving >= 0
                                    ? 'text-emerald-600'
                                    : 'text-red-600'
                            }
                        ">

                        ${formatIDR(totalSaving)}

                    </div>

                    <div
                        class="
                            text-xs
                            text-slate-400
                        ">

                        Avg ${formatIDR(avgSaving)}

                    </div>

                </div>

            </div>

            <div
                class="
                    text-xs
                    font-semibold
                    text-slate-500
                    mb-2
                ">

                Semua Material
                (${materials.length})

            </div>

            <div
                class="
                    cat-part-list
                    space-y-2
                ">

                ${materials.map((m, i) => {

                    const pct =
                        Math.abs(m.saving) /
                        maxAbsSaving *
                        100;

                    const isSaving =
                        m.saving >= 0;

                    const rankClass =
                        i === 0
                            ? 'rank-1'
                            : i === 1
                            ? 'rank-2'
                            : i === 2
                            ? 'rank-3'
                            : 'rank-n';

                    return `

                    <div>

                        <div
                            class="
                                flex
                                items-center
                                justify-between
                                mb-1
                            ">

                            <div
                                class="
                                    flex
                                    items-center
                                    gap-1.5
                                "
                                style="
                                    flex:1;
                                    min-width:0;
                                ">

                                <span
                                    class="
                                        rank-badge
                                        ${rankClass}
                                    ">

                                    ${i + 1}

                                </span>

                                <span
                                    class="
                                        text-xs
                                        text-slate-700
                                        font-medium
                                        cursor-pointer
                                        hover:text-blue-600
                                        hover:underline
                                    "
                                    onclick="showMaterialDetail(
                                        '${encodeURIComponent(m.material)}',
                                        '${encodeURIComponent(m.kategori)}'
                                    )"
                                    style="
                                        flex:1;
                                        min-width:0;
                                        overflow:hidden;
                                        white-space:nowrap;
                                        text-overflow:ellipsis;
                                    "
                                    title="${m.material}">

                                    [${m.kategori}]
                                    ${m.material}

                                </span>

                            </div>

                            <span
                                class="
                                    text-xs
                                    font-mono
                                    font-semibold
                                    ${
                                        isSaving
                                            ? 'text-emerald-600'
                                            : 'text-red-600'
                                    }
                                ">

                                ${formatIDR(m.saving)}

                            </span>

                        </div>

                        <div class="progress-bar">

                            <div
                                class="progress-bar-fill"
                                style="
                                    width:${pct}%;
                                    background:${
                                        isSaving
                                            ? '#10b981'
                                            : '#ef4444'
                                    };
                                ">
                            </div>

                        </div>

                    </div>

                    `;

                }).join('')}

            </div>

        </div>

    `;

}

function showMaterialDetail(
    encodedMaterial,
    encodedKategori
) {

    const material =
        decodeURIComponent(
            encodedMaterial
        );

    const kategori =
        decodeURIComponent(
            encodedKategori
        );

    // =====================================
    // GROUP PART + KATEGORI
    // =====================================

    const partKategoriMap = {};

    APP.filteredData.forEach(r => {

        const part =
            r.part_name || '-';

        if (
            r.kategori !== kategori
        ) {
            return;
        }

        if (!partKategoriMap[part]) {

            partKategoriMap[part] = {

                standard: [],
                alternative: []

            };

        }

        if (
            r.scenario === 'Standard'
        ) {

            partKategoriMap[part]
                .standard.push(r);

        }

        if (
            r.scenario === 'Alternative'
        ) {

            partKategoriMap[part]
                .alternative.push(r);

        }

    });

    // =====================================
    // ANALISA PART
    // =====================================

    const detailRows = [];

    Object.entries(
        partKategoriMap
    ).forEach(([partName, obj]) => {

        if (
            !obj.standard.length
        ) {
            return;
        }

        const mainMaterialRow =
            obj.standard.find(x => {

                const name =
                    (x.material || '')
                    .toUpperCase();

                return (
                    !name.includes('REGRIND') &&
                    !name.includes('PELETIZING')
                );

            });

        if (!mainMaterialRow) {

            return;

        }

        // hanya material yg sedang dibuka
        if (
            mainMaterialRow.material !==
            material
        ) {

            return;

        }

        // ==========================
        // COST STANDARD
        // ==========================

        let standardCost = 0;

        obj.standard.forEach(r => {

            standardCost +=
                (
                    Number(r.qty_g || 0)
                    *
                    Number(r.harga || 0)
                    *
                    Number(r.qty_prod || 0)
                ) / 1000;

        });

        // ==========================
        // COST ALTERNATIVE
        // ==========================

        let alternativeCost = 0;

        obj.alternative.forEach(r => {

            alternativeCost +=
                (
                    Number(r.qty_g || 0)
                    *
                    Number(r.harga || 0)
                    *
                    Number(r.qty_prod || 0)
                ) / 1000;

        });

        const saving =
            standardCost -
            alternativeCost;

        detailRows.push({

            partName,

            standardCost,

            alternativeCost,

            saving,

            standardMaterials:
                obj.standard
                .map(x => x.material),

            alternativeMaterials:
                obj.alternative
                .map(x => x.material)

        });

    });

    detailRows.sort(
        (a, b) =>
            b.saving - a.saving
    );

    const totalSaving =
        detailRows.reduce(
            (sum, x) =>
                sum + x.saving,
            0
        );

    const bestPart =
        detailRows[0];

    const worstPart =
        [...detailRows]
        .sort(
            (a, b) =>
                a.saving - b.saving
        )[0];

    // =====================================
    // ALTERNATIVE MATERIAL SUMMARY
    // =====================================

    const alternativeMap = {};

    detailRows.forEach(r => {

        r.alternativeMaterials
            .forEach(mat => {

                if (
                    !alternativeMap[mat]
                ) {

                    alternativeMap[mat] = {

                        material: mat,

                        saving: 0,

                        count: 0

                    };

                }

                alternativeMap[mat]
                    .saving +=
                    r.saving;

                alternativeMap[mat]
                    .count++;

            });

    });

    const alternatives =
        Object.values(
            alternativeMap
        ).sort(
            (a, b) =>
                b.saving - a.saving
        );

    // =====================================
    // HTML
    // =====================================

    let html = '';

    html += `

    <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">

        <div class="kpi-card">

            <div class="text-xs text-slate-500">
                Kategori
            </div>

            <div class="mt-2">

                <span class="badge badge-blue">
                    ${kategori}
                </span>

            </div>

        </div>

        <div class="kpi-card">

            <div class="text-xs text-slate-500">
                Material Standard
            </div>

            <div class="mt-2 text-sm font-semibold">
                ${material}
            </div>

        </div>

        <div class="kpi-card">

            <div class="text-xs text-slate-500">
                Jumlah Part
            </div>

            <div class="mt-2 text-2xl font-bold">
                ${detailRows.length}
            </div>

        </div>

        <div class="kpi-card ${totalSaving >= 0 ? 'green' : 'red'}">

            <div class="text-xs text-slate-500">
                Total Saving
            </div>

            <div class="mt-2 text-2xl font-bold ${totalSaving >= 0 ? 'text-emerald-700' : 'text-red-700'}">
                ${formatIDR(totalSaving)}
            </div>

        </div>

    </div>

    `;

    html += `

    <div class="grid md:grid-cols-2 gap-4 mb-6">

        <div class="glass-card p-4">

            <div class="section-header">
                Saving Terbesar
            </div>

            <div class="mt-2">

                <div class="font-semibold">
                    ${bestPart?.partName || '-'}
                </div>

                <div class="text-emerald-600 font-bold mt-1">
                    ${bestPart ? formatIDR(bestPart.saving) : '-'}
                </div>

            </div>

        </div>

        <div class="glass-card p-4">

            <div class="section-header">
                Kerugian Terbesar
            </div>

            <div class="mt-2">

                <div class="font-semibold">
                    ${worstPart?.partName || '-'}
                </div>

                <div class="text-red-600 font-bold mt-1">
                    ${worstPart ? formatIDR(worstPart.saving) : '-'}
                </div>

            </div>

        </div>

    </div>

    `;

    html += `

    <div class="glass-card p-4 mb-6">

        <div class="section-header">
            Part Impact Analysis
        </div>

        <table class="data-table">

            <thead>

                <tr>

                    <th>Part</th>
                    <th>Standard</th>
                    <th>Alternative</th>
                    <th>Saving</th>

                </tr>

            </thead>

            <tbody>

                ${detailRows.map(r => `

                <tr>

                    <td>

                        <span
                            class="cursor-pointer hover:text-blue-600 hover:underline"
                            onclick="showPartDetail('${encodeURIComponent(r.partName)}')">

                            ${r.partName}

                        </span>

                    </td>

                    <td>

                        ${r.standardMaterials.join('<br>')}

                    </td>

                    <td>

                        ${r.alternativeMaterials.join('<br>')}

                    </td>

                    <td class="text-right">

                        <span class="badge ${r.saving >= 0 ? 'badge-green' : 'badge-red'}">

                            ${formatIDR(r.saving)}

                        </span>

                    </td>

                </tr>

                `).join('')}

            </tbody>

        </table>

    </div>

    `;

    html += `

    <div class="glass-card p-4">

        <div class="section-header">

            Alternative Material Summary

        </div>

        <table class="data-table">

            <thead>

                <tr>

                    <th>Alternative Material</th>
                    <th>Part</th>
                    <th>Total Impact</th>

                </tr>

            </thead>

            <tbody>

                ${alternatives.map(a => `

                <tr>

                    <td>
                        ${a.material}
                    </td>

                    <td>
                        ${a.count}
                    </td>

                    <td class="text-right">
                        ${formatIDR(a.saving)}
                    </td>

                </tr>

                `).join('')}

            </tbody>

        </table>

    </div>

    `;

    document.getElementById(
        'part-detail-title'
    ).textContent =
        material;

    document.getElementById(
        'part-detail-subtitle'
    ).textContent =
        `Kategori ${kategori}`;

    document.getElementById(
        'part-detail-body'
    ).innerHTML =
        html;

    document.getElementById(
        'part-detail-modal'
    ).classList.remove(
        'hidden'
    );

}

// ===== RENDER TOP MATERIAL CHART =====
function renderMonthlySavingChart() {

    const categories = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'Mei',
        'Jun',
        'Jul',
        'Ags',
        'Sep',
        'Okt',
        'Nov',
        'Des'
    ];

    // =====================================
    // KATEGORI
    // =====================================

    const allKategori = [
        ...new Set(
            APP.filteredData.map(
                r => r.kategori || '-'
            )
        )
    ].sort();

    const monthlyKategori = {};

    allKategori.forEach(k => {

        monthlyKategori[k] =
            Array(12).fill(0);

    });

    // =====================================
    // GROUP PART + KATEGORI + BULAN
    // =====================================

    const partKategoriMap = {};

    APP.filteredData.forEach(r => {

        const part =
            r.part_name || '-';

        const kategori =
            r.kategori || '-';

        const bulan =
            r.bulan;

        if (!bulan)
            return;

        const key =
            `${part}|${kategori}|${bulan}`;

        if (!partKategoriMap[key]) {

            partKategoriMap[key] = {

                bulan,

                kategori,

                standard: [],

                alternative: []

            };

        }

        if (
            r.scenario ===
            'Standard'
        ) {

            partKategoriMap[key]
                .standard
                .push(r);

        }

        if (
            r.scenario ===
            'Alternative'
        ) {

            partKategoriMap[key]
                .alternative
                .push(r);

        }

    });

    // =====================================
    // HITUNG SAVING
    // =====================================

    Object.values(
        partKategoriMap
    ).forEach(group => {

        const standardCost =
            group.standard.reduce(
                (sum, r) =>
                    sum + r.total_cost,
                0
            );

        const alternativeCost =
            group.alternative.reduce(
                (sum, r) =>
                    sum + r.total_cost,
                0
            );

        const saving =
            standardCost -
            alternativeCost;

        monthlyKategori[
            group.kategori
        ][
            group.bulan - 1
        ] += saving;

    });

    // =====================================
    // PLAN & TARGET
    // =====================================

    const plan =
        APP.planTarget?.plan ||
        Array(12).fill(0);

    const target =
        APP.planTarget?.target ||
        Array(12).fill(0);

    // =====================================
    // TOTAL ACTUAL
    // =====================================

    const actualTotal =
        Array(12).fill(0);

    Object.values(
        monthlyKategori
    ).forEach(arr => {

        arr.forEach(
            (v, i) => {

                actualTotal[i] += v;

            }
        );

    });

    // =====================================
    // SERIES
    // =====================================

    const series = [

        {
            name: 'Plan',
            type: 'column',
            data: plan,
            group: 'plan'
        }

    ];

    allKategori.forEach(k => {

        series.push({

            name: k,

            type: 'column',

            data:
                monthlyKategori[k],

            group: 'actual'

        });

    });

    series.push({

        name: 'Target',

        type: 'line',

        data: target

    });

    // =====================================
    // DESTROY CHART
    // =====================================

    if (
        APP.charts.monthlySaving
    ) {

        APP.charts.monthlySaving
            .destroy();

    }

    const options = {

        chart: {

            height: 360,

            type: 'line',

            stacked: true,

            toolbar: {

                show: false

            }

        },

        series,

        stroke: {

            width:
                series.map(
                    s =>
                        s.name === 'Target'
                        ? 4
                        : 0
                ),

            curve: 'smooth'

        },

        plotOptions: {

            bar: {

                columnWidth: '55%'

            }

        },

        dataLabels: {

            enabled: false

        },

        xaxis: {

            categories

        },

        yaxis: {

            labels: {

                formatter: function(v) {

                    return formatIDR(
                        v,
                        false
                    );

                }

            }

        },

        tooltip: {

            shared: true,

            intersect: false,

            custom: function({

            dataPointIndex

        }) {

            const actual =
                actualTotal[
                    dataPointIndex
                ];

            const planVal =
                plan[
                    dataPointIndex
                ];

            const targetVal =
                target[
                    dataPointIndex
                ];

            const achievement =
                targetVal > 0
                ? actual / targetVal * 100
                : 0;

            const achColor =
                actual >= targetVal
                ? '#10b981'
                : '#ef4444';

            let html = `

            <div
                style="
                    min-width:220px;
                    background:white;
                    border-radius:10px;
                    overflow:hidden;
                    box-shadow:
                        0 6px 18px rgba(0,0,0,.12);
                    font-size:11px;
                ">

                <div
                    style="
                        background:#1e293b;
                        color:white;
                        padding:8px 10px;
                        font-weight:700;
                        font-size:12px;
                    ">

                    ${categories[dataPointIndex]}

                </div>

                <div
                    style="
                        padding:8px 10px;
                    ">

                    <div style="display:flex;justify-content:space-between;">
                        <span>Plan</span>
                        <b>${formatIDR(planVal)}</b>
                    </div>

                    <div style="display:flex;justify-content:space-between;">
                        <span>Actual</span>
                        <b style="color:#10b981;">
                            ${formatIDR(actual)}
                        </b>
                    </div>

                    <div style="display:flex;justify-content:space-between;">
                        <span>Target</span>
                        <b style="color:#dc2626;">
                            ${formatIDR(targetVal)}
                        </b>
                    </div>

                    <div
                        style="
                            display:flex;
                            justify-content:space-between;
                            margin-top:4px;
                            padding-top:4px;
                            border-top:1px solid #e2e8f0;
                        ">

                        <span>Achv.</span>

                        <b
                            style="
                                color:${achColor};
                            ">

                            ${achievement.toFixed(1)}%

                        </b>

                    </div>

                    <div
                        style="
                            margin-top:6px;
                            padding-top:6px;
                            border-top:1px solid #e2e8f0;
                        ">

            `;

            Object.keys(
                monthlyKategori
            ).forEach(k => {

                const value =
                    monthlyKategori[k][
                        dataPointIndex
                    ] || 0;

                if (!value)
                    return;

                html += `

                <div
                    style="
                        display:flex;
                        justify-content:space-between;
                        margin-top:2px;
                    ">

                    <span>
                        ${k}
                    </span>

                    <span>
                        ${formatIDR(value)}
                    </span>

                </div>

                `;

            });

            html += `
                    </div>
                </div>
            </div>
            `;

            return html;

        }

        },

        colors: [

            '#f59e0b', // PLAN

            ...CATEGORY_COLORS,

            '#dc2626' // TARGET

        ],

        legend: {

            position: 'top'

        }

    };

    const container =
        document.getElementById(
            'chart-monthly-saving'
        );

    container.innerHTML = '';

    APP.charts.monthlySaving =
        new ApexCharts(
            container,
            options
        );

    APP.charts.monthlySaving
        .render();

}

// ===== RENDER CATEGORY CARDS =====
function renderCategoryCards() {
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

  // Sort categories by abs, pp, ppgf
  const priorityOrder = {
    ABS: 1,
    PP: 2,
    PPGF: 3
  };

  const sortedKat =
    Object.entries(byKat)
    .sort((a, b) => {

      const orderA =
        priorityOrder[a[0]] || 999;

      const orderB =
        priorityOrder[b[0]] || 999;

      if (orderA !== orderB)
        return orderA - orderB;

      const savA =
        a[1].std - a[1].alt;

      const savB =
        b[1].std - b[1].alt;

      return savB - savA;

    });

  container.innerHTML = sortedKat.map(([kat, val], idx) => {
    const saving = val.std - val.alt;
    const partCount = val.parts.size;
    const avgSaving = partCount > 0 ? saving / partCount : 0;
    const isSav = saving >= 0;
    const color = CATEGORY_COLORS[idx % CATEGORY_COLORS.length];

    // All parts sorted by saving descending (highest saving first, loss at bottom)
    const partSaving = {};
    val.rows.forEach(r => {
      if (!partSaving[r.part_name]) partSaving[r.part_name] = { std: 0, alt: 0 };
      if (r.scenario === 'Standard') partSaving[r.part_name].std += r.total_cost;
      if (r.scenario === 'Alternative') partSaving[r.part_name].alt += r.total_cost;
    });
    const allParts = Object.entries(partSaving)
      .map(([name, v]) => ({ name, saving: v.std - v.alt }))
      .sort((a, b) => b.saving - a.saving);

    const maxAbsSaving = Math.max(...allParts.map(p => Math.abs(p.saving)), 1);

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

          <div class="mb-2">
            <div class="text-xs font-semibold text-slate-500 mb-2">Semua Part (${allParts.length})</div>
            <div class="cat-part-list space-y-2">
              ${allParts.map((p, i) => {
                const pct = Math.abs(p.saving) / maxAbsSaving * 100;
                const pSav = p.saving >= 0;
                const rankClass = i === 0 ? 'rank-1' : i === 1 ? 'rank-2' : i === 2 ? 'rank-3' : 'rank-n';
                return `
                  <div>
                    <div class="flex items-center justify-between mb-1">
                      <div
                        class="flex items-center gap-1.5"
                        style="
                          flex:1;
                          min-width:0;
                          overflow:hidden;
                        ">

                        <span class="rank-badge ${rankClass}">
                          ${i+1}
                        </span>

                        <span
                          class="
                            text-xs
                            text-slate-700
                            font-medium
                            cursor-pointer
                            hover:text-blue-600
                            hover:underline
                          "
                          style="
                            flex:1;
                            min-width:0;
                            overflow:hidden;
                            white-space:nowrap;
                            text-overflow:ellipsis;
                            display:block;
                          "
                          title="${p.name}"
                          onclick="showPartDetailKategori(
                              '${encodeURIComponent(p.name)}',
                              '${encodeURIComponent(kat)}'
                          )">

                          ${p.name}

                        </span>

                      </div>
                      <span class="text-xs font-mono font-semibold ${pSav ? 'text-emerald-600' : 'text-red-600'}" style="margin-left:8px; white-space:nowrap;">${formatIDR(p.saving)}</span>
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

function showPartDetailKategori(
    encodedPartName,
    encodedKategori
) {

    const partName =
        decodeURIComponent(
            encodedPartName
        );

    const kategori =
        decodeURIComponent(
            encodedKategori
        );

    const rows =
        APP.filteredData.filter(
            x =>
                x.part_name === partName &&
                (x.kategori || '-') === kategori
        );

    if (!rows.length)
        return;

    // =====================================
    // SUMMARY
    // =====================================

    const totalStd =
        rows
        .filter(
            r => r.scenario === 'Standard'
        )
        .reduce(
            (sum, r) =>
                sum + (r.total_cost || 0),
            0
        );

    const totalAlt =
        rows
        .filter(
            r => r.scenario === 'Alternative'
        )
        .reduce(
            (sum, r) =>
                sum + (r.total_cost || 0),
            0
        );

    const saving =
        totalStd - totalAlt;

    // Standard dulu
    rows.sort((a, b) => {

        if (
            a.scenario === b.scenario
        )
            return 0;

        return a.scenario === 'Standard'
            ? -1
            : 1;

    });

    let html = '';

    // =====================================
    // KPI
    // =====================================

    html += `

    <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">

        <div class="kpi-card">

            <div class="text-xs text-slate-500">
                Kategori Material
            </div>

            <div class="mt-2">

                <span class="
                    badge
                    badge-blue
                    text-sm
                ">

                    ${kategori}

                </span>

            </div>

        </div>

        <div class="kpi-card blue">

            <div class="text-xs text-slate-500">
                Total Standard
            </div>

            <div class="mt-2 font-bold text-2xl text-blue-700">

                ${formatIDR(totalStd)}

            </div>

        </div>

        <div class="kpi-card cyan">

            <div class="text-xs text-slate-500">
                Total Alternative
            </div>

            <div class="mt-2 font-bold text-2xl text-cyan-700">

                ${formatIDR(totalAlt)}

            </div>

        </div>

        <div class="kpi-card ${
            saving >= 0
                ? 'green'
                : 'red'
        }">

            <div class="text-xs text-slate-500">
                Total Saving
            </div>

            <div
                class="
                    mt-2
                    font-bold
                    text-2xl
                    ${
                        saving >= 0
                            ? 'text-emerald-700'
                            : 'text-red-700'
                    }
                ">

                ${formatIDR(saving)}

            </div>

        </div>

    </div>

    `;

    // =====================================
    // FORMULA
    // =====================================

    html += `

    <div class="glass-card p-3 mb-6">

        <div class="section-header">

            Formula Perhitungan

        </div>

        <div
            class="
                text-sm
                text-slate-700
            ">

            Total Cost =
            (Qty_g / 1000)
            × Harga
            × OK_Prod

        </div>

    </div>

    `;

    // =====================================
    // DETAIL SOURCE DATA
    // =====================================

    html += `

    <div class="glass-card p-4">

        <div class="section-header">

            Detail Source Data

        </div>

        <div class="overflow-auto">

            <table class="data-table">

                <thead>

                    <tr>

                        <th>No</th>
                        <th>Scenario</th>
                        <th>Material</th>
                        <th>Component</th>
                        <th>Qty</th>
                        <th>Harga</th>
                        <th>OK Prod</th>
                        <th>Total Cost</th>

                    </tr>

                </thead>

                <tbody>

    `;

    rows.forEach((r, idx) => {

        html += `

        <tr>

            <td class="text-center">

                ${idx + 1}

            </td>

            <td>

                <span class="
                    badge
                    ${
                        r.scenario === 'Standard'
                            ? 'badge-blue'
                            : 'badge-yellow'
                    }
                ">

                    ${r.scenario}

                </span>

            </td>

            <td>

                ${r.material}

            </td>

            <td>

                ${r.component}

            </td>

            <td class="text-right">

                ${formatNumber(
                    r.qty_g
                )}

            </td>

            <td class="text-right">

                ${formatIDR(
                    r.harga
                )}

            </td>

            <td class="text-right">

                ${formatNumber(
                    r.ok_prod
                )}

            </td>

            <td
                class="
                    text-right
                    font-semibold
                ">

                ${formatIDR(
                    r.total_cost
                )}

            </td>

        </tr>

        `;

    });

    html += `

                </tbody>

            </table>

        </div>

    </div>

    `;

    document.getElementById(
        'part-detail-title'
    ).textContent =
        partName;

    document.getElementById(
        'part-detail-subtitle'
    ).textContent =
        `Kategori ${kategori} • ${rows.length} Record`;

    document.getElementById(
        'part-detail-body'
    ).innerHTML =
        html;

    document.getElementById(
        'part-detail-modal'
    ).classList.remove(
        'hidden'
    );

}

// ===== RENDER NG CHARTS =====
function renderNGCharts() {
  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(r => r.scenario === 'Alternative');

  // Group by part
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
      ng_rate: v.qty_prod > 0 ? (v.ng_prod / v.qty_prod) * 100 : 0
    }))
    .sort((a, b) => b.ng_loss - a.ng_loss)
    .slice(0, 10);

  // Top 10 NG Part Bar
  destroyChart('ngTop');
  APP.charts.ngTop = new ApexCharts(
    document.getElementById('chart-ng-top'),
    {
      series: [{ name: 'NG Loss', data: parts.map(p => Math.round(p.ng_loss)) }],
      chart: {
        type: 'bar',
        height: 320,
        toolbar: { show: false },
        fontFamily: 'Plus Jakarta Sans, sans-serif',
        animations: { speed: 500 }
      },
      plotOptions: { bar: { horizontal: true, borderRadius: 6, barHeight: '60%' } },
      colors: ['#ef4444'],
      dataLabels: {
        enabled: true,
        formatter: v => formatIDR(v),
        style: { fontSize: '11px', colors: ['#fff'], fontWeight: '600' }
      },
      tooltip: { y: { formatter: v => formatIDR(v, true) } },
      xaxis: {
        categories: parts.map(p => p.part_name.length > 22 ? p.part_name.substring(0, 20) + '…' : p.part_name),
        labels: { formatter: v => formatIDR(v), style: { fontSize: '10px', colors: '#94a3b8' } },
        axisBorder: { show: false },
        axisTicks: { show: false },
      },
      yaxis: { labels: { style: { fontSize: '11px', fontWeight: '600', colors: '#475569' } } },
      grid: { borderColor: '#f1f5f9', strokeDashArray: 3 },
    }
  );
  APP.charts.ngTop.render();

  // NG Loss by Kategori Donut
  const byKat = {};
  altData.forEach(r => {
    byKat[r.kategori] = (byKat[r.kategori] || 0) + r.ng_loss;
  });
  const katLabels = Object.keys(byKat);
  const katVals = Object.values(byKat).map(v => Math.round(v));

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
        formatter: (v, o) => o.w.globals.labels[o.seriesIndex] + '\n' + v.toFixed(1) + '%'
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
                formatter: () => formatIDR(katVals.reduce((a, b) => a + b, 0))
              }
            }
          }
        }
      },
      tooltip: { y: { formatter: v => formatIDR(v, true) } },
      legend: { position: 'bottom', fontSize: '11px' },
    }
  );
  APP.charts.ngDonut.render();

  // NG Rate Gauge
  const totalNG = altData.reduce((s, r) => s + r.ng_prod, 0);
  const totalQty = altData.reduce((s, r) => s + r.qty_prod, 0);
  const ngRate = totalQty > 0 ? (totalNG / totalQty) * 100 : 0;

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
          hollow: { size: '60%' },
          dataLabels: {
            name: { show: true, fontSize: '13px', color: '#64748b', offsetY: -6 },
            value: { fontSize: '24px', fontWeight: '700', color: '#1e293b', offsetY: 8, formatter: v => v + '%' },
          },
          track: { background: '#f1f5f9', strokeWidth: '97%' },
        }
      },
      fill: {
        type: 'gradient',
        gradient: {
          shade: 'light',
          type: 'horizontal',
          gradientToColors: ngRate < 3 ? ['#10b981'] : ngRate < 8 ? ['#f59e0b'] : ['#ef4444'],
          stops: [0, 100]
        }
      },
      colors: ngRate < 3 ? ['#34d399'] : ngRate < 8 ? ['#fbbf24'] : ['#f87171'],
      labels: ['NG Rate'],
    }
  );
  APP.charts.ngGauge.render();
  document.getElementById('ng-rate-label').textContent = ngRate.toFixed(2) + '%';

  // Worst Material
  const byMat = {};
  altData.forEach(r => {
    if (!byMat[r.material]) byMat[r.material] = { ng: 0, loss: 0 };
    byMat[r.material].ng += r.ng_prod;
    byMat[r.material].loss += r.ng_loss;
  });
  const worstMats = Object.entries(byMat)
    .map(([mat, v]) => ({ mat, ...v }))
    .sort((a, b) => b.loss - a.loss)
    .slice(0, 5);

  const wc = document.getElementById('worst-material-list');
  if (worstMats.length === 0) {
    wc.innerHTML = `<div class="empty-state" style="padding: 30px 20px;"><div class="text-sm">Data kosong</div></div>`;
  } else {
    wc.innerHTML = worstMats.map((m, i) => `
      <div class="flex items-center justify-between p-3 rounded-xl"
           style="background: ${i === 0 ? '#fef2f2' : '#f8fafc'}; border: 1px solid ${i === 0 ? '#fecaca' : '#e2e8f0'};">
        <div class="flex items-center gap-2">
          <div class="rank-badge ${['rank-1','rank-2','rank-3','rank-n','rank-n'][i]}">${i + 1}</div>
          <div>
            <div class="text-sm font-semibold text-slate-700">${m.mat || '—'}</div>
            <div class="text-xs text-slate-400">NG: ${formatNumber(m.ng)}</div>
          </div>
        </div>
        <div class="text-right">
          <div class="text-sm font-bold font-mono text-red-600">${formatIDR(m.loss)}</div>
          <div class="text-xs text-slate-400">NG Loss</div>
        </div>
      </div>
    `).join('');
  }
}

// ===== RENDER NG TABLE =====
let ngTableData = [];

function renderNGTable() {
  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(r => r.scenario === 'Alternative');

  const byPart = {};
  altData.forEach(r => {
    const key = r.part_name + '|' + r.material;
    if (!byPart[key]) {
      byPart[key] = { part_name: r.part_name, material: r.material, qty_prod: 0, ng_prod: 0, ng_loss: 0 };
    }
    byPart[key].qty_prod += r.qty_prod;
    byPart[key].ng_prod += r.ng_prod;
    byPart[key].ng_loss += r.ng_loss;
  });

  ngTableData = Object.values(byPart).map(r => ({
    ...r,
    ng_rate: r.qty_prod > 0 ? (r.ng_prod / r.qty_prod) * 100 : 0
  }));

  renderNGTablePage(1);
}

function renderNGTablePage(page) {
  APP.ngTablePage = page;
  const search = (document.getElementById('ng-search')?.value || '').toLowerCase();
  const { col, dir } = APP.ngTableSort;

  let data = ngTableData.filter(r =>
    r.part_name.toLowerCase().includes(search) ||
    r.material.toLowerCase().includes(search)
  );

  data.sort((a, b) => {
    const av = a[col], bv = b[col];
    return dir === 'asc' ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
  });

  const total = data.length;
  const start = (page - 1) * APP.PAGE_SIZE;
  const slice = data.slice(start, start + APP.PAGE_SIZE);

  const tbody = document.getElementById('ng-table-body');
  if (slice.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-8 text-slate-400">Tidak ada data</td></tr>`;
  } else {
    tbody.innerHTML = slice.map(r => {
      const ngClass = r.ng_rate > 8 ? 'badge-red' : r.ng_rate > 3 ? 'badge-yellow' : 'badge-green';
      return `
        <tr>
          <td class="font-medium text-slate-700">${r.part_name}</td>
          <td class="text-slate-500">${r.material}</td>
          <td class="font-mono text-slate-600">${formatNumber(r.qty_prod)}</td>
          <td class="font-mono text-red-600 font-semibold">${formatNumber(r.ng_prod)}</td>
          <td><span class="badge ${ngClass}">${r.ng_rate.toFixed(2)}%</span></td>
          <td class="font-mono font-bold text-red-700">${formatIDR(r.ng_loss)}</td>
        </tr>
      `;
    }).join('');
  }

  renderPagination('ng-pagination', page, Math.ceil(total / APP.PAGE_SIZE), renderNGTablePage);
}

function sortNGTable(col) {
  if (APP.ngTableSort.col === col) {
    APP.ngTableSort.dir = APP.ngTableSort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    APP.ngTableSort.col = col;
    APP.ngTableSort.dir = 'desc';
  }
  renderNGTablePage(1);
}

// ===== RENDER MATERIAL CHARTS =====
function renderMaterialAnalysis() {
  // HANYA DATA ALTERNATIVE
  const altData = APP.filteredData.filter(r => r.scenario === 'Alternative');

  // Composition Donut by usage_kg
  const byKat = {};
  altData.forEach(r => {
    byKat[r.kategori] = (byKat[r.kategori] || 0) + r.usage_kg;
  });
  const katLabels = Object.keys(byKat);
  const katVals = Object.values(byKat).map(v => +v.toFixed(2));

  destroyChart('matComposition');
  APP.charts.matComposition = new ApexCharts(
    document.getElementById('chart-mat-composition'),
    {
      series: katVals,
      labels: katLabels,
      chart: { type: 'donut', height: 300, fontFamily: 'Plus Jakarta Sans, sans-serif', animations: { speed: 500 } },
      colors: CATEGORY_COLORS,
      dataLabels: {
        enabled: true,
        formatter: (v, o) => o.w.globals.labels[o.seriesIndex] + '\n' + v.toFixed(1) + '%'
      },
      plotOptions: {
        pie: {
          donut: {
            size: '65%',
            labels: {
              show: true,
              total: { show: true, label: 'Total Usage', formatter: () => formatKg(katVals.reduce((a, b) => a + b, 0)) }
            }
          }
        }
      },
      tooltip: { y: { formatter: v => formatKg(v) } },
      legend: { position: 'bottom', fontSize: '11px' },
    }
  );
  APP.charts.matComposition.render();

  // Top Material Cost Bar
  const byMat = {};
  altData.forEach(r => {
    if (!byMat[r.material]) byMat[r.material] = 0;
    byMat[r.material] += r.total_cost;
  });
  const matSorted = Object.entries(byMat).sort((a, b) => b[1] - a[1]).slice(0, 10);

  destroyChart('matCost');
  APP.charts.matCost = new ApexCharts(
    document.getElementById('chart-mat-cost'),
    {
      series: [{ name: 'Total Cost', data: matSorted.map(m => Math.round(m[1])) }],
      chart: { type: 'bar', height: 300, toolbar: { show: false }, fontFamily: 'Plus Jakarta Sans, sans-serif', animations: { speed: 500 } },
      plotOptions: { bar: { horizontal: true, borderRadius: 6, barHeight: '55%' } },
      colors: ['#1d4ed8'],
      dataLabels: {
        enabled: true,
        formatter: v => formatIDR(v),
        style: { fontSize: '11px', colors: ['#fff'], fontWeight: '600' }
      },
      tooltip: { y: { formatter: v => formatIDR(v, true) } },
      xaxis: {
        categories: matSorted.map(([m]) => m.length > 22 ? m.substring(0, 20) + '…' : m),
        labels: { formatter: v => formatIDR(v), style: { fontSize: '10px', colors: '#94a3b8' } },
        axisBorder: { show: false },
        axisTicks: { show: false },
      },
      yaxis: { labels: { style: { fontSize: '11px', fontWeight: '600', colors: '#475569' } } },
      grid: { borderColor: '#f1f5f9', strokeDashArray: 3 },
    }
  );
  APP.charts.matCost.render();

  renderMatTable();
}

let matTableData = [];

function renderMatTable() {
  const altData = APP.filteredData.filter(r => r.scenario === 'Alternative');
  const byMat = {};
  altData.forEach(r => {
    if (!byMat[r.material]) {
      byMat[r.material] = { material: r.material, kategori: r.kategori, usage_kg: 0, total_cost: 0, parts: new Set() };
    }
    byMat[r.material].usage_kg += r.usage_kg;
    byMat[r.material].total_cost += r.total_cost;
    byMat[r.material].parts.add(r.part_name);
  });

  matTableData = Object.values(byMat).map(r => ({ ...r, total_part: r.parts.size }));
  renderMatTablePage(1);
}

function renderMatTablePage(page) {
  APP.matTablePage = page;
  const search = (document.getElementById('mat-search')?.value || '').toLowerCase();
  const { col, dir } = APP.matTableSort;

  let data = matTableData.filter(r =>
    r.material.toLowerCase().includes(search) ||
    r.kategori.toLowerCase().includes(search)
  );

  data.sort((a, b) => {
    const av = a[col], bv = b[col];
    return dir === 'asc' ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
  });

  const maxUsage = Math.max(...data.map(r => r.usage_kg), 1);
  const total = data.length;
  const start = (page - 1) * APP.PAGE_SIZE;
  const slice = data.slice(start, start + APP.PAGE_SIZE);

  const tbody = document.getElementById('mat-table-body');
  if (slice.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="text-center py-8 text-slate-400">Tidak ada data</td></tr>`;
  } else {
    tbody.innerHTML = slice.map(r => `
      <tr>
        <td class="font-medium text-slate-700">${r.material}</td>
        <td><span class="badge badge-blue">${r.kategori}</span></td>
        <td>
          <div class="flex items-center gap-2">
            <div style="width:60px;"><div class="progress-bar"><div class="progress-bar-fill" style="width:${(r.usage_kg / maxUsage * 100).toFixed(1)}%; background:#1d4ed8;"></div></div></div>
            <span class="font-mono text-xs text-slate-600">${formatKg(r.usage_kg)}</span>
          </div>
        </td>
        <td class="font-mono font-semibold text-slate-700">${formatIDR(r.total_cost)}</td>
        <td class="font-mono text-slate-600">${r.total_part}</td>
      </tr>
    `).join('');
  }

  renderPagination('mat-pagination', page, Math.ceil(total / APP.PAGE_SIZE), renderMatTablePage);
}

function sortMatTable(col) {
  if (APP.matTableSort.col === col) {
    APP.matTableSort.dir = APP.matTableSort.dir === 'asc' ? 'desc' : 'asc';
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
    renderTopMaterialChart();
    renderMonthlySavingChart()
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

// ==============================
// POPULATE FILTERS
// ==============================
function populateFilters() {
  const periodeSet = new Set();
  const katSet = new Set();
  const partSet = new Set();
  const materialSet = new Set();

  APP.rawData.forEach(r => {
    if (r.periode) periodeSet.add(r.periode);
    if (r.kategori) katSet.add(r.kategori);
    if (r.part_name) partSet.add(r.part_name);
    if (r.material) materialSet.add(r.material);
  });

  // Sort periode: by tahun then bulan
  const sortedPeriodes = [...periodeSet].sort((a, b) => {
    // "Apr 2026" -> parse
    const parseP = str => {
      const [mon, yr] = str.split(' ');
      const mIdx = BULAN_NAMES.indexOf(mon);
      return +yr * 100 + mIdx;
    };
    return parseP(a) - parseP(b);
  });

  $('#filter-periode').html(
    sortedPeriodes.map(p => `<option value="${p}">${p}</option>`).join('')
  );

  $('#filter-kategori').html(
    [...katSet].sort().map(k => `<option value="${k}">${k}</option>`).join('')
  );

  $('#filter-part').html(
    [...partSet].sort().map(p => `<option value="${p}">${p}</option>`).join('')
  );

  $('#filter-material').html(
    [...materialSet].sort().map(m => `<option value="${m}">${m}</option>`).join('')
  );

  // Destroy old Select2
  $('.filter-multi').each(function () {
    if ($(this).hasClass('select2-hidden-accessible')) {
      $(this).select2('destroy');
    }
  });

  // Init Select2
  $('.filter-multi').select2({
    width: '100%',
    closeOnSelect: false,
    allowClear: true,
    placeholder: 'Pilih Data',
    language: { noResults: () => 'Data tidak ditemukan' }
  });

  // Update counter display
  function updateCounter(id, emptyText) {
    const val = $(id).val();
    const total = Array.isArray(val) ? val.length : 0;
    const rendered = $(id).next('.select2-container').find('.select2-selection__rendered');
    rendered.find('li').hide();
    if (total === 0) {
      rendered.attr('data-count', emptyText);
    } else {
      rendered.attr('data-count', total + ' dipilih');
    }
  }

  updateCounter('#filter-periode', 'Semua Periode');
  updateCounter('#filter-scenario', 'Semua Scenario');
  updateCounter('#filter-kategori', 'Semua Kategori');
  updateCounter('#filter-part', 'Semua Part');
  updateCounter('#filter-material', 'Semua Material');

  // Remove old events
  $('.filter-multi').off('change');

  // Change event
  $('.filter-multi').on('change', function () {
    const id = '#' + $(this).attr('id');
    const emptyMap = {
      '#filter-periode': 'Semua Periode',
      '#filter-scenario': 'Semua Scenario',
      '#filter-kategori': 'Semua Kategori',
      '#filter-part': 'Semua Part',
      '#filter-material': 'Semua Material',
    };
    updateCounter(id, emptyMap[id] || 'Pilih Data');
    debounceFilter();
  });

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
  $('#filter-periode').val(null).trigger('change');
  $('#filter-scenario').val(null).trigger('change');
  $('#filter-kategori').val(null).trigger('change');
  $('#filter-part').val(null).trigger('change');
  $('#filter-material').val(null).trigger('change');
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
    usage_kg: +r.usage_kg.toFixed(6),
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
    usage_kg: +r.usage_kg.toFixed(6),
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
  document.getElementById('header-date').textContent = new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });

  // Init Select2 with empty data
  $('#filter-periode').select2({ placeholder: 'Semua Periode', allowClear: true });
  $('#filter-scenario').select2({ placeholder: 'Semua Scenario', allowClear: true });
  $('#filter-kategori').select2({ placeholder: 'Semua Kategori', allowClear: true });
  $('#filter-part').select2({ placeholder: 'Semua Part', allowClear: true });
  $('#filter-material').select2({ placeholder: 'Semua Material', allowClear: true });

  initDragDrop();
  populateFilters();

  if (window.lucide) lucide.createIcons();
});

function toggleFilters() {
  const panel = document.getElementById('filter-panel');
  const chevron = document.getElementById('filter-chevron');

  if (!panel) return;

  panel.classList.toggle('show');

  if (chevron) {
    chevron.classList.toggle('fa-chevron-down');
    chevron.classList.toggle('fa-chevron-up');
  }
}

$('.filter-multi').on('change', function(){
  if(window.innerWidth <= 992){
    document.getElementById('filter-panel').classList.remove('show');

    const chevron = document.getElementById('filter-chevron');
    chevron.classList.remove('fa-chevron-up');
    chevron.classList.add('fa-chevron-down');
  }
});
