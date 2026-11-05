/* EZPICKUP WMS — Management Reporting Center
 * Read-only reporting UI. Backend action: GET_REPORT_DATA.
 * Exports: CSV, XLSX (SheetJS CE), PDF via browser print, PPTX (PptxGenJS).
 */
(function () {
  "use strict";
  var reportState = { data: null, charts: {}, initialized: false, loading: false };
  var $ = function (id) { return document.getElementById(id); };

  function reportEscape(value) {
    return String(value == null ? "" : value).replace(/[&<>"']/g, function (ch) {
      return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch];
    });
  }
  function reportNumber(value) {
    return Number(value || 0).toLocaleString("id-ID", { maximumFractionDigits: 2 });
  }
  function reportMoney(value) {
    return "Rp " + Number(value || 0).toLocaleString("id-ID", { maximumFractionDigits: 0 });
  }
  function reportPct(value) {
    if (value === null || value === undefined || !isFinite(Number(value))) return "—";
    return (Number(value) > 0 ? "+" : "") + Number(value).toFixed(1) + "%";
  }
  function reportDateToday() {
    return typeof getWIBDateString === "function" ? getWIBDateString() : new Date().toISOString().slice(0, 10);
  }
  function reportDateShift(dateString, days) {
    var d = new Date(dateString + "T12:00:00+07:00");
    d.setDate(d.getDate() + days);
    return d.toLocaleDateString("en-CA", { timeZone: "Asia/Jakarta" });
  }
  function reportFilename() {
    var d = reportState.data || {};
    return "EZPICKUP-Laporan-" + (d.startDate || "periode") + "-sampai-" + (d.endDate || "") ;
  }
  function reportToast(message, kind) {
    if (typeof showToast === "function") showToast(message, kind || "info");
    else window.alert(message);
  }
  function loadScriptOnce(id, src) {
    return new Promise(function (resolve, reject) {
      if (document.getElementById(id)) {
        var existing = document.getElementById(id);
        if (existing.dataset.loaded === "1") return resolve();
        existing.addEventListener("load", resolve, { once: true });
        existing.addEventListener("error", reject, { once: true });
        return;
      }
      var script = document.createElement("script");
      script.id = id;
      script.src = src;
      script.async = true;
      script.onload = function () { script.dataset.loaded = "1"; resolve(); };
      script.onerror = function () { reject(new Error("Library ekspor gagal dimuat: " + src)); };
      document.head.appendChild(script);
    });
  }

  function buildReportsMarkup() {
    var root = $("tab-reports");
    if (!root || reportState.initialized) return;
    root.innerHTML = [
      '<div class="flex flex-wrap items-start justify-between gap-3 mb-5">',
        '<div><h2 class="font-display text-2xl font-bold text-slate-900">Laporan &amp; Analisa</h2>',
        '<p class="text-sm text-slate-500 mt-1">Management Reporting Center — rekap mingguan, bulanan, analisa tren, dan ekspor laporan.</p></div>',
        '<div class="flex flex-wrap gap-2 no-print">',
          '<button id="report-btn-xlsx" onclick="exportReportXlsx()" disabled class="px-3 py-2 rounded-lg bg-emerald-600 text-white text-xs font-semibold disabled:opacity-40"><i data-lucide="file-spreadsheet" class="w-4 h-4 inline mr-1"></i>Excel</button>',
          '<button id="report-btn-pptx" onclick="exportReportPptx()" disabled class="px-3 py-2 rounded-lg bg-red-600 text-white text-xs font-semibold disabled:opacity-40"><i data-lucide="presentation" class="w-4 h-4 inline mr-1"></i>PowerPoint</button>',
          '<button id="report-btn-pdf" onclick="printReportPdf()" disabled class="px-3 py-2 rounded-lg border border-slate-300 bg-white text-slate-700 text-xs font-semibold disabled:opacity-40"><i data-lucide="file-pdf" class="w-4 h-4 inline mr-1"></i>PDF / Cetak</button>',
          '<button id="report-btn-csv" onclick="exportReportCsv()" disabled class="px-3 py-2 rounded-lg border border-slate-300 bg-white text-slate-700 text-xs font-semibold disabled:opacity-40"><i data-lucide="download" class="w-4 h-4 inline mr-1"></i>CSV Detail</button>',
        '</div>',
      '</div>',
      '<div class="bg-white rounded-xl border border-slate-200 shadow-sm p-4 mb-5 no-print">',
        '<div class="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-5 gap-3 items-end">',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Jenis Periode</label>',
            '<select id="report-period" onchange="setReportPreset(this.value)" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"><option value="weekly">Mingguan (Senin–hari ini)</option><option value="monthly">Bulanan (awal bulan–hari ini)</option><option value="custom">Rentang khusus</option></select></div>',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Dari Tanggal</label><input id="report-from" type="date" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"></div>',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Sampai Tanggal</label><input id="report-to" type="date" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"></div>',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Gudang</label><select id="report-warehouse" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"><option value="">Semua Gudang</option></select></div>',
          '<button id="report-btn-run" onclick="loadReportData()" class="w-full px-4 py-2 rounded-lg bg-slate-900 text-white text-sm font-semibold hover:bg-slate-700"><i data-lucide="refresh-cw" class="w-4 h-4 inline mr-1"></i>Generate Laporan</button>',
        '</div>',
        '<div id="report-status" class="text-xs text-slate-400 mt-3">Pilih periode dan gudang, lalu generate laporan. Periode pembanding otomatis menggunakan durasi yang sama.</div>',
      '</div>',
      '<div id="report-print-area">',
        '<div class="hidden" id="report-empty"></div>',
        '<div id="report-results" class="hidden space-y-5">',
          '<div class="flex flex-wrap justify-between items-center gap-2"><div><h3 class="font-display text-lg font-bold text-slate-900">Ringkasan Eksekutif</h3><p id="report-period-label" class="text-xs text-slate-500 mt-0.5"></p></div><span id="report-wh-label" class="text-xs font-semibold text-slate-600 bg-slate-100 rounded-full px-3 py-1"></span></div>',
          '<div id="report-kpis" class="grid grid-cols-2 xl:grid-cols-4 gap-3"></div>',
          '<div class="grid grid-cols-1 xl:grid-cols-2 gap-4">',
            '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Tren Inbound &amp; Outbound</h3><div class="h-64"><canvas id="report-chart-movement"></canvas></div></div>',
            '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Volume Paket &amp; Produktivitas</h3><div class="h-64"><canvas id="report-chart-volume"></canvas></div></div>',
          '</div>',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><div class="flex items-center justify-between gap-2 mb-3"><h3 class="font-bold text-sm text-slate-800">Analisa &amp; Sinyal Operasional</h3><span class="text-[10px] text-slate-400">Berbasis perhitungan data</span></div><div id="report-insights" class="grid grid-cols-1 md:grid-cols-2 gap-2"></div></div>',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Perbandingan Performa Gudang</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 uppercase bg-slate-50"><tr><th class="p-2 text-left">Gudang</th><th class="p-2 text-right">Inbound</th><th class="p-2 text-right">Outbound</th><th class="p-2 text-right">Paket</th><th class="p-2 text-right">Produktivitas</th><th class="p-2 text-right report-money-col">Cost Packing</th></tr></thead><tbody id="report-warehouse-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
          '<div class="grid grid-cols-1 xl:grid-cols-2 gap-4">',
            '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Top Konsumsi Material (OUT)</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 bg-slate-50"><tr><th class="p-2 text-left">Material</th><th class="p-2 text-left">Gudang</th><th class="p-2 text-right">Qty OUT</th></tr></thead><tbody id="report-material-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
            '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Kategori Packing</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 bg-slate-50"><tr><th class="p-2 text-left">Kategori</th><th class="p-2 text-right">AWB</th><th class="p-2 text-right">Paket</th><th class="p-2 text-right report-money-col">Cost</th></tr></thead><tbody id="report-category-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
          '</div>',
          '<div class="grid grid-cols-1 xl:grid-cols-2 gap-4">',
            '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Produktivitas per Peran</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 bg-slate-50"><tr><th class="p-2 text-left">Peran</th><th class="p-2 text-right">Paket</th><th class="p-2 text-right">Pcs</th><th class="p-2 text-right">Entri</th></tr></thead><tbody id="report-role-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
            '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">PO &amp; Case Operasional</h3><div id="report-ops-cards" class="grid grid-cols-2 gap-2"></div></div>',
          '</div>',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-2">Catatan &amp; Batas Data</h3><ul id="report-notes" class="text-xs text-slate-500 list-disc pl-5 space-y-1"></ul><p id="report-truncation-note" class="hidden text-xs text-amber-700 mt-2"></p></div>',
          '<div class="text-[10px] text-slate-400 text-right">EZPICKUP WMS • Management Reporting Center • Internal Use</div>',
        '</div>',
      '</div>'
    ].join("");
    reportState.initialized = true;
    initReportWarehouseOptions();
    setReportPreset("weekly");
    if (typeof lucide !== "undefined") lucide.createIcons();
  }

  function initReportWarehouseOptions() {
    var select = $("report-warehouse");
    if (!select) return;
    var current = currentUser && currentUser.resolvedWarehouseId ? currentUser.resolvedWarehouseId : "";
    var options = ['<option value="">Semua Gudang</option>'];
    (db && db.warehouses ? db.warehouses : []).forEach(function (w) {
      var id = w["Warehouse ID"];
      if (!id) return;
      options.push('<option value="' + reportEscape(id) + '">' + reportEscape(id + " — " + (w["Warehouse Name"] || "")) + "</option>");
    });
    select.innerHTML = options.join("");
    if (current) {
      select.value = current;
      select.disabled = true;
      select.classList.add("bg-slate-100");
    }
  }

  window.setReportPreset = function (preset) {
    if (!reportState.initialized) buildReportsMarkup();
    var today = reportDateToday();
    var from = today;
    if (preset === "weekly") {
      var d = new Date(today + "T12:00:00+07:00");
      var daysFromMonday = (d.getDay() + 6) % 7;
      from = reportDateShift(today, -daysFromMonday);
    } else if (preset === "monthly") {
      from = today.slice(0, 8) + "01";
    } else if (preset === "custom") {
      from = $("report-from").value || today;
    }
    $("report-period").value = preset;
    if (preset !== "custom") $("report-from").value = from;
    if (preset !== "custom" || !$("report-to").value) $("report-to").value = today;
    if (preset !== "custom" && reportState.data) loadReportData();
  };

  window.loadReportData = async function () {
    if (!reportState.initialized) buildReportsMarkup();
    if (reportState.loading) return;
    var from = $("report-from").value;
    var to = $("report-to").value;
    if (!from || !to || from > to) {
      reportToast("Periksa rentang tanggal laporan.", "error");
      return;
    }
    var btn = $("report-btn-run");
    var status = $("report-status");
    reportState.loading = true;
    if (btn) { btn.disabled = true; btn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 inline mr-1 animate-spin"></i>Memproses...'; }
    if (status) status.textContent = "Mengambil data laporan dari backend...";
    if (typeof lucide !== "undefined") lucide.createIcons();
    try {
      var response = await apiPost({
        action: "GET_REPORT_DATA",
        startDate: from,
        endDate: to,
        warehouseId: $("report-warehouse").value || ""
      });
      if (!response || !response.success || !response.report) throw new Error((response && (response.message || response.error)) || "Data laporan tidak dapat dimuat.");
      reportState.data = response.report;
      renderReportData(response.report);
      if (status) status.textContent = "Laporan berhasil dibuat • " + reportNumber(response.report.days) + " hari • dibandingkan dengan " + response.report.previousStartDate + " s.d. " + response.report.previousEndDate + ".";
      ["report-btn-xlsx", "report-btn-pptx", "report-btn-pdf", "report-btn-csv"].forEach(function (id) { if ($(id)) $(id).disabled = false; });
      reportToast("Laporan berhasil dibuat.", "success");
    } catch (err) {
      if (status) status.textContent = "Gagal membuat laporan: " + err.message;
      reportToast("Laporan gagal: " + err.message, "error");
    } finally {
      reportState.loading = false;
      if (btn) { btn.disabled = false; btn.innerHTML = '<i data-lucide="refresh-cw" class="w-4 h-4 inline mr-1"></i>Generate Laporan'; }
      if (typeof lucide !== "undefined") lucide.createIcons();
    }
  };

  function renderKpi(label, value, compare, note) {
    var cls = compare === null || compare === undefined ? "text-slate-400" : Number(compare) > 0 ? "text-emerald-600" : Number(compare) < 0 ? "text-red-600" : "text-slate-500";
    return '<div class="bg-white rounded-xl border border-slate-200 p-4 shadow-sm"><p class="text-[10px] font-bold uppercase tracking-wide text-slate-400">' + reportEscape(label) + '</p><p class="font-display text-xl sm:text-2xl font-bold text-slate-900 mt-2 break-words">' + reportEscape(value) + '</p><p class="text-[10px] mt-1 ' + cls + '">' + reportEscape(compare === null || compare === undefined ? "Tidak ada pembanding" : reportPct(compare) + " vs periode sebelumnya") + '</p>' + (note ? '<p class="text-[10px] text-slate-400 mt-1">' + reportEscape(note) + '</p>' : "") + '</div>';
  }

  function renderTableRows(id, rows, colSpan, emptyText) {
    var el = $(id);
    if (!el) return;
    el.innerHTML = rows.length ? rows.join("") : '<tr><td colspan="' + colSpan + '" class="p-4 text-center text-slate-400">' + reportEscape(emptyText || "Tidak ada data pada periode ini.") + "</td></tr>";
  }

  function renderReportData(r) {
    var metrics = r.metrics || {};
    var movement = metrics.movement || {};
    var prev = r.previousMetrics || {};
    var changes = r.changes || {};
    var inv = r.inventoryNow || {};
    var po = r.purchaseOrders || {};
    var cases = r.cases || {};
    var prod = r.productivity || {};
    $("report-results").classList.remove("hidden");
    $("report-period-label").textContent = r.startDate + " s.d. " + r.endDate + " • " + r.days + " hari";
    $("report-wh-label").textContent = r.warehouseId === "ALL" ? "Semua Gudang" : "Gudang " + r.warehouseId;

    var kpis = [
      renderKpi("Inbound Qty", reportNumber(movement.inboundQty), changes.inboundPct, "Transaksi IN"),
      renderKpi("Outbound Qty", reportNumber(movement.outboundQty), changes.outboundPct, "Transaksi OUT"),
      renderKpi("Paket Packing", reportNumber(metrics.totalPackages), changes.packagesPct, reportNumber(metrics.totalAWB) + " AWB"),
      renderKpi("Produktivitas", reportNumber(metrics.productivityPackages) + " pkt", changes.productivityPct, reportNumber(metrics.productivityPcs) + " pcs")
    ];
    if (r.financialVisibility) kpis.push(renderKpi("Cost Packing", reportMoney(metrics.totalPackingCost), changes.packingCostPct, "Berdasarkan Pack Log"));
    kpis.push(renderKpi("SKU Critical", reportNumber(inv.criticalCount), null, reportNumber(inv.skuCount) + " SKU posisi sekarang"));
    kpis.push(renderKpi("PO Terlambat", reportNumber(po.overdueNow), null, "Outstanding saat ini"));
    kpis.push(renderKpi("Case Terbuka", reportNumber(cases.openNow), null, reportNumber(cases.newInPeriod) + " case baru pada periode"));
    if (prod.attainmentPct !== null && prod.attainmentPct !== undefined) kpis.push(renderKpi("Pencapaian Target", Number(prod.attainmentPct).toFixed(1) + "%", null, "Target aktif × jumlah hari"));
    $("report-kpis").innerHTML = kpis.join("");
    document.querySelectorAll("#tab-reports .report-money-col").forEach(function (el) { el.classList.toggle("hidden", !r.financialVisibility); });

    renderReportCharts(r);
    $("report-insights").innerHTML = (r.insights || []).map(function (x) {
      var style = x.level === "warning" ? "border-amber-200 bg-amber-50" : x.level === "success" ? "border-emerald-200 bg-emerald-50" : "border-blue-100 bg-blue-50";
      var icon = x.level === "warning" ? "alert-triangle" : x.level === "success" ? "check-circle-2" : "info";
      return '<div class="border rounded-lg p-3 ' + style + '"><p class="text-xs font-bold text-slate-800"><i data-lucide="' + icon + '" class="w-3.5 h-3.5 inline mr-1"></i>' + reportEscape(x.title) + '</p><p class="text-xs text-slate-600 mt-1">' + reportEscape(x.detail) + "</p></div>";
    }).join("");

    renderTableRows("report-warehouse-rows", (r.warehousePerformance || []).map(function (w) {
      return "<tr><td class=\"p-2 font-semibold\">" + reportEscape(w.warehouseName || w.warehouseId) + '</td><td class="p-2 text-right">' + reportNumber(w.inboundQty) + '</td><td class="p-2 text-right">' + reportNumber(w.outboundQty) + '</td><td class="p-2 text-right">' + reportNumber(w.packages) + '</td><td class="p-2 text-right">' + reportNumber(w.productivityPackages) + '</td><td class="p-2 text-right report-money-col">' + (r.financialVisibility ? reportMoney(w.packingCost) : "—") + "</td></tr>";
    }), 6);
    renderTableRows("report-material-rows", (r.topMaterials || []).map(function (m) {
      return "<tr><td class=\"p-2 font-semibold\">" + reportEscape(m.materialName) + '<div class="text-[10px] text-slate-400">' + reportEscape(m.materialId) + '</div></td><td class="p-2">' + reportEscape(m.warehouseId) + '</td><td class="p-2 text-right font-semibold">' + reportNumber(m.qtyOut) + "</td></tr>";
    }), 3);
    renderTableRows("report-category-rows", (r.topCategories || []).map(function (c) {
      return "<tr><td class=\"p-2 font-semibold\">" + reportEscape(c.category) + '</td><td class="p-2 text-right">' + reportNumber(c.awb) + '</td><td class="p-2 text-right">' + reportNumber(c.packages) + '</td><td class="p-2 text-right report-money-col">' + (r.financialVisibility ? reportMoney(c.packingCost) : "—") + "</td></tr>";
    }), 4);
    renderTableRows("report-role-rows", (prod.byRole || []).map(function (x) {
      return "<tr><td class=\"p-2 font-semibold\">" + reportEscape(x.role) + '</td><td class="p-2 text-right">' + reportNumber(x.packages) + '</td><td class="p-2 text-right">' + reportNumber(x.pcs) + '</td><td class="p-2 text-right">' + reportNumber(x.entries) + "</td></tr>";
    }), 4);

    var ops = [
      ["PO dibuat periode ini", po.createdInPeriod],
      ["PO selesai periode ini", po.completedInPeriod],
      ["PO outstanding sekarang", po.outstandingNow],
      ["PO terlambat sekarang", po.overdueNow],
      ["Case baru periode ini", cases.newInPeriod],
      ["Case terbuka sekarang", cases.openNow],
      ["Case prioritas tinggi terbuka", cases.highPriorityOpenNow],
      ["PO on-time / delayed", reportNumber(po.onTimeInPeriod) + " / " + reportNumber(po.delayedInPeriod)]
    ];
    $("report-ops-cards").innerHTML = ops.map(function (x) {
      return '<div class="rounded-lg bg-slate-50 border border-slate-100 p-3"><p class="text-[10px] text-slate-400">' + reportEscape(x[0]) + '</p><p class="text-lg font-bold text-slate-800 mt-1">' + reportEscape(reportNumber(x[1])) + "</p></div>";
    }).join("");
    $("report-notes").innerHTML = (r.dataNotes || []).map(function (x) { return "<li>" + reportEscape(x) + "</li>"; }).join("");
    var truncTotal = Object.keys(r.truncatedRows || {}).reduce(function (sum, k) { return sum + Number(r.truncatedRows[k] || 0); }, 0);
    $("report-truncation-note").classList.toggle("hidden", !truncTotal);
    if (truncTotal) $("report-truncation-note").textContent = "Peringatan: " + reportNumber(truncTotal) + " baris detail melebihi batas respons ekspor. Ringkasan KPI tetap menghitung seluruh data periode; detail pada file dibatasi untuk menjaga stabilitas aplikasi.";
    if (typeof lucide !== "undefined") lucide.createIcons();
  }

  function destroyReportChart(key) {
    if (reportState.charts[key]) {
      try { reportState.charts[key].destroy(); } catch (e) {}
      reportState.charts[key] = null;
    }
  }
  function renderReportCharts(r) {
    if (typeof Chart === "undefined") return;
    var trend = r.dailyTrend || [];
    var labels = trend.map(function (d) { return d.date.slice(5); });
    var movement = $("report-chart-movement");
    var volume = $("report-chart-volume");
    destroyReportChart("movement");
    destroyReportChart("volume");
    reportState.charts.movement = new Chart(movement, {
      type: "line",
      data: { labels: labels, datasets: [
        { label: "Inbound", data: trend.map(function (d) { return d.inboundQty; }), borderColor: "#2563eb", backgroundColor: "rgba(37,99,235,.08)", tension: .25, fill: false },
        { label: "Outbound", data: trend.map(function (d) { return d.outboundQty; }), borderColor: "#dc2626", backgroundColor: "rgba(220,38,38,.08)", tension: .25, fill: false }
      ] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } }, scales: { y: { beginAtZero: true } } }
    });
    reportState.charts.volume = new Chart(volume, {
      type: "bar",
      data: { labels: labels, datasets: [
        { label: "Paket Packing", data: trend.map(function (d) { return d.packages; }), backgroundColor: "#dc2626" },
        { label: "Produktivitas (Paket)", data: trend.map(function (d) { return d.productivityPackages; }), backgroundColor: "#0f766e" }
      ] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } }, scales: { y: { beginAtZero: true } } }
    });
  }

  function reportSheetFromJson(XLSX, rows) {
    return XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ Keterangan: "Tidak ada data pada periode ini" }]);
  }
  window.exportReportXlsx = async function () {
    var r = reportState.data;
    if (!r) return reportToast("Generate laporan terlebih dahulu.", "info");
    try {
      if (!window.XLSX) await loadScriptOnce("sheetjs-ce", "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js");
      var XLSX = window.XLSX;
      var wb = XLSX.utils.book_new();
      var summary = [
        { Indikator: "Periode", Nilai: r.startDate + " s.d. " + r.endDate },
        { Indikator: "Gudang", Nilai: r.warehouseId },
        { Indikator: "Inbound Qty", Nilai: r.metrics.movement.inboundQty },
        { Indikator: "Outbound Qty", Nilai: r.metrics.movement.outboundQty },
        { Indikator: "Transfer Qty", Nilai: r.metrics.movement.transferQty },
        { Indikator: "Adjustment Net Qty", Nilai: r.metrics.movement.adjustmentNetQty },
        { Indikator: "AWB", Nilai: r.metrics.totalAWB },
        { Indikator: "Paket Packing", Nilai: r.metrics.totalPackages },
        { Indikator: "Produktivitas Paket", Nilai: r.metrics.productivityPackages },
        { Indikator: "Produktivitas Pcs", Nilai: r.metrics.productivityPcs },
        { Indikator: "SKU Posisi Sekarang", Nilai: r.inventoryNow.skuCount },
        { Indikator: "SKU Critical Sekarang", Nilai: r.inventoryNow.criticalCount },
        { Indikator: "PO Outstanding Sekarang", Nilai: r.purchaseOrders.outstandingNow },
        { Indikator: "PO Terlambat Sekarang", Nilai: r.purchaseOrders.overdueNow },
        { Indikator: "Case Terbuka Sekarang", Nilai: r.cases.openNow }
      ];
      if (r.financialVisibility) {
        summary.push({ Indikator: "Cost Packing", Nilai: r.metrics.totalPackingCost });
        summary.push({ Indikator: "Valuasi Stok Posisi Sekarang", Nilai: r.inventoryNow.valuation });
      }
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, summary), "Executive Summary");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, r.dailyTrend), "Daily Trend");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, r.warehousePerformance), "Warehouse");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, r.topMaterials), "Top Material");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, r.topCategories), "Packing Category");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, r.productivity.byRole), "Produktivitas Role");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, r.topClients), "Top Client");
      var details = r.details || {};
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, details.transactions), "Detail Transaksi");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, details.packLogs), "Detail Packing");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, details.productivityLogs), "Detail Produktivitas");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, details.purchaseOrders), "Detail PO");
      XLSX.utils.book_append_sheet(wb, reportSheetFromJson(XLSX, details.cases), "Detail Case");
      XLSX.writeFile(wb, reportFilename() + ".xlsx");
      reportToast("File Excel berhasil dibuat.", "success");
    } catch (err) {
      reportToast("Gagal membuat Excel: " + err.message, "error");
    }
  };

  window.exportReportCsv = function () {
    var r = reportState.data;
    if (!r) return reportToast("Generate laporan terlebih dahulu.", "info");
    var rows = (r.details && r.details.transactions) || [];
    if (!rows.length) return reportToast("Tidak ada detail transaksi untuk diekspor.", "info");
    var headers = ["Trx ID", "Date", "Trx Type", "Origin WH ID", "Dest WH ID", "Material ID", "Material Name", "Qty", "Reference No", "Operator"];
    var csv = [headers].concat(rows.map(function (row) { return headers.map(function (key) { return row[key] == null ? "" : row[key]; }); }));
    var text = csv.map(function (line) { return line.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(","); }).join("\r\n");
    var blob = new Blob(["\ufeff" + text], { type: "text/csv;charset=utf-8;" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = reportFilename() + "-detail-transaksi.csv";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    reportToast(reportNumber(rows.length) + " transaksi diekspor.", "success");
  };

  window.printReportPdf = function () {
    if (!reportState.data) return reportToast("Generate laporan terlebih dahulu.", "info");
    var style = document.getElementById("report-print-style");
    if (!style) {
      style = document.createElement("style");
      style.id = "report-print-style";
      style.textContent = "@media print{body{overflow:visible!important;height:auto!important;background:#fff!important}body *{visibility:hidden!important}#tab-reports,#tab-reports *{visibility:visible!important}#tab-reports{display:block!important;position:absolute!important;left:0!important;top:0!important;width:100%!important;padding:0!important;background:#fff!important}#tab-reports .no-print{display:none!important}.bg-white{box-shadow:none!important}canvas{max-height:230px!important} .tab-content{display:none!important}#tab-reports{break-inside:auto!important}}";
      document.head.appendChild(style);
    }
    window.print();
  };

  function addPptText(slide, text, x, y, w, h, options) {
    slide.addText(String(text == null ? "" : text), Object.assign({
      x: x, y: y, w: w, h: h, fontFace: "Aptos", fontSize: 12, color: "1F2937",
      margin: 0.04, breakLine: false, valign: "mid"
    }, options || {}));
  }
  function addPptHeader(slide, title, subtitle, pptx) {
    slide.background = { color: "F8FAFC" };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.16, line: { color: "B91C1C", transparency: 100 }, fill: { color: "B91C1C" } });
    addPptText(slide, title, 0.55, 0.35, 12.1, 0.42, { fontSize: 22, bold: true, color: "111827" });
    addPptText(slide, subtitle, 0.57, 0.82, 12, 0.22, { fontSize: 9, color: "64748B" });
    addPptText(slide, "EZPICKUP WMS  •  INTERNAL USE", 0.57, 7.18, 12, 0.16, { fontSize: 7, color: "94A3B8" });
  }
  function pptCard(slide, x, y, w, h, label, value, pptx, note) {
    slide.addShape(pptx.ShapeType.roundRect, { x: x, y: y, w: w, h: h, rectRadius: 0.08, line: { color: "E2E8F0", width: 1 }, fill: { color: "FFFFFF" } });
    addPptText(slide, label, x + 0.16, y + 0.14, w - 0.3, 0.22, { fontSize: 9, color: "64748B", bold: true });
    addPptText(slide, value, x + 0.16, y + 0.48, w - 0.3, 0.42, { fontSize: 20, color: "111827", bold: true });
    if (note) addPptText(slide, note, x + 0.16, y + h - 0.32, w - 0.3, 0.16, { fontSize: 7, color: "94A3B8" });
  }
  window.exportReportPptx = async function () {
    var r = reportState.data;
    if (!r) return reportToast("Generate laporan terlebih dahulu.", "info");
    try {
      if (!window.PptxGenJS) await loadScriptOnce("pptxgenjs-4", "https://cdn.jsdelivr.net/gh/gitbrent/PptxGenJS@4.0.1/dist/pptxgen.bundle.js");
      if (!window.PptxGenJS) throw new Error("Library PowerPoint tidak tersedia pada browser.");
      var pptx = new window.PptxGenJS();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "EZPICKUP WMS";
      pptx.subject = "Laporan Operasional";
      pptx.title = "EZPICKUP WMS — " + r.startDate + " s.d. " + r.endDate;
      pptx.company = "EZPICKUP";
      pptx.lang = "id-ID";
      var m = r.metrics, move = m.movement, po = r.purchaseOrders, inv = r.inventoryNow, cases = r.cases, prod = r.productivity;
      var subtitle = r.startDate + " s.d. " + r.endDate + "  |  " + (r.warehouseId === "ALL" ? "Semua Gudang" : r.warehouseId);

      var slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.22, line: { color: "B91C1C", transparency: 100 }, fill: { color: "B91C1C" } });
      addPptText(slide, "EZPICKUP WMS", 0.75, 1.25, 7.5, 0.4, { fontSize: 18, bold: true, color: "B91C1C" });
      addPptText(slide, "LAPORAN OPERASIONAL", 0.75, 2.05, 11.8, 0.85, { fontSize: 34, bold: true, color: "111827" });
      addPptText(slide, r.startDate + " — " + r.endDate, 0.78, 3.05, 11, 0.4, { fontSize: 20, color: "475569" });
      addPptText(slide, r.warehouseId === "ALL" ? "Seluruh Gudang" : "Gudang " + r.warehouseId, 0.8, 3.58, 10, 0.3, { fontSize: 13, color: "64748B" });
      addPptText(slide, "Management Reporting Center  •  Internal Use", 0.8, 6.75, 11, 0.25, { fontSize: 9, color: "94A3B8" });

      slide = pptx.addSlide(); addPptHeader(slide, "Executive Summary", subtitle, pptx);
      var cards = [
        ["Inbound Qty", reportNumber(move.inboundQty)],
        ["Outbound Qty", reportNumber(move.outboundQty)],
        ["Paket Packing", reportNumber(m.totalPackages)],
        ["Produktivitas", reportNumber(m.productivityPackages) + " pkt"]
      ];
      cards.forEach(function (c, i) { pptCard(slide, 0.6 + (i % 2) * 6.1, 1.45 + Math.floor(i / 2) * 1.45, 5.75, 1.15, c[0], c[1], pptx); });
      addPptText(slide, "Perbandingan terhadap periode sebelumnya (" + r.previousStartDate + " s.d. " + r.previousEndDate + ")", 0.7, 4.65, 12, 0.25, { fontSize: 11, bold: true, color: "334155" });
      addPptText(slide, "Inbound " + reportPct(r.changes.inboundPct) + "   |   Outbound " + reportPct(r.changes.outboundPct) + "   |   Paket " + reportPct(r.changes.packagesPct) + "   |   Produktivitas " + reportPct(r.changes.productivityPct), 0.7, 5.05, 12, 0.55, { fontSize: 12, color: "475569" });

      slide = pptx.addSlide(); addPptHeader(slide, "Inventory & Pergerakan Stok", subtitle, pptx);
      pptCard(slide, 0.65, 1.45, 3.85, 1.2, "SKU Posisi Sekarang", reportNumber(inv.skuCount), pptx);
      pptCard(slide, 4.75, 1.45, 3.85, 1.2, "SKU Critical", reportNumber(inv.criticalCount), pptx);
      pptCard(slide, 8.85, 1.45, 3.85, 1.2, "SKU Reorder", reportNumber(inv.reorderCount), pptx);
      addPptText(slide, "Inbound: " + reportNumber(move.inboundQty) + "   |   Outbound: " + reportNumber(move.outboundQty) + "   |   Transfer: " + reportNumber(move.transferQty) + "   |   Adjustment net: " + reportNumber(move.adjustmentNetQty), 0.8, 3.2, 12, 0.5, { fontSize: 13 });
      if (r.financialVisibility) addPptText(slide, "Valuasi stok posisi saat laporan dibuat: " + reportMoney(inv.valuation), 0.8, 4.0, 12, 0.3, { fontSize: 12, bold: true });
      addPptText(slide, "Catatan: angka inventory adalah posisi saat laporan dibuat, bukan saldo historis akhir periode.", 0.8, 5.1, 11.7, 0.5, { fontSize: 10, color: "64748B", italic: true });

      slide = pptx.addSlide(); addPptHeader(slide, "Outbound & Packing", subtitle, pptx);
      pptCard(slide, 0.65, 1.4, 3.85, 1.2, "Total AWB", reportNumber(m.totalAWB), pptx);
      pptCard(slide, 4.75, 1.4, 3.85, 1.2, "Total Paket", reportNumber(m.totalPackages), pptx);
      if (r.financialVisibility) pptCard(slide, 8.85, 1.4, 3.85, 1.2, "Cost Packing", reportMoney(m.totalPackingCost), pptx);
      addPptText(slide, "Kategori packing terbanyak", 0.8, 3.1, 5, 0.3, { fontSize: 13, bold: true });
      (r.topCategories || []).slice(0, 7).forEach(function (x, i) {
        addPptText(slide, x.category + "  —  " + reportNumber(x.packages) + " paket", 0.85, 3.55 + i * 0.42, 6.2, 0.24, { fontSize: 10 });
      });

      slide = pptx.addSlide(); addPptHeader(slide, "Performa Antar Gudang", subtitle, pptx);
      var whs = (r.warehousePerformance || []).slice(0, 8);
      addPptText(slide, "Gudang", 0.75, 1.35, 2.4, 0.25, { bold: true, fontSize: 10, color: "64748B" });
      addPptText(slide, "Inbound", 3.25, 1.35, 1.6, 0.25, { bold: true, fontSize: 10, color: "64748B", align: "right" });
      addPptText(slide, "Outbound", 5.0, 1.35, 1.6, 0.25, { bold: true, fontSize: 10, color: "64748B", align: "right" });
      addPptText(slide, "Paket", 6.75, 1.35, 1.6, 0.25, { bold: true, fontSize: 10, color: "64748B", align: "right" });
      addPptText(slide, "Produktivitas", 8.55, 1.35, 1.8, 0.25, { bold: true, fontSize: 10, color: "64748B", align: "right" });
      if (r.financialVisibility) addPptText(slide, "Cost Packing", 10.55, 1.35, 2, 0.25, { bold: true, fontSize: 10, color: "64748B", align: "right" });
      whs.forEach(function (w, i) {
        var y = 1.8 + i * 0.55;
        slide.addShape(pptx.ShapeType.line, { x: 0.7, y: y + 0.36, w: 12, h: 0, line: { color: "E2E8F0", width: 0.6 } });
        addPptText(slide, w.warehouseName || w.warehouseId, 0.75, y, 2.4, 0.28, { fontSize: 10, bold: true });
        addPptText(slide, reportNumber(w.inboundQty), 3.25, y, 1.6, 0.28, { fontSize: 10, align: "right" });
        addPptText(slide, reportNumber(w.outboundQty), 5.0, y, 1.6, 0.28, { fontSize: 10, align: "right" });
        addPptText(slide, reportNumber(w.packages), 6.75, y, 1.6, 0.28, { fontSize: 10, align: "right" });
        addPptText(slide, reportNumber(w.productivityPackages), 8.55, y, 1.8, 0.28, { fontSize: 10, align: "right" });
        if (r.financialVisibility) addPptText(slide, reportMoney(w.packingCost), 10.55, y, 2, 0.28, { fontSize: 9, align: "right" });
      });

      slide = pptx.addSlide(); addPptHeader(slide, "Produktivitas Tim", subtitle, pptx);
      pptCard(slide, 0.65, 1.4, 3.85, 1.2, "Paket Tercatat", reportNumber(m.productivityPackages), pptx);
      pptCard(slide, 4.75, 1.4, 3.85, 1.2, "Pcs Tercatat", reportNumber(m.productivityPcs), pptx);
      pptCard(slide, 8.85, 1.4, 3.85, 1.2, "Pencapaian Target", prod.attainmentPct == null ? "N/A" : Number(prod.attainmentPct).toFixed(1) + "%", pptx);
      (prod.byRole || []).slice(0, 8).forEach(function (x, i) {
        addPptText(slide, x.role + "  —  " + reportNumber(x.packages) + " paket / " + reportNumber(x.pcs) + " pcs", 0.8, 3.25 + i * 0.42, 11.5, 0.24, { fontSize: 10 });
      });

      slide = pptx.addSlide(); addPptHeader(slide, "Purchase Order & Case Operasional", subtitle, pptx);
      var ops = [
        ["PO dibuat periode ini", po.createdInPeriod], ["PO selesai periode ini", po.completedInPeriod],
        ["PO outstanding sekarang", po.outstandingNow], ["PO lewat estimasi tiba", po.overdueNow],
        ["Case baru periode ini", cases.newInPeriod], ["Case masih terbuka", cases.openNow],
        ["Case prioritas tinggi", cases.highPriorityOpenNow], ["PO on-time / delayed", reportNumber(po.onTimeInPeriod) + " / " + reportNumber(po.delayedInPeriod)]
      ];
      ops.forEach(function (x, i) { pptCard(slide, 0.65 + (i % 2) * 6.1, 1.35 + Math.floor(i / 2) * 1.25, 5.75, 0.95, x[0], String(x[1]), pptx); });

      slide = pptx.addSlide(); addPptHeader(slide, "Analisa & Rekomendasi Operasional", subtitle, pptx);
      (r.insights || []).slice(0, 8).forEach(function (x, i) {
        var y = 1.35 + i * 0.62;
        slide.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: y, w: 11.95, h: 0.48, rectRadius: 0.04, line: { color: x.level === "warning" ? "FDE68A" : "E2E8F0", width: 0.8 }, fill: { color: x.level === "warning" ? "FFFBEB" : "FFFFFF" } });
        addPptText(slide, x.title + " — " + x.detail, 0.9, y + 0.08, 11.5, 0.28, { fontSize: 9, color: "334155" });
      });
      addPptText(slide, (r.dataNotes || []).join(" "), 0.8, 6.55, 11.8, 0.38, { fontSize: 7, color: "64748B", italic: true });

      await pptx.writeFile({ fileName: reportFilename() + ".pptx" });
      reportToast("PowerPoint berhasil dibuat (" + pptx._slides.length + " slide).", "success");
    } catch (err) {
      reportToast("Gagal membuat PowerPoint: " + err.message, "error");
    }
  };

  // Called by switchTab('reports') after the core data is ready.
  window.initReportsTab = function () {
    buildReportsMarkup();
    if (!reportState.data && !reportState.loading) window.loadReportData();
  };

  document.addEventListener("DOMContentLoaded", function () {
    buildReportsMarkup();
  });
  // The inline WMS script is loaded before this file, so this also covers pages
  // where the DOMContentLoaded event has already fired.
  buildReportsMarkup();
})();
