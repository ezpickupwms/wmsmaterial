/* EZPICKUP WMS — Frontend-only Management Reporting Center
 * Uses only existing GET ?only=... datasets already supported by the current app.
 * No new Apps Script action, no backend schema/code changes.
 */
(function () {
  "use strict";
  var state = { initialized: false, report: null, charts: {}, busy: false };
  var $ = function (id) { return document.getElementById(id); };
  var ACTIVE_PO = ["OPEN", "APPROVED", "PARTIAL"];
  var DONE_CASE = ["Selesai", "Ditutup"];

  function escHtml(v) {
    return String(v == null ? "" : v).replace(/[&<>"']/g, function (c) {
      return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c];
    });
  }
  function n(v) { var x = Number(v); return isFinite(x) ? x : 0; }
  function fmt(v) { return n(v).toLocaleString("id-ID", { maximumFractionDigits: 2 }); }
  function money(v) { return "Rp " + Math.round(n(v)).toLocaleString("id-ID"); }
  function pct(v) {
    if (v === null || v === undefined || !isFinite(Number(v))) return "—";
    return (Number(v) > 0 ? "+" : "") + Number(v).toFixed(1) + "%";
  }
  function dateOnly(v) {
    if (typeof toWIBDateStr === "function") return toWIBDateStr(v);
    return v ? String(v).slice(0, 10) : "";
  }
  function today() {
    return typeof getWIBDateString === "function" ? getWIBDateString() : new Date().toISOString().slice(0, 10);
  }
  function shiftDate(value, amount) {
    var d = new Date(value + "T12:00:00+07:00");
    d.setDate(d.getDate() + amount);
    return [d.getFullYear(), String(d.getMonth() + 1).padStart(2, "0"), String(d.getDate()).padStart(2, "0")].join("-");
  }
  function monthStart(value) { return value.slice(0, 8) + "01"; }
  function monthEnd(value) { return shiftDate(shiftDate(value.slice(0, 8) + "01", 31), -1); }
  function inRange(dateValue, from, to) {
    var d = dateOnly(dateValue);
    return !!d && d >= from && d <= to;
  }
  function selectedWarehouse() {
    return $("report-warehouse") ? $("report-warehouse").value : "";
  }
  function userWarehouse() {
    return currentUser && currentUser.resolvedWarehouseId ? currentUser.resolvedWarehouseId : "";
  }
  function matchWarehouse(row, fields) {
    var wh = selectedWarehouse() || userWarehouse();
    if (!wh) return true;
    return fields.some(function (field) { return row && row[field] === wh; });
  }
  function fileBase() {
    var r = state.report;
    return "EZPICKUP-Laporan-" + (r ? r.startDate + "-sd-" + r.endDate : today());
  }
  function toast(msg, kind) {
    if (typeof showToast === "function") showToast(msg, kind || "info");
    else window.alert(msg);
  }
  function ensureScript(id, src) {
    return new Promise(function (resolve, reject) {
      var old = document.getElementById(id);
      if (old) {
        if (old.dataset.ready === "1") return resolve();
        old.addEventListener("load", resolve, { once: true });
        old.addEventListener("error", reject, { once: true });
        return;
      }
      var s = document.createElement("script");
      s.id = id;
      s.src = src;
      s.async = true;
      s.onload = function () { s.dataset.ready = "1"; resolve(); };
      s.onerror = function () { reject(new Error("Library ekspor tidak berhasil dimuat.")); };
      document.head.appendChild(s);
    });
  }
  function rangeDays(from, to) {
    var out = [], cursor = from, guard = 0;
    while (cursor <= to && guard < 370) {
      out.push(cursor);
      cursor = shiftDate(cursor, 1);
      guard++;
    }
    return out;
  }
  function previousRange(from, to, preset) {
    var fromMonth = from.slice(0, 7), toMonth = to.slice(0, 7);
    var isFullMonth = from === monthStart(from) && to === monthEnd(from);
    if (isFullMonth) {
      var firstPrev = shiftDate(from, -1);
      return { from: monthStart(firstPrev), to: monthEnd(firstPrev) };
    }
    if (preset === "monthly" && from === monthStart(from) && to === today()) {
      var prevMonthDay = Math.min(Number(to.slice(8, 10)), Number(monthEnd(shiftDate(from, -1)).slice(8, 10)));
      var prevTo = shiftDate(monthStart(shiftDate(from, -1)), prevMonthDay - 1);
      return { from: monthStart(prevTo), to: prevTo };
    }
    var count = rangeDays(from, to).length;
    return { from: shiftDate(from, -count), to: shiftDate(from, -1) };
  }
  function byDateRows(rows, field, from, to, whFields) {
    return (rows || []).filter(function (x) {
      return inRange(x[field], from, to) && (!whFields || matchWarehouse(x, whFields));
    });
  }
  function packQty(row) {
    if (typeof getPackQty === "function") return getPackQty(row);
    var q = n(row && row["Qty Paket"]);
    return q > 0 ? q : 1;
  }
  function buildMarkup() {
    var root = $("tab-reports");
    if (!root || state.initialized) return;
    root.innerHTML = [
      '<div class="flex flex-wrap items-start justify-between gap-3 mb-5">',
        '<div><h2 class="font-display text-2xl font-bold text-slate-900">Laporan &amp; Analisa</h2>',
        '<p class="text-sm text-slate-500 mt-1">Laporan mingguan dan bulanan langsung dari data WMS yang sudah tersedia.</p></div>',
        '<div class="flex flex-wrap gap-2 no-print">',
          '<button id="report-xlsx" onclick="exportReportXlsx()" disabled class="px-3 py-2 rounded-lg bg-emerald-600 text-white text-xs font-semibold disabled:opacity-40"><i data-lucide="file-spreadsheet" class="w-4 h-4 inline mr-1"></i>Excel</button>',
          '<button id="report-pptx" onclick="exportReportPptx()" disabled class="px-3 py-2 rounded-lg bg-red-600 text-white text-xs font-semibold disabled:opacity-40"><i data-lucide="presentation" class="w-4 h-4 inline mr-1"></i>PowerPoint</button>',
          '<button id="report-pdf" onclick="printReportPdf()" disabled class="px-3 py-2 rounded-lg border border-slate-300 bg-white text-slate-700 text-xs font-semibold disabled:opacity-40"><i data-lucide="file-pdf" class="w-4 h-4 inline mr-1"></i>PDF / Cetak</button>',
          '<button id="report-csv" onclick="exportReportCsv()" disabled class="px-3 py-2 rounded-lg border border-slate-300 bg-white text-slate-700 text-xs font-semibold disabled:opacity-40"><i data-lucide="download" class="w-4 h-4 inline mr-1"></i>CSV Transaksi</button>',
        '</div>',
      '</div>',
      '<div class="bg-white rounded-xl border border-slate-200 shadow-sm p-4 mb-5 no-print">',
        '<div class="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-5 gap-3 items-end">',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Jenis Laporan</label><select id="report-period" onchange="setReportPreset(this.value)" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"><option value="monthly">Bulan berjalan</option><option value="previous-month">Bulan lalu (penuh)</option><option value="weekly">Minggu berjalan</option><option value="previous-week">Minggu lalu</option><option value="custom">Rentang khusus</option></select></div>',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Dari Tanggal</label><input id="report-from" type="date" onchange="setReportCustom()" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"></div>',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Sampai Tanggal</label><input id="report-to" type="date" onchange="setReportCustom()" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"></div>',
          '<div><label class="block text-[10px] uppercase tracking-wide font-bold text-slate-400 mb-1">Gudang</label><select id="report-warehouse" onchange="generateFrontendReport()" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm"><option value="">Semua Gudang</option></select></div>',
          '<button id="report-run" onclick="generateFrontendReport()" class="w-full px-4 py-2 rounded-lg bg-slate-900 text-white text-sm font-semibold hover:bg-slate-700"><i data-lucide="refresh-cw" class="w-4 h-4 inline mr-1"></i>Generate Laporan</button>',
        '</div>',
        '<div id="report-status" class="text-xs text-slate-400 mt-3">Perhitungan dilakukan di frontend menggunakan data WMS yang dimuat melalui jalur data yang sudah ada.</div>',
      '</div>',
      '<div id="report-print-area"><div id="report-results" class="hidden space-y-5">',
        '<div class="flex flex-wrap justify-between items-center gap-2"><div><h3 class="font-display text-lg font-bold text-slate-900">Ringkasan Eksekutif</h3><p id="report-period-label" class="text-xs text-slate-500 mt-0.5"></p></div><span id="report-wh-label" class="text-xs font-semibold text-slate-600 bg-slate-100 rounded-full px-3 py-1"></span></div>',
        '<div id="report-kpis" class="grid grid-cols-2 xl:grid-cols-4 gap-3"></div>',
        '<div class="grid grid-cols-1 xl:grid-cols-2 gap-4">',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Tren Inbound &amp; Outbound</h3><div class="h-64"><canvas id="report-chart-movement"></canvas></div></div>',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Paket Packing &amp; Produktivitas</h3><div class="h-64"><canvas id="report-chart-volume"></canvas></div></div>',
        '</div>',
        '<div class="bg-white rounded-xl border border-slate-200 p-4"><div class="flex items-center justify-between gap-2 mb-3"><h3 class="font-bold text-sm text-slate-800">Analisa &amp; Sinyal Operasional</h3><span class="text-[10px] text-slate-400">Dihitung otomatis dari data</span></div><div id="report-insights" class="grid grid-cols-1 md:grid-cols-2 gap-2"></div></div>',
        '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Perbandingan Performa Gudang</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 uppercase bg-slate-50"><tr><th class="p-2 text-left">Gudang</th><th class="p-2 text-right">Inbound</th><th class="p-2 text-right">Outbound</th><th class="p-2 text-right">Paket</th><th class="p-2 text-right">Produktivitas</th><th class="p-2 text-right report-money-col">Cost Packing</th></tr></thead><tbody id="report-warehouse-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
        '<div class="grid grid-cols-1 xl:grid-cols-2 gap-4">',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Top Konsumsi Material (OUT)</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 bg-slate-50"><tr><th class="p-2 text-left">Material</th><th class="p-2 text-left">Gudang</th><th class="p-2 text-right">Qty OUT</th></tr></thead><tbody id="report-material-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Kategori Packing</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 bg-slate-50"><tr><th class="p-2 text-left">Kategori</th><th class="p-2 text-right">AWB</th><th class="p-2 text-right">Paket</th><th class="p-2 text-right report-money-col">Cost</th></tr></thead><tbody id="report-category-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
        '</div>',
        '<div class="grid grid-cols-1 xl:grid-cols-2 gap-4">',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Produktivitas per Peran</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead class="text-slate-400 bg-slate-50"><tr><th class="p-2 text-left">Peran</th><th class="p-2 text-right">Paket</th><th class="p-2 text-right">Pcs</th><th class="p-2 text-right">Entri</th></tr></thead><tbody id="report-role-rows" class="divide-y divide-slate-100"></tbody></table></div></div>',
          '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-3">Purchase Order &amp; Case</h3><div id="report-ops-cards" class="grid grid-cols-2 gap-2"></div></div>',
        '</div>',
        '<div class="bg-white rounded-xl border border-slate-200 p-4"><h3 class="font-bold text-sm text-slate-800 mb-2">Catatan Data &amp; Metodologi</h3><ul id="report-notes" class="text-xs text-slate-500 list-disc pl-5 space-y-1"></ul></div>',
        '<div class="text-[10px] text-slate-400 text-right">EZPICKUP WMS • Management Reporting Center • Internal Use</div>',
      '</div></div>'
    ].join("");
    state.initialized = true;
    initWarehouseOptions();
    window.setReportPreset("monthly");
    if (typeof lucide !== "undefined") lucide.createIcons();
  }
  function initWarehouseOptions() {
    var el = $("report-warehouse");
    if (!el) return;
    var keep = el.value;
    var rows = (db && db.warehouses) || [];
    var scoped = userWarehouse();
    var options = ['<option value="">Semua Gudang</option>'];
    rows.forEach(function (w) {
      if (scoped && w["Warehouse ID"] !== scoped) return;
      var id = w["Warehouse ID"];
      if (id) options.push('<option value="' + escHtml(id) + '">' + escHtml(id + " — " + (w["Warehouse Name"] || "")) + "</option>");
    });
    el.innerHTML = options.join("");
    if (scoped) { el.value = scoped; el.disabled = true; el.classList.add("bg-slate-100"); }
    else if (keep && Array.from(el.options).some(function (o) { return o.value === keep; })) el.value = keep;
  }

  window.setReportCustom = function () {
    if (!state.initialized) return;
    $("report-period").value = "custom";
  };
  window.setReportPreset = function (preset) {
    if (!state.initialized) buildMarkup();
    var t = today(), from = t, to = t;
    if (preset === "monthly") {
      from = monthStart(t);
    } else if (preset === "previous-month") {
      var prevDate = shiftDate(monthStart(t), -1);
      from = monthStart(prevDate); to = monthEnd(prevDate);
    } else if (preset === "weekly") {
      var dow = (new Date(t + "T12:00:00+07:00").getDay() + 6) % 7;
      from = shiftDate(t, -dow);
    } else if (preset === "previous-week") {
      var todayDow = (new Date(t + "T12:00:00+07:00").getDay() + 6) % 7;
      to = shiftDate(t, -todayDow - 1); from = shiftDate(to, -6);
    } else if (preset === "custom") {
      from = $("report-from").value || t; to = $("report-to").value || t;
    }
    $("report-period").value = preset;
    $("report-from").value = from;
    $("report-to").value = to;
    if (preset !== "custom" && state.report) window.generateFrontendReport();
  };

  function getPeriodData(from, to) {
    var wh = selectedWarehouse() || userWarehouse();
    var tx = byDateRows(db.transactionLogs, "Date", from, to, ["Origin WH ID", "Dest WH ID"]);
    var packs = byDateRows(db.packLogs, "Date", from, to, ["Warehouse ID"]);
    var prod = byDateRows(db.productivityLogs, "Date", from, to, ["Warehouse ID"]);
    var pos = (db.purchaseOrders || []).filter(function (p) {
      return inRange(p["Order Date"], from, to) && (!wh || p["Destination WH ID"] === wh);
    });
    var casesNew = (db.operationalCases || []).filter(function (c) {
      return inRange(c["Tanggal Lapor"], from, to) && (!wh || c["Warehouse ID"] === wh);
    });
    var casesClosed = (db.operationalCases || []).filter(function (c) {
      return inRange(c["Waktu Selesai"], from, to) && (!wh || c["Warehouse ID"] === wh) && DONE_CASE.indexOf(c["Status"]) !== -1;
    });
    var roleMap = {};
    var totalProdPackages = 0, totalProdPcs = 0;
    prod.forEach(function (p) {
      var role = String(p["Role"] || "Lainnya");
      if (!roleMap[role]) roleMap[role] = { role: role, packages: 0, pcs: 0, entries: 0 };
      roleMap[role].packages += n(p["Qty Paket"]);
      roleMap[role].pcs += n(p["Qty Pcs"]);
      roleMap[role].entries++;
      if (role === "Admin Print") {
        totalProdPackages += n(p["Qty Paket"]);
        totalProdPcs += n(p["Qty Pcs"]);
      }
    });
    var movement = { inboundQty: 0, outboundQty: 0, transferQty: 0, adjustmentQty: 0, inboundCount: 0, outboundCount: 0 };
    var byDay = {};
    rangeDays(from, to).forEach(function (d) { byDay[d] = { date: d, inboundQty: 0, outboundQty: 0, packages: 0, productivityPackages: 0 }; });
    var whStats = {};
    function whStat(id) {
      id = id || "—";
      if (!whStats[id]) {
        var nameRow = (db.warehouses || []).find(function (w) { return w["Warehouse ID"] === id; });
        whStats[id] = { warehouseId: id, warehouseName: nameRow ? (nameRow["Warehouse Name"] || id) : id, inboundQty: 0, outboundQty: 0, packages: 0, productivityPackages: 0, packingCost: 0 };
      }
      return whStats[id];
    }
    (db.warehouses || []).forEach(function (w) {
      if (!wh || w["Warehouse ID"] === wh) whStat(w["Warehouse ID"]);
    });
    tx.forEach(function (t) {
      var type = String(t["Trx Type"] || "").toUpperCase();
      var qty = n(t["Qty"]);
      var ds = dateOnly(t["Date"]);
      if (type === "IN" && (!wh || t["Dest WH ID"] === wh)) {
        movement.inboundQty += qty; movement.inboundCount++;
        whStat(t["Dest WH ID"]).inboundQty += qty;
        if (byDay[ds]) byDay[ds].inboundQty += qty;
      } else if (type === "OUT" && (!wh || t["Origin WH ID"] === wh)) {
        movement.outboundQty += qty; movement.outboundCount++;
        whStat(t["Origin WH ID"]).outboundQty += qty;
        if (byDay[ds]) byDay[ds].outboundQty += qty;
      } else if (type === "TRANSFER" && (!wh || t["Origin WH ID"] === wh || t["Dest WH ID"] === wh)) {
        movement.transferQty += qty;
      } else if (type === "ADJUSTMENT" && (!wh || t["Origin WH ID"] === wh || t["Dest WH ID"] === wh)) {
        movement.adjustmentQty += qty;
      }
    });
    var totalPackages = 0, totalAwb = 0, totalPackingCost = 0, categoryMap = {}, materialMap = {};
    packs.forEach(function (p) {
      var qty = packQty(p), cost = n(p["Total Packing Cost (Rp)"]);
      totalPackages += qty; totalAwb++;
      totalPackingCost += cost;
      var whId = p["Warehouse ID"];
      whStat(whId).packages += qty;
      whStat(whId).packingCost += cost;
      var cat = p["Pack Category Code"] || "Tanpa Kategori";
      if (!categoryMap[cat]) categoryMap[cat] = { category: cat, awb: 0, packages: 0, packingCost: 0 };
      categoryMap[cat].awb++; categoryMap[cat].packages += qty; categoryMap[cat].packingCost += cost;
    });
    var roleRows = Object.keys(roleMap).map(function (k) { return roleMap[k]; }).sort(function (a, b) { return b.packages - a.packages; });
    prod.forEach(function (p) {
      var whId = p["Warehouse ID"];
      var qty = n(p["Qty Paket"]);
      if (p["Role"] === "Admin Print") {
        whStat(whId).productivityPackages += qty;
        var ds = dateOnly(p["Date"]);
        if (byDay[ds]) byDay[ds].productivityPackages += qty;
      }
    });
    var topMaterials = {};
    tx.forEach(function (t) {
      if (String(t["Trx Type"] || "").toUpperCase() !== "OUT") return;
      var id = t["Material ID"] || "—", whId = t["Origin WH ID"] || "—";
      var key = id + "||" + whId;
      if (!topMaterials[key]) topMaterials[key] = { materialId: id, materialName: t["Material Name"] || id, warehouseId: whId, qtyOut: 0 };
      topMaterials[key].qtyOut += n(t["Qty"]);
    });
    var poStatuses = { OPEN: 0, APPROVED: 0, PARTIAL: 0, COMPLETED: 0, CANCELLED: 0 };
    (db.purchaseOrders || []).forEach(function (p) {
      if (wh && p["Destination WH ID"] !== wh) return;
      var st = String(p["PO Status"] || "OPEN").toUpperCase();
      if (poStatuses[st] !== undefined) poStatuses[st]++;
    });
    var openPos = (db.purchaseOrders || []).filter(function (p) {
      return (!wh || p["Destination WH ID"] === wh) && ACTIVE_PO.indexOf(String(p["PO Status"] || "OPEN").toUpperCase()) !== -1;
    });
    var overduePos = openPos.filter(function (p) { return typeof poIsOverdue === "function" ? poIsOverdue(p) : (p["Expected Date"] && dateOnly(p["Expected Date"]) < today()); });
    var poCompleted = (db.purchaseOrders || []).filter(function (p) {
      return (!wh || p["Destination WH ID"] === wh) && String(p["PO Status"]).toUpperCase() === "COMPLETED" && inRange(p["Actual Date"], from, to);
    });
    var allCases = (db.operationalCases || []).filter(function (c) { return !wh || c["Warehouse ID"] === wh; });
    var openCases = allCases.filter(function (c) { return ["Terbuka", "Diproses"].indexOf(c["Status"]) !== -1; });
    var highCases = openCases.filter(function (c) { return ["Tinggi", "High", "Critical", "Kritis"].indexOf(String(c["Prioritas"] || "")) !== -1; });
    var inventory = (db.inventory || []).filter(function (i) { return !wh || i["Warehouse ID"] === wh; });
    var invSummary = {
      skuCount: inventory.length,
      criticalCount: inventory.filter(function (i) { return i["Status Alert"] === "CRITICAL"; }).length,
      reorderCount: inventory.filter(function (i) { return i["Status Alert"] === "REORDER"; }).length,
      valuation: inventory.reduce(function (s, i) { return s + n(i["Total Valuation (Rp)"]); }, 0)
    };
    return {
      startDate: from, endDate: to, warehouseId: wh || "ALL",
      transactions: tx, packLogs: packs, productivityLogs: prod, purchaseOrders: pos, newCases: casesNew,
      metrics: {
        movement: movement, totalAWB: totalAwb, totalPackages: totalPackages, totalPackingCost: totalPackingCost,
        productivityPackages: totalProdPackages, productivityPcs: totalProdPcs, productivityEntries: prod.length
      },
      inventoryNow: invSummary,
      purchaseOrderSummary: {
        createdInPeriod: pos.length, completedInPeriod: poCompleted.length,
        outstandingNow: openPos.length, overdueNow: overduePos.length, statusesNow: poStatuses
      },
      caseSummary: { newInPeriod: casesNew.length, closedInPeriod: casesClosed.length, openNow: openCases.length, highPriorityOpenNow: highCases.length },
      warehousePerformance: Object.keys(whStats).map(function (k) { return whStats[k]; }).sort(function (a, b) { return a.warehouseName.localeCompare(b.warehouseName); }),
      topMaterials: Object.keys(topMaterials).map(function (k) { return topMaterials[k]; }).sort(function (a, b) { return b.qtyOut - a.qtyOut; }).slice(0, 10),
      topCategories: Object.keys(categoryMap).map(function (k) { return categoryMap[k]; }).sort(function (a, b) { return b.packages - a.packages; }).slice(0, 10),
      byRole: roleRows, dailyTrend: Object.keys(byDay).map(function (k) { return byDay[k]; }),
      previousRange: previousRange(from, to, $("report-period").value),
      visibility: { financial: typeof canSeeNominal === "function" ? canSeeNominal() : true }
    };
  }

  function pctChange(current, prior) {
    if (!prior && !current) return 0;
    if (!prior) return null;
    return ((current - prior) / Math.abs(prior)) * 100;
  }
  function compareMetric(nowValue, prevValue) { return pctChange(n(nowValue), n(prevValue)); }
  function makeInsights(r, prev) {
    var m = r.metrics, p = prev.metrics, insights = [];
    function add(level, title, detail) { insights.push({ level: level, title: title, detail: detail }); }
    var inbound = m.movement.inboundQty, outbound = m.movement.outboundQty;
    if (outbound > inbound && outbound > 0) add("warning", "Outbound lebih besar dari inbound", "Qty OUT " + fmt(outbound) + " lebih tinggi " + fmt(outbound - inbound) + " dibanding Qty IN pada periode ini. Pantau stok dan rencana replenishment.");
    else if (inbound > outbound) add("success", "Inbound lebih tinggi dari outbound", "Qty IN melampaui Qty OUT sebesar " + fmt(inbound - outbound) + ". Pastikan kapasitas penyimpanan dan kebutuhan pemakaian tetap seimbang.");
    else add("info", "Arus inbound dan outbound seimbang", "Qty IN dan OUT tercatat sama atau belum ada pergerakan pada periode ini.");
    if (r.purchaseOrderSummary.overdueNow > 0) add("warning", "PO melewati estimasi kedatangan", r.purchaseOrderSummary.overdueNow + " PO aktif sudah melewati Expected Date. Prioritaskan konfirmasi supplier dan dampak ke ketersediaan material.");
    else add("success", "Tidak ada PO aktif yang terlambat", "Tidak ditemukan PO berstatus OPEN, APPROVED, atau PARTIAL yang melewati Expected Date saat laporan dibuat.");
    if (r.caseSummary.highPriorityOpenNow > 0) add("warning", "Case prioritas tinggi masih terbuka", r.caseSummary.highPriorityOpenNow + " case prioritas tinggi masih berstatus Terbuka/Diproses. Periksa PIC dan target penyelesaian.");
    if (r.inventoryNow.criticalCount > 0) add("warning", "SKU berstatus critical", r.inventoryNow.criticalCount + " SKU pada posisi stok saat ini berstatus CRITICAL.");
    if (p && p.metrics) {
      var dp = compareMetric(m.totalPackages, p.metrics.totalPackages);
      if (dp !== null) {
        if (dp >= 10) add("success", "Volume packing meningkat", "Total paket " + pct(dp) + " dibanding periode pembanding.");
        else if (dp <= -10) add("warning", "Volume packing menurun", "Total paket " + pct(dp) + " dibanding periode pembanding. Telusuri volume order, kapasitas, dan kelengkapan input.");
        else add("info", "Volume packing relatif stabil", "Perubahan total paket " + pct(dp) + " dibanding periode pembanding.");
      } else add("info", "Data pembanding belum tersedia", "Periode sebelumnya belum memiliki data paket sehingga persentase perubahan tidak dihitung.");
    }
    if (!insights.length) add("info", "Data periode tersedia", "Belum ada sinyal anomali dari indikator sederhana yang dihitung.");
    return insights;
  }
  function computeReport(from, to) {
    var current = getPeriodData(from, to);
    var prevRange = current.previousRange;
    var previous = getPeriodData(prevRange.from, prevRange.to);
    current.previousStartDate = prevRange.from;
    current.previousEndDate = prevRange.to;
    current.previous = previous;
    current.changes = {
      inbound: compareMetric(current.metrics.movement.inboundQty, previous.metrics.movement.inboundQty),
      outbound: compareMetric(current.metrics.movement.outboundQty, previous.metrics.movement.outboundQty),
      packages: compareMetric(current.metrics.totalPackages, previous.metrics.totalPackages),
      productivity: compareMetric(current.metrics.productivityPackages, previous.metrics.productivityPackages),
      packingCost: current.visibility.financial ? compareMetric(current.metrics.totalPackingCost, previous.metrics.totalPackingCost) : null
    };
    current.insights = makeInsights(current, previous);
    return current;
  }
  function kpi(label, value, change, note) {
    var cls = change === null || change === undefined ? "text-slate-400" : change > 0 ? "text-emerald-600" : change < 0 ? "text-red-600" : "text-slate-500";
    return '<div class="bg-white rounded-xl border border-slate-200 p-4 shadow-sm"><p class="text-[10px] font-bold uppercase tracking-wide text-slate-400">' + escHtml(label) + '</p><p class="font-display text-xl sm:text-2xl font-bold text-slate-900 mt-2 break-words">' + escHtml(value) + '</p><p class="text-[10px] mt-1 ' + cls + '">' + escHtml(change === null || change === undefined ? "Pembanding tidak tersedia" : pct(change) + " vs periode sebelumnya") + '</p>' + (note ? '<p class="text-[10px] text-slate-400 mt-1">' + escHtml(note) + "</p>" : "") + "</div>";
  }
  function tableRows(id, rows, cols) {
    var el = $(id);
    if (!el) return;
    el.innerHTML = rows.length ? rows.join("") : '<tr><td colspan="' + cols + '" class="p-4 text-center text-slate-400">Tidak ada data dalam periode ini.</td></tr>';
  }
  function render(r) {
    var m = r.metrics, ps = r.purchaseOrderSummary, cs = r.caseSummary, inv = r.inventoryNow;
    var changes = r.changes, vis = r.visibility.financial;
    $("report-results").classList.remove("hidden");
    $("report-period-label").textContent = r.startDate + " s.d. " + r.endDate + " • pembanding " + r.previousStartDate + " s.d. " + r.previousEndDate;
    $("report-wh-label").textContent = r.warehouseId === "ALL" ? "Semua Gudang" : "Gudang " + r.warehouseId;
    var cards = [
      kpi("Inbound Qty", fmt(m.movement.inboundQty), changes.inbound, "Transaksi IN"),
      kpi("Outbound Qty", fmt(m.movement.outboundQty), changes.outbound, "Transaksi OUT"),
      kpi("Paket Packing", fmt(m.totalPackages), changes.packages, fmt(m.totalAWB) + " AWB"),
      kpi("Produktivitas Admin Print", fmt(m.productivityPackages) + " pkt", changes.productivity, fmt(m.productivityPcs) + " pcs")
    ];
    if (vis) cards.push(kpi("Cost Packing", money(m.totalPackingCost), changes.packingCost, "Nilai yang tersimpan di Pack Log"));
    cards.push(kpi("SKU Critical", fmt(inv.criticalCount), null, fmt(inv.skuCount) + " SKU posisi saat ini"));
    cards.push(kpi("PO Terlambat", fmt(ps.overdueNow), null, "Outstanding saat ini"));
    cards.push(kpi("Case Terbuka", fmt(cs.openNow), null, fmt(cs.newInPeriod) + " case baru di periode"));
    $("report-kpis").innerHTML = cards.join("");
    document.querySelectorAll("#tab-reports .report-money-col").forEach(function (el) { el.classList.toggle("hidden", !vis); });
    renderCharts(r);
    $("report-insights").innerHTML = r.insights.map(function (x) {
      var style = x.level === "warning" ? "border-amber-200 bg-amber-50" : x.level === "success" ? "border-emerald-200 bg-emerald-50" : "border-blue-100 bg-blue-50";
      var icon = x.level === "warning" ? "alert-triangle" : x.level === "success" ? "check-circle-2" : "info";
      return '<div class="border rounded-lg p-3 ' + style + '"><p class="text-xs font-bold text-slate-800"><i data-lucide="' + icon + '" class="w-3.5 h-3.5 inline mr-1"></i>' + escHtml(x.title) + '</p><p class="text-xs text-slate-600 mt-1">' + escHtml(x.detail) + "</p></div>";
    }).join("");
    tableRows("report-warehouse-rows", r.warehousePerformance.map(function (w) {
      return "<tr><td class=\"p-2 font-semibold\">" + escHtml(w.warehouseName) + '</td><td class="p-2 text-right">' + fmt(w.inboundQty) + '</td><td class="p-2 text-right">' + fmt(w.outboundQty) + '</td><td class="p-2 text-right">' + fmt(w.packages) + '</td><td class="p-2 text-right">' + fmt(w.productivityPackages) + '</td><td class="p-2 text-right report-money-col">' + (vis ? money(w.packingCost) : "—") + "</td></tr>";
    }), 6);
    tableRows("report-material-rows", r.topMaterials.map(function (x) {
      return "<tr><td class=\"p-2 font-semibold\">" + escHtml(x.materialName) + '<div class="text-[10px] text-slate-400">' + escHtml(x.materialId) + '</div></td><td class="p-2">' + escHtml(x.warehouseId) + '</td><td class="p-2 text-right font-semibold">' + fmt(x.qtyOut) + "</td></tr>";
    }), 3);
    tableRows("report-category-rows", r.topCategories.map(function (x) {
      return "<tr><td class=\"p-2 font-semibold\">" + escHtml(x.category) + '</td><td class="p-2 text-right">' + fmt(x.awb) + '</td><td class="p-2 text-right">' + fmt(x.packages) + '</td><td class="p-2 text-right report-money-col">' + (vis ? money(x.packingCost) : "—") + "</td></tr>";
    }), 4);
    tableRows("report-role-rows", r.byRole.map(function (x) {
      return "<tr><td class=\"p-2 font-semibold\">" + escHtml(x.role) + '</td><td class="p-2 text-right">' + fmt(x.packages) + '</td><td class="p-2 text-right">' + fmt(x.pcs) + '</td><td class="p-2 text-right">' + fmt(x.entries) + "</td></tr>";
    }), 4);
    var ops = [
      ["PO dibuat periode ini", ps.createdInPeriod], ["PO selesai periode ini", ps.completedInPeriod],
      ["PO outstanding sekarang", ps.outstandingNow], ["PO terlambat sekarang", ps.overdueNow],
      ["Case baru periode ini", cs.newInPeriod], ["Case selesai periode ini", cs.closedInPeriod],
      ["Case masih terbuka", cs.openNow], ["Case prioritas tinggi terbuka", cs.highPriorityOpenNow]
    ];
    $("report-ops-cards").innerHTML = ops.map(function (x) {
      return '<div class="rounded-lg bg-slate-50 border border-slate-100 p-3"><p class="text-[10px] text-slate-400">' + escHtml(x[0]) + '</p><p class="text-lg font-bold text-slate-800 mt-1">' + fmt(x[1]) + "</p></div>";
    }).join("");
    var notes = [
      "Semua perhitungan dilakukan di frontend dari dataset aktif yang dimuat melalui jalur GET ?only= yang sudah tersedia pada aplikasi. Tidak ada action backend baru.",
      "Data yang sudah dipindahkan ke sheet arsip hanya ikut jika endpoint data WMS saat ini memang mengembalikannya. Laporan ini tidak menggabungkan arsip secara terpisah.",
      "SKU, status critical/reorder, dan valuasi inventory adalah posisi saat laporan dibuat, bukan saldo historis pada akhir bulan.",
      "Produktivitas paket memakai entri peran Admin Print sebagai titik awal pipeline. Rekap peran lain ditampilkan terpisah dan tidak dijumlahkan ulang sebagai paket unik.",
      "Cost Packing berasal dari nilai Total Packing Cost (Rp) yang tercatat di Pack Log. Biaya historis transaksi OUT tidak dihitung memakai harga material hari ini."
    ];
    $("report-notes").innerHTML = notes.map(function (x) { return "<li>" + escHtml(x) + "</li>"; }).join("");
    if (typeof lucide !== "undefined") lucide.createIcons();
  }

  function destroyChart(key) {
    if (state.charts[key]) { try { state.charts[key].destroy(); } catch (e) {} state.charts[key] = null; }
  }
  function renderCharts(r) {
    if (typeof Chart === "undefined") return;
    var rows = r.dailyTrend;
    var labels = rows.map(function (x) { return x.date.slice(5); });
    destroyChart("movement"); destroyChart("volume");
    state.charts.movement = new Chart($("report-chart-movement"), {
      type: "line",
      data: { labels: labels, datasets: [
        { label: "Inbound", data: rows.map(function (x) { return x.inboundQty; }), borderColor: "#2563eb", backgroundColor: "rgba(37,99,235,.08)", tension: .25 },
        { label: "Outbound", data: rows.map(function (x) { return x.outboundQty; }), borderColor: "#dc2626", backgroundColor: "rgba(220,38,38,.08)", tension: .25 }
      ] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } }, scales: { y: { beginAtZero: true } } }
    });
    state.charts.volume = new Chart($("report-chart-volume"), {
      type: "bar",
      data: { labels: labels, datasets: [
        { label: "Paket Packing", data: rows.map(function (x) { return x.packages; }), backgroundColor: "#dc2626" },
        { label: "Produktivitas Admin Print", data: rows.map(function (x) { return x.productivityPackages; }), backgroundColor: "#0f766e" }
      ] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } }, scales: { y: { beginAtZero: true } } }
    });
  }

  window.generateFrontendReport = function () {
    if (!state.initialized) buildMarkup();
    if (state.busy) return;
    var from = $("report-from").value, to = $("report-to").value;
    if (!from || !to || from > to) return toast("Periksa rentang tanggal laporan.", "error");
    var days = rangeDays(from, to).length;
    if (days > 366) return toast("Pilih rentang maksimal 12 bulan agar proses laporan tetap ringan.", "error");
    state.busy = true;
    var btn = $("report-run"), status = $("report-status");
    btn.disabled = true;
    btn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 inline mr-1 animate-spin"></i>Menghitung...';
    try {
      state.report = computeReport(from, to);
      render(state.report);
      ["report-xlsx", "report-pptx", "report-pdf", "report-csv"].forEach(function (id) { if ($(id)) $(id).disabled = false; });
      if (status) status.textContent = "Laporan dihitung dari " + fmt((db.transactionLogs || []).length) + " transaksi, " + fmt((db.packLogs || []).length) + " log packing, dan dataset WMS yang sudah dimuat. Tidak ada permintaan API laporan baru.";
      toast("Laporan berhasil dihitung di frontend.", "success");
    } catch (err) {
      if (status) status.textContent = "Perhitungan laporan gagal: " + err.message;
      toast("Gagal menghitung laporan: " + err.message, "error");
    } finally {
      state.busy = false;
      btn.disabled = false;
      btn.innerHTML = '<i data-lucide="refresh-cw" class="w-4 h-4 inline mr-1"></i>Generate Laporan';
      if (typeof lucide !== "undefined") lucide.createIcons();
    }
  };

  function reportRowsToSheet(XLSX, rows) {
    return XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ Keterangan: "Tidak ada data dalam periode ini" }]);
  }
  function exportSafeRows(rows, financialVisible) {
    if (financialVisible) return rows || [];
    return (rows || []).map(function (row) {
      var clean = {};
      Object.keys(row || {}).forEach(function (key) {
        if (/(rp|cost|harga|nominal|valuation|nilai pembelian|landed)/i.test(key)) return;
        clean[key] = row[key];
      });
      return clean;
    });
  }
  window.exportReportXlsx = async function () {
    var r = state.report;
    if (!r) return toast("Generate laporan terlebih dahulu.", "info");
    try {
      if (!window.XLSX) await ensureScript("sheetjs-report", "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js");
      var XLSX = window.XLSX, wb = XLSX.utils.book_new(), m = r.metrics, vis = r.visibility.financial;
      var summary = [
        { Indikator: "Periode", Nilai: r.startDate + " s.d. " + r.endDate },
        { Indikator: "Gudang", Nilai: r.warehouseId },
        { Indikator: "Inbound Qty", Nilai: m.movement.inboundQty },
        { Indikator: "Outbound Qty", Nilai: m.movement.outboundQty },
        { Indikator: "Transfer Qty", Nilai: m.movement.transferQty },
        { Indikator: "Adjustment Qty", Nilai: m.movement.adjustmentQty },
        { Indikator: "AWB", Nilai: m.totalAWB },
        { Indikator: "Paket Packing", Nilai: m.totalPackages },
        { Indikator: "Produktivitas Admin Print", Nilai: m.productivityPackages },
        { Indikator: "Pcs Produktivitas Admin Print", Nilai: m.productivityPcs },
        { Indikator: "PO dibuat", Nilai: r.purchaseOrderSummary.createdInPeriod },
        { Indikator: "PO selesai", Nilai: r.purchaseOrderSummary.completedInPeriod },
        { Indikator: "PO outstanding saat ini", Nilai: r.purchaseOrderSummary.outstandingNow },
        { Indikator: "PO overdue saat ini", Nilai: r.purchaseOrderSummary.overdueNow },
        { Indikator: "Case baru", Nilai: r.caseSummary.newInPeriod },
        { Indikator: "Case terbuka saat ini", Nilai: r.caseSummary.openNow },
        { Indikator: "SKU saat ini", Nilai: r.inventoryNow.skuCount },
        { Indikator: "SKU Critical saat ini", Nilai: r.inventoryNow.criticalCount }
      ];
      if (vis) {
        summary.push({ Indikator: "Cost Packing", Nilai: m.totalPackingCost });
        summary.push({ Indikator: "Valuasi Stok Saat Ini", Nilai: r.inventoryNow.valuation });
      }
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, summary), "Executive Summary");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, r.dailyTrend), "Daily Trend");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, exportSafeRows(r.warehousePerformance, vis)), "Warehouse");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, r.topMaterials), "Top Material");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, exportSafeRows(r.topCategories, vis)), "Packing Category");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, r.byRole), "Produktivitas Role");
      if (vis) XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, r.purchaseOrders), "Detail PO");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, r.transactions), "Detail Transaksi");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, exportSafeRows(r.packLogs, vis)), "Detail Packing");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, r.productivityLogs), "Detail Produktivitas");
      XLSX.utils.book_append_sheet(wb, reportRowsToSheet(XLSX, exportSafeRows(r.newCases, vis)), "Detail Case");
      XLSX.writeFile(wb, fileBase() + ".xlsx");
      toast("Excel berhasil dibuat.", "success");
    } catch (err) { toast("Gagal membuat Excel: " + err.message, "error"); }
  };
  window.exportReportCsv = function () {
    var r = state.report;
    if (!r) return toast("Generate laporan terlebih dahulu.", "info");
    var keys = ["Trx ID", "Date", "Trx Type", "Origin WH ID", "Dest WH ID", "Material ID", "Material Name", "Qty", "Reference No", "Operator"];
    var rows = [keys].concat(r.transactions.map(function (x) { return keys.map(function (k) { return x[k] == null ? "" : x[k]; }); }));
    var csv = rows.map(function (row) { return row.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(","); }).join("\r\n");
    var blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8;" });
    var url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url; a.download = fileBase() + "-transaksi.csv";
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    toast(fmt(r.transactions.length) + " transaksi diekspor.", "success");
  };
  window.printReportPdf = function () {
    if (!state.report) return toast("Generate laporan terlebih dahulu.", "info");
    var style = $("report-print-style");
    if (!style) {
      style = document.createElement("style"); style.id = "report-print-style";
      style.textContent = "@media print{body{overflow:visible!important;background:#fff!important}body *{visibility:hidden!important}#tab-reports,#tab-reports *{visibility:visible!important}#tab-reports{display:block!important;position:absolute!important;left:0!important;top:0!important;width:100%!important;padding:0!important;background:#fff!important}#tab-reports .no-print{display:none!important}.bg-white{box-shadow:none!important}canvas{max-height:220px!important}.tab-content{display:none!important}#tab-reports{break-inside:auto!important}}";
      document.head.appendChild(style);
    }
    window.print();
  };

  function addSlideText(slide, text, x, y, w, h, opts) {
    slide.addText(String(text == null ? "" : text), Object.assign({ x: x, y: y, w: w, h: h, fontFace: "Aptos", fontSize: 11, color: "1F2937", margin: 0.04, valign: "mid" }, opts || {}));
  }
  function pptHeader(slide, title, sub, pptx) {
    slide.background = { color: "F8FAFC" };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.15, line: { color: "B91C1C", transparency: 100 }, fill: { color: "B91C1C" } });
    addSlideText(slide, title, 0.55, 0.34, 12.1, 0.4, { fontSize: 22, bold: true, color: "111827" });
    addSlideText(slide, sub, 0.57, 0.8, 12, 0.22, { fontSize: 9, color: "64748B" });
    addSlideText(slide, "EZPICKUP WMS • INTERNAL USE", 0.57, 7.18, 12, 0.16, { fontSize: 7, color: "94A3B8" });
  }
  function pptCard(slide, x, y, w, label, value, pptx, note) {
    slide.addShape(pptx.ShapeType.roundRect, { x: x, y: y, w: w, h: 1.1, rectRadius: 0.08, line: { color: "E2E8F0", width: 1 }, fill: { color: "FFFFFF" } });
    addSlideText(slide, label, x + 0.15, y + 0.12, w - 0.3, 0.2, { fontSize: 8, bold: true, color: "64748B" });
    addSlideText(slide, value, x + 0.15, y + 0.42, w - 0.3, 0.38, { fontSize: 19, bold: true });
    if (note) addSlideText(slide, note, x + 0.15, y + 0.84, w - 0.3, 0.13, { fontSize: 7, color: "94A3B8" });
  }
  window.exportReportPptx = async function () {
    var r = state.report;
    if (!r) return toast("Generate laporan terlebih dahulu.", "info");
    try {
      if (!window.PptxGenJS) await ensureScript("pptxgen-report", "https://cdn.jsdelivr.net/gh/gitbrent/PptxGenJS@4.0.1/dist/pptxgen.bundle.js");
      var Pptx = window.PptxGenJS;
      if (!Pptx) throw new Error("Library PowerPoint tidak tersedia.");
      var pptx = new Pptx();
      pptx.layout = "LAYOUT_WIDE"; pptx.author = "EZPICKUP WMS"; pptx.company = "EZPICKUP";
      pptx.title = "Laporan Operasional " + r.startDate + " s.d. " + r.endDate; pptx.lang = "id-ID";
      var m = r.metrics, move = m.movement, sub = r.startDate + " s.d. " + r.endDate + " | " + (r.warehouseId === "ALL" ? "Semua Gudang" : r.warehouseId);
      var slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.2, line: { color: "B91C1C", transparency: 100 }, fill: { color: "B91C1C" } });
      addSlideText(slide, "EZPICKUP WMS", 0.8, 1.3, 8, 0.4, { fontSize: 18, bold: true, color: "B91C1C" });
      addSlideText(slide, "LAPORAN OPERASIONAL", 0.8, 2.05, 11.8, 0.8, { fontSize: 34, bold: true, color: "111827" });
      addSlideText(slide, r.startDate + " — " + r.endDate, 0.82, 3.05, 11, 0.4, { fontSize: 20, color: "475569" });
      addSlideText(slide, sub, 0.84, 3.55, 11, 0.3, { fontSize: 12, color: "64748B" });
      addSlideText(slide, "Management Reporting Center • Internal Use", 0.84, 6.75, 11, 0.25, { fontSize: 9, color: "94A3B8" });

      slide = pptx.addSlide(); pptHeader(slide, "Executive Summary", sub, pptx);
      var items = [
        ["Inbound Qty", fmt(move.inboundQty)], ["Outbound Qty", fmt(move.outboundQty)],
        ["Paket Packing", fmt(m.totalPackages)], ["Produktivitas Admin Print", fmt(m.productivityPackages) + " pkt"]
      ];
      items.forEach(function (x, i) { pptCard(slide, 0.65 + (i % 2) * 6.1, 1.4 + Math.floor(i / 2) * 1.35, 5.75, x[0], x[1], pptx); });
      addSlideText(slide, "Perbandingan dengan " + r.previousStartDate + " s.d. " + r.previousEndDate, 0.75, 4.4, 11.8, 0.25, { fontSize: 11, bold: true });
      addSlideText(slide, "Inbound " + pct(r.changes.inbound) + " | Outbound " + pct(r.changes.outbound) + " | Paket " + pct(r.changes.packages) + " | Produktivitas " + pct(r.changes.productivity), 0.75, 4.8, 11.8, 0.5, { fontSize: 12, color: "475569" });

      slide = pptx.addSlide(); pptHeader(slide, "Pergerakan Inventory", sub, pptx);
      pptCard(slide, 0.65, 1.4, 3.85, "SKU Posisi Saat Ini", fmt(r.inventoryNow.skuCount), pptx);
      pptCard(slide, 4.75, 1.4, 3.85, "SKU Critical", fmt(r.inventoryNow.criticalCount), pptx);
      pptCard(slide, 8.85, 1.4, 3.85, "SKU Reorder", fmt(r.inventoryNow.reorderCount), pptx);
      addSlideText(slide, "Inbound: " + fmt(move.inboundQty) + " | Outbound: " + fmt(move.outboundQty) + " | Transfer: " + fmt(move.transferQty) + " | Adjustment: " + fmt(move.adjustmentQty), 0.8, 3.2, 12, 0.5, { fontSize: 13 });
      if (r.visibility.financial) addSlideText(slide, "Valuasi stok posisi saat laporan dibuat: " + money(r.inventoryNow.valuation), 0.8, 4.0, 12, 0.3, { fontSize: 12, bold: true });
      addSlideText(slide, "Catatan: inventory menunjukkan posisi saat laporan dibuat, bukan saldo historis akhir bulan.", 0.8, 5.0, 11.7, 0.45, { fontSize: 10, color: "64748B", italic: true });

      slide = pptx.addSlide(); pptHeader(slide, "Outbound & Packing", sub, pptx);
      pptCard(slide, 0.65, 1.4, 3.85, "Total AWB", fmt(m.totalAWB), pptx);
      pptCard(slide, 4.75, 1.4, 3.85, "Total Paket", fmt(m.totalPackages), pptx);
      if (r.visibility.financial) pptCard(slide, 8.85, 1.4, 3.85, "Cost Packing", money(m.totalPackingCost), pptx);
      addSlideText(slide, "Kategori packing terbanyak", 0.8, 3.1, 5, 0.3, { fontSize: 13, bold: true });
      r.topCategories.slice(0, 7).forEach(function (x, i) { addSlideText(slide, x.category + " — " + fmt(x.packages) + " paket", 0.85, 3.5 + i * 0.42, 6.2, 0.22, { fontSize: 10 }); });

      slide = pptx.addSlide(); pptHeader(slide, "Performa Antar Gudang", sub, pptx);
      var wlist = r.warehousePerformance.slice(0, 8);
      var cols = [{ x: .75, w: 2.3, label: "Gudang" }, { x: 3.2, w: 1.6, label: "Inbound" }, { x: 4.95, w: 1.6, label: "Outbound" }, { x: 6.7, w: 1.5, label: "Paket" }, { x: 8.35, w: 1.8, label: "Produktivitas" }];
      if (r.visibility.financial) cols.push({ x: 10.35, w: 2.1, label: "Cost Packing" });
      cols.forEach(function (c) { addSlideText(slide, c.label, c.x, 1.35, c.w, .25, { fontSize: 9, bold: true, color: "64748B", align: c.label === "Gudang" ? "left" : "right" }); });
      wlist.forEach(function (w, i) {
        var y = 1.8 + i * .55;
        slide.addShape(pptx.ShapeType.line, { x: .7, y: y + .34, w: 12, h: 0, line: { color: "E2E8F0", width: .6 } });
        var vals = [w.warehouseName, fmt(w.inboundQty), fmt(w.outboundQty), fmt(w.packages), fmt(w.productivityPackages)];
        cols.forEach(function (c, ci) { addSlideText(slide, ci === 0 ? vals[ci] : vals[ci], c.x, y, c.w, .26, { fontSize: 9, align: ci === 0 ? "left" : "right", bold: ci === 0 }); });
        if (r.visibility.financial) addSlideText(slide, money(w.packingCost), 10.35, y, 2.1, .26, { fontSize: 9, align: "right" });
      });

      slide = pptx.addSlide(); pptHeader(slide, "Produktivitas Tim", sub, pptx);
      pptCard(slide, 0.65, 1.4, 3.85, "Admin Print (Paket)", fmt(m.productivityPackages), pptx);
      pptCard(slide, 4.75, 1.4, 3.85, "Admin Print (Pcs)", fmt(m.productivityPcs), pptx);
      pptCard(slide, 8.85, 1.4, 3.85, "Total Entri Produktivitas", fmt(m.productivityEntries), pptx);
      r.byRole.slice(0, 9).forEach(function (x, i) { addSlideText(slide, x.role + " — " + fmt(x.packages) + " paket / " + fmt(x.pcs) + " pcs", .8, 3.1 + i * .4, 11.5, .22, { fontSize: 10 }); });

      slide = pptx.addSlide(); pptHeader(slide, "Purchase Order & Case", sub, pptx);
      var ps = r.purchaseOrderSummary, cs = r.caseSummary;
      var ops = [
        ["PO dibuat periode ini", ps.createdInPeriod], ["PO selesai periode ini", ps.completedInPeriod],
        ["PO outstanding", ps.outstandingNow], ["PO overdue", ps.overdueNow],
        ["Case baru", cs.newInPeriod], ["Case selesai", cs.closedInPeriod],
        ["Case terbuka", cs.openNow], ["Case prioritas tinggi", cs.highPriorityOpenNow]
      ];
      ops.forEach(function (x, i) { pptCard(slide, .65 + (i % 2) * 6.1, 1.35 + Math.floor(i / 2) * 1.2, 5.75, x[0], fmt(x[1]), pptx); });

      slide = pptx.addSlide(); pptHeader(slide, "Analisa & Rekomendasi", sub, pptx);
      r.insights.slice(0, 8).forEach(function (x, i) {
        var y = 1.35 + i * .62;
        slide.addShape(pptx.ShapeType.roundRect, { x: .7, y: y, w: 11.95, h: .48, rectRadius: .04, line: { color: x.level === "warning" ? "FDE68A" : "E2E8F0", width: .8 }, fill: { color: x.level === "warning" ? "FFFBEB" : "FFFFFF" } });
        addSlideText(slide, x.title + " — " + x.detail, .9, y + .07, 11.5, .3, { fontSize: 8.5, color: "334155" });
      });
      addSlideText(slide, "Analisa dihitung dari data yang tersedia di frontend. Inventory adalah posisi saat ini; riwayat yang sudah dipindahkan ke arsip hanya tercakup jika dikirim oleh data loader yang tersedia.", .8, 6.55, 11.8, .36, { fontSize: 7.5, color: "64748B", italic: true });
      await pptx.writeFile({ fileName: fileBase() + ".pptx" });
      toast("PowerPoint berhasil dibuat (" + pptx._slides.length + " slide).", "success");
    } catch (err) { toast("Gagal membuat PowerPoint: " + err.message, "error"); }
  };

  window.initReportsTab = function () {
    buildMarkup();
    initWarehouseOptions();
    if (!state.report) window.generateFrontendReport();
  };
  document.addEventListener("DOMContentLoaded", buildMarkup);
  buildMarkup();
})();
