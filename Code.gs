// ============ TRANSFERS DASHBOARD ============
// Sheet tab name: Shipments
// SLA :
//   Picking 2h | Manifest 4h | Handover 3h | Last-mile 3h
//   Transit SLA = 1.25 x route median transit, rounded up to nearest 0.5h
//   Total route SLA = Transit SLA + 10h (stage targets 9.5h + 0.5h buffer)
//   Critical = 1.5 x route Total SLA

var SLA = { picking: 2, manifest: 4, handover: 3, lastMile: 3, fixedHrs: 10, critMult: 1.5, transitBuffer: 0 };
var COLS = ['Shipment_ID','Company','Origin_Hub','Destination_Hub','Qty','Created_At','Picked_At','Manifested_At','Handover_At','Arrived_At','Delivered_At','Picked_Lag_Hr','Manifest_Lag_Hr','Transit_Lag_Hr','Delay_Hrs','Route','Handover_Lag_Hr','LastMile_Lag_Hr','Total_Cycle_Hr','Status'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Transfers Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function safeNum(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var n = Number(String(v).replace(/,/g, '').trim());
  return isFinite(n) ? n : null;
}

function toMs(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (v instanceof Date) { var t = v.getTime(); return isNaN(t) ? null : t; }
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d.getTime();
}

function avg(sum, n) { return n > 0 ? sum / n : null; }
function pct(a, b) { return b > 0 ? (a / b) * 100 : 0; }
function r2(v) { return v === null || v === undefined ? null : Math.round(v * 100) / 100; }
function median(arr) {
  if (!arr.length) return null;
  var s = arr.slice().sort(function(a, b) { return a - b; });
  var mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

function getDashboardData(filters, force) {
  filters = filters || {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.openById('');
    var sh = ss.getSheetByName('Shipments');
    if (!sh) throw new Error('Sheet "Shipments" not found. Please name the tab "Shipments".');
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      return { empty: true, meta: { totalRowsInSheet: 0, rowsInView: 0, generatedAtLabel: 'No data', refTimeLabel: '-', sla: SLA }, options: { origins: [], destinations: [], companies: [], statuses: [], routes: [] } };
    }
    var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var header = values[0].map(function(x) { return String(x).trim(); });
    var ix = {};
    COLS.forEach(function(c) { ix[c] = header.indexOf(c); });
    if (ix['Shipment_ID'] < 0 || ix['Status'] < 0) throw new Error('Shipment_ID or Status column is missing.');

    var get = function(row, c) { return ix[c] >= 0 ? row[ix[c]] : ''; };
    var records = [];
    var seen = {};
    var maxTs = 0;
    var dup = 0;

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var id = String(get(row, 'Shipment_ID')).trim();
      var origin = String(get(row, 'Origin_Hub')).trim() || 'Unknown';
      var dest = String(get(row, 'Destination_Hub')).trim() || 'Unknown';
      var status = String(get(row, 'Status')).trim();
      if (!id) continue;
      if (seen[id]) { dup++; continue; }
      seen[id] = 1;

      var created = toMs(get(row, 'Created_At'));
      var picked = toMs(get(row, 'Picked_At'));
      var manifested = toMs(get(row, 'Manifested_At'));
      var handover = toMs(get(row, 'Handover_At'));
      var arrived = toMs(get(row, 'Arrived_At'));
      var delivered = toMs(get(row, 'Delivered_At'));
      [created, picked, manifested, handover, arrived, delivered].forEach(function(x) { if (x) maxTs = Math.max(maxTs, x); });

      var lagPick = safeNum(get(row, 'Picked_Lag_Hr'));
      var lagMan = safeNum(get(row, 'Manifest_Lag_Hr'));
      var lagHand = safeNum(get(row, 'Handover_Lag_Hr'));
      var lagTrans = safeNum(get(row, 'Transit_Lag_Hr'));
      var lagLast = safeNum(get(row, 'LastMile_Lag_Hr'));
      // If a lag is blank, calculate it from timestamps
      if (lagPick === null && created !== null && picked !== null) lagPick = (picked - created) / 3600000;
      if (lagMan === null && picked !== null && manifested !== null) lagMan = (manifested - picked) / 3600000;
      if (lagHand === null && manifested !== null && handover !== null) lagHand = (handover - manifested) / 3600000;
      if (lagTrans === null && handover !== null && arrived !== null) lagTrans = (arrived - handover) / 3600000;
      if (lagLast === null && arrived !== null && delivered !== null) lagLast = (delivered - arrived) / 3600000;

      var total = safeNum(get(row, 'Total_Cycle_Hr'));
      if (total === null && created !== null && delivered !== null) total = (delivered - created) / 3600000;

      var isDel = status.toLowerCase().indexOf('deliver') === 0;
      var isOpen = status.toLowerCase().indexOf('open') === 0;
      var route = String(get(row, 'Route')).trim() || (origin + ' \u2192 ' + dest);
      var qty = safeNum(get(row, 'Qty'));

      records.push({
        id: id, company: String(get(row, 'Company')).trim() || 'Unknown',
        origin: origin, dest: dest, route: route, qty: qty, status: status,
        created: created, picked: picked, manifested: manifested, handover: handover, arrived: arrived, delivered: delivered,
        lagPick: lagPick, lagMan: lagMan, lagHand: lagHand, lagTrans: lagTrans, lagLast: lagLast,
        total: isDel ? total : null, isDel: isDel, isOpen: isOpen
      });
    }

    // ---- ROUTE-WISE SLA (calculated on ALL data, so filters don't change the limits) ----
    var rtMap = {}, allT = [], transSla = {};
    records.forEach(function(r) {
      if (r.lagTrans !== null && r.arrived !== null) (rtMap[r.route] = rtMap[r.route] || []).push(r.lagTrans);
    });
    Object.keys(rtMap).forEach(function(k) {
      allT = allT.concat(rtMap[k]);
      transSla[k] = Math.ceil(median(rtMap[k]) * 1.25 * (1 + SLA.transitBuffer) * 2) / 2;
    });
    var gT = Math.ceil((median(allT) || 12) * 1.25 * (1 + SLA.transitBuffer) * 2) / 2;
    var getTransitLimit = function(rt) { return transSla[rt] || gT; };
    records.forEach(function(r) {
      r.tsla = getTransitLimit(r.route);
      r.sla = r.tsla + SLA.fixedHrs;
      r.crit = r.sla * SLA.critMult;
    });

    // ---- FILTERS ----
    var f = {
      origin: filters.origin || '', destination: filters.destination || '',
      company: filters.company || '', status: filters.status || '', route: filters.route || '',
      search: (filters.search || '').toString().toLowerCase(), comparison: filters.comparison || 'WoW'
    };
    var filtered = records.filter(function(r) {
      if (f.origin && r.origin !== f.origin) return false;
      if (f.destination && r.dest !== f.destination) return false;
      if (f.company && r.company !== f.company) return false;
      if (f.status) {
        if (f.status === '__OPEN__') { if (!r.isOpen) return false; }
        else if (r.status !== f.status) return false;
      }
      if (f.route && r.route !== f.route) return false;
      if (f.search) {
        var hay = (r.id + '|' + r.origin + '|' + r.dest + '|' + r.route + '|' + r.company).toLowerCase();
        if (hay.indexOf(f.search) < 0) return false;
      }
      return true;
    });

    // ---- OPTIONS (from full data) ----
    var oSet = {}, dSet = {}, cSet = {}, sSet = {}, rSet = {};
    records.forEach(function(r) { oSet[r.origin] = 1; dSet[r.dest] = 1; cSet[r.company] = 1; sSet[r.status] = 1; rSet[r.route] = 1; });
    var keys = function(o) { return Object.keys(o).sort(); };

    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Kolkata';
    var out = {
      empty: filtered.length === 0,
      meta: {
        totalRowsInSheet: records.length, rowsInView: filtered.length, sla: SLA,
        generatedAtLabel: Utilities.formatDate(new Date(), tz, 'dd MMM yyyy, hh:mm a'),
        refTimeLabel: Utilities.formatDate(new Date(maxTs || Date.now()), tz, 'dd MMM yyyy, hh:mm a'),
        cached: false
      },
      filters: f,
      options: { origins: keys(oSet), destinations: keys(dSet), companies: keys(cSet), statuses: keys(sSet), routes: keys(rSet) },
      quality: { duplicateIds: dup, rowsAnalyzed: records.length }
    };
    if (filtered.length === 0) return out;

    // ---- KPIs + SLA ----
    var del = filtered.filter(function(r) { return r.isDel && r.total !== null; });
    var open = filtered.filter(function(r) { return r.isOpen; });
    var onTime = del.filter(function(r) { return r.total <= r.sla; }).length;
    var delayed = del.length - onTime;
    var critical = del.filter(function(r) { return r.total > r.crit; }).length;
    var totalDelaySum = 0, delayLateSum = 0, maxDelay = 0, cycSum = 0;
    del.forEach(function(r) {
      cycSum += r.total;
      if (r.total > r.sla) { var d = r.total - r.sla; totalDelaySum += d; delayLateSum += d; if (d > maxDelay) maxDelay = d; }
    });
    var totalQty = 0;
    filtered.forEach(function(r) { if (r.qty) totalQty += r.qty; });

    // ---- STAGES ----
    function stageStats(getLag, limitFn, doneFn) {
      var vals = [], breach = 0, excess = 0, sum = 0;
      filtered.forEach(function(r) {
        var v = getLag(r);
        if (v === null || v < 0) return;
        if (doneFn && !doneFn(r)) return;
        vals.push(v); sum += v;
        var lim = limitFn(r);
        if (lim !== null && v > lim) { breach++; excess += (v - lim); }
      });
      return { n: vals.length, avg: r2(avg(sum, vals.length)), median: r2(median(vals)), breachCount: breach, breachPct: r2(pct(breach, vals.length)), delayHrs: r2(excess) };
    }
    var stPick = stageStats(function(r) { return r.lagPick; }, function() { return SLA.picking; }, function(r) { return r.picked !== null; });
    var stMan = stageStats(function(r) { return r.lagMan; }, function() { return SLA.manifest; }, function(r) { return r.manifested !== null; });
    var stHand = stageStats(function(r) { return r.lagHand; }, function() { return SLA.handover; }, function(r) { return r.handover !== null; });
    var stTrans = stageStats(function(r) { return r.lagTrans; }, function(r) { return r.tsla; }, function(r) { return r.arrived !== null; });
    var stLast = stageStats(function(r) { return r.lagLast; }, function() { return SLA.lastMile; }, function(r) { return r.isDel; });

    function mk(key, label, st, limit) {
      return { key: key, label: label, avg: st.avg, median: st.median, breachCount: st.breachCount, breachPct: st.breachPct, delayHrs: st.delayHrs, limit: limit, n: st.n };
    }
    var stagesArr = [
      mk('picking', 'Picking', stPick, SLA.picking),
      mk('manifest', 'Manifest', stMan, SLA.manifest),
      mk('handover', 'Handover', stHand, SLA.handover),
      mk('transit', 'Transit', stTrans, null),
      mk('lastMile', 'Last Mile', stLast, SLA.lastMile)
    ];
    var totalExcess = stagesArr.reduce(function(a, s) { return a + (s.delayHrs || 0); }, 0);
    stagesArr.forEach(function(s) { s.delaySharePct = r2(pct(s.delayHrs, totalExcess)); });
    var sortedStages = stagesArr.slice().sort(function(a, b) { return b.delayHrs - a.delayHrs; });
    sortedStages.forEach(function(s, i) { s.rank = i + 1; });
    var primary = sortedStages[0];

    // ---- OPEN RISK ----
    var openStages = [
      { status: 'Open - Picked', open: 0, within: 0, atRisk: 0 },
      { status: 'Open - Manifested', open: 0, within: 0, atRisk: 0 },
      { status: 'Open - In Transit', open: 0, within: 0, atRisk: 0 },
      { status: 'Open - Arrived', open: 0, within: 0, atRisk: 0 }
    ];
    var openMap = {}; openStages.forEach(function(s) { openMap[s.status] = s; });
    var refTime = maxTs || Date.now();
    var openAtRisk = 0, openWithin = 0;
    open.forEach(function(r) {
      var row = openMap[r.status];
      if (!row) return;
      row.open++;
      var el = null, lim = null;
      if (r.status === 'Open - Picked') { el = r.lagPick; lim = SLA.picking; }
      if (r.status === 'Open - Manifested') { el = r.lagMan; lim = SLA.manifest; }
      if (r.status === 'Open - In Transit') { el = r.lagTrans !== null ? r.lagTrans : (r.handover !== null ? (refTime - r.handover) / 3600000 : null); lim = r.tsla; }
      if (r.status === 'Open - Arrived') { el = r.arrived !== null ? (refTime - r.arrived) / 3600000 : null; lim = SLA.lastMile; }
      if (el === null || lim === null) return;
      if (el > lim) { row.atRisk++; openAtRisk++; } else { row.within++; openWithin++; }
    });
    openStages.forEach(function(s) { s.riskPct = r2(pct(s.atRisk, s.within + s.atRisk)); });
    var openRisk = { open: open.length, within: openWithin, atRisk: openAtRisk, riskPct: r2(pct(openAtRisk, openWithin + openAtRisk)), stages: openStages };

    // ---- KPIs ----
    var onTimePct = r2(pct(onTime, del.length));
    var breachPct = r2(pct(delayed, del.length));
    var avgCycle = r2(avg(cycSum, del.length));
    var avgDelay = r2(avg(totalDelaySum, del.length));
    function tr(cur, good) { return { cur: cur, prev: cur, diff: 0, pct: 0, good: good }; }
    out.kpis = {
      total: filtered.length, delivered: del.length, open: open.length,
      onTimePct: onTimePct, breachPct: breachPct,
      critical: critical, criticalPct: r2(pct(critical, del.length)),
      avgCycle: avgCycle, avgDelay: avgDelay,
      avgDelayWhenLate: r2(avg(delayLateSum, delayed)), totalQty: totalQty, delayed: delayed,
      avgPicking: stPick.avg, avgManifest: stMan.avg, avgHandover: stHand.avg, avgTransit: stTrans.avg, avgLastMile: stLast.avg,
      primaryBottleneck: primary ? { label: primary.label, share: primary.delaySharePct } : null,
      openRisk: openAtRisk, openRiskPct: openRisk.riskPct,
      avgQty: r2(avg(totalQty, filtered.length)),
      trends: { total: tr(filtered.length, 'neutral'), delivered: tr(del.length, 'up'), onTimePct: tr(onTimePct, 'up'), breachPct: tr(breachPct, 'down'), critical: tr(critical, 'down'), avgCycle: tr(avgCycle, 'down'), avgDelay: tr(avgDelay, 'down'), totalQty: tr(totalQty, 'neutral'), delayed: tr(delayed, 'down') }
    };

    // ---- SLA ----
    var cycVals = del.map(function(r) { return r.total; });
    out.sla = {
      n: del.length, onTime: onTime, delayedAll: delayed, critical: critical, warning: delayed - critical,
      onTimePct: onTimePct, breachPct: breachPct, criticalPct: out.kpis.criticalPct,
      avgCycle: avgCycle, medianCycle: r2(median(cycVals)), p90Cycle: null, maxCycle: del.length ? r2(Math.max.apply(null, cycVals)) : null,
      avgDelay: avgDelay, avgDelayWhenLate: r2(avg(delayLateSum, delayed)), maxDelay: r2(maxDelay), totalDelayHrs: r2(totalDelaySum),
      buckets: [{ key: 'ok', label: 'Within SLA', count: onTime }, { key: 'warn', label: 'Late (<1.5x SLA)', count: delayed - critical }, { key: 'crit', label: 'Critical (>1.5x SLA)', count: critical }],
      target: null, criticalAt: null
    };

    out.stages = stagesArr;
    out.openRisk = openRisk;
    out.periods = { comparison: f.comparison, spec: { label: '7 days', prevLabel: 'previous period' }, curFrom: '-', curTo: '-', prevFrom: '-', prevTo: '-', current: { onTimePct: onTimePct, avgCycle: avgCycle, n: del.length }, previous: { onTimePct: onTimePct, avgCycle: avgCycle, n: del.length }, enough: false };

    // ---- HUBS / ROUTES / COMPANIES ----
    var STAGE_LABELS = ['Picking', 'Manifest', 'Handover', 'Transit', 'Last Mile'];
    function aggBy(keyFn) {
      var map = {};
      filtered.forEach(function(r) {
        var k = keyFn(r);
        if (!map[k]) map[k] = { name: k, transfers: 0, qty: 0, delivered: 0, open: 0, evalN: 0, ok: 0, late: 0, crit: 0, cycSum: 0, delaySum: 0, transitSum: 0, transitN: 0, se: [0, 0, 0, 0, 0], sn: [0, 0, 0, 0, 0], sb: [0, 0, 0, 0, 0] };
        var a = map[k];
        a.transfers++; a.qty += r.qty || 0;
        if (r.isDel) a.delivered++;
        if (r.isOpen) a.open++;
        if (r.total !== null) {
          a.evalN++; a.cycSum += r.total;
          if (r.total <= r.sla) a.ok++;
          else { a.late++; a.delaySum += r.total - r.sla; if (r.total > r.crit) a.crit++; }
        }
        if (r.lagTrans !== null && r.arrived !== null) { a.transitSum += r.lagTrans; a.transitN++; }
        var lags = [r.lagPick, r.lagMan, r.lagHand, r.lagTrans, r.lagLast];
        var lims = [SLA.picking, SLA.manifest, SLA.handover, r.tsla, SLA.lastMile];
        var dones = [r.picked !== null, r.manifested !== null, r.handover !== null, r.arrived !== null, r.isDel];
        for (var s = 0; s < 5; s++) {
          if (!dones[s] || lags[s] === null) continue;
          a.sn[s]++;
          if (lags[s] > lims[s]) { a.sb[s]++; a.se[s] += lags[s] - lims[s]; }
        }
      });
      return Object.keys(map).map(function(k) {
        var a = map[k];
        var bi = -1, bv = 0;
        for (var s = 0; s < 5; s++) { if (a.se[s] > bv) { bv = a.se[s]; bi = s; } }
        var totalEx = a.se.reduce(function(x, y) { return x + y; }, 0);
        return {
          name: a.name, transfers: a.transfers, qty: a.qty, delivered: a.delivered, open: a.open, evaluated: a.evalN,
          avgCycle: r2(avg(a.cycSum, a.evalN)), avgDelay: r2(avg(a.delaySum, a.evalN)),
          onTimePct: r2(pct(a.ok, a.evalN)), breachPct: r2(pct(a.late, a.evalN)), breachCount: a.late,
          critical: a.crit, criticalPct: r2(pct(a.crit, a.evalN)), avgTransit: r2(avg(a.transitSum, a.transitN)),
          bottleneck: bi >= 0 ? STAGE_LABELS[bi] : null, bottleneckShare: bi >= 0 ? r2(pct(bv, totalEx)) : null,
          stageBreachPct: a.sn.map(function(n, s2) { return r2(pct(a.sb[s2], n)); }),
          stageDelayHrs: a.se.map(function(v) { return r2(v); })
        };
      }).sort(function(a, b) { return (b.avgDelay || 0) - (a.avgDelay || 0); });
    }
    var originList = aggBy(function(r) { return r.origin; });
    var destList = aggBy(function(r) { return r.dest; });
    out.hubs = { origin: originList, destination: destList, stageLabels: STAGE_LABELS };

    var routeList = aggBy(function(r) { return r.route; });
    routeList.forEach(function(x) { x.transitSla = transSla[x.name] || gT; x.totalSla = x.transitSla + SLA.fixedHrs; });
    var qRoutes = routeList.filter(function(x) { return x.evaluated >= 5; });
    if (qRoutes.length < 2) qRoutes = routeList;
    var byScore = qRoutes.slice().sort(function(a, b) { return b.breachPct - a.breachPct; });
    var byBest = qRoutes.slice().sort(function(a, b) { return b.onTimePct - a.onTimePct; });
    out.routes = { all: routeList, worst: byScore.slice(0, 5), best: byBest.slice(0, 5), minSample: 5, routeCount: routeList.length };

    out.companies = aggBy(function(r) { return r.company; }).map(function(x) {
      return { name: x.name, transfers: x.transfers, qty: x.qty, delivered: x.delivered, open: x.open, onTimePct: x.onTimePct, breachPct: x.breachPct, avgCycle: x.avgCycle, avgDelay: x.avgDelay };
    });

    // ---- STATUS ----
    var statusMap = {};
    filtered.forEach(function(r) { statusMap[r.status] = (statusMap[r.status] || 0) + 1; });
    out.status = Object.keys(statusMap).map(function(k) { return { status: k, count: statusMap[k], pct: r2(pct(statusMap[k], filtered.length)) }; });

    // ---- TRENDS (last 14 days by Created) ----
    var trendMap = {}, trendOrder = [];
    filtered.slice().sort(function(a, b) { return (a.created || 0) - (b.created || 0); }).forEach(function(r) {
      if (!r.created) return;
      var d = Utilities.formatDate(new Date(r.created), tz, 'dd MMM');
      if (!trendMap[d]) { trendMap[d] = { created: 0, delivered: 0, delayed: 0, cycSum: 0, cycN: 0 }; trendOrder.push(d); }
      trendMap[d].created++;
      if (r.isDel) {
        trendMap[d].delivered++;
        if (r.total !== null) { trendMap[d].cycSum += r.total; trendMap[d].cycN++; if (r.total > r.sla) trendMap[d].delayed++; }
      }
    });
    var tKeys = trendOrder.slice(-14);
    out.trends = {
      grain: 'day', labels: tKeys,
      created: tKeys.map(function(k) { return trendMap[k].created; }),
      delivered: tKeys.map(function(k) { return trendMap[k].delivered; }),
      delayed: tKeys.map(function(k) { return trendMap[k].delayed; }),
      avgCycle: tKeys.map(function(k) { return r2(avg(trendMap[k].cycSum, trendMap[k].cycN)); }),
      target: null
    };

    // ---- DELAY DISTRIBUTION ----
    var buckets = [{ label: '0-2h', count: 0 }, { label: '2-5h', count: 0 }, { label: '5-10h', count: 0 }, { label: '10h+', count: 0 }];
    del.forEach(function(r) {
      if (r.total <= r.sla) return;
      var d = r.total - r.sla;
      if (d <= 2) buckets[0].count++;
      else if (d <= 5) buckets[1].count++;
      else if (d <= 10) buckets[2].count++;
      else buckets[3].count++;
    });
    buckets.forEach(function(b) { b.pct = r2(pct(b.count, delayed)); });
    var dom = buckets.slice().sort(function(a, b) { return b.count - a.count; })[0];
    out.delayDist = { buckets: buckets, delayedTotal: delayed, dominant: delayed ? dom.label : null, note: 'Hours beyond the route-wise total SLA' };

    // ---- QUANTITY ----
    var qRows = filtered.filter(function(r) { return r.qty !== null && r.total !== null; });
    var qVals = qRows.map(function(r) { return r.qty; }).sort(function(a, b) { return a - b; });
    var bands = [];
    if (qRows.length >= 10) {
      var cuts = [0.2, 0.4, 0.6, 0.8].map(function(p) { return qVals[Math.min(Math.floor(qVals.length * p), qVals.length - 1)]; });
      var bandArr = [0, 1, 2, 3, 4].map(function() { return { n: 0, qMin: Infinity, qMax: -Infinity, delay: 0, late: 0 }; });
      qRows.forEach(function(r) {
        var b = 0; while (b < cuts.length && r.qty > cuts[b]) b++;
        var bb = bandArr[b]; var d = Math.max(0, r.total - r.sla);
        bb.n++; bb.delay += d; if (d > 0) bb.late++;
        if (r.qty < bb.qMin) bb.qMin = r.qty;
        if (r.qty > bb.qMax) bb.qMax = r.qty;
      });
      bands = bandArr.filter(function(b) { return b.n > 0; }).map(function(b) {
        return { label: b.qMin + '-' + b.qMax, n: b.n, avgDelay: r2(avg(b.delay, b.n)), breachPct: r2(pct(b.late, b.n)) };
      });
    }
    out.qty = { total: totalQty, avg: r2(avg(totalQty, filtered.length)), median: r2(median(qVals)), max: qVals.length ? qVals[qVals.length - 1] : 0, n: filtered.length, bands: bands, verdict: bands.length ? 'Quantity bands show how delay changes with shipment size.' : 'Not enough data for quantity bands.' };

    // ---- INSIGHTS + HEADLINE + ALERTS ----
    out.insights = [];
    if (primary) {
      out.insights.push({ tone: 'red', title: 'Primary Bottleneck: ' + primary.label, headline: primary.label + ' is the primary bottleneck, contributing ' + primary.delaySharePct + '% of stage delay.', facts: [['Breach rate', primary.breachPct + '%'], ['Delay share', primary.delaySharePct + '%'], ['Avg lag', (primary.avg || 0) + 'h'], ['Breaches', primary.breachCount]], note: 'Fix this stage first - it gives the biggest cycle time saving.' });
    }
    if (byScore[0]) {
      out.insights.push({ tone: 'orange', title: 'Weakest Route', headline: byScore[0].name + ' has the highest breach rate at ' + byScore[0].breachPct + '%.', facts: [['Breach %', byScore[0].breachPct + '%'], ['Avg delay', (byScore[0].avgDelay || 0) + 'h'], ['Transfers', byScore[0].transfers]], note: 'Check this route for transit and handover issues.' });
    }
    if (byBest[0]) {
      out.insights.push({ tone: 'green', title: 'Strongest Route', headline: byBest[0].name + ' is the best route with ' + byBest[0].onTimePct + '% on-time.', facts: [['On-time %', byBest[0].onTimePct + '%'], ['Avg delay', (byBest[0].avgDelay || 0) + 'h'], ['Transfers', byBest[0].transfers]], note: 'Copy what works on this route to weaker routes.' });
    }
    out.insights.push({ tone: openAtRisk > 0 ? 'amber' : 'green', title: 'Open Transfer Risk', headline: openAtRisk + ' open transfers are currently beyond their stage SLA.', facts: [['Open', open.length], ['At risk', openAtRisk], ['Risk %', openRisk.riskPct + '%']], note: 'Open transfers are judged on their current stage only.' });
    out.insights.push({ tone: 'orange', title: 'Delay Pattern', headline: delayed ? 'Most delayed deliveries are late by ' + dom.label + ' beyond SLA.' : 'No delayed deliveries in this view.', facts: [['Delayed', delayed], ['Avg delay when late', (r2(avg(delayLateSum, delayed)) || 0) + 'h'], ['Max delay', r2(maxDelay) + 'h']], note: 'Focus on the dominant bucket first.' });

    out.headline = { text: onTimePct + '% on-time', delta: 0, vs: 'previous period', context: open.length + ' open, ' + openAtRisk + ' at risk' };
    out.alerts = [
      { tone: breachPct > 50 ? 'red' : breachPct > 25 ? 'amber' : 'green', label: 'SLA breach rate', value: breachPct + '%', sub: delayed + ' delivered beyond route SLA', go: 'sla' },
      { tone: critical > 0 ? 'red' : 'green', label: 'Critical transfers', value: String(critical), sub: 'cycle above 1.5x route SLA', go: 'sla' },
      { tone: openAtRisk > 0 ? 'amber' : 'green', label: 'Open transfers at risk', value: String(openAtRisk), sub: 'beyond current stage SLA', go: 'bottlenecks' },
      { tone: 'red', label: 'Primary bottleneck', value: primary ? primary.label : 'None', sub: primary ? primary.delaySharePct + '% of stage delay' : '', go: 'bottlenecks' }
    ];

    return out;
  } catch (e) {
    throw new Error(e.message || String(e));
  }
}