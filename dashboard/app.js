/* VM Reporting Platform — client app (all projects, all sheets).
 * Data flow: Google Sheet (every tab) + imported Excel/CSV -> parseSheets() (header detection,
 * auto-correction, cross-tab de-duplication) -> filters -> views and exports.
 * Without API_URL it runs on window.SAMPLE_DATA (offline snapshot). */
(function () {
  'use strict';
  var CFG = window.CONFIG || {};
  var LS = {
    get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } }
  };
  var API_URL = (LS.get('vm.apiUrl') || CFG.API_URL || '').trim();
  var $ = function (id) { return document.getElementById(id); };
  var fmt = function (n) { return (n == null || isNaN(n)) ? '–' : Math.round(n).toLocaleString('en-US'); };
  var pct = function (n) { return (n == null || !isFinite(n)) ? '–' : Math.round(n * 100) + '%'; };
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MONTHS_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var dstr = function (d) { return d ? d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') : ''; };
  var dnice = function (d) { return d ? String(d.getDate()).padStart(2, '0') + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear() : '–'; };
  var dsheet = function (d) { return d ? d.getDate() + ' ' + MONTHS_LONG[d.getMonth()] + ' ' + d.getFullYear() : ''; }; // as in the sheet: "10 March 2026"
  var key = function (s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function uniq(arr) { return arr.filter(function (v, i, a) { return v && a.indexOf(v) === i; }).sort(function (a, b) { return String(a).localeCompare(String(b)); }); }
  function lev(a, b) { // Levenshtein distance
    if (a === b) return 0;
    var m = a.length, n = b.length, d = [], i, j;
    for (i = 0; i <= m; i++) d[i] = [i];
    for (j = 1; j <= n; j++) d[0][j] = j;
    for (i = 1; i <= m; i++) for (j = 1; j <= n; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[m][n];
  }
  function closeEnough(a, b) { var l = Math.min(a.length, b.length); return l >= 4 && lev(a, b) <= (l >= 8 ? 2 : 1); }

  // =====================================================================
  // Gazetteer (64 zilas + every upazila, from geo.js) and auto-correction
  // =====================================================================
  var DISTRICTS = {}, UPZ = {}, UPZ_BY_DIST = {}, UPZ_LIST = [];
  var UPZ_DISPLAY = { 'Fulchhari': 'Fulchari', 'Kala Para': 'Kalapara', 'Dharampasha': 'Dharmapasha' };
  (window.GEO ? GEO.districts.features : []).forEach(function (f) { DISTRICTS[key(f.properties.name)] = f.properties.name; });
  (window.GEO ? GEO.upazilas.features : []).forEach(function (f) {
    var name = UPZ_DISPLAY[f.properties.name] || f.properties.name;
    f.properties.display = name;
    var u = { name: name, district: f.properties.district, geo: f };
    UPZ_LIST.push(u);
    if (!UPZ[key(name)]) UPZ[key(name)] = u;
    if (!UPZ[key(f.properties.name)]) UPZ[key(f.properties.name)] = u;
    (UPZ_BY_DIST[f.properties.district] = UPZ_BY_DIST[f.properties.district] || []).push(u);
  });
  var CITY = { name: 'Chattogram City Corp.', district: 'Chittagong', geo: null };
  var DIST_ALIAS = { ctg: 'Chittagong', chattogram: 'Chittagong', chittagongcity: 'Chittagong', coxbazar: "Cox's Bazar", coxsbazar: "Cox's Bazar",
    bagura: 'Bogra', bogura: 'Bogra', netrokona: 'Netrakona', bandor: 'Bandarban', bandorban: 'Bandarban', bandarbon: 'Bandarban',
    barishal: 'Barisal', jashore: 'Jessore', cumilla: 'Comilla', moulvibazar: 'Maulvibazar', chapainawabganj: 'Nawabganj',
    chapainababganj: 'Nawabganj', brahmanbaria: 'Brahamanbaria', jhalakathi: 'Jhalokati', patu: 'Patuakhali' };
  var UPZ_ALIAS = { naikhomchari: 'Naikhongchhari', naikhongchari: 'Naikhongchhari', bashkhali: 'Banshkhali', shariakandi: 'Sariakandi',
    dhormopasha: 'Dharmapasha', dharmapasha: 'Dharmapasha', gowaninghat: 'Gowainghat', kalapara: 'Kalapara', fulchari: 'Fulchari' };

  function matchDistrict(text) {
    var k = key(text);
    if (!k) return null;
    if (DISTRICTS[k]) return DISTRICTS[k];
    if (DIST_ALIAS[k]) return DIST_ALIAS[k];
    var ks = Object.keys(DISTRICTS), pre = ks.filter(function (d) { return k.length >= 4 && d.indexOf(k) === 0; });
    if (pre.length === 1) return DISTRICTS[pre[0]];
    var best = null;
    ks.forEach(function (d) { if (!best && closeEnough(k, d)) best = DISTRICTS[d]; });
    return best;
  }
  function matchUpazila(text, district) {
    var k = key(text);
    if (!k) return null;
    if (/^(ccc|ctgcitycorp\w*|citycorp\w*|chittagongcitycorp\w*|chattogramcitycorp\w*)$/.test(k)) return CITY;
    if (UPZ_ALIAS[k] && UPZ[key(UPZ_ALIAS[k])]) return UPZ[key(UPZ_ALIAS[k])];
    var k2 = k.replace(/sador$/, 'sadar');
    if (k2 === 'sadar') return district ? (UPZ[key(district + ' sadar')] || null) : null;
    var find = function (list) {
      var hit = null;
      list.forEach(function (u) { if (!hit && (key(u.name) === k2 || key(u.geo.properties.name) === k2)) hit = u; });
      list.forEach(function (u) { if (!hit && closeEnough(k2, key(u.name))) hit = u; });
      return hit;
    };
    var pool = district && UPZ_BY_DIST[district];
    return (pool && find(pool)) || UPZ[k2] || find(UPZ_LIST);
  }

  var CONTENT_RULES = [
    [/flash/, 'Flash Flood'], [/flood|bonna/, 'Flood'], [/light|thunder|bajro|bojro/, 'Lightning'], [/r+a+i+n|brishti|precip/, 'Rainfall'],
    [/cold|sheet/, 'Cold Wave'], [/heat|tap/, 'Heat Wave'], [/land\s*slide|landslid/, 'Landslide'], [/cyclone|storm|ghurni/, 'Cyclone'], [/fog|kuasha/, 'Fog']
  ];
  var CONTENT_CANON = ['Rainfall', 'Lightning', 'Flood', 'Flash Flood', 'Cold Wave', 'Heat Wave', 'Landslide', 'Cyclone', 'Fog'];
  function normContent(s) {
    var t = String(s || '').toLowerCase().replace(/[_\-]+/g, ' ').trim();
    if (!t) return '';
    for (var i = 0; i < CONTENT_RULES.length; i++) if (CONTENT_RULES[i][0].test(t)) return CONTENT_RULES[i][1];
    var k = key(t), hit = null;
    CONTENT_CANON.forEach(function (c) { if (!hit && closeEnough(k, key(c))) hit = c; });
    return hit || t.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  }
  var PROJECTS = { gffo: 'GFFO', rimes: 'RIMES', unicef: 'UNICEF', echo: 'ECHO', caritas: 'Caritas', concern: 'Concern', enrich: 'Enrich',
    protistha: 'Protistha', protista: 'Protistha', provati: 'PROVATi3', provati3: 'PROVATi3', wfp: 'WFP', fao: 'FAO', undp: 'UNDP', ifrc: 'IFRC', brac: 'BRAC' };
  function normProject(s) {
    var k = key(s);
    if (!k) return '';
    if (PROJECTS[k]) return PROJECTS[k];
    var hit = null; Object.keys(PROJECTS).forEach(function (p) { if (!hit && closeEnough(k, p)) hit = PROJECTS[p]; });
    return hit || String(s).trim();
  }
  function normMedia(s) {
    var k = key(s);
    if (!k) return 'Not recorded';
    if (/^gp|grameen/.test(k)) return 'GP';
    if (/kotha/.test(k)) return 'Kotha';
    if (/robi/.test(k)) return 'Robi';
    if (/banglalink|^bl$/.test(k)) return 'Banglalink';
    if (/teletalk/.test(k)) return 'Teletalk';
    return String(s).trim();
  }
  function monthIdx(s) { return MONTHS.map(function (x) { return x.toLowerCase(); }).indexOf(String(s).slice(0, 3).toLowerCase()); }
  function parseDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) { var d0 = new Date(Date.UTC(1899, 11, 30) + v * 864e5); return new Date(d0.getUTCFullYear(), d0.getUTCMonth(), d0.getUTCDate()); }
    var s = String(v).trim(), m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
    if ((m = s.match(/(\d{1,2})(?:st|nd|rd|th)?[\s\-]+([A-Za-z]{3,})[\s,\-]+(\d{4})/)) && monthIdx(m[2]) >= 0) return new Date(+m[3], monthIdx(m[2]), +m[1]);
    if ((m = s.match(/([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})/)) && monthIdx(m[1]) >= 0) return new Date(+m[3], monthIdx(m[1]), +m[2]);
    if ((m = s.match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
    if ((m = s.match(/_(\d{1,2})_(\d{1,2})_(\d{2,4})/))) { var y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1]); }
    return null;
  }
  var num = function (v) { if (v == null || v === '') return null; var n = Number(String(v).replace(/[, ]/g, '')); return isNaN(n) ? null : n; };

  // =====================================================================
  // Sheet parsing: header detection -> rows -> auto-correct -> de-duplicate
  // =====================================================================
  var COLS = ['date', 'district', 'upazila', 'content', 'total', 'success', 'failed', 'project', 'media']; // sheet column order
  var HEAD_RULES = { date: /^date/, district: /district|zila|zilla|area/, upazila: /upazila|upozila|thana|union/, content: /message|content|hazard|alert|^vm/,
    total: /total|dissem|dissim|calls|sent/, success: /success|deliver|answer|reach/, failed: /fail|unsucc/, project: /project|donor|partner|program/, media: /media|channel|operator|platform/ };
  function detectHeader(values) {
    for (var r = 0; r < Math.min(values.length, 12); r++) {
      var row = values[r] || [], map = {}, hasDate = false;
      row.forEach(function (c, i) {
        var t = String(c == null ? '' : c).trim().toLowerCase();
        if (!t) return;
        Object.keys(HEAD_RULES).forEach(function (f) { if (map[f] == null && HEAD_RULES[f].test(t) && !(f === 'success' && /fail/.test(t))) map[f] = i; });
        if (/^date/.test(t)) hasDate = true;
      });
      if (hasDate) {
        if (Object.keys(map).length < 4) { map = {}; COLS.forEach(function (c, i) { map[c] = i; }); } // e.g. "Date | Area | …" -> sheet column order
        return { row: r, map: map };
      }
    }
    return null;
  }
  function defaultProjectFor(name) {
    var k = key(name), hit = '';
    Object.keys(PROJECTS).forEach(function (p) { if (!hit && k.indexOf(p) >= 0) hit = PROJECTS[p]; });
    return hit;
  }
  function parseSheets(raw) {
    var out = { rows: [], ben: [], fixes: [], issues: [], dupes: [], sheets: [] };
    var sheets = (raw.sheets || []).slice();
    if (raw.broadcasts) sheets.push({ name: 'Sheet1', source: 'API', values: objectsToValues(raw.broadcasts) });
    var occ = {}; // broadcast key -> {sheetIndex: count}
    sheets.forEach(function (sh, si) {
      var values = (sh.values || []).map(function (r) { return Array.isArray(r) ? r : []; });
      var label = (sh.source ? sh.source + ' › ' : '') + sh.name;
      out.ben = out.ben.concat(readBeneficiaries(values, sh));
      var h = detectHeader(values);
      if (!h) { out.sheets.push({ label: label, rows: 0, kind: 'summary / other (no Date column)' }); return; }
      var defProject = defaultProjectFor(sh.name) || defaultProjectFor(sh.source || '');
      var count = 0;
      for (var r = h.row + 1; r < values.length; r++) {
        var row = values[r], g = function (f) { return h.map[f] == null ? '' : row[h.map[f]]; };
        if (!row.some(function (c) { return c !== '' && c != null; })) continue;
        var rec = normalizeRow({ date: g('date'), district: g('district'), upazila: g('upazila'), content: g('content'), total: g('total'),
          success: g('success'), failed: g('failed'), project: g('project'), media: g('media') }, defProject);
        if (rec.calls == null && !rec.date) continue;
        rec.src = label; rec.rowNo = r + 1; rec.source = sh.source || 'Google Sheet';
        // the same broadcast logged in two tabs is counted once (repeats inside one tab are real repeats and kept)
        var dk = [dstr(rec.date), rec.calls, rec.success, rec.failed, rec.project, rec.media].join('|');
        var o = occ[dk] || (occ[dk] = {});
        o[si] = (o[si] || 0) + 1;
        var other = Math.max.apply(null, [0].concat(Object.keys(o).filter(function (s) { return +s !== si; }).map(function (s) { return o[s]; })));
        if (o[si] <= other) { out.dupes.push(rec); continue; }
        rec.fixes.forEach(function (f) { out.fixes.push({ rec: rec, text: f }); });
        out.rows.push(rec); count++;
      }
      out.sheets.push({ label: label, rows: count, kind: 'dissemination log' });
    });
    var byDay = {};
    out.rows.forEach(function (r) {
      if (!r.district) out.issues.push({ rec: r, text: 'Zila not recognised: "' + (r.raw.district || '') + '"' });
      if (!r.contentRaw && !r.campaign) out.issues.push({ rec: r, text: 'Message content is empty (shown as "Not specified")' });
      if (r.success != null && r.failed != null && r.calls != null && r.success + r.failed > r.calls) out.issues.push({ rec: r, text: 'Success + Failed (' + fmt(r.success + r.failed) + ') is more than Total (' + fmt(r.calls) + ')' });
      if (r.project && !PROJECTS[key(r.project)] && (matchDistrict(r.project) || UPZ[key(r.project)])) out.issues.push({ rec: r, text: 'Project "' + r.project + '" looks like a place name' });
      var k = dstr(r.date) + '|' + r.calls + '|' + r.media;
      (byDay[k] = byDay[k] || []).push(r);
    });
    Object.keys(byDay).forEach(function (k) {
      var g = byDay[k], ps = uniq(g.map(function (r) { return r.project; }));
      if (ps.length > 1 && g[0].calls && g[0].date) out.issues.push({ rec: g[0], text: 'Possible duplicate across projects (' + ps.join(' / ') + '): same date, total (' + fmt(g[0].calls) + ') and media' });
    });
    out.rows.forEach(function (r, i) { r.id = i + 1; });
    return out;
  }
  function objectsToValues(objs) {
    var hdr = []; objs.forEach(function (o) { Object.keys(o).forEach(function (k) { if (hdr.indexOf(k) < 0) hdr.push(k); }); });
    return [hdr].concat(objs.map(function (o) { return hdr.map(function (k) { return o[k]; }); }));
  }
  function readBeneficiaries(values, sh) {
    var out = [], defProject = defaultProjectFor(sh.name) || defaultProjectFor(sh.source || '') || 'Unassigned';
    var hr = -1, cols = {};
    values.slice(0, 5).forEach(function (row, r) { // (a) a table: Project | District | Upazila | Unique Beneficiaries
      if (hr >= 0) return;
      var c = {}; row.forEach(function (v, i) { var t = String(v || '').toLowerCase(); if (/benefic/.test(t)) c.count = i; if (/project/.test(t)) c.project = i; if (/district|zila/.test(t)) c.district = i; if (/upazila/.test(t)) c.upazila = i; });
      if (c.count != null && c.upazila != null) { hr = r; cols = c; }
    });
    if (hr >= 0) {
      values.slice(hr + 1).forEach(function (row) { var n = num(row[cols.count]); if (n != null) out.push(benRec(row[cols.project] || defProject, row[cols.district], row[cols.upazila], n)); });
      return out;
    }
    var on = false; // (b) a block under a "Unique Beneficiary" title, like the Protistha Sheet2
    values.forEach(function (row) {
      if (/^unique\s*beneficiar/i.test(String(row[0] || '').trim())) { on = true; return; }
      if (on && row[0] && row[1] && num(row[2]) != null) out.push(benRec(defProject, row[0], row[1], num(row[2])));
    });
    return out;
  }
  function benRec(project, district, upazila, n) {
    var d = matchDistrict(district), u = matchUpazila(upazila, d);
    return { project: normProject(project) || 'Unassigned', district: d || (u && u.district) || String(district || ''), upazila: u ? u.name : String(upazila || ''), count: n };
  }

  function normalizeRow(c, defProject) {
    var fixes = [], note = function (field, from, to) { if (String(from == null ? '' : from).trim() && to && key(from) !== key(to)) fixes.push(field + ': "' + String(from).trim() + '" → "' + to + '"'); };
    var rec = { raw: c, fixes: fixes, campaign: '', river: '' };
    rec.date = parseDate(c.date);
    // old Protistha rows hold a campaign name in the Date cell, e.g. "22-07-2026_Ulipur_Brahma_PROVATi3"
    if (typeof c.date === 'string' && /[a-z]/i.test(c.date) && !/^\d{1,2}\s+[A-Za-z]+,?\s+\d{4}$/.test(c.date.trim()) && !/^\d{4}-\d{2}-\d{2}/.test(c.date)) rec.campaign = c.date.trim();
    var text = rec.campaign;
    // --- place ---
    var district = matchDistrict(c.district), ups = [], label = '', upText = String(c.upazila == null ? '' : c.upazila).trim();
    if (!district && String(c.district || '').trim()) {
      var u0 = matchUpazila(c.district, null); // the Zila cell may hold an upazila ("Derai", "Teknaf", "naikhomchari")
      var parts = String(c.district).trim().split(/\s+/), d1 = parts.length > 1 ? matchDistrict(parts[0]) : null;
      if (d1) { district = d1; if (!upText) upText = parts.slice(1).join(' '); fixes.push('Zila: "' + String(c.district).trim() + '" → "' + d1 + '"'); }
      else if (u0) { district = u0.district; if (!upText) upText = u0.name; fixes.push('Zila: "' + String(c.district).trim() + '" is an upazila → "' + u0.district + '" / "' + u0.name + '"'); }
    } else note('Zila', c.district, district);
    if (/^all$/i.test(upText)) label = 'All upazilas';
    else if (upText) {
      upText.split(/\s*(?:,|;|\/|&|\band\b)\s*/i).forEach(function (p) {
        if (!p.trim()) return;
        var u = matchUpazila(p, district);
        if (u) { if (ups.indexOf(u) < 0) ups.push(u); note('Upazila', p, u.name); }
        else ups.push({ name: p.trim(), district: district, geo: null, unknown: true });
      });
    }
    if (!ups.length && !label && text) { // infer upazila from the campaign text
      (district ? (UPZ_BY_DIST[district] || []) : UPZ_LIST).forEach(function (u) { if (!ups.length && key(u.name).length >= 5 && key(text).indexOf(key(u.name)) >= 0) ups.push(u); });
      if (!ups.length && /fulch/i.test(text) && UPZ.fulchari) ups.push(UPZ.fulchari);
    }
    if (!district && ups.length && ups[0].district) district = ups[0].district;
    rec.district = district || '';
    rec.upazilas = ups.map(function (u) { return u.name; });
    rec.upazilaLabel = label || rec.upazilas.slice().sort().join(', ');
    // --- content ---
    rec.contentRaw = String(c.content == null ? '' : c.content).trim();
    rec.content = normContent(rec.contentRaw);
    if (!rec.content && text) rec.content = /flood|rain|light|flash/i.test(text) ? normContent(text) : /teesta|tista|brahma/i.test(text) ? 'River Flood Bulletin' : '';
    if (!rec.content) rec.content = 'Not specified';
    note('Message content', rec.contentRaw, rec.content);
    // --- project / media / numbers ---
    rec.project = normProject(c.project);
    if (!rec.project) rec.project = /provat/i.test(text) ? 'PROVATi3' : /protist/i.test(text) ? 'Protistha' : (defProject || 'Unassigned');
    else note('Project', c.project, rec.project);
    rec.media = normMedia(c.media); note('Media', c.media, rec.media);
    rec.river = /teesta|tista/i.test(text) ? 'Teesta' : /brahma/i.test(text) ? 'Brahmaputra' : '';
    rec.calls = num(c.total); rec.success = num(c.success); rec.failed = num(c.failed);
    return rec;
  }

  // =====================================================================
  // State, colours, loading
  // =====================================================================
  var DATA = { rows: [], ben: [], fixes: [], issues: [], dupes: [], sheets: [] }, RAW = { sheets: [] }, IMPORTED = [];
  var VIEW = [], COLORS = {}, charts = {}, map, distLayer, upzLayer, legendCtl, sortState = {};
  function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
  function seriesColor(i) { return cssVar('--s' + ((i % 8) + 1)); }
  var LIGHT_SERIES = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'];
  var SEQ = ['#cde2fb', '#9ec5f4', '#6da7ec', '#3987e5', '#256abf', '#184f95', '#0d366b'];
  function assignColors() { // colour follows the entity, fixed from the full dataset
    COLORS = {};
    ['content', 'media', 'project', 'district'].forEach(function (dim) {
      var tot = {}; COLORS[dim] = {};
      DATA.rows.forEach(function (r) { tot[r[dim]] = (tot[r[dim]] || 0) + (r.calls || 0); });
      Object.keys(tot).sort(function (a, b) { return tot[b] - tot[a]; }).forEach(function (k, i) { COLORS[dim][k] = Math.min(i, 7); });
    });
  }
  function setStatus(cls, text, tip) { $('dot').className = 'dot ' + cls; $('statusText').textContent = text; $('statusText').parentNode.title = tip || text; }
  function rebuild() {
    DATA = parseSheets({ sheets: (RAW.sheets || []).concat(IMPORTED), broadcasts: RAW.broadcasts });
    assignColors(); buildFilterOptions(); apply(); renderDataCheck();
  }
  function load() {
    $('btnRefresh').disabled = true;
    var p = !API_URL ? Promise.resolve(window.SAMPLE_DATA || { sheets: [] }) :
      fetch(API_URL + (API_URL.indexOf('?') < 0 ? '?' : '&') + 'action=data&t=' + Date.now()).then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); });
    if (API_URL) setStatus('', 'Loading…');
    return p.then(function (raw) {
      if (raw.ok === false) throw new Error(raw.error);
      RAW = raw; rebuild();
      var tip = DATA.rows.length + ' broadcasts from ' + DATA.sheets.filter(function (s) { return s.rows; }).length + ' sheets';
      if (API_URL) setStatus('live', 'Live · ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }), tip);
      else setStatus('', 'Snapshot', 'Offline snapshot. ' + tip + '. Set API_URL in config.js to go live.');
    }).catch(function (e) {
      setStatus('err', 'Load failed', e.message);
      if (!DATA.rows.length && window.SAMPLE_DATA) { RAW = SAMPLE_DATA; rebuild(); }
    }).then(function () { $('btnRefresh').disabled = false; });
  }

  // =====================================================================
  // Filters
  // =====================================================================
  var F = {};
  function fillSelect(el, values, allLabel) {
    var cur = el.value;
    el.innerHTML = '<option value="">' + allLabel + '</option>' + values.map(function (v) { return '<option>' + esc(v) + '</option>'; }).join('');
    if (values.indexOf(cur) >= 0) el.value = cur;
  }
  function buildFilterOptions() {
    var R = DATA.rows;
    fillSelect($('fProject'), uniq(R.map(function (r) { return r.project; })), 'All projects');
    fillSelect($('fDistrict'), uniq(R.map(function (r) { return r.district; })), 'All zilas');
    fillUpazilas();
    fillSelect($('fContent'), uniq(R.map(function (r) { return r.content; })), 'All VM content');
    fillSelect($('fMedia'), uniq(R.map(function (r) { return r.media; })), 'All media');
    var dates = R.map(function (r) { return r.date; }).filter(Boolean);
    var months = uniq(dates.map(function (d) { return dstr(d).slice(0, 7); }));
    var years = uniq(dates.map(function (d) { return String(d.getFullYear()); }));
    var opts = [['', 'All time'], ['last30', 'Last 30 days'], ['last90', 'Last 90 days']];
    years.forEach(function (y) {
      opts.push(['y' + y, 'Year ' + y], ['m' + y, 'Monsoon ' + y + ' (Jun–Sep)']);
      [1, 2, 3, 4].forEach(function (q) { opts.push(['q' + q + y, 'Q' + q + ' ' + y]); });
    });
    months.forEach(function (m) { opts.push(['M' + m, MONTHS_LONG[+m.slice(5) - 1] + ' ' + m.slice(0, 4)]); });
    opts.push(['custom', 'Custom range']);
    var cur = $('fPreset').value;
    $('fPreset').innerHTML = opts.map(function (o) { return '<option value="' + o[0] + '">' + o[1] + '</option>'; }).join('');
    $('fPreset').value = cur || '';
    var dl = function (id, vals) { $(id).innerHTML = vals.map(function (v) { return '<option value="' + esc(v) + '">'; }).join(''); };
    dl('dlProject', uniq(R.map(function (r) { return r.project; }).concat(['GFFO', 'RIMES', 'UNICEF', 'ECHO', 'Caritas', 'Concern', 'Enrich', 'Protistha'])));
    dl('dlUpazila', uniq(UPZ_LIST.map(function (u) { return u.name; })));
    dl('dlDistrict', uniq(Object.keys(DISTRICTS).map(function (k) { return DISTRICTS[k]; })));
    dl('dlContent', CONTENT_CANON);
    dl('dlMedia', ['GP', 'Kotha', 'Robi', 'Banglalink', 'Teletalk']);
  }
  function fillUpazilas() {
    var d = $('fDistrict').value, s = [];
    DATA.rows.forEach(function (r) { if (!d || r.district === d) s = s.concat(r.upazilas); });
    fillSelect($('fUpazila'), uniq(s), 'All upazilas');
  }
  function presetRange(p) {
    var now = new Date(), t = new Date(now.getFullYear(), now.getMonth(), now.getDate()), m;
    if (p === 'last30' || p === 'last90') { var f = new Date(t); f.setDate(f.getDate() - (p === 'last30' ? 30 : 90)); return [f, t]; }
    if ((m = p.match(/^y(\d{4})$/))) return [new Date(+m[1], 0, 1), new Date(+m[1], 11, 31)];
    if ((m = p.match(/^m(\d{4})$/))) return [new Date(+m[1], 5, 1), new Date(+m[1], 8, 30)];
    if ((m = p.match(/^q(\d)(\d{4})$/))) return [new Date(+m[2], (m[1] - 1) * 3, 1), new Date(+m[2], m[1] * 3, 0)];
    if ((m = p.match(/^M(\d{4})-(\d{2})$/))) return [new Date(+m[1], m[2] - 1, 1), new Date(+m[1], +m[2], 0)];
    return null;
  }
  function apply() {
    F = { project: $('fProject').value, district: $('fDistrict').value, upazila: $('fUpazila').value, content: $('fContent').value, media: $('fMedia').value, from: $('fFrom').value, to: $('fTo').value };
    var from = F.from ? parseDate(F.from) : null, to = F.to ? parseDate(F.to) : null;
    VIEW = DATA.rows.filter(function (r) {
      if (F.project && r.project !== F.project) return false;
      if (F.district && r.district !== F.district) return false;
      if (F.upazila && r.upazilas.indexOf(F.upazila) < 0) return false;
      if (F.content && r.content !== F.content) return false;
      if (F.media && r.media !== F.media) return false;
      if ((from || to) && !r.date) return false;
      if (from && r.date < from) return false;
      if (to && r.date > to) return false;
      return true;
    }).sort(function (a, b) { return (a.date ? +a.date : 0) - (b.date ? +b.date : 0); });
    $('chips').textContent = filterText() + '  ·  ' + VIEW.length + ' of ' + DATA.rows.length + ' broadcasts';
    renderKpis(); renderCharts(); renderMap(); renderDetail(); renderSummary();
  }
  function filterText() {
    return ['Project: ' + (F.project || 'All'),
      'Timeframe: ' + (F.from || F.to ? (F.from ? dnice(parseDate(F.from)) : 'start') + ' – ' + (F.to ? dnice(parseDate(F.to)) : 'latest') : 'All time'),
      'Zila: ' + (F.district || 'All'), 'Upazila: ' + (F.upazila || 'All'), 'VM content: ' + (F.content || 'All'), 'Media: ' + (F.media || 'All')].join('  ·  ');
  }

  // =====================================================================
  // Aggregates
  // =====================================================================
  var sum = function (rows, f) { return rows.reduce(function (a, r) { return a + (r[f] || 0); }, 0); };
  function groupBy(rows, fn) {
    var g = {}, order = [];
    rows.forEach(function (r) { var k = fn(r); if (!g[k]) { g[k] = []; order.push(k); } g[k].push(r); });
    return order.map(function (k) { return { key: k, rows: g[k] }; });
  }
  function benList() {
    var ps = uniq(VIEW.map(function (r) { return r.project; }));
    return DATA.ben.filter(function (b) { return ps.indexOf(b.project) >= 0 && (!F.district || b.district === F.district) && (!F.upazila || b.upazila === F.upazila); });
  }
  function totals() {
    var withS = VIEW.filter(function (r) { return r.success != null && r.calls; });
    return { broadcasts: VIEW.length, calls: sum(VIEW, 'calls'), success: sum(withS, 'success'), failed: sum(VIEW, 'failed'),
      rate: sum(withS, 'calls') ? sum(withS, 'success') / sum(withS, 'calls') : null,
      zilas: uniq(VIEW.map(function (r) { return r.district; })).length, projects: uniq(VIEW.map(function (r) { return r.project; })).length,
      ben: benList().reduce(function (a, b) { return a + b.count; }, 0) };
  }
  function summaryData() {
    var areas = groupBy(VIEW, function (r) { return r.district + '|' + r.upazilaLabel; }).map(function (g) {
      var p = g.key.split('|');
      return { district: p[0] || '(unknown)', upazila: p[1] || '', broadcasts: g.rows.length, calls: sum(g.rows, 'calls') };
    }).sort(function (a, b) { return a.district.localeCompare(b.district) || a.upazila.localeCompare(b.upazila); });
    var projects = groupBy(VIEW, function (r) { return r.project; }).map(function (g) {
      var ws = g.rows.filter(function (r) { return r.success != null && r.calls; });
      var hasF = g.rows.some(function (r) { return r.failed != null; });
      return { project: g.key, broadcasts: g.rows.length, calls: sum(g.rows, 'calls'), success: ws.length ? sum(ws, 'success') : null, failed: hasF ? sum(g.rows, 'failed') : null, rate: sum(ws, 'calls') ? sum(ws, 'success') / sum(ws, 'calls') : null };
    }).sort(function (a, b) { return b.calls - a.calls; });
    return { t: totals(), areas: areas, projects: projects, ben: benList().slice().sort(function (a, b) { return a.district.localeCompare(b.district); }) };
  }

  // =====================================================================
  // KPI cards + hero ring
  // =====================================================================
  function compact(n) { if (n >= 1e6) return (n / 1e6).toFixed(n >= 1e7 ? 0 : 1) + 'M'; if (n >= 1e4) return Math.round(n / 1e3) + 'K'; return fmt(n); }
  function monthSeries() {
    var m = {};
    VIEW.forEach(function (r) {
      if (!r.date) return;
      var k = dstr(r.date).slice(0, 7), o = m[k] || (m[k] = { key: k, b: 0, c: 0, s: 0, sb: 0 });
      o.b++; o.c += r.calls || 0; if (r.success != null && r.calls) { o.s += r.success; o.sb += r.calls; }
    });
    return Object.keys(m).sort().map(function (k) { m[k].label = MONTHS[+k.slice(5) - 1]; return m[k]; });
  }
  function sparkLine(vals, color) {
    var W = 300, H = 74, n = vals.length;
    if (!n) return '<svg class="spark" viewBox="0 0 300 74"></svg>';
    var mx = Math.max.apply(null, vals), mn = Math.min.apply(null, vals.concat([0]));
    var pts = vals.map(function (v, i) { return [n === 1 ? W / 2 : 8 + i * (W - 16) / (n - 1), H - 10 - ((v - mn) / ((mx - mn) || 1)) * (H - 22)]; });
    var d = 'M' + pts[0][0] + ',' + pts[0][1];
    for (var i = 1; i < n; i++) {
      var p0 = pts[i - 2] || pts[i - 1], p1 = pts[i - 1], p2 = pts[i], p3 = pts[i + 1] || p2;
      d += ' C' + (p1[0] + (p2[0] - p0[0]) / 6) + ',' + (p1[1] + (p2[1] - p0[1]) / 6) + ' ' + (p2[0] - (p3[0] - p1[0]) / 6) + ',' + (p2[1] - (p3[1] - p1[1]) / 6) + ' ' + p2[0] + ',' + p2[1];
    }
    var id = 'g' + Math.random().toString(36).slice(2, 7), last = pts[n - 1];
    return '<svg class="spark" viewBox="0 0 300 74" preserveAspectRatio="none" aria-hidden="true"><defs><linearGradient id="' + id + '" x1="0" y1="0" x2="0" y2="1">' +
      '<stop offset="0" stop-color="' + color + '" stop-opacity=".28"/><stop offset="1" stop-color="' + color + '" stop-opacity="0"/></linearGradient></defs>' +
      '<path d="' + d + ' L' + last[0] + ',' + H + ' L' + pts[0][0] + ',' + H + 'Z" fill="url(#' + id + ')"/>' +
      '<path d="' + d + '" fill="none" stroke="' + color + '" stroke-width="2" vector-effect="non-scaling-stroke"/>' +
      '<circle cx="' + last[0] + '" cy="' + last[1] + '" r="9" fill="' + color + '" opacity=".18"/><circle cx="' + last[0] + '" cy="' + last[1] + '" r="4.5" fill="#fff" stroke="' + color + '" stroke-width="2"/></svg>';
  }
  function sparkBars(vals, color, flat) {
    var W = 300, H = 74, n = vals.length, mx = Math.max.apply(null, vals.concat([1]));
    if (!n) return '<svg class="spark" viewBox="0 0 300 74"></svg>';
    var bw = Math.min(16, (W - 16) / n * 0.55);
    return '<svg class="spark" viewBox="0 0 300 74" preserveAspectRatio="none" aria-hidden="true">' + vals.map(function (v, i) {
      var x = n === 1 ? W / 2 : 8 + i * (W - 16) / (n - 1), h = Math.max(3, (v / mx) * (H - 14));
      return '<rect x="' + (x - bw / 2) + '" y="' + (H - 4 - h) + '" width="' + bw + '" height="' + h + '" rx="' + Math.min(4, bw / 2) + '" fill="' + color + '" opacity="' + (flat || i === n - 1 ? 1 : .55) + '"><title>' + fmt(v) + '</title></rect>';
    }).join('') + '</svg>';
  }
  function axis(ms) {
    if (!ms.length) return '<div class="axis"><span>no dated rows</span></div>';
    var pick = ms.length <= 4 ? ms : [ms[0], ms[Math.floor(ms.length / 3)], ms[Math.floor(2 * ms.length / 3)], ms[ms.length - 1]];
    return '<div class="axis">' + pick.map(function (m) { return '<span>' + m.label + '</span>'; }).join('') + '</div>';
  }
  function renderKpis() {
    var t = totals(), ms = monthSeries();
    var zl = groupBy(VIEW, function (r) { return r.district; }).map(function (g) { return { k: g.key, v: sum(g.rows, 'calls') }; }).sort(function (a, b) { return b.v - a.v; }).slice(0, 8);
    var rates = ms.filter(function (m) { return m.sb; });
    var cards = [
      { k: 'VM broadcasts', badge: ms.length + ' months', v: fmt(t.broadcasts), viz: sparkBars(ms.map(function (m) { return m.b; }), cssVar('--s2')), ax: axis(ms) },
      { k: 'Total dissemination', badge: 'calls', v: compact(t.calls), viz: sparkLine(ms.map(function (m) { return m.c; }), cssVar('--s1')), ax: axis(ms) },
      { k: 'Success rate', badge: fmt(t.success) + ' reached', v: pct(t.rate), viz: sparkLine(rates.map(function (m) { return m.s / m.sb; }), cssVar('--accent')), ax: axis(rates) },
      { k: 'Zilas covered', badge: t.projects + ' projects', v: fmt(t.zilas), viz: sparkBars(zl.map(function (z) { return z.v; }), cssVar('--s7'), true),
        ax: '<div class="axis">' + zl.slice(0, 5).map(function (z) { return '<span>' + esc(String(z.k).slice(0, 6)) + '</span>'; }).join('') + '</div>' }
    ];
    $('kpis').innerHTML = cards.map(function (c) {
      return '<div class="kpi glass"><div class="k">' + c.k + '<span class="badge">' + c.badge + '</span></div><div class="v">' + c.v + '</div>' + c.viz + c.ax + '</div>';
    }).join('');
    renderHero(t);
  }
  function renderHero(t) { // ring: total calls in the centre, one arc per project
    var h = new Date().getHours();
    $('greet').innerHTML = (h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening') + ', team 👋<span>here is where your voice messages reached</span>';
    var su = groupBy(VIEW, function (r) { return r.project; }).map(function (g) { return { name: g.key, calls: sum(g.rows, 'calls') }; }).filter(function (s) { return s.calls > 0; }).sort(function (a, b) { return b.calls - a.calls; });
    var C = 240, R = 158, SW = 30, total = su.reduce(function (a, s) { return a + s.calls; }, 0) || 1, gap = su.length > 1 ? 0.035 : 0, a0 = -Math.PI / 2;
    var pt = function (r, a) { return [C + r * Math.cos(a), C + r * Math.sin(a)]; };
    var darkUI = getComputedStyle(document.documentElement).colorScheme === 'dark';
    var svg = '<defs><radialGradient id="ringGlass" cx="40%" cy="35%" r="70%"><stop offset="0" stop-color="#fff" stop-opacity=".9"/><stop offset="1" stop-color="#fff" stop-opacity=".35"/></radialGradient></defs>' +
      '<circle cx="' + C + '" cy="' + C + '" r="222" fill="none" stroke="currentColor" stroke-opacity=".18"/>' +
      '<circle cx="' + C + '" cy="' + C + '" r="' + R + '" fill="none" stroke="currentColor" stroke-opacity=".07" stroke-width="' + SW + '"/>';
    su.forEach(function (s) {
      var sweep = s.calls / total * Math.PI * 2, a1 = a0 + gap / 2, a2 = a0 + sweep - gap / 2, tiny = a2 <= a1;
      if (tiny) { var mid = a0 + sweep / 2; a1 = mid - 0.012; a2 = mid + 0.012; }
      var p1 = pt(R, a1), p2 = pt(R, a2), col = seriesColor(COLORS.project[s.name] || 0);
      svg += '<path d="M' + p1[0] + ',' + p1[1] + ' A' + R + ',' + R + ' 0 ' + (a2 - a1 > Math.PI ? 1 : 0) + ' 1 ' + p2[0] + ',' + p2[1] + '" fill="none" stroke="' + col + '" stroke-width="' + SW + '" stroke-linecap="' + (tiny ? 'butt' : 'round') + '"><title>' + esc(s.name) + ': ' + fmt(s.calls) + ' calls (' + Math.round(s.calls / total * 100) + '%)</title></path>';
      if (s.calls / total >= 0.02) {
        var am = (a1 + a2) / 2, lp = pt(R + SW / 2 + 22, am), anchor = Math.cos(am) > 0.25 ? 'start' : Math.cos(am) < -0.25 ? 'end' : 'middle';
        svg += '<text x="' + lp[0] + '" y="' + (lp[1] - 2) + '" text-anchor="' + anchor + '" font-size="13" font-weight="700" fill="currentColor">' + esc(s.name) + '</text>' +
          '<text x="' + lp[0] + '" y="' + (lp[1] + 14) + '" text-anchor="' + anchor + '" font-size="12" fill="currentColor" fill-opacity=".65">' + Math.round(s.calls / total * 100) + '% · ' + compact(s.calls) + '</text>';
      }
      a0 += sweep;
    });
    svg += '<circle cx="' + C + '" cy="' + C + '" r="122" fill="url(#ringGlass)" fill-opacity="' + (darkUI ? '.10' : '.55') + '" stroke="#fff" stroke-opacity="' + (darkUI ? '.18' : '.7') + '"/>' +
      '<text x="' + C + '" y="' + (C + 10) + '" text-anchor="middle" font-size="58" font-weight="300" fill="currentColor" letter-spacing="-1">' + compact(t.calls) + '</text>' +
      '<text x="' + C + '" y="' + (C + 42) + '" text-anchor="middle" font-size="14" fill="currentColor" fill-opacity=".7">Total dissemination</text>' +
      '<text x="' + C + '" y="' + (C + 62) + '" text-anchor="middle" font-size="12" fill="currentColor" fill-opacity=".55">' + fmt(t.calls) + ' calls · ' + su.length + ' projects</text>';
    $('ring').innerHTML = svg; $('ring').style.color = cssVar('--text');
    var last = VIEW.reduce(function (a, r) { return r.date && (!a || r.date > a) ? r.date : a; }, null);
    $('herofoot').innerHTML =
      '<div class="mini">Top project<b>' + (su[0] ? esc(su[0].name) + ' · ' + compact(su[0].calls) : '–') + '</b></div>' +
      '<div class="mini">Latest broadcast<b>' + dnice(last) + '</b></div>' +
      '<div class="mini">Unique beneficiaries<b>' + (t.ben ? fmt(t.ben) : 'not recorded') + '</b></div>';
  }

  // =====================================================================
  // Charts
  // =====================================================================
  function chartConfigs(light) {
    var col = light ? function (i) { return LIGHT_SERIES[i]; } : seriesColor;
    var ink = light ? '#52514e' : cssVar('--text-2'), grid = light ? '#ecebe7' : cssVar('--grid'), surface = light ? '#ffffff' : cssVar('--surface');
    var base = function (extra) {
      var o = { responsive: !light, maintainAspectRatio: false, animation: light ? false : { duration: 250 },
        plugins: { legend: { position: 'bottom', labels: { color: ink, boxWidth: 8, boxHeight: 8, usePointStyle: true, pointStyle: 'circle', padding: 14 } },
          tooltip: { mode: 'index', intersect: false, backgroundColor: 'rgba(31,29,43,.92)', padding: 10, cornerRadius: 12, boxPadding: 4 } },
        scales: { x: { ticks: { color: ink }, grid: { display: false } }, y: { ticks: { color: ink }, grid: { color: grid }, beginAtZero: true } } };
      return Object.assign(o, extra || {});
    };
    var monthKeys = uniq(VIEW.filter(function (r) { return r.date; }).map(function (r) { return dstr(r.date).slice(0, 7); }));
    var contents = Object.keys(COLORS.content).filter(function (c) { return VIEW.some(function (r) { return r.content === c; }); });
    var monthly = { type: 'bar', data: { labels: monthKeys.map(function (m) { return MONTHS[+m.slice(5) - 1] + ' ' + m.slice(2, 4); }),
      datasets: contents.map(function (c) {
        return { label: c, backgroundColor: c === 'Not specified' ? '#a3a29b' : col(COLORS.content[c]), borderColor: surface, borderWidth: { top: 2 }, borderRadius: 4, borderSkipped: 'bottom', maxBarThickness: 38,
          data: monthKeys.map(function (m) { return sum(VIEW.filter(function (r) { return r.content === c && r.date && dstr(r.date).slice(0, 7) === m; }), 'calls'); }) };
      }) }, options: base() };
    monthly.options.scales.x.stacked = true; monthly.options.scales.y.stacked = true;
    var zl = groupBy(VIEW, function (r) { return r.district || '(unknown)'; }).map(function (g) { return { k: g.key, v: sum(g.rows, 'calls') }; }).sort(function (a, b) { return b.v - a.v; });
    var byZila = { type: 'bar', data: { labels: zl.map(function (z) { return z.k; }), datasets: [{ label: 'Total calls', data: zl.map(function (z) { return z.v; }), backgroundColor: col(0), borderRadius: 4, maxBarThickness: 22 }] }, options: base({ indexAxis: 'y' }) };
    byZila.options.plugins.legend.display = false;
    byZila.options.scales = { x: { ticks: { color: ink }, grid: { color: grid }, beginAtZero: true }, y: { ticks: { color: ink, autoSkip: false }, grid: { display: false } } };
    var media = Object.keys(COLORS.media).filter(function (m) { return VIEW.some(function (r) { return r.media === m; }); });
    var byMedia = { type: 'bar', data: { labels: media, datasets: [{ label: 'Broadcasts', data: media.map(function (m) { return VIEW.filter(function (r) { return r.media === m; }).length; }),
      backgroundColor: media.map(function (m) { return m === 'Not recorded' ? '#a3a29b' : col(COLORS.media[m]); }), borderRadius: 4, maxBarThickness: 60 }] }, options: base() };
    byMedia.options.plugins.legend.display = false;
    var pr = summaryData().projects, stackDs = function (label, data, color) { return { label: label, data: data, backgroundColor: color, borderColor: surface, borderWidth: { top: 2 }, borderRadius: 4, borderSkipped: 'bottom', maxBarThickness: 40 }; };
    var byProject = { type: 'bar', data: { labels: pr.map(function (p) { return p.project; }), datasets: [
      stackDs('Success', pr.map(function (p) { return p.success || 0; }), col(2)), stackDs('Failed', pr.map(function (p) { return p.failed || 0; }), col(1)),
      stackDs('Not reported', pr.map(function (p) { return Math.max(0, p.calls - (p.success || 0) - (p.failed || 0)); }), '#b9b8b1')] }, options: base() };
    byProject.options.scales.x.stacked = true; byProject.options.scales.y.stacked = true;
    return { chMonthly: monthly, chZila: byZila, chMedia: byMedia, chProject: byProject };
  }
  function renderCharts() {
    if (!window.Chart) return;
    Chart.defaults.font.family = 'Manrope, "Segoe UI", sans-serif';
    var cfgs = chartConfigs(false);
    $('chZila').parentNode.style.height = Math.max(290, 24 * cfgs.chZila.data.labels.length + 40) + 'px';
    Object.keys(cfgs).forEach(function (id) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart($(id), cfgs[id]); });
  }
  function chartImage(id, w, h, type) {
    var cfg = chartConfigs(true)[id], cv = document.createElement('canvas');
    cv.width = w; cv.height = h; cv.style.cssText = 'position:fixed;left:-9999px;width:' + w + 'px;height:' + h + 'px';
    document.body.appendChild(cv);
    cfg.options.devicePixelRatio = 2;
    cfg.plugins = [{ id: 'bg', beforeDraw: function (c) { var x = c.ctx; x.save(); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.restore(); } }];
    var ch = new Chart(cv, cfg), url = ch.toBase64Image(type || 'image/png', 0.92);
    ch.destroy(); cv.remove();
    return url;
  }

  // =====================================================================
  // Map (zila choropleth + upazilas named in broadcasts)
  // =====================================================================
  // Project area = every upazila named in the filtered broadcasts. A zila whose broadcasts name no
  // upazila ("All" or blank) and has no named upazila is shown as a zila-wide area (light blue).
  var PA_BLUE = '#0b4ea2', PA_ZILA = '#9dbbe3';
  function projectArea() {
    var byKey = {}, list = [];
    function add(f, r) {
      var k = f.properties.display + '|' + f.properties.district;
      if (!byKey[k]) { byKey[k] = { f: f, rows: [] }; list.push(byKey[k]); }
      if (byKey[k].rows.indexOf(r) < 0) byKey[k].rows.push(r);
    }
    VIEW.forEach(function (r) {
      if (r.upazilas.length) {
        r.upazilas.forEach(function (n) {
          var all = GEO.upazilas.features.filter(function (f) { return f.properties.display === n; });
          var same = all.filter(function (f) { return f.properties.district === r.district; });
          (same.length ? same : all.slice(0, 1)).forEach(function (f) { add(f, r); });
        });
      }
    });
    var zilas = uniq(list.map(function (a) { return a.f.properties.district; }));
    var wide = uniq(VIEW.filter(function (r) { return r.district && zilas.indexOf(r.district) < 0; }).map(function (r) { return r.district; }));
    var wideF = GEO.districts.features.filter(function (f) { return wide.indexOf(f.properties.name) >= 0; });
    return { list: list, zilas: uniq(zilas.concat(wide)), wide: wideF };
  }
  function areaPopup(a) {
    var rows = a.rows, ps = groupBy(rows, function (r) { return r.project; }).map(function (g) { return esc(g.key) + ' (' + g.rows.length + ')'; }).join(', ');
    return '<b>' + esc(a.f.properties.display) + '</b> Upazila, ' + esc(a.f.properties.district) + ' Zila<br>Broadcasts: <b>' + rows.length + '</b><br>Total calls: <b>' + fmt(sum(rows, 'calls')) +
      '</b><br>Success: <b>' + fmt(sum(rows, 'success')) + '</b><br>Projects: ' + ps + '<br>VM content: ' + esc(uniq(rows.map(function (r) { return r.content; }).filter(Boolean)).join(', '));
  }
  var NORTH_SVG = '<svg viewBox="0 0 40 70" width="30" height="52"><text x="20" y="17" text-anchor="middle" font-family="Times New Roman,serif" font-size="19">N</text><path d="M20 22 L33 66 L20 57 L7 66 Z" fill="#111"/></svg>';
  function renderMap() {
    if (!window.L || !window.GEO) return;
    var A = projectArea(), inArea = {};
    A.zilas.forEach(function (z) { inArea[z] = 1; });
    if (!map) {
      map = L.map('map', { zoomSnap: 0.25 });
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 18, attribution: '© OpenStreetMap contributors' }).addTo(map);
      map.setView([23.7, 90.3], 7);
      map.on('zoomend', function () { $('map').classList.toggle('zoomed', map.getZoom() >= 8.5); });
      L.control.scale({ position: 'bottomright', imperial: false, maxWidth: 180 }).addTo(map);
      var north = L.control({ position: 'topright' });
      north.onAdd = function () { var d = L.DomUtil.create('div', 'north-arrow'); d.innerHTML = NORTH_SVG; return d; };
      north.addTo(map);
    }
    [distLayer, upzLayer, legendCtl].forEach(function (l) { if (l) l.remove(); });
    distLayer = L.geoJSON(GEO.districts, {
      style: function (f) { var w = A.wide.indexOf(f) >= 0; return { color: '#6b6b6b', weight: 0.9, fillColor: w ? PA_ZILA : '#e6e6e6', fillOpacity: w ? 0.9 : 0.6 }; },
      onEachFeature: function (f, layer) {
        var n = f.properties.name;
        if (A.wide.indexOf(f) >= 0) { var zr = VIEW.filter(function (r) { return r.district === n; }); layer.bindPopup(areaPopup({ f: { properties: { display: 'All upazilas', district: n } }, rows: zr })); }
        layer.bindTooltip(n, { sticky: true, direction: 'top', className: 'dist-tip' });
        if (!inArea[n]) return;
        layer.on('click', function () { $('fDistrict').value = $('fDistrict').value === n ? '' : n; fillUpazilas(); $('fUpazila').value = ''; apply(); });
      }
    }).addTo(map);
    var lookup = new Map(A.list.map(function (a) { return [a.f, a]; }));
    upzLayer = L.geoJSON({ type: 'FeatureCollection', features: A.list.map(function (a) { return a.f; }) }, {
      style: function (f) { var sel = F.upazila === f.properties.display; return { color: sel ? '#eb6834' : '#ffffff', weight: sel ? 3 : 1.2, fillColor: PA_BLUE, fillOpacity: 0.95 }; },
      onEachFeature: function (f, layer) {
        var n = f.properties.display;
        layer.bindPopup(areaPopup(lookup.get(f)));
        layer.bindTooltip(n, { permanent: true, direction: 'center', className: 'upz-label' });
        layer.on('click', function (e) { L.DomEvent.stopPropagation(e); $('fDistrict').value = f.properties.district; fillUpazilas(); $('fUpazila').value = $('fUpazila').value === n ? '' : n; apply(); });
      }
    }).addTo(map);
    legendCtl = L.control({ position: 'bottomleft' });
    legendCtl.onAdd = function () {
      var d = L.DomUtil.create('div', 'legend map-legend');
      d.innerHTML = '<b>Legend</b><div><i class="sw-pa"></i>Project area</div><div><i class="sw-db"></i>District boundary</div>' +
        (A.wide.length ? '<div><i class="sw-zw"></i>Zila-wide broadcast</div>' : '') + '<div class="lg-note">' + A.list.length + ' upazilas in ' + A.zilas.length + ' zilas</div>';
      return d;
    };
    legendCtl.addTo(map);
    var fitKey = F.project + '|' + F.district + '|' + F.upazila;
    if ($('p-map').classList.contains('on') && map._fitKey !== fitKey) {
      if (F.district || F.upazila) map.fitBounds(L.geoJSON({ type: 'FeatureCollection', features: A.list.length ? A.list.map(function (a) { return a.f; }) : GEO.districts.features }).getBounds(), { padding: [30, 30], maxZoom: 10 });
      else map.fitBounds(L.geoJSON(GEO.districts).getBounds(), { padding: [10, 10] });
      map._fitKey = fitKey;
    }
  }

  // Print map in a GIS layout: lat/long frame, grey land & sea, blue project area, legend, north arrow, scale bar.
  function mercY(lat) { return 180 / Math.PI * Math.log(Math.tan(Math.PI / 4 + lat * Math.PI / 360)); }
  function invMercY(y) { return (2 * Math.atan(Math.exp(y * Math.PI / 180)) - Math.PI / 2) * 180 / Math.PI; }
  function dms(v, pos, neg) { var a = Math.abs(v), d = Math.floor(a + 1e-9), m = Math.round((a - d) * 60); if (m === 60) { d++; m = 0; } return d + '°' + (m ? m + "'" : '') + (v >= 0 ? pos : neg); }
  function drawMapCanvas(W, H) {
    var A = projectArea();
    var cv = document.createElement('canvas'); cv.width = W; cv.height = H;
    var x = cv.getContext('2d'), u = W / 1654, F_ = function (sz, w, fam) { return (w || '') + ' ' + Math.round(sz * u) + 'px ' + (fam || 'Segoe UI, Arial'); };
    x.fillStyle = '#fff'; x.fillRect(0, 0, W, H);
    // title
    x.textAlign = 'center'; x.fillStyle = '#111'; x.font = F_(34, 'bold'); x.fillText('Project Area Map', W / 2, 58 * u);
    x.font = F_(17); x.fillStyle = '#444'; x.fillText(filterText(), W / 2, 90 * u);
    // frame + projection (Mercator) fitted to Bangladesh
    var R = { x: 92 * u, y: 150 * u, w: W - 184 * u, h: H - 150 * u - 110 * u };
    var lon0 = 87.85, lon1 = 92.85, lat0 = 20.35, lat1 = 26.95, my0 = mercY(lat0), my1 = mercY(lat1);
    var k = Math.min(R.w / (lon1 - lon0), R.h / (my1 - my0)), cx = (lon0 + lon1) / 2, cmy = (my0 + my1) / 2;
    var P = function (p) { return [R.x + R.w / 2 + (p[0] - cx) * k, R.y + R.h / 2 - (mercY(p[1]) - cmy) * k]; };
    var vLon0 = cx - R.w / 2 / k, vLon1 = cx + R.w / 2 / k, vLat0 = invMercY(cmy - R.h / 2 / k), vLat1 = invMercY(cmy + R.h / 2 / k);
    x.save(); x.beginPath(); x.rect(R.x, R.y, R.w, R.h); x.clip();
    x.fillStyle = '#d1d5d8'; x.fillRect(R.x, R.y, R.w, R.h);                       // sea
    if (GEO.neighbours) drawFC(x, GEO.neighbours, P, function () { return { fill: '#eeeeee', stroke: '#c6c6c6', w: 2.2 * u, dash: [10 * u, 5 * u] }; });
    drawFC(x, GEO.districts, P, function () { return { fill: null, stroke: 'rgba(255,255,255,.85)', w: 9 * u }; });   // coastal glow
    drawFC(x, GEO.districts, P, function (f) { return { fill: A.wide.indexOf(f) >= 0 ? PA_ZILA : '#e2e2e2', stroke: '#6e6e6e', w: 1.3 * u }; });
    drawFC(x, { features: A.list.map(function (a) { return a.f; }) }, P, function () { return { fill: PA_BLUE, stroke: '#ffffff', w: 1.5 * u }; });
    x.restore();
    // graticule ticks + labels
    x.strokeStyle = '#111'; x.lineWidth = 2 * u; x.strokeRect(R.x, R.y, R.w, R.h);
    x.lineWidth = 1 * u; x.fillStyle = '#222'; x.font = F_(17);
    var tick = function (x1, y1, x2, y2) { x.beginPath(); x.moveTo(x1, y1); x.lineTo(x2, y2); x.stroke(); };
    for (var lo = Math.ceil(vLon0 * 10) / 10; lo <= vLon1; lo = Math.round((lo + 0.1) * 10) / 10) {
      var px = P([lo, 0])[0], major = Math.abs(lo * 2 - Math.round(lo * 2)) < 1e-6, t = (major ? 12 : 5) * u;
      tick(px, R.y, px, R.y - t); tick(px, R.y + R.h, px, R.y + R.h + t);
      if (major) { x.textAlign = 'center'; x.fillText(dms(lo, 'E', 'W'), px, R.y - 20 * u); x.fillText(dms(lo, 'E', 'W'), px, R.y + R.h + 34 * u); }
    }
    for (var la = Math.ceil(vLat0 * 10) / 10; la <= vLat1; la = Math.round((la + 0.1) * 10) / 10) {
      var py = P([0, la])[1], majorL = Math.abs(la * 2 - Math.round(la * 2)) < 1e-6, tl = (majorL ? 12 : 5) * u;
      tick(R.x, py, R.x - tl, py); tick(R.x + R.w, py, R.x + R.w + tl, py);
      if (majorL) {
        x.textAlign = 'center';
        x.save(); x.translate(R.x - 22 * u, py); x.rotate(-Math.PI / 2); x.fillText(dms(la, 'N', 'S'), 0, 0); x.restore();
        x.save(); x.translate(R.x + R.w + 22 * u, py); x.rotate(Math.PI / 2); x.fillText(dms(la, 'N', 'S'), 0, 0); x.restore();
      }
    }
    // north arrow
    var nx = R.x + R.w - 95 * u, ny = R.y + 70 * u;
    x.fillStyle = '#111'; x.textAlign = 'center'; x.font = F_(52, '', 'Times New Roman, serif'); x.fillText('N', nx, ny);
    x.beginPath(); x.moveTo(nx, ny + 18 * u); x.lineTo(nx + 32 * u, ny + 150 * u); x.lineTo(nx, ny + 120 * u); x.lineTo(nx - 32 * u, ny + 150 * u); x.closePath(); x.fill();
    // legend
    var lw = 340 * u, lh = (A.wide.length ? 244 : 200) * u, lx = R.x + 50 * u, ly = R.y + R.h - lh - 40 * u;
    x.fillStyle = 'rgba(226,227,229,.95)'; x.fillRect(lx, ly, lw, lh);
    x.fillStyle = '#111'; x.font = F_(30, 'bold'); x.fillText('Legend', lx + lw / 2, ly + 46 * u);
    x.textAlign = 'left'; x.font = F_(26);
    x.fillStyle = PA_BLUE; x.fillRect(lx + 26 * u, ly + 74 * u, 56 * u, 26 * u); x.fillStyle = '#111'; x.fillText('Project area', lx + 98 * u, ly + 97 * u);
    x.fillStyle = '#e2e2e2'; x.fillRect(lx + 26 * u, ly + 118 * u, 56 * u, 26 * u); x.strokeStyle = '#6e6e6e'; x.lineWidth = 1.5 * u; x.strokeRect(lx + 26 * u, ly + 118 * u, 56 * u, 26 * u);
    x.fillStyle = '#111'; x.fillText('District boundary', lx + 98 * u, ly + 141 * u);
    var ny2 = ly + 180 * u;
    if (A.wide.length) { x.fillStyle = PA_ZILA; x.fillRect(lx + 26 * u, ly + 162 * u, 56 * u, 26 * u); x.fillStyle = '#111'; x.fillText('Zila-wide broadcast', lx + 98 * u, ly + 185 * u); ny2 += 44 * u; }
    x.font = F_(17); x.fillStyle = '#444'; x.fillText(A.list.length + ' upazilas in ' + A.zilas.length + ' zilas', lx + 26 * u, ny2);
    // scale bar 0-25-50-100-150 km
    var kmPx = k / (111.32 * Math.cos(((vLat0 + vLat1) / 2) * Math.PI / 180)), marks = [0, 25, 50, 100, 150];
    var sx = R.x + R.w - 150 * kmPx - 170 * u, sy = R.y + R.h - 70 * u, bh = 14 * u;
    x.lineWidth = 1.5 * u; x.strokeStyle = '#111';
    for (var i = 0; i < marks.length - 1; i++) {
      var a0 = sx + marks[i] * kmPx, a1 = sx + marks[i + 1] * kmPx;
      x.fillStyle = i % 2 ? '#d9d9d9' : '#111'; x.fillRect(a0, sy, a1 - a0, bh); x.strokeRect(a0, sy, a1 - a0, bh);
    }
    x.fillStyle = '#111'; x.font = F_(24); x.textAlign = 'center';
    marks.forEach(function (m) { x.fillText(String(m), sx + m * kmPx, sy - 10 * u); });
    x.textAlign = 'left'; x.fillText('Kilometers', sx + 150 * kmPx + 10 * u, sy + bh + 2 * u);
    // footer
    x.fillStyle = '#666'; x.font = F_(15); x.textAlign = 'left';
    x.fillText('Boundaries: geoBoundaries (CC BY 4.0), Natural Earth. Data: ' + (API_URL ? 'Google Sheet (live)' : 'snapshot') + '. Generated ' + dnice(new Date()) + (CFG.ORG ? ' · ' + CFG.ORG : ''), R.x, H - 40 * u);
    return cv;
  }
  function eachCoord(fc, fn) { fc.features.forEach(function (f) { var g = f.geometry, polys = g.type === 'Polygon' ? [g.coordinates] : g.coordinates; polys.forEach(function (pl) { pl.forEach(function (ring) { ring.forEach(fn); }); }); }); }
  function bboxOf(fc) { var bb = [180, 90, -180, -90]; eachCoord(fc, function (p) { bb[0] = Math.min(bb[0], p[0]); bb[1] = Math.min(bb[1], p[1]); bb[2] = Math.max(bb[2], p[0]); bb[3] = Math.max(bb[3], p[1]); }); return bb; }
  function frame(bb, R, padFrac) {
    var kx = Math.cos(((bb[1] + bb[3]) / 2) * Math.PI / 180), w = (bb[2] - bb[0]) * kx, h = bb[3] - bb[1];
    var sc = Math.min(R.w / (w * (1 + 2 * padFrac)), R.h / (h * (1 + 2 * padFrac))), ox = R.x + (R.w - w * sc) / 2, oy = R.y + (R.h - h * sc) / 2;
    var P = function (p) { return [ox + (p[0] - bb[0]) * kx * sc, oy + (bb[3] - p[1]) * sc]; };
    P.pxPerKm = sc / 111.32; return P;
  }
  function centroid(f) {
    var g = f.geometry, polys = g.type === 'Polygon' ? [g.coordinates] : g.coordinates, best = null, bestA = 0;
    polys.forEach(function (pl) {
      var r = pl[0], a = 0, cx = 0, cy = 0;
      for (var i = 0, j = r.length - 1; i < r.length; j = i++) { var t = r[j][0] * r[i][1] - r[i][0] * r[j][1]; a += t; cx += (r[j][0] + r[i][0]) * t; cy += (r[j][1] + r[i][1]) * t; }
      if (Math.abs(a) > bestA) { bestA = Math.abs(a); best = [cx / (3 * a), cy / (3 * a)]; }
    });
    return best;
  }
  function drawFC(x, fc, P, styleFn) {
    fc.features.forEach(function (f) {
      var s = styleFn(f), g = f.geometry, polys = g.type === 'Polygon' ? [g.coordinates] : g.coordinates;
      x.beginPath();
      polys.forEach(function (pl) { pl.forEach(function (ring) { ring.forEach(function (p, i) { var q = P(p); if (i) x.lineTo(q[0], q[1]); else x.moveTo(q[0], q[1]); }); x.closePath(); }); });
      if (s.fill) { x.fillStyle = s.fill; x.fill('evenodd'); }
      x.setLineDash(s.dash || []); x.strokeStyle = s.stroke; x.lineWidth = s.w; x.stroke(); x.setLineDash([]);
    });
  }

  // =====================================================================
  // Report tables — same columns / layout as the Excel workbook
  // =====================================================================
  var SHEET1_HEAD = ['Date', 'Districts', 'Upazila', 'Message content', 'Total Dissimination', 'Success', 'Failed', 'Project', 'Disseminated Media'];
  function sheet1Rows(rows) {
    return rows.map(function (r) {
      return [r.date ? dsheet(r.date) : r.campaign, r.district, r.upazilaLabel === 'All upazilas' ? 'All' : r.upazilaLabel, r.content === 'Not specified' ? '' : r.content,
        r.calls, r.success, r.failed, r.project, r.media === 'Not recorded' ? '' : r.media];
    });
  }
  function sheet1Total(rows) { return ['Total', '', '', rows.length + ' broadcasts', sum(rows, 'calls'), sum(rows, 'success'), sum(rows, 'failed'), '', '']; }
  // "Sheet2" row by row, like the Protistha workbook (+ a per-project block for multi-project data)
  function sheet2Layout() {
    var S = summaryData(), L = [], blank = function () { L.push(['', '', '']); };
    L.push(['Total Broadcasts=', '', S.t.broadcasts]); L.push(['Total Calls=', '', S.t.calls]); blank();
    S.areas.forEach(function (a) { L.push([a.district, a.upazila, a.calls]); });
    blank(); blank(); blank();
    var benTitle = L.length; L.push(['Unique Beneficiary', '', '']); blank();
    var benStart = L.length;
    S.ben.forEach(function (b) { L.push([b.district, b.upazila, b.count]); });
    var benEnd = L.length;
    blank(); L.push(['', 'Total', S.ben.reduce(function (a, b) { return a + b.count; }, 0)]);
    var benTotal = L.length - 1;
    blank(); blank();
    var projTitle = L.length; L.push(['By Project', '', '', '', '', '']);
    L.push(['Project', 'Broadcasts', 'Total Calls', 'Success', 'Failed', 'Success rate']);
    S.projects.forEach(function (p) { L.push([p.project, p.broadcasts, p.calls, p.success == null ? '' : p.success, p.failed == null ? '' : p.failed, p.rate == null ? '' : Math.round(p.rate * 1000) / 10 + '%']); });
    return { rows: L, benTitle: benTitle, benStart: benStart, benEnd: benEnd, benTotal: benTotal, projTitle: projTitle };
  }
  function tag(text, colorIdx) { return '<span class="tag" style="--c:' + (colorIdx == null ? '#a3a29b' : seriesColor(colorIdx)) + '">' + esc(text) + '</span>'; }
  function renderDetail() {
    var q = ($('search').value || '').toLowerCase();
    var rows = VIEW.filter(function (r) { return !q || [dsheet(r.date), r.project, r.district, r.upazilaLabel, r.content, r.media, r.campaign].join(' ').toLowerCase().indexOf(q) >= 0; });
    var s = sortState.detail;
    if (s) rows = rows.slice().sort(function (a, b) { var va = a[s.k], vb = b[s.k]; if (s.k === 'date') { va = va ? +va : 0; vb = vb ? +vb : 0; } return (va > vb ? 1 : va < vb ? -1 : 0) * s.dir; });
    var keys = ['date', 'district', 'upazilaLabel', 'content', 'calls', 'success', 'failed', 'project', 'media'];
    $('detailCount').textContent = rows.length + ' broadcasts · ' + fmt(sum(rows, 'calls')) + ' calls · same columns as the Google Sheet · ✎ = auto-corrected (hover to see)';
    $('tDetail').innerHTML = '<thead><tr>' + SHEET1_HEAD.map(function (h, i) { return '<th data-k="' + keys[i] + '" class="' + (i >= 4 && i <= 6 ? 'num' : '') + '">' + h + '</th>'; }).join('') + '</tr></thead><tbody>' +
      rows.map(function (r) {
        return '<tr><td>' + (r.date ? dsheet(r.date) : '–') + (r.campaign ? '<span class="sub2" style="padding-left:0">' + esc(r.campaign) + '</span>' : '') + '</td><td>' + esc(r.district) +
          '</td><td>' + esc(r.upazilaLabel) + '</td><td>' + tag(r.content, r.content === 'Not specified' ? null : COLORS.content[r.content]) + (r.fixes.length ? ' <span class="fixdot" title="' + esc(r.fixes.join('\n')) + '">✎</span>' : '') +
          '</td><td class="num">' + fmt(r.calls) + '</td><td class="num">' + fmt(r.success) + '</td><td class="num">' + fmt(r.failed) + '</td><td>' + tag(r.project, COLORS.project[r.project]) +
          '</td><td>' + esc(r.media === 'Not recorded' ? '–' : r.media) + '</td></tr>';
      }).join('') + '<tr class="total"><td colspan="4">Total · ' + rows.length + ' broadcasts</td><td class="num">' + fmt(sum(rows, 'calls')) + '</td><td class="num">' + fmt(sum(rows, 'success')) + '</td><td class="num">' + fmt(sum(rows, 'failed')) + '</td><td colspan="2"></td></tr></tbody>';
    $('tDetail').querySelectorAll('th').forEach(function (th) {
      th.onclick = function () { var c = sortState.detail || {}; sortState.detail = { k: th.dataset.k, dir: c.k === th.dataset.k ? -c.dir : 1 }; renderDetail(); };
    });
  }
  function renderSummary() {
    var Lo = sheet2Layout();
    $('tSummary').innerHTML = '<tbody>' + Lo.rows.map(function (r, i) {
      if (!r.some(function (c) { return c !== '' && c != null; })) return '<tr class="gap"><td colspan="6"></td></tr>';
      var cls = i < 2 || i === Lo.benTotal ? 'strong' : i === Lo.benTitle || i === Lo.projTitle ? 'title' : i === Lo.projTitle + 1 ? 'head' : '';
      var cells = r.map(function (c) { return '<td class="' + (typeof c === 'number' || /%$/.test(c) ? 'num' : '') + '">' + (typeof c === 'number' ? fmt(c) : esc(c)) + '</td>'; });
      return '<tr class="' + cls + '">' + cells.join('') + '</tr>';
    }).join('') + '</tbody>';
  }

  // =====================================================================
  // Data check (auto-corrections, duplicates, issues) + Excel/CSV import
  // =====================================================================
  function renderDataCheck() {
    var D = DATA;
    $('dcSummary').innerHTML =
      '<div class="mini">Sheets read<b>' + D.sheets.filter(function (s) { return s.rows; }).length + ' of ' + D.sheets.length + '</b></div>' +
      '<div class="mini">Broadcasts loaded<b>' + fmt(D.rows.length) + '</b></div>' +
      '<div class="mini">Auto-corrections<b>' + fmt(D.fixes.length) + '</b></div>' +
      '<div class="mini">Duplicates skipped<b>' + fmt(D.dupes.length) + '</b></div>' +
      '<div class="mini">Need attention<b>' + fmt(D.issues.length) + '</b></div>';
    $('dcSheets').innerHTML = D.sheets.map(function (s) { return '<li><b>' + esc(s.label) + '</b> — ' + (s.rows ? s.rows + ' broadcasts' : s.kind) + '</li>'; }).join('');
    var where = function (r) { return esc(r.src) + ' · row ' + r.rowNo + ' · ' + (r.date ? dnice(r.date) : '–') + ' · ' + esc(r.project); };
    $('dcIssues').innerHTML = D.issues.length ? D.issues.map(function (i) { return '<tr><td>' + where(i.rec) + '</td><td>' + esc(i.text) + '</td></tr>'; }).join('') : '<tr><td colspan="2">Nothing to fix</td></tr>';
    var grouped = {}; D.fixes.forEach(function (f) { grouped[f.text] = (grouped[f.text] || 0) + 1; });
    $('dcFixes').innerHTML = Object.keys(grouped).sort(function (a, b) { return grouped[b] - grouped[a]; }).map(function (t) { return '<tr><td>' + esc(t) + '</td><td class="num">' + grouped[t] + '×</td></tr>'; }).join('') || '<tr><td>None</td></tr>';
    $('dcDupes').innerHTML = D.dupes.length ? D.dupes.map(function (r) { return '<tr><td>' + where(r) + '</td><td class="num">' + fmt(r.calls) + '</td></tr>'; }).join('') : '<tr><td>None</td></tr>';
  }
  function importedRows() { return DATA.rows.filter(function (r) { return IMPORTED.some(function (s) { return r.src === s.source + ' › ' + s.name; }); }); }
  function importFiles(files) {
    if (!window.XLSX) return msg($('importMsg'), false, 'Excel library did not load (check the internet connection).');
    var list = Array.prototype.slice.call(files), done = 0, before = DATA.rows.length;
    list.forEach(function (file) {
      var rd = new FileReader();
      rd.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(rd.result), { type: 'array', cellDates: true });
          wb.SheetNames.forEach(function (n) { IMPORTED.push({ name: n, source: file.name, values: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) }); });
        } catch (e) { msg($('importMsg'), false, file.name + ': ' + e.message); }
        if (++done === list.length) {
          rebuild();
          msg($('importMsg'), true, list.map(function (f) { return f.name; }).join(', ') + ': ' + (DATA.rows.length - before) + ' new broadcasts added (after removing duplicates). They are in this browser only — press "Save to Google Sheet" to keep them.');
          $('btnSaveImport').hidden = !importedRows().length;
        }
      };
      rd.readAsArrayBuffer(file);
    });
  }
  function saveImported() {
    var rows = importedRows();
    if (!rows.length) return;
    var k = LS.get('vm.key') || ($('entryForm').key.value || '').trim();
    if (!k) return msg($('importMsg'), false, 'Enter the API key in "Add a broadcast" first.');
    post({ key: k, rows: rows.map(function (r) { return { date: dstr(r.date), district: r.district, upazila: r.upazilaLabel === 'All upazilas' ? 'All' : r.upazilaLabel, content: r.content === 'Not specified' ? '' : r.content, count: r.calls, success: r.success, failed: r.failed, project: r.project, media: r.media === 'Not recorded' ? '' : r.media }; }) })
      .then(function (j) { msg($('importMsg'), true, j.added + ' rows saved to the Google Sheet. Reloading…'); IMPORTED = []; $('btnSaveImport').hidden = true; return load(); })
      .catch(function (e) { msg($('importMsg'), false, e.message); });
  }

  // =====================================================================
  // Exports — same layout as the Excel: "Sheet1" log + "Sheet2" summary
  // =====================================================================
  function fileBase() { return 'VM_Dissemination_Report_' + (F.project || 'All_projects').replace(/\W+/g, '_') + '_' + dstr(new Date()); }
  function download(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function busy(btn, fn) {
    btn.disabled = true; var t = btn.innerHTML; btn.textContent = 'Preparing…';
    setTimeout(function () { Promise.resolve().then(fn).catch(function (e) { alert('Export failed: ' + e.message); console.error(e); }).then(function () { btn.disabled = false; btn.innerHTML = t; }); }, 30);
  }
  var cellTxt = function (v) { return v == null ? '' : typeof v === 'number' ? fmt(v) : String(v); };
  var nonEmpty = function (r) { return r.some(function (c) { return c !== '' && c != null; }); };
  var boldSummaryRow = function (r) { return /^(Total|Unique Beneficiary|By Project|Project$)/.test(String(r[0])) || r[1] === 'Total'; };

  function exportXlsx() {
    if (!window.XLSX) throw new Error('Excel library did not load');
    var which = $('exportWhich').value, wb = XLSX.utils.book_new();
    if (which !== 'summary') {
      var ws = XLSX.utils.aoa_to_sheet([SHEET1_HEAD].concat(sheet1Rows(VIEW)));
      ws['!cols'] = [16, 14, 34, 16, 18, 10, 10, 12, 18].map(function (w) { return { wch: w }; });
      ws['!autofilter'] = { ref: 'A1:I' + (VIEW.length + 1) };
      XLSX.utils.book_append_sheet(wb, ws, 'Dissemination log');
    }
    if (which !== 'detail') {
      var Lo = sheet2Layout(), ws2 = XLSX.utils.aoa_to_sheet(Lo.rows);
      if (Lo.benEnd > Lo.benStart) ws2[XLSX.utils.encode_cell({ r: Lo.benTotal, c: 2 })] = { t: 'n', f: 'SUM(C' + (Lo.benStart + 1) + ':C' + Lo.benEnd + ')', v: Lo.rows[Lo.benTotal][2] };
      ws2['!merges'] = [{ s: { r: Lo.benTitle, c: 0 }, e: { r: Lo.benTitle, c: 1 } }];
      ws2['!cols'] = [{ wch: 18 }, { wch: 34 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 12 }];
      XLSX.utils.book_append_sheet(wb, ws2, 'Summary');
    }
    XLSX.writeFile(wb, fileBase() + '.xlsx');
  }
  function exportPdf() {
    var which = $('exportWhich').value, visuals = $('withVisuals').checked;
    var doc = new jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    var W = doc.internal.pageSize.getWidth(), M = 12, y = 35, half = W / 2 - M - 2, ch = half * 560 / 900;
    doc.setFontSize(17); doc.setFont(undefined, 'bold'); doc.text('VM Dissemination Report', M, 15);
    doc.setFontSize(9); doc.setFont(undefined, 'normal'); doc.setTextColor(82);
    doc.text(doc.splitTextToSize(filterText(), W - 2 * M), M, 21);
    doc.text('Generated ' + dnice(new Date()) + (CFG.ORG ? ' · ' + CFG.ORG : ''), M, 30); doc.setTextColor(0);
    var tableOpts = { theme: 'grid', styles: { fontSize: 7.5, cellPadding: 1.4, lineColor: [200, 200, 200], lineWidth: 0.1, textColor: 20 },
      headStyles: { fillColor: [242, 242, 242], textColor: 0, fontStyle: 'bold' }, footStyles: { fillColor: [242, 242, 242], textColor: 0, fontStyle: 'bold' }, showFoot: 'lastPage', margin: { left: M, right: M } };
    if (which !== 'summary') {
      doc.setFontSize(11); doc.setFont(undefined, 'bold'); doc.text('Dissemination log', M, y); doc.setFont(undefined, 'normal');
      doc.autoTable(Object.assign({ startY: y + 2, head: [SHEET1_HEAD], body: sheet1Rows(VIEW).map(function (r) { return r.map(cellTxt); }), foot: [sheet1Total(VIEW).map(cellTxt)],
        columnStyles: { 4: { halign: 'right' }, 5: { halign: 'right' }, 6: { halign: 'right' } } }, tableOpts));
    }
    if (which !== 'detail') {
      var Lo = sheet2Layout();
      if (which !== 'summary') { doc.addPage(); y = 15; }
      doc.setFontSize(11); doc.setFont(undefined, 'bold'); doc.text('Summary', M, y); doc.setFont(undefined, 'normal');
      doc.autoTable(Object.assign({}, tableOpts, { startY: y + 2, tableWidth: 'wrap', body: Lo.rows.filter(nonEmpty).map(function (r) { return [0, 1, 2, 3, 4, 5].map(function (i) { return cellTxt(r[i]); }); }),
        columnStyles: { 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' }, 5: { halign: 'right' } },
        didParseCell: function (dd) { if (boldSummaryRow(dd.row.raw)) dd.cell.styles.fontStyle = 'bold'; } }));
    }
    if (visuals) {
      doc.addPage();
      doc.setFontSize(11); doc.setFont(undefined, 'bold'); doc.text('Project area map', M, 12); doc.setFont(undefined, 'normal');
      doc.addImage(drawMapCanvas(1654, 2339).toDataURL('image/jpeg', 0.9), 'JPEG', (W - 127) / 2, 16, 127, 180);
      doc.addPage();
      doc.setFont(undefined, 'bold'); doc.text('Monthly dissemination by VM content', M, 12); doc.text('Success / failed by project', W / 2 + 2, 12);
      doc.addImage(chartImage('chMonthly', 900, 560, 'image/jpeg'), 'JPEG', M, 15, half, ch);
      doc.addImage(chartImage('chProject', 900, 560, 'image/jpeg'), 'JPEG', W / 2 + 2, 15, half, ch);
      doc.text('Total calls by zila', M, 25 + ch); doc.text('Broadcasts by media', W / 2 + 2, 25 + ch); doc.setFont(undefined, 'normal');
      doc.addImage(chartImage('chZila', 900, 560, 'image/jpeg'), 'JPEG', M, 28 + ch, half, ch);
      doc.addImage(chartImage('chMedia', 900, 560, 'image/jpeg'), 'JPEG', W / 2 + 2, 28 + ch, half, ch);
    }
    var n = doc.getNumberOfPages();
    for (var i = 1; i <= n; i++) { doc.setPage(i); doc.setFontSize(8); doc.setTextColor(138); doc.text('Page ' + i + ' of ' + n, W - M, doc.internal.pageSize.getHeight() - 6, { align: 'right' }); }
    doc.save(fileBase() + '.pdf');
  }
  function b64ToBytes(dataUrl) { var b = atob(dataUrl.split(',')[1]), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
  function exportDocx() {
    var d = window.docx; if (!d) throw new Error('Word library did not load (check internet connection)');
    var which = $('exportWhich').value, visuals = $('withVisuals').checked;
    var P = function (text, run) { return new d.Paragraph({ children: [new d.TextRun(Object.assign({ text: text }, run || {}))] }); };
    var H = function (text) { return new d.Paragraph({ text: text, heading: d.HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 } }); };
    var img = function (url, w, h) { return new d.Paragraph({ children: [new d.ImageRun({ data: b64ToBytes(url), transformation: { width: w, height: h } })] }); };
    var border = { style: d.BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
    var table = function (rows, opts) {
      return new d.Table({ width: { size: opts.width || 100, type: d.WidthType.PERCENTAGE }, rows: rows.map(function (cells, ri) {
        var bold = (opts.header && ri === 0) || opts.boldRow(cells);
        return new d.TableRow({ tableHeader: !!(opts.header && ri === 0), children: cells.map(function (v) {
          return new d.TableCell({ borders: { top: border, bottom: border, left: border, right: border }, shading: bold ? { fill: 'F2F2F2', type: d.ShadingType.CLEAR, color: 'auto' } : undefined,
            children: [new d.Paragraph({ alignment: typeof v === 'number' ? d.AlignmentType.RIGHT : d.AlignmentType.LEFT, children: [new d.TextRun({ text: cellTxt(v), size: 16, bold: !!bold })] })] });
        }) });
      }) });
    };
    var children = [new d.Paragraph({ text: 'VM Dissemination Report', heading: d.HeadingLevel.TITLE }), P(filterText(), { size: 18, color: '52514E' }), P('Generated ' + dnice(new Date()) + (CFG.ORG ? ' · ' + CFG.ORG : ''), { size: 18, color: '52514E' })];
    if (which !== 'summary') children.push(H('Dissemination log'), table([SHEET1_HEAD].concat(sheet1Rows(VIEW)).concat([sheet1Total(VIEW)]), { header: true, boldRow: function (c) { return c[0] === 'Total'; } }));
    if (which !== 'detail') {
      var rows = sheet2Layout().rows.filter(nonEmpty).map(function (r) { return [0, 1, 2, 3, 4, 5].map(function (i) { return r[i] == null ? '' : r[i]; }); });
      children.push(H('Summary'), table(rows, { width: 80, boldRow: boldSummaryRow }));
    }
    if (visuals) {
      children.push(H('Project area map'), img(drawMapCanvas(1654, 2339).toDataURL('image/png'), 300, 424),
        H('Monthly dissemination by VM content'), img(chartImage('chMonthly', 1200, 520), 620, 269),
        H('Success / failed by project'), img(chartImage('chProject', 1200, 520), 620, 269),
        H('Total calls by zila'), img(chartImage('chZila', 1200, 620), 620, 320));
    }
    var doc = new d.Document({ creator: CFG.ORG || 'RIMES', title: 'VM Dissemination Report', sections: [{ properties: { page: { size: { orientation: d.PageOrientation.LANDSCAPE } } }, children: children }] });
    return d.Packer.toBlob(doc).then(function (blob) { download(blob, fileBase() + '.docx'); });
  }

  // =====================================================================
  // Data entry (API)
  // =====================================================================
  function post(body) {
    if (!API_URL) return Promise.reject(new Error('Set API_URL in config.js first (deploy the Apps Script).'));
    return fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json(); }).then(function (j) { if (!j.ok) throw new Error(j.error || 'Failed'); return j; });
  }
  function formObj(form) { var o = {}; new FormData(form).forEach(function (v, k) { o[k] = String(v).trim(); }); return o; }
  function msg(el, ok, text) { el.className = 'msg ' + (ok ? 'ok' : 'err'); el.textContent = text; }
  function wireEntry() {
    var ef = $('entryForm'), bf = $('benForm');
    ef.date.value = dstr(new Date()); ef.key.value = LS.get('vm.key') || '';
    ef.onsubmit = function (e) {
      e.preventDefault(); var o = formObj(ef); LS.set('vm.key', o.key);
      var d = matchDistrict(o.district), u = matchUpazila(o.upazila, d);
      post({ key: o.key, rows: [{ date: o.date, district: d || (u && u.district) || o.district, upazila: u ? u.name : o.upazila, content: normContent(o.content), count: Number(o.count),
        success: o.success === '' ? '' : Number(o.success), failed: o.failed === '' ? '' : Number(o.failed), project: normProject(o.project), media: normMedia(o.media) }] })
        .then(function () { msg($('entryMsg'), true, 'Added. Refreshing…'); ef.count.value = ''; ef.success.value = ''; ef.failed.value = ''; return load(); })
        .catch(function (err) { msg($('entryMsg'), false, err.message); });
    };
    bf.onsubmit = function (e) {
      e.preventDefault(); var o = formObj(bf), k = ef.key.value || LS.get('vm.key');
      if (!k) return msg($('benMsg'), false, 'Enter the API key in the form above.');
      var u = matchUpazila(o.upazila, null);
      post({ key: k, beneficiaries: [{ project: normProject(o.project), district: u ? u.district : '', upazila: u ? u.name : o.upazila, count: Number(o.count) }] })
        .then(function () { msg($('benMsg'), true, 'Saved. Refreshing…'); return load(); })
        .catch(function (err) { msg($('benMsg'), false, err.message); });
    };
    var u = API_URL || 'https://script.google.com/macros/s/XXXX/exec';
    $('apiExamples').textContent =
      '# Read everything (all tabs, JSON)\n' + u + '?action=data\n\n' +
      '# Add one row in the same order as the sheet:\n# date, district, upazila, content, total, success, failed, project, media\n' +
      u + '?action=add&key=YOUR_KEY&q=23-11-2026,Sylhet,Kanaighat,Lightning,1005,454,551,UNICEF,Kotha\n\n' +
      '# Short form: date, upazila, content, total, media, project\n' + u + '?action=add&key=YOUR_KEY&q=23-11-2026,Ulipur,Flood,2579,GP,Protistha\n\n' +
      '# Unique beneficiaries (Summary block)\n' + u + '?action=beneficiary&key=YOUR_KEY&project=Protistha&upazila=Ulipur&count=7303\n\n' +
      '# Many rows (POST, body as text/plain)\ncurl -L -H "Content-Type: text/plain" -d \'{"key":"YOUR_KEY","rows":["23-11-2026,Gaibandha,Fulchari,Flood,1522,900,622,GFFO,GP"]}\' "' + u + '"';
  }

  // =====================================================================
  // Wiring
  // =====================================================================
  function init() {
    if (CFG.TITLE) { $('appTitle').textContent = CFG.TITLE; document.title = CFG.TITLE; }
    if (CFG.ORG) $('orgBadge').textContent = CFG.ORG.slice(0, 5);
    var th = LS.get('vm.theme'); if (th) document.documentElement.dataset.theme = th;
    $('btnTheme').onclick = function () {
      var dark = document.documentElement.dataset.theme ? document.documentElement.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
      document.documentElement.dataset.theme = dark ? 'light' : 'dark'; LS.set('vm.theme', document.documentElement.dataset.theme); renderCharts(); renderKpis(); renderDetail();
    };
    ['fProject', 'fUpazila', 'fContent', 'fMedia'].forEach(function (id) { $(id).onchange = apply; });
    $('fDistrict').onchange = function () { fillUpazilas(); apply(); };
    $('fPreset').onchange = function () {
      var r = presetRange($('fPreset').value);
      if (r) { $('fFrom').value = dstr(r[0]); $('fTo').value = dstr(r[1]); } else if ($('fPreset').value === '') { $('fFrom').value = ''; $('fTo').value = ''; }
      apply();
    };
    ['fFrom', 'fTo'].forEach(function (id) { $(id).onchange = function () { $('fPreset').value = 'custom'; apply(); }; });
    $('btnReset').onclick = function () { ['fProject', 'fPreset', 'fFrom', 'fTo', 'fDistrict', 'fContent', 'fMedia'].forEach(function (id) { $(id).value = ''; }); fillUpazilas(); $('fUpazila').value = ''; apply(); };
    $('search').oninput = renderDetail;
    $('btnRefresh').onclick = load;
    document.querySelectorAll('.nav').forEach(function (t) {
      t.onclick = function () {
        document.querySelectorAll('.nav').forEach(function (x) { x.classList.toggle('on', x.dataset.tab === t.dataset.tab); });
        document.querySelectorAll('.panel').forEach(function (p) { p.classList.toggle('on', p.id === 'p-' + t.dataset.tab); });
        if (t.dataset.tab === 'map' && map) setTimeout(function () { map.invalidateSize(); map._fitKey = null; renderMap(); }, 30);
      };
    });
    $('btnPdf').onclick = function () { busy($('btnPdf'), exportPdf); };
    $('btnDocx').onclick = function () { busy($('btnDocx'), exportDocx); };
    $('btnXlsx').onclick = function () { busy($('btnXlsx'), exportXlsx); };
    $('btnMapPng').onclick = function () { drawMapCanvas(2480, 3508).toBlob(function (b) { download(b, 'Project_Area_Map_' + dstr(new Date()) + '.png'); }); };
    $('btnMapGeo').onclick = function () {
      var fc = { type: 'FeatureCollection', features: projectArea().list.map(function (a) {
        var rows = a.rows;
        return { type: 'Feature', geometry: a.f.geometry, properties: { upazila: a.f.properties.display, zila: a.f.properties.district, broadcasts: rows.length, total_calls: sum(rows, 'calls'), success: sum(rows, 'success'), failed: sum(rows, 'failed'), projects: uniq(rows.map(function (r) { return r.project; })).join(', '), content: uniq(rows.map(function (r) { return r.content; }).filter(Boolean)).join(', ') } };
      }) };
      download(new Blob([JSON.stringify(fc)], { type: 'application/geo+json' }), 'Project_Area_' + dstr(new Date()) + '.geojson');
    };
    var dz = $('dropzone'), fi = $('fileInput');
    dz.onclick = function () { fi.click(); };
    dz.onkeydown = function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } };
    fi.onchange = function () { if (fi.files.length) importFiles(fi.files); fi.value = ''; };
    dz.ondragover = function (e) { e.preventDefault(); dz.classList.add('over'); };
    dz.ondragleave = function () { dz.classList.remove('over'); };
    dz.ondrop = function (e) { e.preventDefault(); dz.classList.remove('over'); if (e.dataTransfer.files.length) importFiles(e.dataTransfer.files); };
    $('btnSaveImport').onclick = saveImported;
    wireEntry();
    load();
    if (API_URL && CFG.REFRESH_MINUTES > 0) setInterval(load, CFG.REFRESH_MINUTES * 60000);
  }
  window.VM = { parseSheets: parseSheets, normContent: normContent, matchDistrict: matchDistrict, matchUpazila: matchUpazila, get data() { return DATA; }, get view() { return VIEW; }, sheet2Layout: sheet2Layout, drawMapCanvas: drawMapCanvas };
  document.addEventListener('DOMContentLoaded', init);
})();
