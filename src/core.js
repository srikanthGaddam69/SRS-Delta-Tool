/* SRS Delta core: pure comparison + workbook builder. Works in browser (window.SRSDelta) and Node (module.exports). */
(function (root) {
  'use strict';
  var PUNCT = /[\s.,;:!?'"`‘’“”\-–—•· ​()\[\]{}_]+/g;
  var LISTSEP = /[\n\r\t,;\/]+/;

  function cellText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (r) { return r.text; }).join('');
      if ('result' in v || 'formula' in v || 'sharedFormula' in v) return cellText(v.result);
      if (v.text !== undefined) return cellText(v.text);
      if (v.error) return '';
      return String(v);
    }
    return String(v);
  }
  function norm(s, ic) { s = String(s || '').replace(PUNCT, ''); return ic ? s.toLowerCase() : s; }
  function listKey(s, ic) {
    var set = {};
    String(s || '').split(LISTSEP).forEach(function (t) { t = norm(t, ic); if (t) set[t] = 1; });
    return Object.keys(set).sort().join('|');
  }
  function keyOf(s, mode, ic) { return mode === 'list' ? listKey(s, ic) : norm(s, ic); }
  function hnorm(h) { return String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

  /* Read an ExcelJS worksheet into {headers, rows} (text values). Header row = first row (of first 10) with >=2 filled cells. */
  function readSheet(ws) {
    var maxC = ws.columnCount, hdrRow = 1;
    for (var r = 1; r <= Math.min(10, ws.rowCount); r++) {
      var n = 0; for (var c = 1; c <= maxC; c++) if (cellText(ws.getRow(r).getCell(c).value).trim()) n++;
      if (n >= 2) { hdrRow = r; break; }
    }
    var headers = [];
    for (c = 1; c <= maxC; c++) headers.push(cellText(ws.getRow(hdrRow).getCell(c).value).trim());
    while (headers.length && !headers[headers.length - 1]) headers.pop();
    var rows = [];
    for (r = hdrRow + 1; r <= ws.rowCount; r++) {
      var row = ws.getRow(r), vals = [], any = false;
      for (c = 1; c <= headers.length; c++) { var t = cellText(row.getCell(c).value); vals.push(t); if (t.trim()) any = true; }
      if (any) rows.push(vals);
    }
    return { headers: headers, rows: rows, headerRow: hdrRow };
  }

  function guessIdCol(headers) {
    for (var i = 0; i < headers.length; i++) if (/req.*id|^id$|requirement\s*id/i.test(headers[i])) return i;
    return 0;
  }
  function isDescHeader(h) { return /^\s*srs/i.test(h) || /description|requirement text|^text$/i.test(h); }
  function guessMode(name) { return /area|stage|migration|subsystem|allocation/i.test(name) ? 'list' : 'text'; }
  function niceName(h) { return isDescHeader(h) ? 'Description' : String(h).trim(); }

  /* Auto-map columns of old -> new by header name, then by position for leftovers. */
  function autoMap(hOld, hNew, idOld, idNew) {
    var usedNew = {}; usedNew[idNew] = 1; var fields = [], pending = [];
    hOld.forEach(function (h, i) {
      if (i === idOld || !h) return;
      var j = -1;
      for (var k = 0; k < hNew.length; k++) if (!usedNew[k] && hnorm(hNew[k]) === hnorm(h)) { j = k; break; }
      if (j < 0 && isDescHeader(h)) for (k = 0; k < hNew.length; k++) if (!usedNew[k] && isDescHeader(hNew[k])) { j = k; break; }
      if (j >= 0) usedNew[j] = 1; else pending.push(fields.length);
      fields.push({ name: niceName(h), oldCol: i, newCol: j, mode: guessMode(h), include: true });
    });
    pending.forEach(function (fi) {
      var f = fields[fi];
      if (f.oldCol < hNew.length && !usedNew[f.oldCol] && hNew[f.oldCol]) { f.newCol = f.oldCol; usedNew[f.oldCol] = 1; }
      else f.include = false;
    });
    return fields;
  }

  function guessLabel(headers, fileName) {
    for (var i = 0; i < headers.length; i++) if (/^\s*srs[\s_\-]*\d/i.test(headers[i])) return headers[i].trim();
    return String(fileName || 'Version').replace(/\.[^.]+$/, '');
  }

  /* Main comparison. */
  function compare(o) {
    var fields = o.fields.filter(function (f) { return f.include && f.newCol >= 0 && f.oldCol >= 0; });
    var ic = o.ignoreCase !== false, warnings = [];
    function index(data, idc, tag) {
      var map = {}, order = [], dups = 0;
      data.rows.forEach(function (r) {
        var id = String(r[idc] || '').trim(); if (!id) return;
        if (map[id]) { dups++; return; } map[id] = r; order.push(id);
      });
      if (dups) warnings.push(dups + ' duplicate ID(s) in ' + tag + ' – first occurrence used.');
      return { map: map, order: order };
    }
    var A = index(o.old, o.idOld, o.labelOld), B = index(o.new, o.idNew, o.labelNew);
    var ids = B.order.concat(A.order.filter(function (i) { return !B.map[i]; }));
    var recs = [], st = { totalOld: A.order.length, totalNew: B.order.length, added: 0, deleted: 0, modified: 0, unchanged: 0,
      perField: fields.map(function () { return 0; }), transitions: fields.map(function () { return {}; }) };
    ids.forEach(function (id) {
      var a = A.map[id], b = B.map[id], vals = [], status, anyYes = false;
      fields.forEach(function (f, fi) {
        var va = a ? String(a[f.oldCol] || '').trim() : null, vb = b ? String(b[f.newCol] || '').trim() : null, flag;
        if (a && b) {
          flag = keyOf(va, f.mode, ic) !== keyOf(vb, f.mode, ic) ? 'Yes' : 'No';
          if (flag === 'Yes') {
            anyYes = true; st.perField[fi]++;
            if (f.name !== 'Description') { var t = (va || '(blank)') + '  →  ' + (vb || '(blank)'); t = t.replace(/\s*\n\s*/g, ', '); st.transitions[fi][t] = (st.transitions[fi][t] || 0) + 1; }
          }
        } else flag = 'N/A';
        vals.push([va, vb, flag]);
      });
      if (!a) { status = 'New'; st.added++; } else if (!b) { status = 'Deleted'; st.deleted++; }
      else if (anyYes) { status = 'Modified'; st.modified++; } else { status = 'Unchanged'; st.unchanged++; }
      recs.push({ id: id, status: status, vals: vals });
    });
    var safetyIdx = -1; fields.forEach(function (f, i) { if (/safety/i.test(f.name)) safetyIdx = i; });
    if (safetyIdx >= 0) {
      var isSR = function (v) { return /^\s*safety\s*related\s*$/i.test(v || '') || /^\s*(yes|y|sr)\s*$/i.test(v || ''); };
      var s = { modifiedSR: 0, newSR: 0, deletedSR: 0, totalOldSR: 0, totalNewSR: 0 };
      recs.forEach(function (r) {
        var v = r.vals[safetyIdx];
        if (isSR(v[0])) s.totalOldSR++; if (isSR(v[1])) s.totalNewSR++;
        if (r.status === 'Modified' && isSR(v[1])) s.modifiedSR++;
        if (r.status === 'New' && isSR(v[1])) s.newSR++;
        if (r.status === 'Deleted' && isSR(v[0])) s.deletedSR++;
      });
      st.safety = s; st.safetyIdx = safetyIdx;
    }
    return { fields: fields, records: recs, stats: st, warnings: warnings, labelOld: o.labelOld, labelNew: o.labelNew, ignoreCase: ic };
  }

  function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  /* Build the output workbook with ExcelJS. opts: {includeUnchanged, rev, revDesc, history:[[rev,date,desc,src]], date, srcNames} */
  function buildWorkbook(ExcelJS, res, opts) {
    var wb = new ExcelJS.Workbook(); wb.creator = 'SRS Delta Tool'; wb.created = new Date();
    var ws = wb.addWorksheet('Delta', { views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }] });
    var L0 = res.labelOld, L1 = res.labelNew, F = res.fields;
    var thin = { style: 'thin', color: { argb: 'FFA6A6A6' } }, border = { top: thin, left: thin, bottom: thin, right: thin };
    var fill = function (c) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + c } }; };
    var heads = [L0, L1], widths = [14, 14];
    F.forEach(function (f) {
      heads.push(L0 + ' ' + f.name, L1 + ' ' + f.name, f.name + ' change (Yes/No)');
      if (f.name === 'Description') widths.push(60, 60, 14); else widths.push(22, 22, 14);
    });
    heads.push('Status'); widths.push(12);
    var hr = ws.addRow(heads); hr.height = 45;
    hr.eachCell(function (c, i) {
      c.font = { bold: true }; c.border = border; c.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      c.fill = fill(i > 2 && (i - 2) % 3 === 0 ? 'D9D9D9' : 'DDEBF7');
    });
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var statusCol = heads.length;
    res.records.forEach(function (r) {
      if (r.status === 'Unchanged' && !opts.includeUnchanged) return;
      var v = [r.status === 'New' ? 'New' : r.id, r.status === 'Deleted' ? 'Deleted' : r.id];
      r.vals.forEach(function (x) { v.push(x[0] || null, x[1] || null, x[2]); }); v.push(r.status);
      var row = ws.addRow(v), rf = r.status === 'New' ? 'E2EFDA' : r.status === 'Deleted' ? 'FCE4D6' : null;
      for (var c = 1; c <= statusCol; c++) {
        var cell = row.getCell(c), isFlag = c > 2 && c < statusCol && (c - 2) % 3 === 0;
        cell.border = border;
        cell.alignment = { wrapText: !isFlag && c !== statusCol, vertical: 'top', horizontal: (isFlag || c === statusCol) ? 'center' : 'left' };
        if (rf) cell.fill = fill(rf);
        else if (isFlag && cell.value === 'Yes') { cell.fill = fill('FFF2CC'); cell.font = { bold: true }; }
      }
    });
    var last = ws.rowCount;
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: statusCol } };

    /* Summary with live formulas + cached results */
    var sm = wb.addWorksheet('Summary'), st = res.stats;
    var R = function (col) { return "Delta!$" + col + "$2:$" + col + "$" + Math.max(last, 2); };
    var flagCol = function (fi) { return colLetter(5 + fi * 3); }, valCol = function (fi, side) { return colLetter(3 + fi * 3 + side); };
    var S = colLetter(statusCol);
    sm.getColumn(1).width = 62; sm.getColumn(2).width = 12;
    sm.getCell('A1').value = 'Delta Summary – ' + L0 + ' vs ' + L1; sm.getCell('A1').font = { bold: true, size: 14 };
    sm.getCell('A2').value = 'Counts are live formulas on the Delta sheet. Whitespace, line-break, punctuation' + (res.ignoreCase ? ' and letter-case' : '') + ' differences are not counted as changes.';
    sm.getCell('A2').font = { italic: true, color: { argb: 'FF595959' } };
    var r = 4;
    function section(title, items) {
      ['A', 'B'].forEach(function (c, i) { var x = sm.getCell(c + r); x.value = i ? 'Count' : title; x.font = { bold: true, color: { argb: 'FFFFFFFF' } }; x.fill = fill('1F4E78'); x.border = border; });
      r++;
      items.forEach(function (it) {
        var a = sm.getCell('A' + r), b = sm.getCell('B' + r);
        a.value = it[0]; b.value = it[1] ? { formula: it[1], result: it[2] } : it[2];
        a.border = b.border = border; b.alignment = { horizontal: 'center' }; r++;
      });
      r++;
    }
    var unchangedFormula = opts.includeUnchanged ? 'COUNTIF(' + R(S) + ',"Unchanged")' : null;
    section('1. Overall requirement count', [
      ['Total requirements in ' + L0, null, st.totalOld],
      ['Total requirements in ' + L1, null, st.totalNew],
      ['New requirements (in ' + L1 + ' only)', 'COUNTIF(' + R(S) + ',"New")', st.added],
      ['Deleted requirements (in ' + L0 + ' only)', 'COUNTIF(' + R(S) + ',"Deleted")', st.deleted],
      ['Modified requirements (at least one change)', 'COUNTIF(' + R(S) + ',"Modified")', st.modified],
      ['Unchanged requirements', unchangedFormula, st.unchanged]
    ]);
    section('2. Modified requirements by change factor', F.map(function (f, fi) {
      return [f.name + ' changed', 'COUNTIF(' + R(flagCol(fi)) + ',"Yes")', st.perField[fi]];
    }));
    if (st.safety) {
      var si = st.safetyIdx, sOld = valCol(si, 0), sNew = valCol(si, 1), s = st.safety;
      var tr = st.transitions[si];
      var items = Object.keys(tr).sort().map(function (k) { return [k, null, tr[k]]; });
      items.push(['Modified requirements that are Safety Related in ' + L1, 'COUNTIFS(' + R(S) + ',"Modified",' + R(sNew) + ',"Safety Related")', s.modifiedSR]);
      items.push(['New requirements that are Safety Related', 'COUNTIFS(' + R(S) + ',"New",' + R(sNew) + ',"Safety Related")', s.newSR]);
      items.push(['Deleted requirements that were Safety Related', 'COUNTIFS(' + R(S) + ',"Deleted",' + R(sOld) + ',"Safety Related")', s.deletedSR]);
      items.push(['Total Safety Related in ' + L0, null, s.totalOldSR]);
      items.push(['Total Safety Related in ' + L1, null, s.totalNewSR]);
      section('3. Safety Related breakdown', items);
    }
    var n = st.safety ? 4 : 3;
    F.forEach(function (f, fi) {
      if (f.name === 'Description' || fi === st.safetyIdx) return;
      var tr = st.transitions[fi], keys = Object.keys(tr);
      if (!keys.length) return;
      keys.sort(function (a, b) { return tr[b] - tr[a]; });
      section((n++) + '. ' + f.name + ' – value changes (' + L0 + ' → ' + L1 + ')', keys.slice(0, 25).map(function (k) { return [k, null, tr[k]]; }));
    });
    sm.getCell('A' + r).value = 'Note: one requirement can change in several factors, so section 2 can add up to more than the modified count.';
    sm.getCell('A' + r).font = { italic: true, color: { argb: 'FF595959' } };
    sm.getColumn(1).alignment = { wrapText: true, vertical: 'top' };

    /* Revision history */
    var rh = wb.addWorksheet('Revision History');
    var hh = rh.addRow(['Rev', 'Date', 'Description of change', 'Source files']);
    hh.eachCell(function (c) { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = fill('1F4E78'); c.border = border; });
    (opts.history || []).concat([[opts.rev, opts.date, opts.revDesc, opts.srcNames]]).forEach(function (h) {
      var row = rh.addRow(h); row.eachCell({ includeEmpty: true }, function (c) { c.border = border; c.alignment = { wrapText: true, vertical: 'top' }; });
    });
    [6, 12, 90, 40].forEach(function (w, i) { rh.getColumn(i + 1).width = w; });
    return wb;
  }

  /* Read Revision History rows from a previous delta workbook (ExcelJS workbook). */
  function readHistory(wb) {
    var ws = wb.getWorksheet('Revision History'); if (!ws) return [];
    var out = [];
    ws.eachRow(function (row, i) {
      if (i === 1) return;
      var v = [1, 2, 3, 4].map(function (c) { return cellText(row.getCell(c).value); });
      if (v.join('').trim()) out.push(v);
    });
    return out;
  }
  function nextRev(history) {
    var max = 0; history.forEach(function (h) { var n = parseInt(String(h[0]).replace(/\D/g, ''), 10); if (n > max) max = n; });
    return String(max + 1).padStart(2, '0');
  }

  var api = { cellText: cellText, readSheet: readSheet, guessIdCol: guessIdCol, autoMap: autoMap, guessLabel: guessLabel,
    compare: compare, buildWorkbook: buildWorkbook, readHistory: readHistory, nextRev: nextRev, colLetter: colLetter };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.SRSDelta = api;
})(this);
