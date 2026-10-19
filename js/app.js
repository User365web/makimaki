/*! makiconvertpro - CSV to Excel converter. Runs entirely in the browser; no dependencies. */
(function () {
  'use strict';

  var MAX_ROWS = 1048576;
  var MAX_COLS = 16384;
  var MAX_BYTES = 80 * 1024 * 1024;
  var NUM = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/;
  var DELIMS = { comma: ',', semicolon: ';', tab: '\t', pipe: '|' };
  var DELIM_NAMES = { ',': 'comma', ';': 'semicolon', '\t': 'tab', '|': 'pipe' };

  /* ---------- Decoding ---------- */
  function decode(buf) {
    var u8 = new Uint8Array(buf);
    if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8);
    if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8);
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(u8);
    } catch (e) {
      return new TextDecoder('windows-1252').decode(u8);
    }
  }

  /* ---------- CSV parsing ---------- */
  function detectDelimiter(text) {
    var sample = text.slice(0, 30000);
    var cands = [',', ';', '\t', '|'];
    var best = ',';
    var bestScore = -1;
    cands.forEach(function (d) {
      var counts = [], inQ = false, c = 0, lines = 0;
      for (var i = 0; i < sample.length && lines < 10; i++) {
        var ch = sample[i];
        if (ch === '"') inQ = !inQ;
        else if (!inQ) {
          if (ch === d) c++;
          else if (ch === '\n') { counts.push(c); c = 0; lines++; }
        }
      }
      if (!counts.length) counts.push(c);
      var first = counts[0];
      if (first === 0) return;
      var same = counts.filter(function (n) { return n === first; }).length / counts.length;
      var score = same * 1000 + first;
      if (score > bestScore) { bestScore = score; best = d; }
    });
    return best;
  }

  function parseCSV(text, delim) {
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    var rows = [], row = [], field = '', inQ = false, n = text.length;
    for (var i = 0; i < n; i++) {
      var ch = text[i];
      if (inQ) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false;
        } else field += ch;
      } else if (ch === '"' && field === '') {
        inQ = true;
      } else if (ch === delim) {
        row.push(field); field = '';
      } else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); field = '';
        rows.push(row); row = [];
      } else field += ch;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    while (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
    return rows;
  }

  /* ---------- XLSX building ---------- */
  function colName(i) {
    var s = '';
    i++;
    while (i > 0) {
      var m = (i - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function esc(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function cleanSheetName(name) {
    var n = String(name || '').replace(/[\[\]:*?\/\\]/g, ' ').replace(/^'+|'+$/g, '').trim();
    return (n || 'Sheet1').slice(0, 31);
  }

  function buildSheet(rows, cols, o) {
    var out = [], widths = [], c, r;
    for (c = 0; c < cols; c++) widths.push(8);
    for (r = 0; r < rows.length; r++) {
      var row = rows[r], cells = [], lim = Math.min(row.length, cols);
      var isHead = o.header && r === 0;
      for (c = 0; c < lim; c++) {
        var v = row[c];
        if (v === '') continue;
        if (r < 300) {
          var w = Math.min(60, v.length + 2);
          if (w > widths[c]) widths[c] = w;
        }
        var ref = colName(c) + (r + 1);
        var s = isHead ? ' s="1"' : '';
        var t = o.numbers && !isHead ? v.trim() : '';
        if (t && NUM.test(t) && t.replace(/[-.]/g, '').length <= 15) {
          cells.push('<c r="' + ref + '"' + s + '><v>' + t + '</v></c>');
        } else {
          cells.push('<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' +
            esc(v.length > 32767 ? v.slice(0, 32767) : v) + '</t></is></c>');
        }
      }
      if (cells.length) out.push('<row r="' + (r + 1) + '">' + cells.join('') + '</row>');
    }
    var colsXml = '<cols>' + widths.map(function (w, i) {
      return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
    }).join('') + '</cols>';
    var pane = o.header
      ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/>'
      : '';
    var filter = o.header ? '<autoFilter ref="A1:' + colName(cols - 1) + rows.length + '"/>' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView workbookViewId="0">' + pane + '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' + colsXml +
      '<sheetData>' + out.join('') + '</sheetData>' + filter + '</worksheet>';
  }

  var XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  var NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var STYLES = XML + '<styleSheet xmlns="' + NS + '">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE8EEFB"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  var CONTENT_TYPES = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
  var ROOT_RELS = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  var WB_RELS = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

  function workbookXml(sheetName, cols, rowCount, header) {
    var names = '';
    if (header) {
      names = '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'' +
        esc(sheetName).replace(/'/g, "''") + '\'!$A$1:$' + colName(cols - 1) + '$' + rowCount + '</definedName></definedNames>';
    }
    return XML + '<workbook xmlns="' + NS + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + esc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets>' + names + '</workbook>';
  }

  /* ---------- ZIP writer (deflate when supported, else stored) ---------- */
  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(buf) {
    var c = 0xffffffff;
    for (var i = 0; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  async function deflateRaw(data) {
    if (typeof CompressionStream === 'undefined') return null;
    try {
      var stream = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    } catch (e) {
      return null;
    }
  }

  async function makeZip(files) {
    var enc = new TextEncoder(), chunks = [], central = [], offset = 0;
    var DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;
    for (var i = 0; i < files.length; i++) {
      var name = enc.encode(files[i].name), raw = files[i].data;
      var crc = crc32(raw), packed = await deflateRaw(raw);
      var method = packed && packed.length < raw.length ? 8 : 0;
      var body = method === 8 ? packed : raw;
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(8, method, true); h.setUint16(10, 0, true); h.setUint16(12, DATE, true);
      h.setUint32(14, crc, true); h.setUint32(18, body.length, true); h.setUint32(22, raw.length, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      chunks.push(new Uint8Array(h.buffer), name, body);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true);
      c.setUint16(8, 0x0800, true); c.setUint16(10, method, true); c.setUint16(12, 0, true); c.setUint16(14, DATE, true);
      c.setUint32(16, crc, true); c.setUint32(20, body.length, true); c.setUint32(24, raw.length, true);
      c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + body.length;
    }
    var size = 0;
    central.forEach(function (p) { size += p.length; });
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, size, true); e.setUint32(16, offset, true);
    return new Blob(chunks.concat(central, [new Uint8Array(e.buffer)]), {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  /* ---------- Public conversion API ---------- */
  async function convertText(text, opts) {
    opts = opts || {};
    var o = {
      delim: opts.delim || 'auto',
      header: opts.header !== false,
      numbers: opts.numbers !== false,
      sheetName: cleanSheetName(opts.sheetName)
    };
    if (!text.trim()) throw new Error('This file is empty.');
    var delim = o.delim === 'auto' ? detectDelimiter(text) : (DELIMS[o.delim] || ',');
    var rows = parseCSV(text, delim), notes = [];
    if (!rows.length) throw new Error('No rows were found in this file.');
    if (rows.length > MAX_ROWS) {
      rows.length = MAX_ROWS;
      notes.push('Excel allows 1,048,576 rows, so extra rows were left out.');
    }
    var cols = 0;
    rows.forEach(function (r) { if (r.length > cols) cols = r.length; });
    if (cols > MAX_COLS) {
      cols = MAX_COLS;
      notes.push('Excel allows 16,384 columns, so extra columns were left out.');
    }
    var files = [
      { name: '[Content_Types].xml', data: new TextEncoder().encode(CONTENT_TYPES) },
      { name: '_rels/.rels', data: new TextEncoder().encode(ROOT_RELS) },
      { name: 'xl/workbook.xml', data: new TextEncoder().encode(workbookXml(o.sheetName, cols, rows.length, o.header)) },
      { name: 'xl/_rels/workbook.xml.rels', data: new TextEncoder().encode(WB_RELS) },
      { name: 'xl/styles.xml', data: new TextEncoder().encode(STYLES) },
      { name: 'xl/worksheets/sheet1.xml', data: new TextEncoder().encode(buildSheet(rows, cols, o)) }
    ];
    return {
      blob: await makeZip(files),
      rows: rows.length,
      cols: cols,
      delimiter: DELIM_NAMES[delim] || 'comma',
      notes: notes,
      preview: rows.slice(0, 6).map(function (r) { return r.slice(0, 8); })
    };
  }

  async function convertFile(file, opts) {
    if (file.size > MAX_BYTES) {
      throw new Error('This file is larger than 80 MB. Split it into smaller files and convert each part.');
    }
    return convertText(decode(await file.arrayBuffer()), opts);
  }

  var API = { parseCSV: parseCSV, detectDelimiter: detectDelimiter, convertText: convertText, convertFile: convertFile };
  if (typeof module !== 'undefined' && module.exports) { module.exports = API; return; }
  window.makiconvertpro = API;

  /* ---------- UI ---------- */
  var dz = document.getElementById('dropzone');
  if (!dz) return;
  var input = document.getElementById('file-input');
  var list = document.getElementById('results');
  var status = document.getElementById('status');
  var clearBtn = document.getElementById('clear-all');
  var urls = [];

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function fmtSize(b) {
    if (b < 1024) return b + ' B';
    if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
    return (b / 1048576).toFixed(1) + ' MB';
  }

  function options() {
    return {
      delim: document.getElementById('opt-delim').value,
      header: document.getElementById('opt-header').checked,
      numbers: document.getElementById('opt-numbers').checked,
      sheetName: document.getElementById('opt-sheet').value
    };
  }

  function say(msg) { status.textContent = msg; }

  function previewTable(rows) {
    var wrap = el('div', 'table-wrap'), table = el('table'), body = el('tbody');
    rows.forEach(function (r) {
      var tr = el('tr');
      r.forEach(function (v) { tr.appendChild(el('td', null, v.length > 40 ? v.slice(0, 40) + '\u2026' : v)); });
      body.appendChild(tr);
    });
    table.appendChild(body);
    wrap.appendChild(table);
    return wrap;
  }

  async function handle(file) {
    var li = el('li', 'result');
    li.setAttribute('data-state', 'busy');
    var main = el('div', 'result-main');
    main.appendChild(el('strong', 'result-name', file.name));
    var meta = el('span', 'result-meta', 'Converting\u2026');
    main.appendChild(meta);
    li.appendChild(main);
    list.appendChild(li);
    clearBtn.hidden = false;
    await new Promise(function (r) { setTimeout(r, 30); });
    try {
      var res = await convertFile(file, options());
      var url = URL.createObjectURL(res.blob);
      urls.push(url);
      var base = file.name.replace(/\.[^.]+$/, '') || 'converted';
      meta.textContent = res.rows.toLocaleString() + ' rows, ' + res.cols.toLocaleString() +
        ' columns, ' + res.delimiter + ' separated, ' + fmtSize(file.size);
      var a = el('a', 'btn btn-small', 'Download Excel file');
      a.href = url;
      a.download = base + '.xlsx';
      li.appendChild(a);
      res.notes.forEach(function (n) { li.appendChild(el('p', 'result-note', n)); });
      var d = el('details', 'result-preview');
      d.appendChild(el('summary', null, 'Preview first rows'));
      d.appendChild(previewTable(res.preview));
      li.appendChild(d);
      li.setAttribute('data-state', 'ok');
      say(file.name + ' is ready to download.');
    } catch (err) {
      meta.textContent = err && err.message ? err.message : 'This file could not be converted.';
      li.setAttribute('data-state', 'error');
      say('Could not convert ' + file.name + '.');
    }
  }

  async function handleAll(files) {
    var arr = Array.prototype.slice.call(files);
    for (var i = 0; i < arr.length; i++) await handle(arr[i]);
  }

  input.addEventListener('change', function () {
    if (input.files.length) handleAll(input.files);
    input.value = '';
  });
  ['dragenter', 'dragover'].forEach(function (t) {
    dz.addEventListener(t, function (e) { e.preventDefault(); dz.classList.add('is-over'); });
  });
  ['dragleave', 'drop'].forEach(function (t) {
    dz.addEventListener(t, function (e) { e.preventDefault(); dz.classList.remove('is-over'); });
  });
  dz.addEventListener('drop', function (e) {
    if (e.dataTransfer && e.dataTransfer.files.length) handleAll(e.dataTransfer.files);
  });
  clearBtn.addEventListener('click', function () {
    urls.forEach(function (u) { URL.revokeObjectURL(u); });
    urls = [];
    list.textContent = '';
    clearBtn.hidden = true;
    say('Cleared all converted files.');
  });
})();
