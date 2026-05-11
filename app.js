/* global XLSX */

const els = {
  fileInput: document.getElementById('fileInput'),
  fileMeta: document.getElementById('fileMeta'),
  status: document.getElementById('status'),
  latSelect: document.getElementById('latSelect'),
  lngSelect: document.getElementById('lngSelect'),
  nameSelect: document.getElementById('nameSelect'),
  descSelect: document.getElementById('descSelect'),
  folderName: document.getElementById('folderName'),
  sheetSelect: document.getElementById('sheetSelect'),
  headerRow: document.getElementById('headerRow'),
  formatSelect: document.getElementById('formatSelect'),
  outName: document.getElementById('outName'),
  downloadBtn: document.getElementById('downloadBtn'),
  previewBtn: document.getElementById('previewBtn'),
  preview: document.getElementById('preview'),
  howToBtn: document.getElementById('howToBtn'),
  howToModal: document.getElementById('howToModal'),
  howToClose: document.getElementById('howToClose'),
  howToOk: document.getElementById('howToOk'),
};

/** @type {{kind: 'excel', workbook: any, rows: any[], headers: string[], sheetName: string} | {kind: 'text', rows: any[], headers: string[], sourceName: string} | null} */
let state = null;

function setStatus(message, kind = 'ok') {
  els.status.textContent = message || '';
  els.status.className = 'status ' + (message ? kind : '');
}

function enableMapping(enabled) {
  for (const key of [
    'latSelect',
    'lngSelect',
    'nameSelect',
    'descSelect',
    'folderName',
    'sheetSelect',
    'headerRow',
    'formatSelect',
    'outName',
    'downloadBtn',
    'previewBtn',
  ]) {
    els[key].disabled = !enabled;
  }
}

function xmlEscape(v) {
  return String(v ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function toNumberLoose(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const s = String(v).trim();
  if (!s) return null;
  const n = Number(s.replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function isValidLatLng(lat, lng) {
  return Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
}

function isValidLat(lat) {
  return Number.isFinite(lat) && Math.abs(lat) <= 90;
}

function isValidLng(lng) {
  return Number.isFinite(lng) && Math.abs(lng) <= 180;
}

function normalizeTextLine(line) {
  return String(line ?? '').replace(/\uFEFF/g, '').trim();
}

function splitFields(line) {
  // Try common separators: comma, tab, semicolon, pipe. Fall back to whitespace.
  const hasStrongSep = /[,\t;|]/.test(line);
  if (hasStrongSep) return line.split(/[,\t;|]+/).map((s) => s.trim()).filter(Boolean);
  return line.split(/\s+/).map((s) => s.trim()).filter(Boolean);
}

function extractCandidateNumbers(line) {
  /** @type {Array<{value: number, index: number, raw: string}>} */
  const nums = [];
  const re = /[-+]?\d{1,3}(?:[.,]\d+)?/g;
  let m;
  while ((m = re.exec(line)) !== null) {
    const raw = m[0];
    const value = toNumberLoose(raw);
    if (value === null) continue;
    nums.push({ value, index: m.index, raw });
  }
  return nums;
}

function findLatLngInLine(line) {
  const nums = extractCandidateNumbers(line);
  if (nums.length < 2) return null;

  // Choose best pair (lat,lng) or (lng,lat) by validity and "lat-likeness"
  /** @type {{lat:number,lng:number, i:number, j:number, score:number} | null} */
  let best = null;
  for (let a = 0; a < nums.length; a++) {
    for (let b = a + 1; b < nums.length; b++) {
      const x = nums[a].value;
      const y = nums[b].value;

      // Option 1: x=lat, y=lng
      if (isValidLat(x) && isValidLng(y)) {
        const score =
          (Math.abs(x) <= 90 ? 2 : 0) +
          (Math.abs(y) <= 180 ? 1 : 0) +
          (Math.abs(y) > 90 ? 1 : 0); // lng often outside lat range
        if (!best || score > best.score) best = { lat: x, lng: y, i: a, j: b, score };
      }

      // Option 2: x=lng, y=lat
      if (isValidLng(x) && isValidLat(y)) {
        const score =
          (Math.abs(y) <= 90 ? 2 : 0) +
          (Math.abs(x) <= 180 ? 1 : 0) +
          (Math.abs(x) > 90 ? 1 : 0);
        if (!best || score > best.score) best = { lat: y, lng: x, i: b, j: a, score };
      }
    }
  }

  if (!best) return null;
  return best;
}

function parseTextToRows(text, sourceName) {
  const lines = String(text ?? '')
    .split(/\r?\n/)
    .map(normalizeTextLine)
    .filter((l) => l && !l.startsWith('#') && !l.startsWith('//'));

  const headers = ['Latitude', 'Longitude', 'Name', 'Description'];
  /** @type {Array<Record<string, any>>} */
  const rows = [];

  for (const line of lines) {
    // If line looks like a header (contains "lat" but no usable lat/lng), skip it.
    const lower = line.toLowerCase();
    const found = findLatLngInLine(line);
    if (!found) {
      if (lower.includes('lat') && (lower.includes('lon') || lower.includes('lng') || lower.includes('long'))) continue;
      continue;
    }

    // Prefer structured parsing when possible
    const fields = splitFields(line);
    let name = '';
    let desc = '';

    // If there are obvious non-numeric fields, use them for name/desc
    const nonNumeric = fields.filter((f) => toNumberLoose(f) === null);
    if (nonNumeric.length) {
      name = nonNumeric[0] || '';
      desc = nonNumeric.slice(1).join(' ') || '';
    } else {
      // Fallback: remove the two matched numbers from the raw line and treat remaining text as name/desc
      const nums = extractCandidateNumbers(line);
      const parts = [];
      for (let k = 0; k < nums.length; k++) {
        if (k === found.i || k === found.j) continue;
        parts.push(nums[k].raw);
      }
      const leftover = line
        .replace(nums[found.i]?.raw ?? '', ' ')
        .replace(nums[found.j]?.raw ?? '', ' ')
        .replace(/\s+/g, ' ')
        .trim();
      name = leftover || `${sourceName || 'Point'} ${rows.length + 1}`;
    }

    rows.push({
      Latitude: found.lat,
      Longitude: found.lng,
      Name: name,
      Description: desc,
    });
  }

  return { headers, rows };
}

function guessHeader(headers, kinds) {
  const norm = headers.map((h) => ({ h, n: String(h).toLowerCase().replace(/\s|_|-|\./g, '') }));
  for (const kind of kinds) {
    for (const item of norm) {
      if (item.n === kind) return item.h;
      if (item.n.includes(kind)) return item.h;
    }
  }
  return '';
}

function setSelectOptions(selectEl, headers, includeNone) {
  selectEl.innerHTML = '';
  if (includeNone) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '(none)';
    selectEl.appendChild(opt);
  }
  for (const h of headers) {
    const opt = document.createElement('option');
    opt.value = h;
    opt.textContent = h;
    selectEl.appendChild(opt);
  }
}

function getHeadersFromSheet(workbook, sheetName, headerRow1Based) {
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const r = Math.max(0, (headerRow1Based || 1) - 1);
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    let h = cell ? String(cell.v).trim() : '';
    if (!h) h = `Column ${XLSX.utils.encode_col(c)}`;
    headers.push(h);
  }
  return headers;
}

function readRows(workbook, sheetName, headerRow1Based) {
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const headerRow = Math.max(0, (headerRow1Based || 1) - 1);

  const headers = getHeadersFromSheet(workbook, sheetName, headerRow1Based);
  const dataStartRow = headerRow + 1;

  const rows = [];
  for (let r = dataStartRow; r <= range.e.r; r++) {
    const rowObj = {};
    let anyValue = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const key = headers[c - range.s.c];
      const v = cell ? cell.v : '';
      if (v !== '' && v !== null && v !== undefined) anyValue = true;
      rowObj[key] = v;
    }
    if (anyValue) rows.push(rowObj);
  }

  return { headers, rows };
}

function rebuildFromState() {
  if (!state) return;

  let headers = [];
  let rows = [];

  if (state.kind === 'excel') {
    const sheetName = els.sheetSelect.value;
    const headerRow = Math.max(1, Number(els.headerRow.value || 1));
    const r = readRows(state.workbook, sheetName, headerRow);
    headers = r.headers;
    rows = r.rows;
    state = { ...state, sheetName, headers, rows };
    setStatus(`Loaded ${rows.length} row(s) from sheet "${sheetName}".`, 'ok');
  } else {
    headers = state.headers;
    rows = state.rows;
    setStatus(`Loaded ${rows.length} row(s) from "${state.sourceName}".`, 'ok');
  }

  setSelectOptions(els.latSelect, headers, false);
  setSelectOptions(els.lngSelect, headers, false);
  setSelectOptions(els.nameSelect, headers, true);
  setSelectOptions(els.descSelect, headers, true);

  const guessLat = guessHeader(headers, ['lat', 'latitude']);
  const guessLng = guessHeader(headers, ['lng', 'lon', 'long', 'longitude']);
  const guessName = guessHeader(headers, ['name', 'title', 'label']);
  const guessDesc = guessHeader(headers, ['desc', 'description', 'note', 'notes', 'comment', 'remarks', 'remark']);

  if (guessLat) els.latSelect.value = guessLat;
  if (guessLng) els.lngSelect.value = guessLng;
  if (guessName) els.nameSelect.value = guessName;
  if (guessDesc) els.descSelect.value = guessDesc;

  els.preview.hidden = true;
  els.preview.innerHTML = '';
}

function validateMapping() {
  if (!state) return { ok: false, message: 'Upload a file first.' };
  const latKey = els.latSelect.value;
  const lngKey = els.lngSelect.value;
  if (!latKey || !lngKey) return { ok: false, message: 'Select both Latitude and Longitude columns.' };
  return { ok: true, message: '' };
}

function buildKml() {
  const v = validateMapping();
  if (!v.ok) throw new Error(v.message);
  if (!state) throw new Error('No data loaded.');

  const latKey = els.latSelect.value;
  const lngKey = els.lngSelect.value;
  const nameKey = els.nameSelect.value;
  const descKey = els.descSelect.value;
  const folderName = els.folderName.value.trim();

  const placemarks = [];
  let kept = 0;
  let skipped = 0;

  for (let i = 0; i < state.rows.length; i++) {
    const row = state.rows[i];
    const lat = toNumberLoose(row[latKey]);
    const lng = toNumberLoose(row[lngKey]);
    if (!isValidLatLng(lat, lng)) {
      skipped++;
      continue;
    }

    const name = nameKey ? row[nameKey] : '';
    const desc = descKey ? row[descKey] : '';

    const safeName = xmlEscape(name || `Point ${kept + 1}`);
    const safeDesc = xmlEscape(desc || '');

    placemarks.push(
      [
        '    <Placemark>',
        `      <name>${safeName}</name>`,
        safeDesc ? `      <description>${safeDesc}</description>` : '      <description/>',
        '      <Point>',
        `        <coordinates>${lng},${lat},0</coordinates>`,
        '      </Point>',
        '    </Placemark>',
      ].join('\n')
    );
    kept++;
  }

  const docName = xmlEscape(folderName || 'Points');
  const body =
    folderName.trim() !== ''
      ? [
          `  <Document>`,
          `    <name>${docName}</name>`,
          `    <Folder>`,
          `      <name>${docName}</name>`,
          placemarks.join('\n'),
          `    </Folder>`,
          `  </Document>`,
        ].join('\n')
      : [`  <Document>`, `    <name>${docName}</name>`, placemarks.join('\n'), `  </Document>`].join('\n');

  const kml =
    [
      '<?xml version="1.0" encoding="UTF-8"?>',
      '<kml xmlns="http://www.opengis.net/kml/2.2">',
      body,
      '</kml>',
      '',
    ].join('\n');

  return { kml, kept, skipped };
}

function buildGpx() {
  const v = validateMapping();
  if (!v.ok) throw new Error(v.message);
  if (!state) throw new Error('No data loaded.');

  const latKey = els.latSelect.value;
  const lngKey = els.lngSelect.value;
  const nameKey = els.nameSelect.value;
  const descKey = els.descSelect.value;

  let kept = 0;
  let skipped = 0;
  const wpts = [];

  for (const row of state.rows) {
    const lat = toNumberLoose(row[latKey]);
    const lng = toNumberLoose(row[lngKey]);
    if (!isValidLatLng(lat, lng)) {
      skipped++;
      continue;
    }

    const name = nameKey ? row[nameKey] : '';
    const desc = descKey ? row[descKey] : '';

    const safeName = xmlEscape(name || `Point ${kept + 1}`);
    const safeDesc = xmlEscape(desc || '');

    wpts.push(
      [
        `  <wpt lat="${lat}" lon="${lng}">`,
        `    <name>${safeName}</name>`,
        safeDesc ? `    <desc>${safeDesc}</desc>` : '    <desc/>',
        '  </wpt>',
      ].join('\n')
    );
    kept++;
  }

  const gpx =
    [
      '<?xml version="1.0" encoding="UTF-8"?>',
      '<gpx version="1.1" creator="kmlconvert" xmlns="http://www.topografix.com/GPX/1/1">',
      wpts.join('\n'),
      '</gpx>',
      '',
    ].join('\n');

  return { gpx, kept, skipped };
}

function downloadText(filename, text, mimeType) {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function ensureExtension(name, ext) {
  const s = String(name || '').trim();
  const base = s || `points.${ext}`;
  const safe = base.replace(/[<>:\"/\\\\|?*\\x00-\\x1F]/g, '_').replace(/\\s+/g, ' ');
  const lower = safe.toLowerCase();
  if (lower.endsWith(`.${ext}`)) return safe;
  if (lower.match(/\.[a-z0-9]{1,6}$/)) return safe;
  return `${safe}.${ext}`;
}

function stripKnownOutputExt(name) {
  const s = String(name || '').trim();
  if (!s) return '';
  return s.replace(/\.(kml|gpx)$/i, '');
}

function currentDownloadConfig() {
  const fmt = els.formatSelect?.value === 'gpx' ? 'gpx' : 'kml';
  if (fmt === 'gpx') {
    return { fmt, ext: 'gpx', mime: 'application/gpx+xml;charset=utf-8' };
  }
  return { fmt, ext: 'kml', mime: 'application/vnd.google-earth.kml+xml;charset=utf-8' };
}

function sanitizeFilename(name) {
  // Keep the input as a base name (no .kml/.gpx shown).
  const s = stripKnownOutputExt(name);
  return String(s || '')
    .trim()
    .replace(/[<>:\"/\\\\|?*\\x00-\\x1F]/g, '_')
    .replace(/\\s+/g, ' ');
}

function syncDownloadUi() {
  // Don't auto-append extensions in the input; we add it only on download.
  els.outName.value = sanitizeFilename(els.outName.value);
}

function setHowToOpen(open) {
  if (!els.howToModal) return;
  els.howToModal.hidden = !open;
  els.howToModal.setAttribute('aria-hidden', open ? 'false' : 'true');
  if (open) {
    els.howToClose?.focus?.();
  } else {
    els.howToBtn?.focus?.();
  }
}

function refreshButtons() {
  const ok = validateMapping().ok;
  els.downloadBtn.disabled = !ok;
  els.previewBtn.disabled = !ok;
}

async function onFile(file) {
  setStatus('', 'ok');
  els.fileMeta.textContent = '';
  els.preview.hidden = true;
  els.preview.innerHTML = '';
  enableMapping(false);

  if (!file) return;

  els.fileMeta.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;

  try {
    const ext = (file.name.split('.').pop() || '').toLowerCase();

    if (ext === 'txt' || ext === 'csv') {
      const text = await file.text();
      const parsed = parseTextToRows(text, file.name);
      state = { kind: 'text', rows: parsed.rows, headers: parsed.headers, sourceName: file.name };

      // Disable Excel-only controls
      els.sheetSelect.innerHTML = '';
      els.headerRow.value = '1';
      enableMapping(true);
      els.sheetSelect.disabled = true;
      els.headerRow.disabled = true;

      rebuildFromState();
      refreshButtons();
      syncDownloadUi();
      return;
    }

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetNames = wb.SheetNames || [];
    if (!sheetNames.length) throw new Error('No sheets found in workbook.');

    els.sheetSelect.innerHTML = '';
    for (const s of sheetNames) {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      els.sheetSelect.appendChild(opt);
    }

    state = { kind: 'excel', workbook: wb, rows: [], headers: [], sheetName: sheetNames[0] };
    els.sheetSelect.value = sheetNames[0];
    els.headerRow.value = '1';
    enableMapping(true);
    els.sheetSelect.disabled = false;
    els.headerRow.disabled = false;

    rebuildFromState();
    refreshButtons();
    syncDownloadUi();
  } catch (err) {
    state = null;
    enableMapping(false);
    setStatus(err?.message || 'Failed to read file.', 'bad');
  }
}

els.fileInput.addEventListener('change', (e) => {
  const file = e.target.files && e.target.files[0];
  onFile(file);
});

els.sheetSelect.addEventListener('change', () => {
  if (!state) return;
  rebuildFromState();
  refreshButtons();
});

els.headerRow.addEventListener('change', () => {
  if (!state) return;
  rebuildFromState();
  refreshButtons();
});

for (const el of [els.latSelect, els.lngSelect, els.nameSelect, els.descSelect]) {
  el.addEventListener('change', refreshButtons);
}

els.outName.addEventListener('change', () => {
  els.outName.value = sanitizeFilename(els.outName.value);
});

els.formatSelect.addEventListener('change', () => {
  syncDownloadUi();
});

els.howToBtn.addEventListener('click', () => setHowToOpen(true));
els.howToClose.addEventListener('click', () => setHowToOpen(false));
els.howToOk.addEventListener('click', () => setHowToOpen(false));

els.howToModal.addEventListener('click', (e) => {
  if (e.target === els.howToModal) setHowToOpen(false);
});

window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && els.howToModal && !els.howToModal.hidden) {
    setHowToOpen(false);
  }
});

els.downloadBtn.addEventListener('click', () => {
  try {
    const cfg = currentDownloadConfig();
    const filename = ensureExtension(sanitizeFilename(els.outName.value), cfg.ext);

    if (cfg.fmt === 'gpx') {
      const { gpx, kept, skipped } = buildGpx();
      downloadText(filename, gpx, cfg.mime);
      setStatus(`Generated ${kept} waypoint(s). Skipped ${skipped} row(s) with invalid lat/lng.`, kept ? 'ok' : 'bad');
    } else {
      const { kml, kept, skipped } = buildKml();
      downloadText(filename, kml, cfg.mime);
      setStatus(`Generated ${kept} placemark(s). Skipped ${skipped} row(s) with invalid lat/lng.`, kept ? 'ok' : 'bad');
    }
  } catch (err) {
    setStatus(err?.message || 'Failed to generate file.', 'bad');
  }
});

els.previewBtn.addEventListener('click', () => {
  try {
    const v = validateMapping();
    if (!v.ok) throw new Error(v.message);
    if (!state) throw new Error('No data loaded.');

    const latKey = els.latSelect.value;
    const lngKey = els.lngSelect.value;
    const nameKey = els.nameSelect.value;
    const descKey = els.descSelect.value;

    /** @type {Array<{name: string, lat: number, lng: number, desc: string}>} */
    const out = [];
    for (const row of state.rows) {
      if (out.length >= 5) break;
      const lat = toNumberLoose(row[latKey]);
      const lng = toNumberLoose(row[lngKey]);
      if (!isValidLatLng(lat, lng)) continue;
      out.push({
        name: String((nameKey ? row[nameKey] : '') ?? '').trim(),
        lat,
        lng,
        desc: String((descKey ? row[descKey] : '') ?? '').trim(),
      });
    }

    els.preview.innerHTML = '';
    if (!out.length) {
      const div = document.createElement('div');
      div.className = 'previewEmpty';
      div.textContent = 'No valid points found (check your mapping).';
      els.preview.appendChild(div);
      els.preview.hidden = false;
      return;
    }

    const table = document.createElement('table');
    table.className = 'previewTable';
    const thead = document.createElement('thead');
    thead.innerHTML = `
      <tr>
        <th>#</th>
        <th>Name</th>
        <th>Latitude</th>
        <th>Longitude</th>
        <th>Description</th>
      </tr>
    `;
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    out.forEach((p, idx) => {
      const tr = document.createElement('tr');
      const name = p.name || `Point ${idx + 1}`;
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${xmlEscape(name)}</td>
        <td>${p.lat}</td>
        <td>${p.lng}</td>
        <td>${xmlEscape(p.desc || '')}</td>
      `;
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    els.preview.appendChild(table);
    els.preview.hidden = false;
  } catch (err) {
    setStatus(err?.message || 'Failed to preview points.', 'bad');
  }
});

setStatus('Upload a file to begin.', 'ok');
