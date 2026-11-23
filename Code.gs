/**
 * FILAMENTARIUM – Google Sheets háttér (Apps Script)
 * ---------------------------------------------------
 * Ez a szkript a "3D Munkavezetés" táblázatod mögött fut, és egy privát
 * webcímen keresztül engedi, hogy a Filamentarium app olvasson és új sorokat
 * írjon. A táblázat elrendezéséhez NEM nyúl: csak sorokat fűz a meglévő
 * táblák végére, a meglévő képletmintáidat másolva.
 *
 * EGYETLEN dolog, amit át kell írnod: a JELSZÓ lent (SECRET).
 * Ugyanezt kell majd beírni az app Beállítások oldalán.
 */

const CONFIG = {
  SECRET: 'IRD-AT-EGY-SAJAT-JELSZORA',   // <-- ezt írd át!

  // Munkalapok nevei (pont úgy, ahogy a fülön szerepelnek)
  LOG:   '3D Gyártás',
  PROJ:  'Projektek',
  STOCK: 'Filamentkészletek',
  CAL:   'Kalibrációk',

  // Új filamentkövetőnél ettől a naplósortól számolja a felhasználást
  // (a 2026-os követőid mind a 101. sortól számolnak).
  STOCK_START_ROW: 101,
  // A tartományok nyitottak lefelé (pl. H101:H), mint a táblázatod többi képlete,
  // így sosem fogynak ki a sorokból.
  STOCK_END_ROW: '',
  PROJ_END_ROW: ''
};

/* ======================= belépési pontok ======================= */

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return handle_(body);
}

function handle_(p) {
  try {
    if (String(p.key || '') !== CONFIG.SECRET) throw new Error('Hibás jelszó – ellenőrizd a Beállításokban.');
    const fn = ACTIONS[p.action];
    if (!fn) throw new Error('Ismeretlen művelet: ' + p.action);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const result = fn(p.data || {});
      return json_({ ok: true, data: result });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  ping: function () { return { pong: true, name: ss_().getName() }; },
  getAll: function () { return getAll_(); },
  addPrint: function (d) { addPrint_(d); return getAll_(); },
  updatePrint: function (d) { updatePrint_(d); return getAll_(); },
  addDelivery: function (d) { addDelivery_(d); return getAll_(); },
  addTracker: function (d) { addTracker_(d); return getAll_(); },
  updateProjectItem: function (d) { updateProjectItem_(d); return getAll_(); },
  addProjectItem: function (d) { addProjectItem_(d); return getAll_(); },
  addProject: function (d) { addProject_(d); return getAll_(); },
  addCal: function (d) { saveCal_(d, true); return getAll_(); },
  updateCal: function (d) { saveCal_(d, false); return getAll_(); },
  setProjectColor: function (d) { return { projectColors: setProjectColor_(d) }; }
};

/* ======================= segédek ======================= */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Nem találom ezt a munkalapot: "' + name + '"');
  return sh;
}
function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function num_(v) {
  if (typeof v === 'number') return v;
  const s = str_(v).replace(/\s/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function q_(s) { return '"' + String(s).replace(/"/g, '""') + '"'; }   // képletbe tett szöveg
function quoted_(formula) {                                              // képletből kiolvasott szövegek
  const out = []; const re = /"((?:[^"]|"")*)"/g; let m;
  while ((m = re.exec(formula || ''))) out.push(m[1].replace(/""/g, '"'));
  return out;
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }

function dateOut_(raw, disp) {
  if (Object.prototype.toString.call(raw) === '[object Date]' && !isNaN(raw)) {
    return Utilities.formatDate(raw, tz_(), 'yyyy-MM-dd');
  }
  const m = String(disp || raw || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}
function dateIn_(iso) {
  const m = String(iso || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return new Date();
  return Utilities.parseDate(m[1] + '-' + m[2] + '-' + m[3], tz_(), 'yyyy-M-d');
}
function hoursOut_(raw, disp) {
  if (typeof raw === 'number') return raw * 24;
  const m = String(disp || '').match(/^\s*(\d+):(\d{1,2})(?::(\d{1,2}))?/);
  if (m) return +m[1] + (+m[2]) / 60 + (m[3] ? +m[3] / 3600 : 0);
  return num_(disp);
}

/** Utolsó sor, ahol a megadott oszlopok bármelyikében van érték. */
function lastDataRow_(sh, cols, firstRow) {
  const last = sh.getLastRow();
  if (last < firstRow) return firstRow - 1;
  let best = firstRow - 1;
  cols.forEach(function (c) {
    const vals = sh.getRange(firstRow, c, last - firstRow + 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (str_(vals[i][0]) !== '') { best = Math.max(best, firstRow + i); break; }
    }
  });
  return best;
}

/** Ha a cella legördülő listájában még nincs benne az új érték, hozzáadja (innentől lefelé). */
function extendList_(sh, row, col, value) {
  value = str_(value);
  if (!value) return;
  const cell = sh.getRange(row, col);
  const rule = cell.getDataValidation();
  if (!rule) return;
  if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return;
  const cv = rule.getCriteriaValues();
  const list = (cv[0] || []).slice();
  if (list.indexOf(value) >= 0) return;
  list.push(value);
  const nr = rule.copy().requireValueInList(list, cv[1] !== false).build();
  sh.getRange(row, col, sh.getMaxRows() - row + 1, 1).setDataValidation(nr);
}


/* ---- Képletek írása a táblázat nyelvi beállításához igazítva ----
   A magyar beállítású táblázat pontosvesszőt (;) vár, az angol vesszőt (,).
   Először a legutóbb bevált változattal próbálkozik, ha hibát ad, a másikkal,
   és megjegyzi, melyik működött. */
function semi_(f) {           // vesszők → pontosvessző (idézőjeleken belül nem)
  let out = '', q = false;
  for (let i = 0; i < f.length; i++) {
    const c = f.charAt(i);
    if (c === '"') q = !q;
    out += (!q && c === ',') ? ';' : c;
  }
  return out;
}
function setF_(cell, f) {
  const props = PropertiesService.getScriptProperties();
  const pref = props.getProperty('formulaSep') || ';';
  const order = pref === ';' ? [semi_(f), f] : [f, semi_(f)];
  for (let i = 0; i < order.length; i++) {
    cell.setFormula(order[i]);
    SpreadsheetApp.flush();
    if (String(cell.getDisplayValue()).charAt(0) !== '#') {
      props.setProperty('formulaSep', order[i] === f ? ',' : ';');
      return;
    }
  }
}
/** Egy sor celláinak írása: a "=" kezdetűek képletként, a többi értékként. */
function writeRow_(sh, row, col, vals) {
  const plain = vals.map(function (v) { return (typeof v === 'string' && v.charAt(0) === '=') ? null : v; });
  // az értékeket egyben írjuk, a képlethelyeket kihagyva
  let start = -1;
  for (let i = 0; i <= plain.length; i++) {
    const isVal = i < plain.length && plain[i] !== null;
    if (isVal && start < 0) start = i;
    if (!isVal && start >= 0) { sh.getRange(row, col + start, 1, i - start).setValues([plain.slice(start, i)]); start = -1; }
  }
  vals.forEach(function (v, i) { if (plain[i] === null) setF_(sh.getRange(row, col + i), v); });
}

/* ======================= OLVASÁS ======================= */

function getAll_() {
  SpreadsheetApp.flush();
  return {
    log: readLog_(),
    projects: readProjects_(),
    stock: readStock_(),
    cal: readCal_(),
    meta: { sheetUrl: ss_().getUrl(), name: ss_().getName(), at: new Date().toISOString(), projectColors: projectColors_() }
  };
}

function readLog_() {
  const sh = sheet_(CONFIG.LOG);
  const last = lastDataRow_(sh, [1, 4], 2);
  if (last < 2) return [];
  const rng = sh.getRange(2, 1, last - 1, 15);
  const v = rng.getValues(), d = rng.getDisplayValues();
  const out = [];
  for (let i = 0; i < v.length; i++) {
    const r = v[i];
    if (str_(r[0]) === '' && str_(r[3]) === '') continue;
    out.push({
      row: i + 2,
      date: dateOut_(r[0], d[i][0]),
      client: str_(r[1]), project: str_(r[2]), model: str_(r[3]),
      pcs: num_(r[4]), g: num_(r[5]), waste: num_(r[6]),
      hours: hoursOut_(r[8], d[i][8]),
      mat: str_(r[9]), color: str_(r[10]), status: str_(r[11]),
      settings: str_(r[12]), support: str_(r[13]), printer: str_(r[14])
    });
  }
  return out;
}

function readProjects_() {
  const sh = sheet_(CONFIG.PROJ);
  const last = sh.getLastRow();
  if (last < 3) return [];
  const rng = sh.getRange(3, 2, last - 2, 7);           // B..H
  const v = rng.getValues(), d = rng.getDisplayValues(), f = rng.getFormulas();
  const groups = [];
  let g = null;
  for (let i = 0; i < v.length; i++) {
    const row = i + 3;
    const name = str_(d[i][0]);
    if (!name) { g = null; continue; }                   // üres sor = csoport vége
    const doneF = f[i][3];
    const isModel = !!doneF || str_(v[i][1]) !== '';
    if (!isModel) {
      g = { name: name, headerRow: row, lastRow: row, items: [] };
      groups.push(g);
      continue;
    }
    if (!g) { g = { name: '', headerRow: null, lastRow: row, items: [] }; groups.push(g); }
    const left = v[i][5];
    g.items.push({
      row: row, name: name,
      crit: quoted_(doneF)[0] || '',
      ordered: num_(v[i][1]), orderedG: num_(v[i][2]),
      done: num_(v[i][3]), doneG: num_(v[i][4]),
      left: typeof left === 'number' ? left : str_(d[i][5]),
      note: str_(v[i][6])
    });
    g.lastRow = row;
  }
  return groups;
}

function readStock_() {
  const sh = sheet_(CONFIG.STOCK);
  const last = sh.getLastRow();
  const trackers = [], deliveries = [];
  if (last >= 3) {
    const rng = sh.getRange(3, 2, last - 2, 4);         // B..E
    const v = rng.getValues(), f = rng.getFormulas();
    for (let i = 0; i < v.length - 1; i++) {
      const label = str_(v[i][0]);
      if (label.indexOf('Felhasznált') !== 0) continue;
      const vr = v[i + 1], fr = f[i + 1];
      const usedQ = quoted_(fr[0]);
      const availQ = quoted_(fr[3]);
      trackers.push({
        labelRow: i + 3, row: i + 4,
        name: label.replace(/^Felhasznált\s*/, '').replace(/:\s*$/, ''),
        used: num_(vr[0]), waste: num_(vr[1]), wastePct: num_(vr[2]), avail: num_(vr[3]),
        logMat: usedQ[0] || '', logColor: usedQ[1] || '',
        deliv: availQ.length >= 4 ? { brand: availQ[0], type: availQ[1], spec: availQ[2], color: availQ[3] } : null
      });
      i++;
    }
  }
  if (last >= 5) {
    const rng = sh.getRange(5, 10, last - 4, 6);        // J..O
    const v = rng.getValues(), d = rng.getDisplayValues();
    for (let i = 0; i < v.length; i++) {
      if (str_(v[i][0]) === '') continue;
      deliveries.push({
        row: i + 5, brand: str_(v[i][0]), type: str_(v[i][1]), spec: str_(v[i][2]),
        color: str_(v[i][3]), grams: num_(v[i][4]), date: dateOut_(v[i][5], d[i][5])
      });
    }
  }
  return { trackers: trackers, deliveries: deliveries };
}

function readCal_() {
  const sh = sheet_(CONFIG.CAL);
  const last = lastDataRow_(sh, [1, 3], 4);
  const notes = [str_(sh.getRange(1, 1).getValue()), str_(sh.getRange(1, 3).getValue()), str_(sh.getRange(1, 4).getValue())]
    .filter(function (s) { return s; });
  const rows = [];
  if (last >= 4) {                                     // 1. sor: jegyzetek, 3. sor: fejléc, 4.-től adatok
    const v = sh.getRange(4, 1, last - 3, 11).getValues();
    for (let i = 0; i < v.length; i++) {
      const r = v[i];
      if (str_(r[0]) === '' && str_(r[2]) === '') continue;
      rows.push({
        row: i + 4, maker: str_(r[0]), type: str_(r[1]), product: str_(r[2]), desc: str_(r[3]),
        calibrated: str_(r[4]), extr: str_(r[5]), plate: str_(r[6]), flow: str_(r[8]), pa: str_(r[10])
      });
    }
  }
  return { notes: notes, rows: rows };
}

/* ======================= projektszínek ======================= */
// A projektek színeit a szkript saját tárhelyén tartjuk (nem a táblázatban),
// így a telefon és a PC is ugyanazt látja, a táblázat pedig érintetlen marad.
function projectColors_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('projectColors') || '{}'); }
  catch (e) { return {}; }
}
function setProjectColor_(d) {
  const key = str_(d.project).toLowerCase();
  if (!key) throw new Error('Hiányzó projektnév.');
  const all = projectColors_();
  if (str_(d.color)) all[key] = str_(d.color); else delete all[key];
  PropertiesService.getScriptProperties().setProperty('projectColors', JSON.stringify(all));
  return all;
}

/* ======================= ÍRÁS: napló ======================= */

function printRow_(d, r) {
  return [
    dateIn_(d.date), str_(d.client), str_(d.project), str_(d.model),
    num_(d.pcs), num_(d.g), num_(d.waste),
    '=SUM(F' + r + ':G' + r + ')',
    num_(d.hours) / 24,
    str_(d.mat), str_(d.color), str_(d.status) || 'Kész', str_(d.settings),
    str_(d.support) || 'Nem', str_(d.printer)
  ];
}

function addPrint_(d) {
  if (!str_(d.model)) throw new Error('A modell neve kötelező.');
  const sh = sheet_(CONFIG.LOG);
  const last = lastDataRow_(sh, [1, 4], 2);
  const r = last + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  if (last >= 2) sh.getRange(last, 1, 1, 15).copyTo(sh.getRange(r, 1, 1, 15), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  writeRow_(sh, r, 1, printRow_(d, r));
  sh.getRange(r, 9).setNumberFormat('[h]":"mm');
  [[2, d.client], [3, d.project], [10, d.mat], [11, d.color], [12, d.status], [15, d.printer]]
    .forEach(function (p) { extendList_(sh, r, p[0], p[1]); });
}

function updatePrint_(d) {
  const sh = sheet_(CONFIG.LOG);
  const r = Number(d.row);
  if (!r || r < 2) throw new Error('Hibás sorszám.');
  const current = str_(sh.getRange(r, 4).getValue());
  if (current !== str_(d.expectModel)) {
    throw new Error('A táblázat időközben megváltozott (a ' + r + '. sorban most „' + current + '” van). Frissíts, és próbáld újra.');
  }
  const vals = printRow_(d, r);
  sh.getRange(r, 1, 1, 7).setValues([vals.slice(0, 7)]);
  sh.getRange(r, 9, 1, 7).setValues([vals.slice(8)]);
  [[2, d.client], [3, d.project], [10, d.mat], [11, d.color], [12, d.status], [15, d.printer]]
    .forEach(function (p) { extendList_(sh, r, p[0], p[1]); });
}

/* ======================= ÍRÁS: készlet ======================= */

function addDelivery_(d) {
  if (!num_(d.grams)) throw new Error('A mennyiség kötelező.');
  const sh = sheet_(CONFIG.STOCK);
  const vals = sh.getRange(5, 10, Math.max(sh.getMaxRows() - 4, 1), 6).getValues();
  let r = 5;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i].every(function (x) { return str_(x) === ''; })) { r = 5 + i; break; }
    r = 6 + i;
  }
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
  if (r > 5) sh.getRange(r - 1, 10, 1, 6).copyTo(sh.getRange(r, 10, 1, 6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(r, 10, 1, 6).setValues([[str_(d.brand), str_(d.type), str_(d.spec) || 'Basic', str_(d.color), num_(d.grams), dateIn_(d.date)]]);
}

function addTracker_(d) {
  ['name', 'logMat', 'logColor', 'brand', 'type', 'color'].forEach(function (k) {
    if (!str_(d[k])) throw new Error('Hiányzó mező: ' + k);
  });
  const sh = sheet_(CONFIG.STOCK);
  const st = readStock_();
  const lastT = st.trackers[st.trackers.length - 1];
  let r = lastT ? lastT.row + 1 : 3;
  const maxR = sh.getMaxRows();
  while (r + 1 <= maxR) {
    const block = sh.getRange(r, 2, 2, 4).getValues();
    if (block.every(function (row) { return row.every(function (x) { return str_(x) === ''; }); })) break;
    r++;
  }
  if (r + 1 > maxR) sh.insertRowsAfter(maxR, 10);
  const v = r + 1, S = CONFIG.STOCK_START_ROW, E = CONFIG.STOCK_END_ROW, L = "'" + CONFIG.LOG + "'!";   // E üres = nyitott tartomány
  if (lastT) sh.getRange(lastT.labelRow, 2, 2, 4).copyTo(sh.getRange(r, 2, 2, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  const crit = L + 'J' + S + ':J' + E + ',' + q_(d.logMat) + ',' + L + 'K' + S + ':K' + E + ',' + q_(d.logColor);
  writeRow_(sh, r, 2, ['Felhasznált ' + str_(d.name) + ':', 'Ebből hulladék', '', 'Elérhető még:']);
  writeRow_(sh, r + 1, 2,
    [
      '=SUMIFS(' + L + 'H' + S + ':H' + E + ',' + crit + ')',
      '=SUMIFS(' + L + 'G' + S + ':G' + E + ',' + crit + ')',
      '=IFERROR(C' + v + '/(B' + v + '/100),0)',
      '=SUMIFS(N:N,J:J,' + q_(d.brand) + ',K:K,' + q_(d.type) + ',L:L,' + q_(str_(d.spec) || 'Basic') + ',M:M,' + q_(d.color) + ')-B' + v
    ]);
}

/* ======================= ÍRÁS: projektek ======================= */

function updateProjectItem_(d) {
  const sh = sheet_(CONFIG.PROJ);
  const r = Number(d.row);
  const current = str_(sh.getRange(r, 2).getDisplayValue());
  if (current !== str_(d.expectName)) throw new Error('A Projektek lap időközben megváltozott. Frissíts, és próbáld újra.');
  if (d.ordered !== undefined && d.ordered !== '') sh.getRange(r, 3).setValue(num_(d.ordered));
  if (d.note !== undefined) sh.getRange(r, 8).setValue(str_(d.note));
}

function addProjectItem_(d) {
  const name = str_(d.name), crit = str_(d.crit) || name;
  if (!name) throw new Error('A modell neve kötelező.');
  const sh = sheet_(CONFIG.PROJ);
  const after = Number(d.afterRow);
  const current = str_(sh.getRange(after, 2).getDisplayValue());
  if (current !== str_(d.expectName)) throw new Error('A Projektek lap időközben megváltozott. Frissíts, és próbáld újra.');

  // Honnan számoljon a napló? Az első olyan naplósortól, ahol ez a modellnév szerepel, különben a következő új sortól.
  const log = sheet_(CONFIG.LOG);
  const lastLog = lastDataRow_(log, [1, 4], 2);
  let S = lastLog + 1;
  if (lastLog >= 2) {
    const models = log.getRange(2, 4, lastLog - 1, 1).getValues();
    for (let i = 0; i < models.length; i++) { if (str_(models[i][0]) === crit) { S = i + 2; break; } }
  }
  const E = CONFIG.PROJ_END_ROW, L = "'" + CONFIG.LOG + "'!";

  sh.insertRowAfter(after);
  const r = after + 1;
  // formátum: egy meglévő modellsorról
  const fmtRow = d.formatRow ? Number(d.formatRow) + (Number(d.formatRow) > after ? 1 : 0) : after;
  sh.getRange(fmtRow, 2, 1, 7).copyTo(sh.getRange(r, 2, 1, 7), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  writeRow_(sh, r, 2, [
    name, num_(d.ordered), '',
    '=SUMIF(' + L + 'D' + S + ':D' + E + ',' + q_(crit) + ',' + L + 'E' + S + ':E' + E + ')',
    '=SUMIF(' + L + 'D' + S + ':D' + E + ',' + q_(crit) + ',' + L + 'F' + S + ':F' + E + ')',
    '=IF(C' + r + '-E' + r + '>0,C' + r + '-E' + r + ',IF(C' + r + '-E' + r + '=0,"Kész!",ABS(C' + r + '-E' + r + ')&" darabbal több"))',
    str_(d.note)
  ]);
}

function addProject_(d) {
  const name = str_(d.name);
  if (!name) throw new Error('A projekt neve kötelező.');
  const sh = sheet_(CONFIG.PROJ);
  const groups = readProjects_();
  const last = lastDataRow_(sh, [2], 3);
  const r = last + 2;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
  const header = groups.filter(function (g) { return g.headerRow; })[0];
  if (header) sh.getRange(header.headerRow, 2, 1, 7).copyTo(sh.getRange(r, 2, 1, 7), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(r, 2).setValue(name);
}

/* ======================= ÍRÁS: kalibráció ======================= */

function saveCal_(d, isNew) {
  if (!str_(d.product)) throw new Error('A terméknév kötelező.');
  const sh = sheet_(CONFIG.CAL);
  let r;
  if (isNew) {
    const last = lastDataRow_(sh, [1, 3], 4);
    r = last + 1;
    if (last >= 4) sh.getRange(last, 1, 1, 11).copyTo(sh.getRange(r, 1, 1, 11), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  } else {
    r = Number(d.row);
    const current = str_(sh.getRange(r, 3).getValue());
    if (current !== str_(d.expectProduct)) throw new Error('A Kalibrációk lap időközben megváltozott. Frissíts, és próbáld újra.');
  }
  const t = function (x) { const s = str_(x); return s !== '' && !isNaN(Number(s.replace(',', '.'))) ? Number(s.replace(',', '.')) : s; };
  sh.getRange(r, 1, 1, 11).setValues([[
    str_(d.maker), str_(d.type), str_(d.product), str_(d.desc), str_(d.calibrated) || 'Nem',
    t(d.extr), t(d.plate), 'Flow ratio:', str_(d.flow) || 'N/A', 'Pressure Advance:', str_(d.pa) || 'N/A'
  ]]);
}

/** Kézi próba: a szerkesztőben futtasd ezt (Futtatás gomb), hogy engedélyt adj és lásd, működik-e. */
function teszt() {
  const all = getAll_();
  Logger.log('Napló sorok: ' + all.log.length + ', projektcsoportok: ' + all.projects.length +
    ', filamentkövetők: ' + all.stock.trackers.length + ', kalibrációk: ' + all.cal.rows.length);
}
