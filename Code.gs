/**
 * Backend Rekap Absensi SMKN 1 Sungai Rumbai.
 * Data disimpan di sheet "data" pada spreadsheet ini: satu baris = satu dokumen (path, json).
 */
const NAMA_SHEET = 'data';

function buku_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Data Absensi SMKN 1 Sungai Rumbai');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// Jalankan sekali dari editor untuk melihat alamat spreadsheet datanya (menu Execution log).
function lihatAlamatSheet() {
  Logger.log(buku_().getUrl());
}

function lembar_() {
  const ss = buku_();
  let s = ss.getSheetByName(NAMA_SHEET);
  if (!s) {
    s = ss.insertSheet(NAMA_SHEET);
    s.appendRow(['path', 'json', 'diubah']);
    s.setFrozenRows(1);
    s.setColumnWidth(1, 260);
    s.setColumnWidth(2, 600);
  }
  return s;
}

function baca_() {
  const s = lembar_();
  const n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, 2).getValues() : [];
}

function kirim_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const docs = {};
  baca_().forEach(function (r) {
    try { docs[r[0]] = JSON.parse(r[1]); } catch (e) {}
  });
  return kirim_({ ok: true, docs: docs });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ops = JSON.parse(e.postData.contents).ops || [];
    const s = lembar_();
    const rows = baca_();
    const index = {};
    rows.forEach(function (r, i) { index[r[0]] = i + 2; });
    const hapus = [];
    const tambah = [];
    const sekarang = new Date();
    ops.forEach(function (o) {
      const path = String(o.path);
      if (o.op === 'set') {
        const json = JSON.stringify(o.data);
        if (json.length > 49000) throw new Error('Dokumen terlalu besar: ' + path);
        if (index[path]) s.getRange(index[path], 1, 1, 3).setValues([[path, json, sekarang]]);
        else { tambah.push([path, json, sekarang]); index[path] = -1; }
      } else if (o.op === 'del' && index[path] > 0) {
        hapus.push(index[path]);
        delete index[path];
      }
    });
    hapus.sort(function (a, b) { return b - a; }).forEach(function (r) { s.deleteRow(r); });
    if (tambah.length) s.getRange(s.getLastRow() + 1, 1, tambah.length, 3).setValues(tambah);
    return kirim_({ ok: true });
  } catch (err) {
    return kirim_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
