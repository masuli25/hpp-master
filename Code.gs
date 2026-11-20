const SPREADSHEET_ID = "1xXjRCHncZ8xvvi6s6qyx6zOpsi6dWHv-idj57DZGEIE";

const TABLES = {
  PENGATURAN: ["key","value","updated_at"],
  BAHAN: ["id","nama","kategori","harga_beli","jumlah_beli","satuan_beli","jumlah_dasar","satuan_dasar","harga_per_satuan_dasar","supplier","catatan","updated_at"],
  KATEGORI: ["id","nama_kategori","jenis","catatan"],
  RESEP: ["id","nama_produk","kategori","hasil_batch","satuan_hasil","catatan","updated_at"],
  DETAIL_RESEP: ["id","resep_id","bahan_id","nama_bahan","jumlah_pakai","satuan_pakai","biaya","jenis_komponen"],
  PRODUKSI: ["id","tanggal","resep_id","nama_produk","jumlah_rencana","hasil_aktual","waste_persen","biaya_tenaga_kerja","catatan"],
  BIAYA_OPERASIONAL: ["id","nama_biaya","kategori","jenis_biaya","nominal","periode","metode_alokasi","target_alokasi","catatan","updated_at"],
  PACKAGING: ["id","nama_kemasan","harga_beli","jumlah_beli","satuan_beli","biaya_per_unit","catatan","updated_at"],
  HPP: ["id","tanggal","resep_id","nama_produk","jumlah_batch","total_bahan","total_tambahan","total_packaging","total_tenaga_kerja","total_overhead","total_waste","total_biaya","hasil_produksi","hpp_per_unit","catatan"],
  RIWAYAT_HARGA: ["id","tanggal","bahan_id","nama_bahan","harga_lama","harga_baru","satuan","catatan"]
};

// Aplikasi disajikan dari Apps Script agar akses Google membatasi pengguna.
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("HPP MASTER — Food Costing")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function SETUP_HPP_MASTER() {
  const result = initializeDatabase_();
  Logger.log(JSON.stringify(result));
  return result;
}

// Fungsi-fungsi api* dipanggil dari halaman Index melalui google.script.run.
// Tidak menggunakan API key di browser.
function getAllData() {
  return getAll_();
}

function addRecord(table, record) {
  return addRecord_(table, record || {});
}

function replaceTable(table, records) {
  return replaceTable_(table, records || []);
}

function initializeDatabase_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const created = [];
  Object.keys(TABLES).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); created.push(name); }
    const headers = TABLES[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#185adb").setFontColor("#ffffff");
      sh.autoResizeColumns(1, headers.length);
    } else {
      const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
      headers.forEach((h, i) => { if (!current[i]) sh.getRange(1, i + 1).setValue(h); });
    }
  });
  return { ok: true, message: "Database HPP MASTER siap.", created: created, tables: Object.keys(TABLES) };
}

function getAll_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = {};
  Object.keys(TABLES).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { data[name] = []; return; }
    const values = sh.getDataRange().getValues();
    const headers = values.shift();
    data[name] = values.filter(row => row.some(v => v !== "")).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const value = row[i];
        obj[h] = value instanceof Date ? value.toISOString() : value;
      });
      return obj;
    });
  });
  return { ok: true, data: data };
}

function addRecord_(table, record) {
  if (!TABLES[table]) throw new Error("Nama tabel tidak diizinkan.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(table);
    if (!sh) throw new Error("Tabel belum dibuat.");
    const headers = TABLES[table];
    const row = headers.map(h => record[h] !== undefined ? record[h] : "");
    sh.appendRow(row);
    return { ok: true, table: table, inserted: 1 };
  } finally {
    lock.releaseLock();
  }
}

function replaceTable_(table, records) {
  if (!TABLES[table]) throw new Error("Nama tabel tidak diizinkan.");
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(table);
  if (!sh) throw new Error("Tabel belum dibuat.");
  const headers = TABLES[table];
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  if (records.length) {
    const rows = records.map(r => headers.map(h => r[h] !== undefined ? r[h] : ""));
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return { ok: true, table: table, count: records.length };
}