const SPREADSHEET_ID = "1xXjRCHncZ8xvvi6s6qyx6zOpsi6dWHv-idj57DZGEIE";
const API_KEY = "GANTI_DENGAN_KUNCI_RAHASIA_ANDA";

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

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "health";
  if (action === "health") return json_({ ok: true, app: "HPP MASTER API", version: "1.0" });
  return json_({ ok: false, message: "Gunakan POST untuk operasi database." });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.key !== API_KEY) return json_({ ok: false, message: "API key tidak valid." });
    const action = body.action || "";
    if (action === "initialize") return json_(initializeDatabase_());
    if (action === "getAll") return json_(getAll_());
    if (action === "add") return json_(addRecord_(body.table, body.record || {}));
    if (action === "replaceTable") return json_(replaceTable_(body.table, body.records || []));
    return json_({ ok: false, message: "Action tidak dikenal." });
  } catch (err) {
    return json_({ ok: false, message: String(err && err.message ? err.message : err) });
  }
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
  return { ok: true, message: "Database HPP MASTER berhasil disiapkan.", spreadsheet: ss.getUrl(), created: created, tables: Object.keys(TABLES) };
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
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  });
  return { ok: true, data: data };
}

function addRecord_(table, record) {
  if (!TABLES[table]) throw new Error("Nama tabel tidak diizinkan.");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(table);
  if (!sh) throw new Error("Tabel belum dibuat. Jalankan initialize terlebih dahulu.");
  const headers = TABLES[table];
  const row = headers.map(h => record[h] !== undefined ? record[h] : "");
  sh.appendRow(row);
  return { ok: true, table: table, inserted: 1 };
}

function replaceTable_(table, records) {
  if (!TABLES[table]) throw new Error("Nama tabel tidak diizinkan.");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(table);
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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Jalankan fungsi ini sekali dari editor Apps Script untuk membuat semua tab.
function SETUP_HPP_MASTER() {
  const result = initializeDatabase_();
  Logger.log(JSON.stringify(result));
}