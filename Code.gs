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

function updateRecord(table, id, record) {
  if (!TABLES[table] || !id) throw new Error("Tabel atau ID tidak valid.");
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(table);
  if (!sh || sh.getLastRow() < 2) throw new Error("Data tidak ditemukan.");
  const headers = TABLES[table];
  const idCol = headers.indexOf("id") + 1;
  if (!idCol) throw new Error("Kolom ID tidak tersedia.");
  const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getDisplayValues();
  const offset = ids.findIndex(r => String(r[0]) === String(id));
  if (offset < 0) throw new Error("Data dengan ID tersebut tidak ditemukan.");
  const rowNumber = offset + 2;
  const old = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const merged = {};
  headers.forEach((key, i) => merged[key] = key === "id" ? old[i] : (record[key] !== undefined ? record[key] : old[i]));
  if (table === "BAHAN") {
    const qty = Number(merged.jumlah_dasar || 0);
    if (!(qty > 0)) throw new Error("Jumlah dalam satuan dasar harus lebih dari 0.");
    merged.harga_per_satuan_dasar = Number(merged.harga_beli || 0) / qty;
    if (Number(old[headers.indexOf("harga_beli")]) !== Number(merged.harga_beli)) {
      const history = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("RIWAYAT_HARGA");
      if (history) history.appendRow(["HRG-" + Date.now(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), id, merged.nama, old[headers.indexOf("harga_beli")], merged.harga_beli, merged.satuan_beli, "Perubahan harga melalui aplikasi"]);
    }
  }
  if (table === "PACKAGING") {
    const qty = Number(merged.jumlah_beli || 0);
    if (!(qty > 0)) throw new Error("Jumlah pembelian packaging harus lebih dari 0.");
    merged.biaya_per_unit = Number(merged.harga_beli || 0) / qty;
  }
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([headers.map(k => merged[k])]);
  return {ok:true,table:table,updated:1};
}

function deleteRecord(table, id) {
  if (!TABLES[table] || !id) throw new Error("Tabel atau ID tidak valid.");
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(table);
  if (!sh || sh.getLastRow() < 2) throw new Error("Data tidak ditemukan.");
  const idCol = TABLES[table].indexOf("id") + 1;
  const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getDisplayValues();
  const offset = ids.findIndex(r => String(r[0]) === String(id));
  if (offset < 0) throw new Error("Data dengan ID tersebut tidak ditemukan.");
  // Lindungi bahan yang masih dipakai dalam komposisi resep.
  if (table === "BAHAN") {
    const detail = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("DETAIL_RESEP");
    if (detail && detail.getLastRow() >= 2) {
      const headers = TABLES.DETAIL_RESEP;
      const bahanIdCol = headers.indexOf("bahan_id") + 1;
      const usedIds = detail.getRange(2, bahanIdCol, detail.getLastRow() - 1, 1).getDisplayValues();
      if (usedIds.some(r => String(r[0]) === String(id))) {
        throw new Error("Bahan tidak dapat dihapus karena masih digunakan dalam komposisi resep. Hapus atau ganti komposisinya terlebih dahulu.");
      }
    }
  }
  sh.deleteRow(offset + 2);
  return {ok:true,table:table,deleted:1};
}

/**
 * Perbaikan satu kali untuk tautan komposisi resep uji.
 * Mengalihkan detail dari ID lama ke ID resep yang terbentuk saat resep diedit.
 */
function PERBAIKI_LINK_RESEP_UJI() {
  const oldRecipeId = "HPP-1795195135721";
  const currentRecipeId = "HPP-1795196304003";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recipeSheet = ss.getSheetByName("RESEP");
  const detailSheet = ss.getSheetByName("DETAIL_RESEP");
  if (!recipeSheet || !detailSheet) throw new Error("Sheet RESEP atau DETAIL_RESEP tidak ditemukan.");
  const recipeValues = recipeSheet.getDataRange().getValues();
  const recipeHeaders = recipeValues.shift();
  const idIndex = recipeHeaders.indexOf("id");
  if (idIndex < 0 || !recipeValues.some(row => String(row[idIndex]) === currentRecipeId)) {
    throw new Error("ID resep tujuan tidak ditemukan. Tidak ada data yang diubah.");
  }
  const detailValues = detailSheet.getDataRange().getValues();
  if (detailValues.length < 2) return {ok:true,updated:0};
  const detailHeaders = detailValues[0];
  const recipeIdIndex = detailHeaders.indexOf("resep_id");
  if (recipeIdIndex < 0) throw new Error("Kolom resep_id tidak ditemukan.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let updated = 0;
    for (let r = 1; r < detailValues.length; r++) {
      if (String(detailValues[r][recipeIdIndex]) === oldRecipeId) {
        detailSheet.getRange(r + 1, recipeIdIndex + 1).setValue(currentRecipeId);
        updated++;
      }
    }
    SpreadsheetApp.flush();
    return {ok:true,updated:updated,recipe_id:currentRecipeId};
  } finally {
    lock.releaseLock();
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
    // Simpan harga per satuan dasar secara permanen, bukan hanya menghitungnya di tampilan.
    if (table === "BAHAN") {
      const harga = Number(record.harga_beli || 0);
      const jumlahDasar = Number(record.jumlah_dasar || 0);
      if (!(jumlahDasar > 0)) throw new Error("Jumlah dalam satuan dasar harus lebih dari 0.");
      record.harga_per_satuan_dasar = harga / jumlahDasar;
    }
    if (table === "PACKAGING") {
      const qty = Number(record.jumlah_beli || 0);
      if (!(qty > 0)) throw new Error("Jumlah pembelian packaging harus lebih dari 0.");
      record.biaya_per_unit = Number(record.harga_beli || 0) / qty;
    }
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

// Mengambil komposisi resep untuk ditampilkan di aplikasi.
function getRecipeDetails() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("DETAIL_RESEP");
  if (!sh || sh.getLastRow() < 2) return { ok: true, data: [] };
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const data = values.filter(row => row.some(v => v !== "")).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { ok: true, data: data };
}


// Konversi satuan resep ke satuan dasar bahan. Satuan yang tidak sekelompok ditolak.
function convertToBaseUnit_(quantity, fromUnit, toUnit) {
  const units = {
    gram: { group: "mass", factor: 1 }, kg: { group: "mass", factor: 1000 },
    ml: { group: "volume", factor: 1 }, liter: { group: "volume", factor: 1000 },
    pcs: { group: "count", factor: 1 }, buah: { group: "count", factor: 1 }
  };
  const from = units[String(fromUnit).toLowerCase()];
  const to = units[String(toUnit).toLowerCase()];
  if (!from || !to || from.group !== to.group) {
    throw new Error("Satuan " + fromUnit + " tidak dapat dikonversi ke " + toUnit + ". Pilih satuan yang sesuai dengan bahan.");
  }
  return quantity * from.factor / to.factor;
}

// Simpan komposisi sekaligus menghitung biaya berdasarkan harga satuan dasar bahan.
// Membentuk data komposisi dan menghitung ulang biaya dari harga bahan terbaru.
function buildRecipeDetailRow_(record, id) {
  record = record || {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const bahanSheet = ss.getSheetByName("BAHAN");
  const resepSheet = ss.getSheetByName("RESEP");
  if (!record.resep_id || !record.bahan_id) throw new Error("Resep dan bahan wajib dipilih.");
  const bahanRows = bahanSheet.getDataRange().getValues();
  const bahanHeaders = bahanRows.shift();
  const bahan = bahanRows.map(row => {
    const o = {}; bahanHeaders.forEach((h, i) => o[h] = row[i]); return o;
  }).find(x => String(x.id) === String(record.bahan_id));
  if (!bahan) throw new Error("Bahan tidak ditemukan.");
  const resepRows = resepSheet.getDataRange().getValues();
  const resepHeaders = resepRows.shift();
  const resep = resepRows.map(row => {
    const o = {}; resepHeaders.forEach((h, i) => o[h] = row[i]); return o;
  }).find(x => String(x.id) === String(record.resep_id));
  if (!resep) throw new Error("Resep tidak ditemukan.");
  const qtyInput = Number(record.jumlah_pakai);
  const baseQty = Number(bahan.jumlah_dasar);
  const price = Number(bahan.harga_beli);
  if (!(qtyInput > 0) || !(baseQty > 0) || !(price >= 0)) throw new Error("Jumlah pemakaian atau data harga bahan tidak valid.");
  const inputUnit = String(record.satuan_pakai || bahan.satuan_dasar);
  const qty = convertToBaseUnit_(qtyInput, inputUnit, String(bahan.satuan_dasar));
  const cost = qty * (price / baseQty);
  return [
    id || ("DET-" + Date.now()), record.resep_id, record.bahan_id, bahan.nama,
    qty, bahan.satuan_dasar, cost, "Bahan utama"
  ];
}

// Tambah komposisi baru.
function addRecipeDetail(record) {
  const row = buildRecipeDetailRow_(record, "");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("DETAIL_RESEP").appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, biaya: row[6], nama_bahan: row[3], satuan: row[5] };
}

// Edit komposisi berdasarkan ID; ID lama dipertahankan dan biaya dihitung ulang.
function updateRecipeDetail(id, record) {
  if (!id) throw new Error("ID komposisi tidak valid.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("DETAIL_RESEP");
    if (!sh || sh.getLastRow() < 2) throw new Error("Komposisi tidak ditemukan.");
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    const offset = ids.findIndex(r => String(r[0]) === String(id));
    if (offset < 0) throw new Error("Komposisi dengan ID tersebut tidak ditemukan.");
    const rowNumber = offset + 2;
    const row = buildRecipeDetailRow_(record, id);
    sh.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    return { ok: true, updated: 1, biaya: row[6], nama_bahan: row[3], satuan: row[5] };
  } finally {
    lock.releaseLock();
  }
}

// Menghitung dan menyimpan HPP batch. Waste diterapkan pada subtotal biaya.
function calculateAndSaveHPP(input) {
  input = input || {};
  const recipeId = String(input.resep_id || "");
  const batches = Number(input.jumlah_batch || 1);
  const actualYield = Number(input.hasil_produksi || 0);
  const packagingPerUnit = Number(input.biaya_packaging_per_unit || 0);
  const labor = Number(input.total_tenaga_kerja || 0);
  const overhead = Number(input.total_overhead || 0);
  const wastePct = Number(input.waste_persen || 0);
  if (!recipeId || !(batches > 0) || !(actualYield > 0)) throw new Error("Pilih resep, jumlah batch, dan hasil produksi yang valid.");
  if ([packagingPerUnit, labor, overhead, wastePct].some(x => !isFinite(x) || x < 0) || wastePct > 100) throw new Error("Biaya dan waste harus bernilai 0 atau lebih; waste maksimal 100%.");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const all = getAll_().data;
  const recipe = all.RESEP.find(x => String(x.id) === recipeId);
  if (!recipe) throw new Error("Resep tidak ditemukan.");
  const details = all.DETAIL_RESEP.filter(x => String(x.resep_id) === recipeId);
  if (!details.length) throw new Error("Resep belum memiliki komposisi bahan.");
  const baseIngredients = details.reduce((sum, x) => sum + Number(x.biaya || 0), 0) * batches;
  const packaging = packagingPerUnit * actualYield;
  const subtotal = baseIngredients + packaging + labor + overhead;
  const wasteCost = subtotal * wastePct / 100;
  const total = subtotal + wasteCost;
  const perUnit = total / actualYield;
  const id = "HPP-" + Date.now();
  const date = input.tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const row = [
    id, date, recipeId, recipe.nama_produk, batches, baseIngredients, 0,
    packaging, labor, overhead, wasteCost, total, actualYield, perUnit,
    input.catatan || ""
  ];
  ss.getSheetByName("HPP").appendRow(row);
  return {
    ok: true, id: id, nama_produk: recipe.nama_produk,
    total_bahan: baseIngredients, total_packaging: packaging,
    total_tenaga_kerja: labor, total_overhead: overhead,
    total_waste: wasteCost, total_biaya: total,
    hasil_produksi: actualYield, hpp_per_unit: perUnit
  };
}


// Mengisi harga_per_satuan_dasar pada seluruh data bahan lama dan baru.
function PERBARUI_HARGA_SATUAN_DASAR() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("BAHAN");
  if (!sh || sh.getLastRow() < 2) return { ok: true, updated: 0 };
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idxHarga = headers.indexOf("harga_beli");
  const idxJumlah = headers.indexOf("jumlah_dasar");
  const idxPerUnit = headers.indexOf("harga_per_satuan_dasar");
  if ([idxHarga, idxJumlah, idxPerUnit].some(i => i < 0)) throw new Error("Kolom harga bahan tidak ditemukan.");
  let updated = 0;
  const out = values.slice(1).map(row => {
    const harga = Number(row[idxHarga] || 0);
    const jumlah = Number(row[idxJumlah] || 0);
    const val = jumlah > 0 ? harga / jumlah : "";
    if (row[idxPerUnit] !== val) updated++;
    row[idxPerUnit] = val;
    return [val];
  });
  if (out.length) sh.getRange(2, idxPerUnit + 1, out.length, 1).setValues(out);
  sh.getRange(2, idxPerUnit + 1, Math.max(out.length, 1), 1).setNumberFormat('"Rp" #,##0.00');
  return { ok: true, updated: updated, rows: out.length };
}

// Merapikan semua tab: lebar kolom, tinggi baris, header, format angka, dan filter.
function RAPIKAN_SHEET_HPP_MASTER() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const currencyHeaders = ["harga_beli","harga_per_satuan_dasar","biaya","biaya_per_unit","nominal","total_bahan","total_tambahan","total_packaging","total_tenaga_kerja","total_overhead","total_waste","total_biaya","hpp_per_unit","harga_lama","harga_baru"];
  const dateHeaders = ["tanggal","updated_at"];
  Object.keys(TABLES).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastCol = Math.max(sh.getLastColumn(), TABLES[name].length);
    const lastRow = Math.max(sh.getLastRow(), 1);
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 38);
    if (lastRow > 1) sh.setRowHeights(2, lastRow - 1, 28);
    const headerRange = sh.getRange(1, 1, 1, lastCol);
    headerRange.setBackground("#185adb").setFontColor("#ffffff")
      .setFontWeight("bold").setFontSize(10).setVerticalAlignment("middle")
      .setWrap(true);
    sh.getRange(1, 1, lastRow, lastCol).setVerticalAlignment("middle");
    if (lastRow > 1) {
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      headers.forEach((h, i) => {
        const col = i + 1;
        if (currencyHeaders.indexOf(String(h)) >= 0) {
          sh.getRange(2, col, lastRow - 1, 1).setNumberFormat('"Rp" #,##0.00');
        } else if (dateHeaders.indexOf(String(h)) >= 0) {
          sh.getRange(2, col, lastRow - 1, 1).setNumberFormat("dd mmm yyyy");
        } else if (["jumlah_beli","jumlah_dasar","hasil_batch","jumlah_pakai","jumlah_rencana","hasil_aktual","jumlah_batch","hasil_produksi"].indexOf(String(h)) >= 0) {
          sh.getRange(2, col, lastRow - 1, 1).setNumberFormat("#,##0.###");
        }
      });
    }
    sh.autoResizeColumns(1, lastCol);
    for (let col = 1; col <= lastCol; col++) {
      const width = sh.getColumnWidth(col);
      sh.setColumnWidth(col, Math.max(105, Math.min(width, 250)));
    }
    if (sh.getFilter()) sh.getFilter().remove();
    if (lastRow > 1) sh.getRange(1, 1, lastRow, lastCol).createFilter();
  });
  const update = PERBARUI_HARGA_SATUAN_DASAR();
  SpreadsheetApp.flush();
  return { ok: true, formatted: Object.keys(TABLES).length, priceRowsUpdated: update.updated };
}
