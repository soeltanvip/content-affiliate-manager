const SPREADSHEET_ID = '1kdxxV_tc5PHsicspDd9k43awfHFgDO6XKGQ1Y3TeIos';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DanDaf Affiliate Content Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Inisialisasi dan sinkronisasi struktur basis data
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const tables = [
    {
      name: 'produk',
      headers: ['id', 'kode_produk', 'nama', 'kategori', 'cover_url', 'link_affiliate', 'link_dashboard', 'created_at']
    },
    {
      name: 'media',
      headers: ['id', 'produk_id', 'batch_id', 'batch_nama', 'media_url', 'media_type', 'content_type', 'catatan', 'created_at']
    },
    {
      name: 'pengaturan',
      headers: ['kunci', 'nilai']
    },
    {
      name: 'postingan',
      headers: ['id', 'produk_id', 'tanggal', 'slot', 'kode_tracking', 'status', 'fb_url', 'caption', 'created_at']
    }
  ];

  tables.forEach(function(tbl) {
    let sheet = ss.getSheetByName(tbl.name);
    if (!sheet) {
      sheet = ss.insertSheet(tbl.name);
    }
    // Update atau set header baris pertama
    sheet.getRange(1, 1, 1, tbl.headers.length).setValues([tbl.headers]);
  });

  return { ok: true, success: true, message: 'Database setup completed.' };
}

// Shared serializer untuk memastikan tipe data JSON-safe untuk google.script.run
function serializeForClient(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) {
    return value.map(function(item) { return serializeForClient(item); });
  }
  if (typeof value === 'object') {
    const sanitized = {};
    for (const key in value) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        sanitized[key] = serializeForClient(value[key]);
      }
    }
    return sanitized;
  }
  return value;
}

// Endpoint pemuatan data awal
function getInitialData() {
  try {
    setupDatabase(); // Pastikan lembar kerja ada
    const settings = getSettingsInternal();
    const products = getProductsInternal();
    return {
      ok: true,
      success: true,
      data: serializeForClient({
        settings: settings,
        products: products
      })
    };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal memuat data awal: ' + err.message };
  }
}

// Ambil daftar produk beserta hitungan aset media
function getProductsInternal() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const prodSheet = ss.getSheetByName('produk');
  const mediaSheet = ss.getSheetByName('media');
  
  const prodData = prodSheet.getDataRange().getValues();
  if (prodData.length <= 1) return [];

  // Hitung jumlah media per produk_id
  const mediaCountMap = {};
  if (mediaSheet) {
    const mediaData = mediaSheet.getDataRange().getValues();
    for (let m = 1; m < mediaData.length; m++) {
      const pId = String(mediaData[m][1]);
      mediaCountMap[pId] = (mediaCountMap[pId] || 0) + 1;
    }
  }

  const products = [];
  for (let i = 1; i < prodData.length; i++) {
    const row = prodData[i];
    const pid = String(row[0]);
    products.push({
      id: pid,
      kode_produk: String(row[1] || ''),
      nama: String(row[2] || ''),
      kategori: String(row[3] || ''),
      cover_url: String(row[4] || ''),
      link_affiliate: String(row[5] || ''),
      link_dashboard: String(row[6] || ''),
      created_at: row[7] instanceof Date ? row[7].toISOString() : String(row[7] || ''),
      media_count: mediaCountMap[pid] || 0
    });
  }
  return products;
}

function getProducts() {
  try {
    return { ok: true, success: true, data: serializeForClient(getProductsInternal()) };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

// Simpan atau perbarui data produk
function saveProduct(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('produk');
    const data = sheet.getDataRange().getValues();
    
    const isEdit = payload.id && payload.id.trim() !== '';
    const nowStr = new Date().toISOString();

    if (isEdit) {
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 2, 1, 6).setValues([[
          payload.kode_produk,
          payload.nama,
          payload.kategori || '',
          payload.cover_url || '',
          payload.link_affiliate || '',
          payload.link_dashboard || ''
        ]]);
        return { ok: true, success: true, message: 'Produk berhasil diubah.' };
      }
    }

    // Tambah Produk Baru
    const newId = 'prod_' + Utilities.getUuid();
    sheet.appendRow([
      newId,
      payload.kode_produk,
      payload.nama,
      payload.kategori || '',
      payload.cover_url || '',
      payload.link_affiliate || '',
      payload.link_dashboard || '',
      nowStr
    ]);

    return { ok: true, success: true, data: { id: newId }, message: 'Produk berhasil ditambahkan.' };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal menyimpan: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// Hapus produk beserta seluruh aset media terkait
function deleteProduct(productId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName('produk');
    const mediaSheet = ss.getSheetByName('media');

    // Hapus dari sheet produk
    const pData = prodSheet.getDataRange().getValues();
    for (let i = pData.length - 1; i >= 1; i--) {
      if (String(pData[i][0]) === String(productId)) {
        prodSheet.deleteRow(i + 1);
        break;
      }
    }

    // Hapus media terkait
    if (mediaSheet) {
      const mData = mediaSheet.getDataRange().getValues();
      for (let j = mData.length - 1; j >= 1; j--) {
        if (String(mData[j][1]) === String(productId)) {
          mediaSheet.deleteRow(j + 1);
        }
      }
    }

    return { ok: true, success: true, message: 'Produk berhasil dihapus.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// Ambil aset media berdasarkan ID produk
function getMediaByProduct(productId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('media');
    if (!sheet) return { ok: true, success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const mediaList = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(productId)) {
        mediaList.push({
          id: String(data[i][0]),
          produk_id: String(data[i][1]),
          batch_id: String(data[i][2]),
          batch_nama: String(data[i][3]),
          media_url: String(data[i][4]),
          media_type: String(data[i][5]),
          content_type: String(data[i][6]),
          catatan: String(data[i][7]),
          created_at: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8])
        });
      }
    }
    return { ok: true, success: true, data: serializeForClient(mediaList) };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

// Simpan batch media yang diunggah dari client ke Cloudinary
function saveMediaBatch(batchPayload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!batchPayload || !batchPayload.length) {
      return { ok: true, success: true, message: 'Tidak ada data untuk disimpan.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('media');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('media');
    }

    const nowStr = new Date().toISOString();
    const rowsToAppend = batchPayload.map(function(item) {
      return [
        'med_' + Utilities.getUuid(),
        item.produk_id,
        item.batch_id,
        item.batch_nama,
        item.media_url,
        item.media_type,
        item.content_type,
        item.catatan || '',
        nowStr
      ];
    });

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 9).setValues(rowsToAppend);

    return { ok: true, success: true, message: 'Batch media tersimpan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// Hapus satuan media
function deleteMedia(mediaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('media');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(mediaId)) {
        sheet.deleteRow(i + 1);
        return { ok: true, success: true, message: 'Media terhapus.' };
      }
    }
    return { ok: false, success: false, message: 'Media tidak ditemukan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// Hapus seluruh batch media
function deleteBatch(batchId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('media');
    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]) === String(batchId)) {
        sheet.deleteRow(i + 1);
      }
    }
    return { ok: true, success: true, message: 'Batch terhapus.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// Pembacaan Pengaturan Internal (Key-Value)
function getSettingsInternal() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('pengaturan');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('pengaturan');
  }

  const data = sheet.getDataRange().getValues();
  const settings = {
    cloudinary_cloud_name: '',
    cloudinary_preset: '',
    cloudinary_api_key: '',
    tracking_prefix: 'DANDAF'
  };

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const val = String(data[i][1]).trim();
    if (key) {
      settings[key] = val;
    }
  }
  return settings;
}

function getSettings() {
  try {
    return { ok: true, success: true, data: serializeForClient(getSettingsInternal()) };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

// Simpan Pengaturan (Key-Value)
function saveSettings(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('pengaturan');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('pengaturan');
    }

    // Bersihkan data lama, tulis ulang header dan pasangan key-value
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 2).setValues([['kunci', 'nilai']]);

    const rows = [];
    for (const key in payload) {
      if (Object.prototype.hasOwnProperty.call(payload, key)) {
        rows.push([key, String(payload[key] || '')]);
      }
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }

    return { ok: true, success: true, message: 'Pengaturan berhasil disimpan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}
