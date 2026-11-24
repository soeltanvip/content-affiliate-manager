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
      headers: ['id', 'kode_produk', 'nama', 'kategori', 'cover_url', 'link_dashboard', 'created_at']
    },
    {
      name: 'media',
      headers: ['id', 'produk_id', 'batch_id', 'batch_nama', 'media_url', 'media_type', 'content_type', 'catatan', 'created_at', 'cloudinary_public_id', 'cloudinary_resource_type']
    },
    {
      name: 'pengaturan',
      headers: ['kunci', 'nilai']
    },
    {
      name: 'akun',
      headers: ['id', 'nama_akun', 'platform', 'prefix', 'batas_harian', 'aktif', 'created_at']
    },
    {
      name: 'postingan',
      headers: ['id', 'produk_id', 'media_ids', 'tipe', 'tanggal_rencana', 'jam_tayang', 'kode_pelacakan', 'link_affiliate', 'caption', 'status', 'url_postingan', 'created_at', 'updated_at', 'akun_id']
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

  // Migrasi pengaturan prefix & kuota lama ke akun tunggal pertama
  migrateSettingsToAccount(ss);

  return { ok: true, success: true, message: 'Database setup completed.' };
}

// Migrasi jika sheet akun kosong dan sheet pengaturan menyimpan tracking_prefix / limit
function migrateSettingsToAccount(ss) {
  const akunSheet = ss.getSheetByName('akun');
  const pengSheet = ss.getSheetByName('pengaturan');
  if (!akunSheet || !pengSheet) return;

  const akunData = akunSheet.getDataRange().getValues();
  if (akunData.length <= 1) {
    const pengData = pengSheet.getDataRange().getValues();
    let prefix = '';
    let batas = 3;
    const remainingRows = [];

    for (let i = 1; i < pengData.length; i++) {
      const k = String(pengData[i][0]).trim();
      const v = String(pengData[i][1]).trim();
      if (k === 'tracking_prefix') {
        prefix = v;
      } else if (k === 'daily_post_limit') {
        batas = parseInt(v, 10) || 3;
      } else if (k) {
        remainingRows.push([k, v]);
      }
    }

    if (prefix || batas) {
      const newAccId = 'acc_' + Utilities.getUuid();
      const nowStr = new Date().toISOString();
      akunSheet.appendRow([
        newAccId,
        'Akun Utama',
        'Facebook Page',
        (prefix || 'DD').toUpperCase().substring(0, 4),
        batas,
        true,
        nowStr
      ]);

      // Tulis ulang pengaturan tanpa tracking_prefix dan daily_post_limit
      pengSheet.clearContents();
      pengSheet.getRange(1, 1, 1, 2).setValues([['kunci', 'nilai']]);
      if (remainingRows.length > 0) {
        pengSheet.getRange(2, 1, remainingRows.length, 2).setValues(remainingRows);
      }
    }
  }
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
    const posts = getPostsInternal();
    const accounts = getAccountsInternal();
    return {
      ok: true,
      success: true,
      data: serializeForClient({
        settings: settings,
        products: products,
        accounts: accounts,
        posts: posts
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
      link_dashboard: String(row[5] || ''),
      created_at: row[6] instanceof Date ? row[6].toISOString() : String(row[6] || ''),
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
        // ponytail: index kolom 2-6 (5 kolom: kode, nama, kategori, cover, dashboard)
        sheet.getRange(rowIndex, 2, 1, 5).setValues([[
          payload.kode_produk,
          payload.nama,
          payload.kategori || '',
          payload.cover_url || '',
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
          created_at: data[i][8] instanceof Date ? data[i][8].toISOString() : String(data[i][8]),
          cloudinary_public_id: String(data[i][9] || ''),
          cloudinary_resource_type: String(data[i][10] || '')
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
        nowStr,
        item.cloudinary_public_id || '',
        item.cloudinary_resource_type || ''
      ];
    });

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 11).setValues(rowsToAppend);

    return { ok: true, success: true, message: 'Batch media tersimpan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// Dapatkan nomor urut batch harian untuk penamaan public_id
function getNextBatchSequence(dateStr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('media');
    let datePart = '';
    if (dateStr) {
      datePart = String(dateStr).replace(/[^0-9]/g, '').slice(2, 8);
    }
    if (datePart.length !== 6) {
      const d = new Date();
      const yy = String(d.getFullYear()).slice(-2);
      const mm = ('0' + (d.getMonth() + 1)).slice(-2);
      const dd = ('0' + d.getDate()).slice(-2);
      datePart = yy + mm + dd;
    }

    let maxSeq = 0;
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const prefixPattern = 'CAM-' + datePart + '-';
      for (let i = 1; i < data.length; i++) {
        const pubId = String(data[i][9] || '');
        const targetStr = pubId.indexOf('/') !== -1 ? pubId.split('/').pop() : pubId;
        if (targetStr.indexOf(prefixPattern) === 0) {
          const parts = targetStr.split('-');
          if (parts.length >= 3) {
            const seq = parseInt(parts[2], 10);
            if (!isNaN(seq) && seq > maxSeq) {
              maxSeq = seq;
            }
          }
        }
      }
    }

    const nextSeq = maxSeq + 1;
    const seqStr = nextSeq < 10 ? '0' + nextSeq : String(nextSeq);
    return {
      ok: true,
      success: true,
      data: {
        batchSeq: seqStr,
        datePart: datePart
      }
    };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal mengambil nomor batch: ' + err.message };
  }
}

// Helper ekstrak public_id dari Cloudinary media URL (untuk media lama)
function extractPublicIdFromUrl(url) {
  if (!url) return '';
  const clean = String(url).split('?')[0];
  const uploadIdx = clean.indexOf('/upload/');
  if (uploadIdx === -1) return '';
  let path = clean.substring(uploadIdx + 8);
  path = path.replace(/^v\d+\//, '');
  const dotIdx = path.lastIndexOf('.');
  if (dotIdx !== -1) {
    path = path.substring(0, dotIdx);
  }
  return path;
}

// Helper penghapusan file di Cloudinary menggunakan SHA-1 Signature
function deleteFromCloudinary(publicId, resourceType) {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiSecret = scriptProps.getProperty('CLOUDINARY_API_SECRET');
  const settings = getSettingsInternal();
  const cloudName = settings.cloudinary_cloud_name;
  const apiKey = settings.cloudinary_api_key;

  if (!apiSecret || !cloudName || !apiKey) {
    throw new Error('Konfigurasi Cloudinary (Cloud Name, API Key, dan API Secret di Pengaturan) belum lengkap.');
  }

  const timestamp = Math.floor(Date.now() / 1000).toString();
  const toSign = 'public_id=' + publicId + '&timestamp=' + timestamp + apiSecret;
  const rawDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, toSign, Utilities.Charset.UTF_8);
  let signature = '';
  for (let i = 0; i < rawDigest.length; i++) {
    let byteVal = rawDigest[i];
    if (byteVal < 0) byteVal += 256;
    let byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = '0' + byteHex;
    signature += byteHex;
  }

  const resType = (resourceType || 'image').toLowerCase() === 'video' ? 'video' : 'image';
  const url = 'https://api.cloudinary.com/v1_1/' + cloudName + '/' + resType + '/destroy';

  const payload = {
    public_id: publicId,
    timestamp: timestamp,
    api_key: apiKey,
    signature: signature
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });

  const resCode = response.getResponseCode();
  const resText = response.getContentText();
  let json = {};
  try {
    json = JSON.parse(resText);
  } catch (e) {
    throw new Error('Gagal memproses respon Cloudinary: ' + resText);
  }

  if (resCode >= 200 && resCode < 300 && (json.result === 'ok' || json.result === 'not found')) {
    return true;
  }
  throw new Error('Cloudinary: ' + (json.error ? json.error.message : resText));
}

// Cek apakah media sedang digunakan di postingan yang akan tayang
function checkMediaUsageInActivePosts(ss, mediaIdsToCheck) {
  const postSheet = ss.getSheetByName('postingan');
  if (!postSheet) return false;
  const data = postSheet.getDataRange().getValues();
  const targetIds = {};
  mediaIdsToCheck.forEach(function(id) { targetIds[id] = true; });

  const activeStatuses = {
    'menunggu persetujuan': true,
    'disetujui': true,
    'terjadwal': true
  };

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][9] || '').toLowerCase().trim();
    if (activeStatuses[status]) {
      const rowMediaIds = String(data[i][2] || '').split(',').map(function(s) { return s.trim(); });
      for (let j = 0; j < rowMediaIds.length; j++) {
        if (targetIds[rowMediaIds[j]]) {
          return true;
        }
      }
    }
  }
  return false;
}

// Hapus satuan media dari Cloudinary lalu hapus baris sheet
function deleteMedia(mediaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('media');
    if (!sheet) return { ok: false, success: false, message: 'Sheet media tidak ditemukan.' };

    if (checkMediaUsageInActivePosts(ss, [mediaId])) {
      return { ok: false, success: false, message: 'Media ini dipakai postingan yang akan tayang.' };
    }

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(mediaId)) {
        rowIndex = i + 1;
        targetRow = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { ok: false, success: false, message: 'Media tidak ditemukan.' };
    }

    let pubId = String(targetRow[9] || '');
    if (!pubId) {
      pubId = extractPublicIdFromUrl(String(targetRow[4] || ''));
    }
    const resType = String(targetRow[10] || (String(targetRow[5] || '').toUpperCase() === 'VIDEO' ? 'video' : 'image'));

    if (pubId) {
      deleteFromCloudinary(pubId, resType);
    }

    sheet.deleteRow(rowIndex);
    return { ok: true, success: true, message: 'Media berhasil dihapus permanen dari Cloudinary dan database.' };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal menghapus media: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// Hapus seluruh batch media dari Cloudinary lalu hapus baris sheet
function deleteBatch(batchId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('media');
    if (!sheet) return { ok: false, success: false, message: 'Sheet media tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    const batchMediaIds = [];
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(batchId)) {
        batchMediaIds.push(String(data[i][0]));
        rowsToDelete.push({
          rowIndex: i + 1,
          pubId: String(data[i][9] || '') || extractPublicIdFromUrl(String(data[i][4] || '')),
          resType: String(data[i][10] || (String(data[i][5] || '').toUpperCase() === 'VIDEO' ? 'video' : 'image'))
        });
      }
    }

    if (rowsToDelete.length === 0) {
      return { ok: true, success: true, message: 'Batch tidak ditemukan atau sudah kosong.' };
    }

    if (checkMediaUsageInActivePosts(ss, batchMediaIds)) {
      return { ok: false, success: false, message: 'Media ini dipakai postingan yang akan tayang.' };
    }

    // Hapus satu per satu dari Cloudinary
    for (let k = 0; k < rowsToDelete.length; k++) {
      const item = rowsToDelete[k];
      if (item.pubId) {
        deleteFromCloudinary(item.pubId, item.resType);
      }
    }

    // Hapus dari sheet mulai dari baris terbawah
    for (let j = rowsToDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(rowsToDelete[j].rowIndex);
    }

    return { ok: true, success: true, message: 'Batch media berhasil dihapus permanen dari Cloudinary dan database.' };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal menghapus batch: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// Pembacaan Pengaturan Internal (Key-Value Khusus Cloudinary)
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
    cloudinary_api_key: ''
  };

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const val = String(data[i][1]).trim();
    if (key && Object.prototype.hasOwnProperty.call(settings, key)) {
      settings[key] = val;
    }
  }

  const scriptProps = PropertiesService.getScriptProperties();
  settings.has_api_secret = !!scriptProps.getProperty('CLOUDINARY_API_SECRET');

  return settings;
}

function getSettings() {
  try {
    return { ok: true, success: true, data: serializeForClient(getSettingsInternal()) };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

// Simpan Pengaturan (Key-Value) dan API Secret terisolasi di Script Properties
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

    // Simpan API Secret ke Script Properties secara terisolasi jika diinput
    if (payload.cloudinary_api_secret && String(payload.cloudinary_api_secret).trim() !== '') {
      PropertiesService.getScriptProperties().setProperty('CLOUDINARY_API_SECRET', String(payload.cloudinary_api_secret).trim());
    }

    // Bersihkan data lama, tulis ulang header dan pasangan key-value Cloudinary publik
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 2).setValues([['kunci', 'nilai']]);

    const allowedKeys = ['cloudinary_cloud_name', 'cloudinary_preset', 'cloudinary_api_key'];
    const rows = [];
    for (let i = 0; i < allowedKeys.length; i++) {
      const k = allowedKeys[i];
      if (payload[k] !== undefined) {
        rows.push([k, String(payload[k] || '')]);
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

// ==============================================
// MODUL AKUN
// ==============================================

function getAccountsInternal() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('akun');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('akun');
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // Hitung penggunaan akun pada postingan
  const postSheet = ss.getSheetByName('postingan');
  const postUsage = {};
  if (postSheet) {
    const postData = postSheet.getDataRange().getValues();
    for (let p = 1; p < postData.length; p++) {
      const aId = String(postData[p][13] || '');
      if (aId) {
        postUsage[aId] = (postUsage[aId] || 0) + 1;
      }
    }
  }

  const accounts = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const accId = String(row[0]);
    accounts.push({
      id: accId,
      nama_akun: String(row[1] || ''),
      platform: String(row[2] || 'Facebook Page'),
      prefix: String(row[3] || '').toUpperCase(),
      batas_harian: parseInt(row[4], 10) || 3,
      aktif: row[5] === true || String(row[5]).toLowerCase() === 'true',
      created_at: row[6] instanceof Date ? row[6].toISOString() : String(row[6] || ''),
      has_posts: (postUsage[accId] || 0) > 0
    });
  }
  return accounts;
}

function getAccounts() {
  try {
    return { ok: true, success: true, data: serializeForClient(getAccountsInternal()) };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

function saveAccount(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('akun');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('akun');
    }

    const data = sheet.getDataRange().getValues();
    const isEdit = payload.id && String(payload.id).trim() !== '';
    const nowStr = new Date().toISOString();

    let cleanPrefix = String(payload.prefix || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!cleanPrefix) {
      return { ok: false, success: false, message: 'Prefix tracking wajib berupa huruf/angka.' };
    }
    if (cleanPrefix.length > 4) {
      cleanPrefix = cleanPrefix.substring(0, 4);
    }

    // Cek keunikan prefix di semua akun
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]);
      const rowPrefix = String(data[i][3] || '').toUpperCase();
      if ((!isEdit || rowId !== payload.id) && rowPrefix === cleanPrefix) {
        return { ok: false, success: false, message: 'Prefix ' + cleanPrefix + ' sudah digunakan akun lain.' };
      }
    }

    if (isEdit) {
      let rowIndex = -1;
      let existingRow = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIndex = i + 1;
          existingRow = data[i];
          break;
        }
      }
      if (rowIndex !== -1) {
        // Cek jika akun sudah dipakai postingan
        const postSheet = ss.getSheetByName('postingan');
        let hasPosts = false;
        if (postSheet) {
          const postData = postSheet.getDataRange().getValues();
          for (let p = 1; p < postData.length; p++) {
            if (String(postData[p][13] || '') === String(payload.id)) {
              hasPosts = true;
              break;
            }
          }
        }

        const existingPrefix = String(existingRow[3] || '').toUpperCase();
        if (hasPosts && existingPrefix !== cleanPrefix) {
          return { ok: false, success: false, message: 'Prefix tidak bisa diubah karena sudah ada postingan yang memakai akun ini.' };
        }

        const aktifVal = payload.aktif !== false && String(payload.aktif) !== 'false';
        sheet.getRange(rowIndex, 2, 1, 5).setValues([[
          payload.nama_akun,
          payload.platform || 'Facebook Page',
          hasPosts ? existingPrefix : cleanPrefix,
          parseInt(payload.batas_harian, 10) || 3,
          aktifVal
        ]]);

        return { ok: true, success: true, message: 'Akun berhasil diperbarui.' };
      }
    }

    // Tambah akun baru
    const newId = 'acc_' + Utilities.getUuid();
    sheet.appendRow([
      newId,
      payload.nama_akun,
      payload.platform || 'Facebook Page',
      cleanPrefix,
      parseInt(payload.batas_harian, 10) || 3,
      payload.aktif !== false && String(payload.aktif) !== 'false',
      nowStr
    ]);

    return { ok: true, success: true, message: 'Akun berhasil ditambahkan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function toggleAccountStatus(accountId, activeStatus) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('akun');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(accountId)) {
        sheet.getRange(i + 1, 6).setValue(activeStatus);
        return { ok: true, success: true, message: 'Status akun berhasil diperbarui.' };
      }
    }
    return { ok: false, success: false, message: 'Akun tidak ditemukan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ==============================================
// MODUL POSTINGAN (BACKEND CRUD & WORKFLOW)
// ==============================================

function getPostsInternal() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('postingan');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('postingan');
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const prodSheet = ss.getSheetByName('produk');
  const prodMap = {};
  if (prodSheet) {
    const pData = prodSheet.getDataRange().getValues();
    for (let p = 1; p < pData.length; p++) {
      prodMap[String(pData[p][0])] = {
        kode: String(pData[p][1] || ''),
        nama: String(pData[p][2] || '')
      };
    }
  }

  const akunSheet = ss.getSheetByName('akun');
  const akunMap = {};
  if (akunSheet) {
    const aData = akunSheet.getDataRange().getValues();
    for (let a = 1; a < aData.length; a++) {
      akunMap[String(aData[a][0])] = {
        nama: String(aData[a][1] || ''),
        prefix: String(aData[a][3] || ''),
        batas: parseInt(aData[a][4], 10) || 3
      };
    }
  }

  const posts = [];
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pid = String(row[1] || '');
    const prodInfo = prodMap[pid] || { kode: '-', nama: 'Produk Tidak Ditemukan' };
    const aid = String(row[13] || '');
    const accInfo = akunMap[aid] || { nama: '-', prefix: '-', batas: 3 };

    let tgl = row[4];
    if (tgl instanceof Date) {
      tgl = Utilities.formatDate(tgl, tz, 'yyyy-MM-dd');
    } else {
      tgl = String(tgl || '');
    }

    let jam = row[5];
    if (jam instanceof Date) {
      jam = Utilities.formatDate(jam, tz, 'HH:mm');
    } else {
      jam = String(jam || '');
    }

    posts.push({
      id: String(row[0] || ''),
      produk_id: pid,
      produk_kode: prodInfo.kode,
      produk_nama: prodInfo.nama,
      akun_id: aid,
      akun_nama: accInfo.nama,
      akun_prefix: accInfo.prefix,
      media_ids: String(row[2] || '').split(',').filter(function(x) { return x.trim() !== ''; }),
      tipe: String(row[3] || 'TEKS'),
      tanggal_rencana: tgl,
      jam_tayang: jam,
      kode_pelacakan: String(row[6] || ''),
      link_affiliate: String(row[7] || ''),
      caption: String(row[8] || ''),
      status: String(row[9] || 'draf'),
      url_postingan: String(row[10] || ''),
      created_at: row[11] instanceof Date ? row[11].toISOString() : String(row[11] || ''),
      updated_at: row[12] instanceof Date ? row[12].toISOString() : String(row[12] || '')
    });
  }
  return posts;
}

function getPosts() {
  try {
    return { ok: true, success: true, data: serializeForClient(getPostsInternal()) };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

function getPostDetail(postId) {
  try {
    const posts = getPostsInternal();
    for (let i = 0; i < posts.length; i++) {
      if (posts[i].id === postId) {
        return { ok: true, success: true, data: serializeForClient(posts[i]) };
      }
    }
    return { ok: false, success: false, message: 'Postingan tidak ditemukan.' };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}

// Mendapatkan nomor urut tracking code berikutnya tanpa menulis ke basis data (berdasarkan akun)
function getNextTrackingCode(akunId, dateStr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('postingan');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('postingan');
    }

    const akunSheet = ss.getSheetByName('akun');
    let prefix = 'DD';
    if (akunSheet && akunId) {
      const aData = akunSheet.getDataRange().getValues();
      for (let a = 1; a < aData.length; a++) {
        if (String(aData[a][0]) === String(akunId)) {
          prefix = String(aData[a][3] || 'DD').toUpperCase();
          break;
        }
      }
    }
    prefix = prefix.replace(/[^A-Z0-9]/g, '');
    if (!prefix) prefix = 'DD';
    if (prefix.length > 4) prefix = prefix.substring(0, 4);

    let datePart = '';
    if (dateStr) {
      datePart = String(dateStr).replace(/[^0-9]/g, '').slice(2, 8);
    }
    if (datePart.length !== 6) {
      const d = new Date();
      const yy = String(d.getFullYear()).slice(-2);
      const mm = ('0' + (d.getMonth() + 1)).slice(-2);
      const dd = ('0' + d.getDate()).slice(-2);
      datePart = yy + mm + dd;
    }

    const tagBase = prefix + datePart;
    const data = sheet.getDataRange().getValues();
    let maxSeq = 0;
    for (let j = 1; j < data.length; j++) {
      const code = String(data[j][6] || '').trim();
      if (code.indexOf(tagBase) === 0) {
        const seqNum = parseInt(code.substring(tagBase.length), 10);
        if (!isNaN(seqNum) && seqNum > maxSeq) {
          maxSeq = seqNum;
        }
      }
    }

    const nextSeq = maxSeq + 1;
    const seqFormatted = nextSeq < 10 ? '0' + nextSeq : String(nextSeq);
    const generatedTrackingCode = tagBase + seqFormatted;

    return {
      ok: true,
      success: true,
      data: {
        code: generatedTrackingCode
      }
    };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal mengambil kode pelacakan: ' + err.message };
  }
}

function savePost(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!payload.produk_id) {
      return { ok: false, success: false, message: 'Pilih produk terlebih dahulu.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('postingan');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('postingan');
    }

    // Validasi Media Terpilih
    const selectedMediaIds = Array.isArray(payload.media_ids) ? payload.media_ids : [];
    let photoCount = 0;
    let videoCount = 0;

    if (selectedMediaIds.length > 0) {
      const mediaSheet = ss.getSheetByName('media');
      const mData = mediaSheet ? mediaSheet.getDataRange().getValues() : [];
      const mTypeMap = {};
      for (let m = 1; m < mData.length; m++) {
        mTypeMap[String(mData[m][0])] = String(mData[m][5] || 'FOTO').toUpperCase();
      }
      for (let k = 0; k < selectedMediaIds.length; k++) {
        const mType = mTypeMap[selectedMediaIds[k]] || 'FOTO';
        if (mType === 'VIDEO') {
          videoCount++;
        } else {
          photoCount++;
        }
      }
    }

    // Aturan Facebook: Jangan campur foto & video
    if (photoCount > 0 && videoCount > 0) {
      return { ok: false, success: false, message: 'Di Facebook, foto dan video harus dipisah menjadi dua postingan.' };
    }

    if (payload.tipe === 'TEKS' && (photoCount > 0 || videoCount > 0)) {
      return { ok: false, success: false, message: 'Tipe postingan teks tidak boleh memiliki media.' };
    }

    if (payload.tipe === 'VIDEO' && (videoCount !== 1 || photoCount !== 0)) {
      return { ok: false, success: false, message: 'Tipe video harus memilih tepat satu berkas video.' };
    }

    if (payload.tipe === 'FOTO' && (photoCount < 1 || photoCount > 10 || videoCount > 0)) {
      return { ok: false, success: false, message: 'Tipe foto harus memilih antara 1 hingga 10 foto.' };
    }

    const data = sheet.getDataRange().getValues();
    const isEdit = payload.id && String(payload.id).trim() !== '';
    const nowStr = new Date().toISOString();
    const mediaIdsJoined = selectedMediaIds.join(',');

    if (isEdit) {
      let rowIndex = -1;
      let existingRow = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIndex = i + 1;
          existingRow = data[i];
          break;
        }
      }

      if (rowIndex !== -1) {
        const existingStatus = String(existingRow[9] || 'draf');
        let finalCaption = payload.caption || '';

        // Kunci caption jika status sudah tayang
        if (existingStatus === 'sudah tayang') {
          finalCaption = String(existingRow[8] || '');
        }

        sheet.getRange(rowIndex, 2, 1, 9).setValues([[
          payload.produk_id,
          mediaIdsJoined,
          payload.tipe,
          payload.tanggal_rencana || '',
          payload.jam_tayang || '',
          existingRow[6], // kode_pelacakan tidak boleh berubah
          payload.link_affiliate || '',
          finalCaption,
          existingStatus
        ]]);
        sheet.getRange(rowIndex, 13).setValue(nowStr);
        if (payload.akun_id) {
          sheet.getRange(rowIndex, 14).setValue(payload.akun_id);
        }

        return {
          ok: true,
          success: true,
          message: 'Postingan berhasil diperbarui.',
          data: { id: payload.id, kode_pelacakan: String(existingRow[6]), status: existingStatus, akun_id: payload.akun_id || existingRow[13] }
        };
      }
    }

    // Simpan Postingan Baru (cek keunikan kode_pelacakan)
    const trackingCode = String(payload.kode_pelacakan || '').trim();
    if (!trackingCode) {
      return { ok: false, success: false, message: 'Kode pelacakan belum dibuat. Silakan klik Tambah Link Affiliate terlebih dahulu.' };
    }

    for (let j = 1; j < data.length; j++) {
      if (String(data[j][6]).trim() === trackingCode) {
        return {
          ok: false,
          success: false,
          message: 'Kode ' + trackingCode + ' sudah dipakai. Tutup form dan buka lagi untuk mendapat kode baru.'
        };
      }
    }

    const newId = 'post_' + Utilities.getUuid();
    sheet.appendRow([
      newId,
      payload.produk_id,
      mediaIdsJoined,
      payload.tipe,
      payload.tanggal_rencana || '',
      payload.jam_tayang || '',
      trackingCode,
      payload.link_affiliate || '',
      payload.caption || '',
      'draf',
      '', // url_postingan
      nowStr,
      nowStr,
      payload.akun_id || ''
    ]);

    return {
      ok: true,
      success: true,
      message: 'Postingan berhasil dibuat.',
      data: {
        id: newId,
        kode_pelacakan: trackingCode,
        akun_id: payload.akun_id || '',
        status: 'draf'
      }
    };
  } catch (err) {
    return { ok: false, success: false, message: 'Gagal menyimpan postingan: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function updatePostStatus(postId, newStatus, extraData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('postingan');
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    let postRow = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(postId)) {
        rowIndex = i + 1;
        postRow = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { ok: false, success: false, message: 'Postingan tidak ditemukan.' };
    }

    let finalStatus = newStatus;
    const nowStr = new Date().toISOString();

    if (newStatus === 'setujui' || newStatus === 'terjadwalkan') {
      const tgl = extraData && extraData.tanggal_rencana ? extraData.tanggal_rencana : String(postRow[4]);
      const jam = extraData && extraData.jam_tayang ? extraData.jam_tayang : String(postRow[5]);

      if (!jam || jam.trim() === '') {
        finalStatus = 'disetujui';
      } else {
        const postAkunId = String(postRow[13] || '');
        let limit = 3;
        const akunSheet = ss.getSheetByName('akun');
        if (akunSheet && postAkunId) {
          const aData = akunSheet.getDataRange().getValues();
          for (let a = 1; a < aData.length; a++) {
            if (String(aData[a][0]) === postAkunId) {
              limit = parseInt(aData[a][4], 10) || 3;
              break;
            }
          }
        }

        let count = 0;
        for (let j = 1; j < data.length; j++) {
          if (j + 1 === rowIndex) continue;
          const st = String(data[j][9] || '');
          const rowTgl = String(data[j][4] || '');
          const rowAkunId = String(data[j][13] || '');
          if ((st === 'terjadwal' || st === 'sudah tayang') && rowTgl.indexOf(tgl) !== -1 && (!postAkunId || rowAkunId === postAkunId)) {
            count++;
          }
        }

        if (count >= limit) {
          return { ok: false, success: false, message: 'Batas kuota postingan harian (' + limit + ') untuk tanggal ' + tgl + ' sudah penuh.' };
        }
        finalStatus = 'terjadwal';
      }

      if (extraData && extraData.jam_tayang) sheet.getRange(rowIndex, 6).setValue(extraData.jam_tayang);
      if (extraData && extraData.tanggal_rencana) sheet.getRange(rowIndex, 5).setValue(extraData.tanggal_rencana);
    } else if (newStatus === 'sudah_tayang') {
      const fbUrl = extraData && extraData.url_postingan ? String(extraData.url_postingan).trim() : '';
      if (!fbUrl) {
        return { ok: false, success: false, message: 'URL postingan Facebook wajib diisi.' };
      }
      sheet.getRange(rowIndex, 11).setValue(fbUrl);
      finalStatus = 'sudah tayang';
    } else if (newStatus === 'draf' || newStatus === 'tolak') {
      finalStatus = 'draf';
    }

    sheet.getRange(rowIndex, 10).setValue(finalStatus);
    sheet.getRange(rowIndex, 13).setValue(nowStr);

    return { ok: true, success: true, message: 'Status berhasil diperbarui.', data: { status: finalStatus } };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function archivePost(postId) {
  return updatePostStatus(postId, 'diarsipkan', {});
}

function getPostingSummaryPerDay(dateStr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('postingan');
    if (!sheet) return { ok: true, success: true, data: { count: 0, limit: 3, available: true } };

    const data = sheet.getDataRange().getValues();
    const settings = getSettingsInternal();
    const limit = parseInt(settings.daily_post_limit, 10) || 3;
    let count = 0;
    const cleanDate = String(dateStr || '').trim();

    for (let i = 1; i < data.length; i++) {
      const st = String(data[i][9] || '');
      const tgl = String(data[i][4] || '');
      if ((st === 'terjadwal' || st === 'sudah tayang') && tgl.indexOf(cleanDate) !== -1) {
        count++;
      }
    }

    return {
      ok: true,
      success: true,
      data: {
        count: count,
        limit: limit,
        available: count < limit
      }
    };
  } catch (err) {
    return { ok: false, success: false, message: err.message };
  }
}
