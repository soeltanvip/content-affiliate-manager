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
      headers: ['id', 'produk_id', 'batch_id', 'batch_nama', 'media_url', 'media_type', 'content_type', 'catatan', 'created_at']
    },
    {
      name: 'pengaturan',
      headers: ['kunci', 'nilai']
    },
    {
      name: 'postingan',
      headers: ['id', 'produk_id', 'media_ids', 'tipe', 'tanggal_rencana', 'jam_tayang', 'kode_pelacakan', 'link_affiliate', 'caption', 'status', 'url_postingan', 'created_at', 'updated_at']
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
    const posts = getPostsInternal();
    return {
      ok: true,
      success: true,
      data: serializeForClient({
        settings: settings,
        products: products,
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
    tracking_prefix: 'DD',
    daily_post_limit: 3
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

  const posts = [];
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pid = String(row[1] || '');
    const prodInfo = prodMap[pid] || { kode: '-', nama: 'Produk Tidak Ditemukan' };

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

// Mendapatkan nomor urut tracking code berikutnya tanpa menulis ke basis data
function getNextTrackingCode(prefixInput, dateStr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('postingan');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('postingan');
    }

    const settings = getSettingsInternal();
    let prefix = String(prefixInput || settings.tracking_prefix || 'DD').toUpperCase().replace(/[^A-Z0-9]/g, '');
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

        return {
          ok: true,
          success: true,
          message: 'Postingan berhasil diperbarui.',
          data: { id: payload.id, kode_pelacakan: String(existingRow[6]), status: existingStatus }
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
      nowStr
    ]);

    return {
      ok: true,
      success: true,
      message: 'Postingan berhasil dibuat.',
      data: {
        id: newId,
        kode_pelacakan: trackingCode,
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
        const settings = getSettingsInternal();
        const limit = parseInt(settings.daily_post_limit, 10) || 3;
        let count = 0;
        for (let j = 1; j < data.length; j++) {
          if (j + 1 === rowIndex) continue;
          const st = String(data[j][9] || '');
          const rowTgl = String(data[j][4] || '');
          if ((st === 'terjadwal' || st === 'sudah tayang') && rowTgl.indexOf(tgl) !== -1) {
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
