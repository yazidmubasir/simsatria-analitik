/**
 * SIM SATRIA - WRITE GATEWAY
 *
 * TUJUAN:
 * User GURU/SISWA/KARYAWAN/WALI_KELAS tetap VIEWER secara fisik pada
 * Spreadsheet/Drive sekolah, tetapi tetap dapat menyimpan data melalui
 * aplikasi. Gateway ini adalah proyek Apps Script TERPISAH yang harus
 * dideploy sebagai:
 *   Execute as: Me (ADMIN_SEKOLAH / akun layanan sekolah)
 *   Who has access: Anyone
 *
 * Jangan menaruh fungsi write gateway ini pada project web utama yang
 * Execute as USER_ACCESSING. Project utama hanya mengirim request terotorisasi.
 *
 * SECURITY:
 * - Hanya operasi yang ada pada ALLOWED_ACTIONS yang diterima.
 * - Spreadsheet ID tidak boleh dikirim bebas dari client; gateway mengambil
 *   spreadsheet berdasarkan NPSN dari request + registry sekolah yang dikontrol.
 * - Setiap request membawa email user, NPSN, role, dan timestamp.
 * - Gateway memverifikasi user pada USERS sekolah sebelum menulis.
 * - Payload dibatasi ukuran dan tidak boleh berisi fungsi/skrip.
 */

const WRITE_GATEWAY_CONFIG = {
  MAX_PAYLOAD_BYTES: 200000,
  MAX_CLOCK_SKEW_MS: 5 * 60 * 1000,
  ACTIVE_STATUS: 'ACTIVE',
  ALLOWED_ROLES: ['GURU', 'WALI_KELAS', 'KARYAWAN', 'SISWA'],
  ALLOWED_ACTIONS: [
    'APPEND_ROW',
    'SET_VALUES',
    'UPLOAD_FILE',
    'CREATE_PDF',
  ],
  // Isi melalui Script Properties gateway, BUKAN hard-code di repository.
  SHARED_SECRET_PROPERTY: 'SIM_SATRIA_WRITE_GATEWAY_SECRET',
};

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      service: 'SIM SATRIA Write Gateway',
      version: '1.0.0',
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents
      ? String(e.postData.contents)
      : '';
    if (!raw) throw new Error('Payload request kosong.');
    if (raw.length > WRITE_GATEWAY_CONFIG.MAX_PAYLOAD_BYTES) {
      throw new Error('Payload terlalu besar.');
    }

    const request = JSON.parse(raw);
    const result = processWriteGatewayRequest_(request);
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({
      success: false,
      error: err && err.message ? err.message : String(err),
    });
  }
}

function processWriteGatewayRequest_(request) {
  validateGatewayRequest_(request);

  const school = getGatewaySchool_(request.npsn);
  if (!school || !school.SPREADSHEET_ID) {
    throw new Error('Sekolah untuk NPSN ' + request.npsn + ' tidak ditemukan.');
  }

  const user = getGatewaySchoolUser_(school.SPREADSHEET_ID, request.email);
  if (!user) throw new Error('Pengguna tidak ditemukan pada USERS sekolah.');

  const role = String(user.ROLE || '').trim().toUpperCase();
  const status = String(user.STATUS || '').trim().toUpperCase();
  if (status !== WRITE_GATEWAY_CONFIG.ACTIVE_STATUS) {
    throw new Error('Akun pengguna tidak aktif.');
  }
  if (role !== String(request.role || '').trim().toUpperCase()) {
    throw new Error('Role request tidak sesuai dengan role pada USERS.');
  }
  if (WRITE_GATEWAY_CONFIG.ALLOWED_ROLES.indexOf(role) < 0) {
    throw new Error('Role tidak diizinkan menggunakan Write Gateway.');
  }

  const action = String(request.action || '').trim().toUpperCase();
  if (WRITE_GATEWAY_CONFIG.ALLOWED_ACTIONS.indexOf(action) < 0) {
    throw new Error('Action Write Gateway tidak diizinkan: ' + action);
  }

  const ss = SpreadsheetApp.openById(String(school.SPREADSHEET_ID));
  const result = executeGatewayAction_(action, request, ss, school, user);

  return Object.assign({
    success: true,
    email: String(request.email).trim().toLowerCase(),
    npsn: String(school.NPSN || request.npsn),
    sekolah: String(school.NAMA_SEKOLAH || ''),
  }, result || {});
}

function executeGatewayAction_(action, request, ss, school, user) {
  switch (action) {
    case 'APPEND_ROW':
      return gatewayAppendRow_(ss, request);
    case 'SET_VALUES':
      return gatewaySetValues_(ss, request);
    case 'UPLOAD_FILE':
      return gatewayUploadFile_(school, user, request);
    case 'CREATE_PDF':
      return gatewayCreatePdf_(school, user, request);
    default:
      throw new Error('Action tidak didukung.');
  }
}

function gatewayAppendRow_(ss, request) {
  const sheet = getGatewaySheet_(ss, request.sheet);
  const row = Array.isArray(request.row) ? request.row : null;
  if (!row || !row.length) throw new Error('ROW untuk APPEND_ROW wajib berupa array.');
  if (row.length > 200) throw new Error('Jumlah kolom terlalu banyak.');
  sheet.appendRow(row.map(sanitizeGatewayValue_));
  return { action: 'APPEND_ROW', sheet: sheet.getName(), rowNumber: sheet.getLastRow() };
}

function gatewaySetValues_(ss, request) {
  const sheet = getGatewaySheet_(ss, request.sheet);
  const values = request.values;
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
    throw new Error('VALUES untuk SET_VALUES wajib berupa array 2 dimensi.');
  }
  const startRow = Number(request.startRow || 1);
  const startCol = Number(request.startCol || 1);
  if (startRow < 1 || startCol < 1) throw new Error('Posisi range tidak valid.');
  const rowCount = values.length;
  const colCount = values[0].length;
  if (!colCount || rowCount > 500 || colCount > 200) throw new Error('Ukuran range terlalu besar.');
  values.forEach(function(row) {
    if (!Array.isArray(row) || row.length !== colCount) throw new Error('Bentuk VALUES tidak konsisten.');
  });
  sheet.getRange(startRow, startCol, rowCount, colCount)
    .setValues(values.map(function(row) { return row.map(sanitizeGatewayValue_); }));
  return { action: 'SET_VALUES', sheet: sheet.getName(), startRow: startRow, startCol: startCol, rows: rowCount, columns: colCount };
}

function gatewayUploadFile_(school, user, request) {
  const folderId = String(request.folderId || school.DRIVE_FOLDER_ID || '').trim();
  if (!folderId) throw new Error('Folder Drive sekolah belum dikonfigurasi.');
  if (!request.file || !request.file.name || !request.file.data) throw new Error('File upload tidak lengkap.');
  if (String(request.file.data).length > 150000) throw new Error('File upload terlalu besar untuk gateway.');

  const root = DriveApp.getFolderById(folderId);
  const moduleName = String(request.module || 'GENERAL').trim().toUpperCase();
  const folder = getOrCreateGatewaySubfolder_(root, moduleName);
  const bytes = Utilities.base64Decode(String(request.file.data));
  const blob = Utilities.newBlob(bytes, String(request.file.mimeType || 'application/octet-stream'), String(request.file.name));
  const file = folder.createFile(blob);

  return {
    action: 'UPLOAD_FILE',
    fileId: file.getId(),
    fileName: file.getName(),
    url: file.getUrl(),
    folderId: folder.getId(),
  };
}

function gatewayCreatePdf_(school, user, request) {
  // PDF gateway generik: menerima HTML yang sudah dibentuk oleh modul.
  // HTML tidak boleh menjalankan JavaScript eksternal.
  const html = String(request.html || '').trim();
  if (!html) throw new Error('HTML PDF kosong.');
  if (html.length > 180000) throw new Error('HTML PDF terlalu besar.');

  const folderId = String(request.folderId || school.DRIVE_FOLDER_ID || '').trim();
  if (!folderId) throw new Error('Folder Drive sekolah belum dikonfigurasi.');

  const safeHtml = html.replace(/<script[\s\S]*?<\/script>/gi, '');
  const blob = Utilities.newBlob(safeHtml, 'text/html', 'satria-pdf-source.html');
  const tempFile = DriveApp.createFile(blob);
  try {
    const pdfBlob = tempFile.getBlob().getAs(MimeType.PDF).setName(String(request.fileName || 'SIM-SATRIA.pdf'));
    const folder = DriveApp.getFolderById(folderId);
    const pdf = folder.createFile(pdfBlob);
    return {
      action: 'CREATE_PDF',
      fileId: pdf.getId(),
      fileName: pdf.getName(),
      url: pdf.getUrl(),
      folderId: folder.getId(),
    };
  } finally {
    try { tempFile.setTrashed(true); } catch (ignore) {}
  }
}

function validateGatewayRequest_(request) {
  if (!request || typeof request !== 'object') throw new Error('Request tidak valid.');
  const email = String(request.email || '').trim().toLowerCase();
  const npsn = String(request.npsn || '').trim();
  const role = String(request.role || '').trim().toUpperCase();
  const timestamp = Number(request.timestamp || 0);
  const secret = String(request.secret || '');
  const expectedSecret = PropertiesService.getScriptProperties().getProperty(WRITE_GATEWAY_CONFIG.SHARED_SECRET_PROPERTY) || '';

  if (!expectedSecret) throw new Error('Write Gateway belum dikonfigurasi: secret belum diisi.');
  if (secret !== expectedSecret) throw new Error('Gateway authentication gagal.');
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email request tidak valid.');
  if (!npspSafe_(npsp)) throw new Error('NPSN request tidak valid.');
  if (WRITE_GATEWAY_CONFIG.ALLOWED_ROLES.indexOf(role) < 0) throw new Error('Role request tidak diizinkan.');
  if (!timestamp || Math.abs(Date.now() - timestamp) > WRITE_GATEWAY_CONFIG.MAX_CLOCK_SKEW_MS) throw new Error('Request sudah kedaluwarsa.');
}

function npspSafe_(npsn) { return /^[0-9A-Za-z._-]{4,32}$/.test(String(npsn || '')); }

function getGatewaySchool_(npsn) {
  const masterId = String(PropertiesService.getScriptProperties().getProperty('SIM_SATRIA_MASTER_SPREADSHEET_ID') || '').trim();
  if (!masterId) throw new Error('Master Spreadsheet ID belum dikonfigurasi pada Gateway.');
  const ss = SpreadsheetApp.openById(masterId);
  const sheet = ss.getSheetByName('SCHOOLS');
  if (!sheet) throw new Error('Sheet SCHOOLS tidak ditemukan pada MASTER.');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
  const idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.NPSN] || '').trim() === String(npsn).trim()) {
      return {
        NPSN: String(values[i][idx.NPSN] || '').trim(),
        NAMA_SEKOLAH: String(values[i][idx.NAMA_SEKOLAH] || '').trim(),
        SPREADSHEET_ID: String(values[i][idx.SPREADSHEET_ID] || '').trim(),
        DRIVE_FOLDER_ID: String(values[i][idx.DRIVE_FOLDER_ID] || '').trim(),
      };
    }
  }
  return null;
}

function getGatewaySchoolUser_(spreadsheetId, email) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('USERS');
  if (!sheet) throw new Error('Sheet USERS tidak ditemukan pada Spreadsheet sekolah.');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
  const idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  const target = String(email).trim().toLowerCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.EMAIL] || '').trim().toLowerCase() === target) {
      return {
        USER_ID: values[i][idx.USER_ID],
        EMAIL: values[i][idx.EMAIL],
        NIP: values[i][idx.NIP],
        NAMA: values[i][idx.NAMA],
        ROLE: values[i][idx.ROLE],
        STATUS: values[i][idx.STATUS],
      };
    }
  }
  return null;
}

function getGatewaySheet_(ss, sheetName) {
  const name = String(sheetName || '').trim().toUpperCase();
  if (!name || !/^[A-Z0-9_ -]{1,80}$/.test(name)) throw new Error('Nama sheet tidak valid.');
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet ' + name + ' tidak ditemukan pada sekolah.');
  return sheet;
}

function getOrCreateGatewaySubfolder_(root, name) {
  const safeName = String(name || 'GENERAL').trim().replace(/[^A-Z0-9_ -]/gi, '_').substring(0, 80) || 'GENERAL';
  const found = root.getFoldersByName(safeName);
  return found.hasNext() ? found.next() : root.createFolder(safeName);
}

function sanitizeGatewayValue_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' && value.length > 50000) throw new Error('Nilai terlalu panjang.');
  return value;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
