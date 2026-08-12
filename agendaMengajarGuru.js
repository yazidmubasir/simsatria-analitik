/* =========================================================
   SIM SATRIA - AGENDA MENGAJAR GURU
   SERVER SIDE / MULTI SCHOOL
   ========================================================= */

const AGENDA_GURU_CONFIG = {
  SHEET_GURU: 'GURU',
  SHEET_KELAS: 'KELAS',
  SHEET_TRX: 'TRX_AGENDA_GURU',
  SHEET_LOG: 'LOG',
  FOLDER_AGENDA: 'AGENDA'
};

/* =========================================================
   SCHOOL CONTEXT
   ========================================================= */
function getAgendaSchoolContext_() {
  let login = {};
  let context = {};

  if (typeof getLoginInfo === 'function') {
    login = getLoginInfo() || {};
    if (login.success === false) {
      throw new Error(login.message || login.error || 'Autentikasi gagal.');
    }
  }

  if (typeof getCurrentUserContext === 'function') {
    context = getCurrentUserContext() || {};
  } else if (typeof getSchoolContextInfo === 'function') {
    context = getSchoolContextInfo() || {};
  } else if (typeof getCurrentSchoolContext === 'function') {
    context = getCurrentSchoolContext() || {};
  } else if (typeof getMySchoolContext === 'function') {
    context = getMySchoolContext() || {};
  } else if (typeof getSchoolContext === 'function') {
    context = getSchoolContext() || {};
  }

  if (context.success === false) {
    throw new Error(context.message || context.error || 'School Context tidak tersedia.');
  }

  const data = context.data || context.context || context.school ? (context.data || context.context || context) : context;
  const school = data.school || context.school || login.school || {};
  const user = data.user || context.user || login.user || {};

  let spreadsheetId =
    data.spreadsheetId || data.spreadsheetID || data.spreadsheet_id ||
    data.idSpreadsheet || data.idSpreadsheetSekolah || data.databaseSpreadsheetId ||
    data.databaseId || school.spreadsheetId || school.spreadsheetID ||
    school.spreadsheet_id || school.idSpreadsheet || context.spreadsheetId || '';

  if (spreadsheetId && typeof spreadsheetId === 'object') {
    spreadsheetId = spreadsheetId.id || spreadsheetId.spreadsheetId || spreadsheetId.ID || '';
  }

  spreadsheetId = String(spreadsheetId || '').trim();
  if (!spreadsheetId) {
    throw new Error(
      'Spreadsheet sekolah aktif belum tersedia pada School Context. ' +
      'Pastikan database sekolah sudah terdaftar.'
    );
  }

  const npsn = String(
    data.npsn || data.NPSN || school.npsn || school.NPSN || login.npsn || ''
  ).trim();

  const namaSekolah = String(
    data.namaSekolah || data.nama_sekolah || data.sekolah ||
    school.namaSekolah || school.nama_sekolah || school.sekolah ||
    login.namaSekolah || login.sekolah || login.schoolName || ''
  ).trim();

  const spreadsheetName = String(
    data.spreadsheetName || data.namaSpreadsheet || data.namaDatabase ||
    school.spreadsheetName || school.namaSpreadsheet || ''
  ).trim();

  const userId = String(
    user.userId || user.id || data.userId || data.user_id ||
    context.userId || login.userId || login.id || ''
  ).trim();

  const email = String(
    user.email || data.email || data.userEmail || context.email || login.email || ''
  ).trim().toLowerCase();

  const namaUser = String(
    user.nama || user.name || data.namaUser || data.userName ||
    context.namaUser || login.nama || login.name || ''
  ).trim();

  const role = String(
    user.role || data.role || context.role || login.role || ''
  ).trim().toUpperCase();

  const nip = String(
    user.nip || data.nip || data.NIP || context.nip || login.nip || ''
  ).trim();

  return {
    success: true,
    npsn: npsn,
    namaSekolah: namaSekolah,
    spreadsheetId: spreadsheetId,
    spreadsheetName: spreadsheetName,
    user: {
      userId: userId,
      email: email,
      nama: namaUser,
      role: role,
      nip: nip
    }
  };
}

function getAgendaSchoolSS_() {
  return SpreadsheetApp.openById(getAgendaSchoolContext_().spreadsheetId);
}

function getAgendaSheet_(name, required) {
  const sh = getAgendaSchoolSS_().getSheetByName(name);
  if (!sh && required !== false) {
    throw new Error('Sheet "' + name + '" tidak ditemukan pada Spreadsheet sekolah aktif.');
  }
  return sh;
}

/* =========================================================
   HEADER UTILITIES
   ========================================================= */
function agendaNormalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9]/g, '');
}

function agendaFindHeader_(headers, names) {
  const normalized = headers.map(agendaNormalizeHeader_);
  for (let i = 0; i < names.length; i++) {
    const idx = normalized.indexOf(agendaNormalizeHeader_(names[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function agendaEnsureHeadersNonDestructive_(sheet, requiredHeaders) {
  let lastCol = sheet.getLastColumn();

  if (sheet.getLastRow() === 0 || lastCol === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return requiredHeaders.slice();
  }

  let headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  const normalized = headers.map(agendaNormalizeHeader_);
  const missing = [];

  requiredHeaders.forEach(function(header) {
    if (normalized.indexOf(agendaNormalizeHeader_(header)) === -1) missing.push(header);
  });

  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }

  sheet.setFrozenRows(1);
  return headers;
}

function agendaRequirePermission_() {
  if (typeof requirePermission === 'function') requirePermission('MANAGE_AGENDA');
}

/* =========================================================
   GET GURU
   ========================================================= */
function getGuruAgenda() {
  agendaRequirePermission_();
  const ctx = getAgendaSchoolContext_();
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_GURU, true);

  if (sh.getLastRow() < 2) {
    return {
      success: true,
      school: { npsn: ctx.npsn, namaSekolah: ctx.namaSekolah, spreadsheetName: ctx.spreadsheetName },
      user: ctx.user,
      data: []
    };
  }

  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0] || [];

  let nipIndex = agendaFindHeader_(headers, ['nip']);
  let namaIndex = agendaFindHeader_(headers, ['nama', 'nama guru']);
  let mapelIndex = agendaFindHeader_(headers, ['mapel', 'mata pelajaran']);
  let emailIndex = agendaFindHeader_(headers, ['email']);
  let statusIndex = agendaFindHeader_(headers, ['status']);

  if (nipIndex === -1) nipIndex = 1;
  if (namaIndex === -1) namaIndex = 2;
  if (mapelIndex === -1) mapelIndex = 3;

  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const nip = String(row[nipIndex] || '').trim();
    const nama = String(row[namaIndex] || '').trim();
    if (!nip && !nama) continue;

    result.push({
      nip: nip,
      nama: nama,
      mapel: String(row[mapelIndex] || '').trim(),
      email: emailIndex > -1 ? String(row[emailIndex] || '').trim().toLowerCase() : '',
      status: statusIndex > -1 ? String(row[statusIndex] || '').trim() : ''
    });
  }

  return {
    success: true,
    school: { npsn: ctx.npsn, namaSekolah: ctx.namaSekolah, spreadsheetName: ctx.spreadsheetName },
    user: ctx.user,
    data: result
  };
}

/* =========================================================
   GET KELAS
   ========================================================= */
function getKelasAgendaGuru() {
  agendaRequirePermission_();
  const ctx = getAgendaSchoolContext_();
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_KELAS, false);

  if (!sh || sh.getLastRow() < 2) {
    return { success: true, sekolah: ctx.namaSekolah, npsn: ctx.npsn, kelas: [] };
  }

  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  let kelasIndex = agendaFindHeader_(headers, ['kelas', 'nama kelas', 'rombel', 'nama rombel']);
  if (kelasIndex === -1) kelasIndex = 0;

  const seen = {};
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const kelas = String(values[i][kelasIndex] || '').trim();
    if (!kelas) continue;
    const key = kelas.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      result.push(kelas);
    }
  }

  result.sort(function(a, b) {
    return a.localeCompare(b, 'id', { numeric: true, sensitivity: 'base' });
  });

  return { success: true, sekolah: ctx.namaSekolah, npsn: ctx.npsn, kelas: result };
}

/* =========================================================
   UPLOAD FOTO
   Semua operasi Drive dipusatkan di DriveService.js.
   ========================================================= */
function uploadFileAgenda(base64, fileName) {
  agendaRequirePermission_();

  if (!base64 || !fileName) {
    throw new Error('Data foto dan nama file wajib diisi.');
  }

  if (typeof uploadFileToSchoolDrive !== 'function') {
    throw new Error('DriveService.js belum tersedia. Fungsi uploadFileToSchoolDrive() tidak ditemukan.');
  }

  try {
    return uploadFileToSchoolDrive(
      base64,
      fileName,
      AGENDA_GURU_CONFIG.FOLDER_AGENDA
    );
  } catch (e) {
    throw new Error('Upload foto Agenda gagal: ' + (e && e.message ? e.message : String(e)));
  }
}

/* =========================================================
   SIMPAN TRANSAKSI
   ========================================================= */
function simpanAgendaGuru(data) {
  agendaRequirePermission_();

  if (!data || typeof data !== 'object') {
    throw new Error('Data agenda tidak valid.');
  }

  const ctx = getAgendaSchoolContext_();
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_TRX, true);
  const transactionId = 'AGD-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();

  const requiredHeaders = [
    'TRANSACTION_ID','TIMESTAMP','NPSN','USER_ID','EMAIL','NIP','NAMA_USER','ROLE',
    'NAMA_GURU','MAPEL','TANGGAL','SESI','KELAS','TUJUAN_PEMBELAJARAN','DPL',
    'PENGALAMAN_BELAJAR','PRINSIP_PEMBELAJARAN','REKAP_MURID_TIDAK_IKUT',
    'MATERI_PEMBELAJARAN','KETERANGAN','BUKTI_FISIK'
  ];

  const headers = agendaEnsureHeadersNonDestructive_(sh, requiredHeaders);
  const row = new Array(headers.length).fill('');

  const tujuanPembelajaran = String(data.tujuanPembelajaran ?? data.tujuan ?? '').trim();
  const materiPembelajaran = String(data.materiPembelajaran ?? data.materi ?? '').trim();
  const pengalamanBelajar = String(data.pengalamanBelajar ?? data.pm ?? '').trim();
  const prinsipPembelajaran = String(data.prinsipPembelajaran ?? data.prinsip ?? '').trim();
  const rekapMuridTidakIkut = String(data.rekapMuridTidakIkut ?? data.siswaTidakMasuk ?? '').trim();
  const buktiFisik = String(data.buktiFisik ?? data.foto ?? '').trim();

  const values = {
    transactionid: transactionId,
    timestamp: new Date(),
    npsn: ctx.npsn,
    userid: ctx.user.userId,
    email: ctx.user.email,
    nip: String(data.nip || '').trim(),
    namauser: ctx.user.nama,
    role: ctx.user.role,
    namaguru: String(data.namaGuru || '').trim(),
    mapel: String(data.mapel || '').trim(),
    tanggal: String(data.tanggal || '').trim(),
    sesi: String(data.sesi || '').trim(),
    kelas: String(data.kelas || '').trim(),
    tujuanpembelajaran: tujuanPembelajaran,
    materipembelajaran: materiPembelajaran,
    pengalamanbelajar: pengalamanBelajar,
    prinsippembelajaran: prinsipPembelajaran,
    rekapmuridtidakikut: rekapMuridTidakIkut,
    buktifisik: buktiFisik,
    tujuan: tujuanPembelajaran,
    materi: materiPembelajaran,
    pm: pengalamanBelajar,
    prinsip: prinsipPembelajaran,
    siswatidakmasuk: rekapMuridTidakIkut,
    foto: buktiFisik,
    dpl: String(data.dpl || '').trim(),
    keterangan: String(data.keterangan || '').trim()
  };

  headers.forEach(function(header, index) {
    const key = agendaNormalizeHeader_(header);
    if (Object.prototype.hasOwnProperty.call(values, key)) row[index] = values[key];
  });

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);

  agendaWriteLog_(ctx, {
    action: 'SIMPAN',
    module: 'AGENDA_MENGAJAR_GURU',
    description: 'Agenda mengajar ' + values.namaguru + ' - ' + values.kelas + ' - ' + values.tanggal,
    transactionId: transactionId
  });

  return {
    success: true,
    transactionId: transactionId,
    npsn: ctx.npsn,
    sekolah: ctx.namaSekolah,
    message: 'Agenda Satria Mengajar berhasil disimpan.'
  };
}

/* =========================================================
   LOG
   ========================================================= */
function agendaWriteLog_(ctx, data) {
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_LOG, true);
  const headers = agendaEnsureHeadersNonDestructive_(sh, [
    'TIMESTAMP','NPSN','USER_ID','EMAIL','NIP','NAMA_USER','ROLE',
    'ACTION','MODULE','DESCRIPTION','TRANSACTION_ID'
  ]);

  const values = {
    timestamp: new Date(),
    npsn: ctx.npsn,
    userid: ctx.user.userId,
    email: ctx.user.email,
    nip: ctx.user.nip,
    namauser: ctx.user.nama,
    role: ctx.user.role,
    action: data.action || '',
    module: data.module || '',
    description: data.description || '',
    transactionid: data.transactionId || ''
  };

  const row = new Array(headers.length).fill('');
  headers.forEach(function(header, index) {
    const key = agendaNormalizeHeader_(header);
    if (Object.prototype.hasOwnProperty.call(values, key)) row[index] = values[key];
  });

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

/* =========================================================
   TEST MODULE
   ========================================================= */
function testAgendaGuruMultiSchool() {
  const ctx = getAgendaSchoolContext_();
  const ss = getAgendaSchoolSS_();

  return {
    success: true,
    module: 'AGENDA_MENGAJAR_GURU',
    npsn: ctx.npsn,
    sekolah: ctx.namaSekolah,
    spreadsheetId: ctx.spreadsheetId,
    spreadsheetName: ss.getName(),
    userId: ctx.user.userId,
    email: ctx.user.email,
    namaUser: ctx.user.nama,
    role: ctx.user.role,
    guruSheet: !!ss.getSheetByName(AGENDA_GURU_CONFIG.SHEET_GURU),
    kelasSheet: !!ss.getSheetByName(AGENDA_GURU_CONFIG.SHEET_KELAS),
    trxSheet: !!ss.getSheetByName(AGENDA_GURU_CONFIG.SHEET_TRX),
    logSheet: !!ss.getSheetByName(AGENDA_GURU_CONFIG.SHEET_LOG)
  };
}
