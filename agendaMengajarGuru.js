/* =========================================================
   SIM SATRIA - AGENDA MENGAJAR GURU
   MULTI SCHOOL / SCHOOL CONTEXT
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
   Tidak ada Spreadsheet ID / Folder ID hard-coded.
   Semua resource mengikuti sekolah yang sedang login.
   ========================================================= */
function getAgendaSchoolContext_() {
  /*
   * PENTING:
   * getLoginInfo() hanya menyediakan identitas user/sekolah pada
   * beberapa versi SIM SATRIA. SpreadsheetId/folderId berada pada
   * School Context. Karena itu modul ini membaca School Context terlebih
   * dahulu, lalu menggunakan getLoginInfo() sebagai fallback identitas.
   */

  let login = null;
  let context = null;

  if (typeof getLoginInfo === 'function') {
    login = getLoginInfo();
    if (login && login.success === false) {
      throw new Error(
        login.message || login.error || 'Autentikasi gagal.'
      );
    }
  }

  /* =====================================================
     AMBIL SCHOOL CONTEXT DARI ENGINE UTAMA
     ===================================================== */
  if (typeof getSchoolContextInfo === 'function') {
    context = getSchoolContextInfo();
  } else if (typeof getCurrentSchoolContext === 'function') {
    context = getCurrentSchoolContext();
  } else if (typeof getMySchoolContext === 'function') {
    context = getMySchoolContext();
  } else if (typeof getSchoolContext === 'function') {
    context = getSchoolContext();
  }

  if (context && context.success === false) {
    throw new Error(
      context.message ||
      context.error ||
      'School Context sekolah aktif tidak tersedia.'
    );
  }

  context = context || {};
  login = login || {};

  /* Beberapa implementasi mengembalikan context di dalam data/context/school. */
  const contextData =
    context.data ||
    context.context ||
    context.school ||
    context;

  const loginSchool = login.school || {};
  const loginUser = login.user || {};
  const contextSchool = contextData.school || {};
  const contextUser = contextData.user || {};

  /* =====================================================
     NORMALISASI SPREADSHEET ID
     Mendukung beberapa nama field yang dipakai engine.
     ===================================================== */
  let spreadsheetId =
    contextData.spreadsheetId ||
    contextData.spreadsheetID ||
    contextData.spreadsheet_id ||
    contextData.idSpreadsheet ||
    contextData.idSpreadsheetSekolah ||
    contextData.databaseSpreadsheetId ||
    contextData.databaseId ||
    contextSchool.spreadsheetId ||
    contextSchool.spreadsheetID ||
    contextSchool.spreadsheet_id ||
    contextSchool.idSpreadsheet ||
    context.spreadsheetId ||
    context.spreadsheetID ||
    '';

  /* spreadsheet kadang berupa object {id, name}. */
  if (spreadsheetId && typeof spreadsheetId === 'object') {
    spreadsheetId =
      spreadsheetId.id ||
      spreadsheetId.spreadsheetId ||
      spreadsheetId.ID ||
      '';
  }

  spreadsheetId = String(spreadsheetId || '').trim();

  /* =====================================================
     NORMALISASI FOLDER DRIVE
     ===================================================== */
  let folderId =
    contextData.folderId ||
    contextData.driveFolderId ||
    contextData.folder_drive_id ||
    contextData.rootFolderId ||
    contextData.idFolder ||
    contextData.driveId ||
    contextSchool.folderId ||
    contextSchool.driveFolderId ||
    contextSchool.folder_drive_id ||
    context.folderId ||
    context.driveFolderId ||
    '';

  if (folderId && typeof folderId === 'object') {
    folderId =
      folderId.id ||
      folderId.folderId ||
      folderId.ID ||
      '';
  }

  folderId = String(folderId || '').trim();

  /* =====================================================
     IDENTITAS SEKOLAH
     ===================================================== */
  const npsn = String(
    contextData.npsn ||
    contextData.NPSN ||
    contextSchool.npsn ||
    contextSchool.NPSN ||
    loginSchool.npsn ||
    login.npsn ||
    ''
  ).trim();

  const namaSekolah = String(
    contextData.namaSekolah ||
    contextData.nama_sekolah ||
    contextData.sekolah ||
    contextSchool.namaSekolah ||
    contextSchool.nama_sekolah ||
    contextSchool.sekolah ||
    loginSchool.namaSekolah ||
    loginSchool.sekolah ||
    login.sekolah ||
    ''
  ).trim();

  const spreadsheetName = String(
    contextData.spreadsheetName ||
    contextData.namaSpreadsheet ||
    contextData.namaDatabase ||
    contextSchool.spreadsheetName ||
    contextSchool.namaSpreadsheet ||
    ''
  ).trim();

  /* =====================================================
     USER CONTEXT
     ===================================================== */
  const userId = String(
    contextUser.userId ||
    contextUser.id ||
    contextData.userId ||
    contextData.user_id ||
    context.userId ||
    loginUser.userId ||
    loginUser.id ||
    login.userId ||
    ''
  ).trim();

  const email = String(
    contextUser.email ||
    contextData.email ||
    contextData.userEmail ||
    context.email ||
    loginUser.email ||
    login.email ||
    ''
  ).trim();

  const namaUser = String(
    contextUser.nama ||
    contextUser.name ||
    contextData.namaUser ||
    contextData.userName ||
    context.namaUser ||
    loginUser.nama ||
    loginUser.name ||
    login.nama ||
    ''
  ).trim();

  const role = String(
    contextUser.role ||
    contextData.role ||
    context.role ||
    loginUser.role ||
    login.role ||
    ''
  ).trim();

  const nip = String(
    contextUser.nip ||
    contextData.nip ||
    contextData.NIP ||
    context.nip ||
    loginUser.nip ||
    login.nip ||
    ''
  ).trim();

  /*
   * Jangan lagi langsung gagal hanya karena getLoginInfo() tidak
   * mengembalikan spreadsheetId. Yang menjadi sumber resource adalah
   * School Context engine.
   */
  if (!spreadsheetId) {
    throw new Error(
      'Spreadsheet sekolah aktif belum tersedia pada School Context. ' +
      'Pastikan School Context sekolah aktif sudah di-refresh/tersedia ' +
      'dan field spreadsheetId sudah terdaftar.'
    );
  }

  return {
    success: true,
    npsn: npsn,
    namaSekolah: namaSekolah,
    spreadsheetId: spreadsheetId,
    spreadsheetName: spreadsheetName,
    folderId: folderId,
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
  const ctx = getAgendaSchoolContext_();
  return SpreadsheetApp.openById(ctx.spreadsheetId);
}

function getAgendaSheet_(name, required) {
  const ss = getAgendaSchoolSS_();
  const sh = ss.getSheetByName(name);

  if (!sh && required !== false) {
    throw new Error(
      'Sheet "' + name + '" tidak ditemukan pada Spreadsheet sekolah aktif.'
    );
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
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return requiredHeaders.slice();
  }

  let headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(function(h) { return String(h || '').trim(); });

  const normalized = headers.map(agendaNormalizeHeader_);
  const missing = [];

  requiredHeaders.forEach(function(header) {
    if (normalized.indexOf(agendaNormalizeHeader_(header)) === -1) {
      missing.push(header);
    }
  });

  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length)
      .setValues([missing]);
    headers = headers.concat(missing);
  }

  sheet.setFrozenRows(1);
  return headers;
}

/* =========================================================
   GET GURU
   ========================================================= */
function getGuruAgenda() {
  const ctx = getAgendaSchoolContext_();
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_GURU, true);

  if (sh.getLastRow() < 2) {
    return {
      success: true,
      school: {
        npsn: ctx.npsn,
        namaSekolah: ctx.namaSekolah,
        spreadsheetName: ctx.spreadsheetName
      },
      user: ctx.user,
      data: []
    };
  }

  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0] || [];

  let nipIndex = agendaFindHeader_(headers, ['nip', 'NIP']);
  let namaIndex = agendaFindHeader_(headers, ['nama', 'nama guru', 'NAMA']);
  let mapelIndex = agendaFindHeader_(headers, ['mapel', 'mata pelajaran', 'Mata Pelajaran']);
  let emailIndex = agendaFindHeader_(headers, ['email', 'Email']);
  let statusIndex = agendaFindHeader_(headers, ['status', 'Status']);

  // Kompatibilitas dengan struktur lama: B,C,D.
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
      email: emailIndex > -1 ? String(row[emailIndex] || '').trim() : '',
      status: statusIndex > -1 ? String(row[statusIndex] || '').trim() : ''
    });
  }

  return {
    success: true,
    school: {
      npsn: ctx.npsn,
      namaSekolah: ctx.namaSekolah,
      spreadsheetName: ctx.spreadsheetName
    },
    user: ctx.user,
    data: result
  };
}

/* =========================================================
   GET KELAS
   Mengambil dari sheet Kelas sekolah aktif.
   Tidak lagi hard-code X.1 dst.
   ========================================================= */
function getKelasAgendaGuru() {
  const ctx = getAgendaSchoolContext_();
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_KELAS, false);

  if (!sh || sh.getLastRow() < 2) {
    return {
      success: true,
      sekolah: ctx.namaSekolah,
      npsn: ctx.npsn,
      kelas: []
    };
  }

  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0] || [];

  let kelasIndex = agendaFindHeader_(headers, [
    'kelas', 'nama kelas', 'rombel', 'nama rombel'
  ]);

  if (kelasIndex === -1) {
    // Kompatibilitas: jika sheet Kelas hanya memiliki satu kolom data.
    kelasIndex = 0;
  }

  const map = {};
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const kelas = String(values[i][kelasIndex] || '').trim();
    if (!kelas) continue;
    const key = kelas.toLowerCase();
    if (!map[key]) {
      map[key] = true;
      result.push(kelas);
    }
  }

  result.sort(function(a, b) {
    return a.localeCompare(b, 'id', { numeric: true, sensitivity: 'base' });
  });

  return {
    success: true,
    sekolah: ctx.namaSekolah,
    npsn: ctx.npsn,
    kelas: result
  };
}

/* =========================================================
   UPLOAD FOTO
   Folder mengikuti sekolah aktif.
   Jika folder AGENDA tersedia di folder sekolah, gunakan.
   Jika tidak tersedia, gunakan folder sekolah utama.
   Tidak ada folder ID hard-code.
   ========================================================= */
function uploadFileAgenda(base64, fileName) {
  const ctx = getAgendaSchoolContext_();

  if (!ctx.folderId) {
    throw new Error('Folder Drive sekolah aktif belum tersedia pada School Context.');
  }

  const root = DriveApp.getFolderById(ctx.folderId);
  let folder = null;
  const folders = root.getFoldersByName(AGENDA_GURU_CONFIG.FOLDER_AGENDA);

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = root.createFolder(AGENDA_GURU_CONFIG.FOLDER_AGENDA);
  }

  const match = String(base64 || '').match(/^data:(.*);base64,/);
  if (!match) {
    throw new Error('Format file upload tidak valid.');
  }

  const mimeType = match[1];
  const bytes = Utilities.base64Decode(String(base64).split(',')[1]);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);

  return {
    success: true,
    url: file.getUrl(),
    fileId: file.getId(),
    fileName: file.getName(),
    folder: folder.getName()
  };
}

/* =========================================================
   TRANSACTION / LOG
   TRX tetap berada di Spreadsheet sekolah aktif.
   Metadata multi-school selalu ikut disimpan bila header tersedia.
   ========================================================= */
function simpanAgendaGuru(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Data agenda tidak valid.');
  }

  const ctx = getAgendaSchoolContext_();
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_TRX, true);

  const transactionId =
    'AGD-' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();

  const requiredHeaders = [
    'TRANSACTION_ID',
    'TIMESTAMP',
    'NPSN',
    'USER_ID',
    'EMAIL',
    'NIP',
    'NAMA_USER',
    'ROLE',
    'NAMA_GURU',
    'MAPEL',
    'TANGGAL',
    'SESI',
    'KELAS',
    'TUJUAN_PEMBELAJARAN',
    'DPL',
    'PENGALAMAN_BELAJAR',
    'PRINSIP_PEMBELAJARAN',
    'REKAP_MURID_TIDAK_IKUT',
    'MATERI_PEMBELAJARAN',
    'KETERANGAN',
    'BUKTI_FISIK'
  ];

  const headers = agendaEnsureHeadersNonDestructive_(sh, requiredHeaders);
  const row = new Array(headers.length).fill('');

  /*
   * Gunakan nama field yang SAMA dengan header TRX.
   * Alias lama tetap disediakan agar kompatibel dengan data/module lama.
   */
  const tujuanPembelajaran = String(
    data.tujuanPembelajaran ?? data.tujuan ?? ''
  ).trim();

  const materiPembelajaran = String(
    data.materiPembelajaran ?? data.materi ?? ''
  ).trim();

  const pengalamanBelajar = String(
    data.pengalamanBelajar ?? data.pm ?? ''
  ).trim();

  const prinsipPembelajaran = String(
    data.prinsipPembelajaran ?? data.prinsip ?? ''
  ).trim();

  const rekapMuridTidakIkut = String(
    data.rekapMuridTidakIkut ?? data.siswaTidakMasuk ?? ''
  ).trim();

  const buktiFisik = String(
    data.buktiFisik ?? data.foto ?? ''
  ).trim();

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

    /* HEADER FINAL */
    tujuanpembelajaran: tujuanPembelajaran,
    materipembelajaran: materiPembelajaran,
    pengalamanbelajar: pengalamanBelajar,
    prinsippembelajaran: prinsipPembelajaran,
    rekapmuridtidakikut: rekapMuridTidakIkut,
    buktifisik: buktiFisik,

    /* ALIAS HEADER LAMA */
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
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      row[index] = values[key];
    }
  });

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);

  agendaWriteLog_(ctx, {
    action: 'SIMPAN',
    module: 'AGENDA_MENGAJAR_GURU',
    description:
      'Agenda mengajar ' +
      values.namaguru +
      ' - ' +
      values.kelas +
      ' - ' +
      values.tanggal,
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
   Urutan wajib:
   TIMESTAMP | NPSN | USER_ID | EMAIL | NIP | NAMA_USER |
   ROLE | ACTION | MODULE | DESCRIPTION | TRANSACTION_ID
   ========================================================= */
function agendaWriteLog_(ctx, data) {
  const sh = getAgendaSheet_(AGENDA_GURU_CONFIG.SHEET_LOG, true);

  const headers = agendaEnsureHeadersNonDestructive_(sh, [
    'TIMESTAMP',
    'NPSN',
    'USER_ID',
    'EMAIL',
    'NIP',
    'NAMA_USER',
    'ROLE',
    'ACTION',
    'MODULE',
    'DESCRIPTION',
    'TRANSACTION_ID'
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
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      row[index] = values[key];
    }
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
