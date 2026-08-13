/**
 * SIM SATRIA - REPAIR AKSES USER SEKOLAH
 *
 * Memperbaiki binding autentikasi dan permission Spreadsheet/Drive
 * untuk user sekolah yang sudah ada pada USERS.
 */

/**
 * Memberikan Editor langsung ke file Spreadsheet sekolah.
 *
 * Penting: jangan hanya mengandalkan akses folder/root Drive. Web App
 * yang berjalan sebagai USER_ACCESSING membutuhkan izin langsung/efektif
 * terhadap file Spreadsheet yang dibuka melalui SpreadsheetApp.openById().
 * Fungsi ini harus dipanggil oleh ADMIN_SEKOLAH.
 */
function grantSchoolSpreadsheetEditor_(spreadsheetId, email) {
  const id = String(spreadsheetId || '').trim();
  const target = normalizeEmail_(email);
  if (!id) throw new Error('Spreadsheet ID sekolah kosong.');
  if (!target) throw new Error('Email pengguna kosong.');

  // Admin yang menjalankan fungsi ini harus sudah memiliki akses ke file.
  const file = DriveApp.getFileById(id);
  file.addEditor(target);

  // Verifikasi akses langsung/efektif melalui daftar editor file.
  const editors = file.getEditors().map(function(user) {
    return normalizeEmail_(user.getEmail());
  });

  if (editors.indexOf(target) < 0) {
    throw new Error(
      'Akses Editor Spreadsheet sekolah belum terdeteksi untuk ' + target +
      '. Pastikan akun admin memiliki hak mengelola akses file dan kebijakan Google Workspace mengizinkan berbagi file.'
    );
  }

  // Verifikasi bahwa file memang dapat dibuka oleh akun eksekusi admin.
  try {
    SpreadsheetApp.openById(id).getName();
  } catch (e) {
    throw new Error(
      'Spreadsheet sekolah tidak dapat dibuka oleh akun aplikasi: ' +
      (e && e.message ? e.message : String(e))
    );
  }

  return {
    success: true,
    spreadsheetId: id,
    email: target,
    fileName: file.getName(),
  };
}

function repairSchoolUserAccess(email) {
  const context = requireUserManager_();
  email = normalizeEmail_(email);
  if (!email) throw new Error('Email pengguna wajib diisi.');
  if (email === normalizeEmail_(context.email)) {
    throw new Error('Gunakan akun admin sendiri tanpa repair user.');
  }

  const sheet = getOrCreateUsersSheet_(context.school.spreadsheetId);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('Sheet USERS belum memiliki data pengguna.');

  const headers = values[0].map(normalizeHeader_);
  const index = {};
  headers.forEach(function(h, i) { index[h] = i; });

  ['EMAIL', 'USER_ID', 'NIP', 'NAMA', 'ROLE', 'STATUS'].forEach(function(header) {
    if (index[header] === undefined) throw new Error('Kolom ' + header + ' tidak ditemukan di USERS.');
  });

  let user = null;
  for (let i = 1; i < values.length; i++) {
    if (normalizeEmail_(values[i][index.EMAIL]) === email) {
      user = {
        USER_ID: String(values[i][index.USER_ID] || '').trim(),
        EMAIL: email,
        NIP: String(values[i][index.NIP] || '').trim(),
        NAMA: String(values[i][index.NAMA] || '').trim(),
        ROLE: normalizeRole_(values[i][index.ROLE]),
        STATUS: normalizeRole_(values[i][index.STATUS]),
      };
      break;
    }
  }

  if (!user) {
    throw new Error(
      'Akun ' + email + ' belum terdaftar pada USERS sekolah ' + context.school.namaSekolah + '.'
    );
  }

  if (['GURU', 'WALI_KELAS', 'KARYAWAN'].indexOf(user.ROLE) < 0) {
    throw new Error('Role ' + user.ROLE + ' tidak dapat diperbaiki melalui repair user sekolah.');
  }

  if (user.STATUS !== 'ACTIVE') {
    removeSchoolUserBinding_(email);
    revokeSchoolSpreadsheetAccess_(context.school.spreadsheetId, email);
    return {
      success: true,
      email: email,
      role: user.ROLE,
      status: user.STATUS,
      binding: false,
      accessGranted: false,
      message: 'User INACTIVE: binding dan akses Spreadsheet tidak diberikan.',
    };
  }

  // 1. Binding selalu dibangun dari USERS sekolah aktif.
  const binding = registerSchoolUserBinding_(context, user);

  // 2. Berikan akses langsung ke file Spreadsheet sekolah.
  const spreadsheetAccess = grantSchoolSpreadsheetEditor_(context.school.spreadsheetId, email);

  // 3. Sinkronkan root + PRESENSI + AGENDA.
  let drivePermission = { success: true, skipped: true };
  try {
    drivePermission = syncTeacherSchoolDrivePermission_(
      context,
      email,
      user.ROLE,
      user.STATUS,
    );
  } catch (e) {
    drivePermission = {
      success: false,
      error: e && e.message ? e.message : String(e),
    };
  }

  clearUserContextCache_(email);

  return {
    success: true,
    email: email,
    userId: user.USER_ID,
    nama: user.NAMA,
    role: user.ROLE,
    status: user.STATUS,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    spreadsheetId: context.school.spreadsheetId,
    binding: {
      success: true,
      npsn: binding.npsn,
      sekolah: binding.namaSekolah,
    },
    spreadsheetAccess: spreadsheetAccess,
    accessGranted: true,
    drivePermission: drivePermission,
    message: 'Binding, akses Spreadsheet, dan permission Drive user berhasil diperbaiki.',
  };
}

/** Repair seluruh user ACTIVE pada USERS sekolah. */
function repairAllActiveSchoolUsersAccess() {
  const context = requireUserManager_();
  const sheet = getOrCreateUsersSheet_(context.school.spreadsheetId);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { success: true, total: 0, repaired: 0, failed: 0, school: context.school.namaSekolah, results: [] };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf('EMAIL');
  const statusIndex = headers.indexOf('STATUS');
  if (emailIndex < 0 || statusIndex < 0) {
    throw new Error('Struktur USERS wajib memiliki EMAIL dan STATUS.');
  }

  const emails = [];
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    const status = normalizeRole_(values[i][statusIndex]);
    if (email && status === 'ACTIVE' && email !== normalizeEmail_(context.email)) {
      emails.push(email);
    }
  }

  const results = [];
  let repaired = 0;
  let failed = 0;

  emails.forEach(function(email) {
    try {
      const result = repairSchoolUserAccess(email);
      results.push(result);
      if (result.success) repaired++;
      else failed++;
    } catch (e) {
      failed++;
      results.push({
        success: false,
        email: email,
        error: e && e.message ? e.message : String(e),
      });
    }
  });

  return {
    success: failed === 0,
    total: emails.length,
    repaired: repaired,
    failed: failed,
    school: context.school.namaSekolah,
    npsn: context.npsn,
    spreadsheetId: context.school.spreadsheetId,
    results: results,
  };
}
