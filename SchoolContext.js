/**
 * SCHOOL CONTEXT
 * Semua modul bisnis mengambil sekolah dari context server.
 * Tidak ada modul bisnis yang boleh memilih Spreadsheet sekolah dari frontend.
 */
function getSchoolContext() {
  const email = getGoogleUserEmail_();

  // Jika akun adalah ADMIN_SEKOLAH, putuskan binding GURU lama sebelum
  // context dipakai. Ini mencegah akun seperti masayid11 tetap membawa
  // identitas guru lama (mis. AYA).
  if (typeof enforceAdminIdentityIsolation_ === "function") {
    enforceAdminIdentityIsolation_();
    clearUserContextCache_(email);
  }

  return getCurrentUserContext().school;
}

function getSchoolSpreadsheet_() {
  const school = getSchoolContext();
  if (!school.spreadsheetId) {
    throw new Error("SPREADSHEET_ID sekolah belum dikonfigurasi.");
  }
  return SpreadsheetApp.openById(school.spreadsheetId);
}

function getSchoolSheet_(sheetName) {
  const ss = getSchoolSpreadsheet_();
  const normalizedName = String(sheetName || "").trim().toUpperCase();
  const sh = ss.getSheetByName(normalizedName);
  if (!sh) {
    throw new Error(
      'Sheet "' + normalizedName + '" tidak ditemukan pada Spreadsheet sekolah aktif.',
    );
  }
  return sh;
}

/**
 * Jalur standar untuk seluruh data master sekolah:
 * GURU, KARYAWAN, SISWA, KELAS, USERS, CONFIG, dan transaksi.
 */
function getSchoolDataSheet_(sheetName, required) {
  const ss = getSchoolSpreadsheet_();
  const normalizedName = String(sheetName || "").trim().toUpperCase();
  if (!normalizedName) throw new Error("Nama sheet sekolah wajib diisi.");

  const sh = ss.getSheetByName(normalizedName);
  if (!sh && required !== false) {
    throw new Error(
      'Sheet "' + normalizedName + '" tidak ditemukan pada Spreadsheet sekolah aktif untuk NPSN ' +
        getCurrentUserContext().npsn + ".",
    );
  }
  return sh;
}

function getSchoolContextInfo() {
  const email = getGoogleUserEmail_();
  clearUserContextCache_(email);
  if (typeof enforceAdminIdentityIsolation_ === "function") enforceAdminIdentityIsolation_();
  clearUserContextCache_(email);

  const c = getCurrentUserContext();
  return {
    success: true,
    email: c.email,
    userId: c.userId,
    nip: c.nip,
    nama: c.nama,
    role: c.role,
    npsn: c.npsn,
    sekolah: c.school.namaSekolah,
    spreadsheetId: c.school.spreadsheetId,
    driveFolderId: c.school.driveFolderId,
  };
}
