/**
 * SCHOOL CONTEXT
 * Semua modul bisnis mengambil sekolah dari context server.
 * GURU/WALI/KARYAWAN/SISWA memakai binding sekolah; ADMIN memakai MASTER.
 */
function getSchoolContext() {
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
    throw new Error('Sheet "' + normalizedName + '" tidak ditemukan pada Spreadsheet sekolah aktif.');
  }
  return sh;
}

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
