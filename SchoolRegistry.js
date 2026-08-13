/**
 * SCHOOL REGISTRY
 *
 * MASTER dibaca langsung oleh akun pemilik aplikasi.
 * Tidak ada fallback proxy untuk autentikasi login karena proxy/cache
 * membuat perubahan ADMIN_SEKOLAH tidak langsung terlihat.
 */
function getMasterSpreadsheet_() {
  const id = getMasterSpreadsheetId_();
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      "Spreadsheet MASTER tidak dapat dibuka oleh akun eksekusi aplikasi. " +
      "Pastikan deployment Web App menggunakan Execute as: Me (pemilik aplikasi). Detail: " + e.message,
    );
  }
}

function getMasterAdminSheet_() {
  const sheet = getMasterSpreadsheet_().getSheetByName("ADMIN_SEKOLAH");
  if (!sheet) throw new Error("Sheet ADMIN_SEKOLAH tidak ditemukan pada MASTER.");
  return sheet;
}

function getMasterSchoolsSheet_() {
  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");
  if (!sheet) throw new Error("Sheet SCHOOLS tidak ditemukan pada MASTER.");
  return sheet;
}

function getSchoolByNpsn(npsn) {
  const school = getSchoolByNpsnAuth_(npsn);
  if (!school) throw new Error("Sekolah tidak ditemukan.");
  return school;
}

function getSchoolByNpsnDirect_(npsn) {
  const target = normalizeNpsn_(npsn);
  if (!target) return null;
  const rows = sheetValuesToObjects_(getMasterSchoolsSheet_());
  return rows.find(row => normalizeNpsn_(row.NPSN) === target) || null;
}
