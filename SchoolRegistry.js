/**
 * SCHOOL REGISTRY
 *
 * MASTER dibaca LANGSUNG pada setiap proses autentikasi.
 * Identitas pengguna berasal dari Session.getActiveUser(), sehingga Web App
 * harus menggunakan "User accessing the web app".
 *
 * Tidak ada fallback registry/proxy yang menjadi sumber kebenaran login.
 * syncMasterAuthRegistry() hanya maintenance/diagnostik dan TIDAK diperlukan
 * agar ADMIN_SEKOLAH dapat login.
 */
function getMasterSpreadsheet_() {
  const id = getMasterSpreadsheetId_();
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      "Spreadsheet MASTER tidak dapat dibuka oleh akun pengguna. Pastikan akun pengguna memiliki akses minimal Viewer ke Spreadsheet MASTER. Detail: " + e.message,
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
