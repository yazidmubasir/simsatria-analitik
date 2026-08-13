/**
 * SCHOOL REGISTRY
 *
 * DEPLOYMENT: Execute as = User accessing the web app.
 *
 * MASTER ACCESS POLICY:
 * - ADMIN_SEKOLAH / SUPERADMIN may read MASTER because their Google accounts
 *   are explicitly granted access to MASTER.
 * - GURU / WALI_KELAS / KARYAWAN / SISWA must NEVER read MASTER directly.
 * - Non-admin users receive their school Spreadsheet ID through the binding
 *   created by ADMIN_SEKOLAH and then access only the school database.
 *
 * Session.getActiveUser() remains the identity source for every caller.
 */
function getMasterSpreadsheet_() {
  const id = getMasterSpreadsheetId_();
  if (!id) throw new Error("MASTER_SPREADSHEET_ID belum dikonfigurasi.");
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      "Akun administrator tidak dapat membaca Spreadsheet MASTER. Pastikan akun ADMIN_SEKOLAH memiliki akses minimal Viewer ke MASTER. Detail: " + e.message,
    );
  }
}

/* These helpers are ADMIN/SUPERADMIN-only by architecture. */
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
  const rows = sheetRowsAsObjects_(getMasterSchoolsSheet_());
  return rows.find(row => normalizeNpsn_(row.NPSN) === target) || null;
}
