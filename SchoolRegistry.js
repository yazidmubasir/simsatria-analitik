/**
 * SCHOOL REGISTRY
 *
 * Deployment: Execute as = User accessing the web app.
 *
 * MASTER hanya digunakan pada jalur ADMIN_SEKOLAH/SUPERADMIN.
 * GURU/WALI_KELAS/KARYAWAN/SISWA tidak boleh membaca MASTER.
 * Mereka memperoleh Spreadsheet sekolah melalui binding yang dibuat oleh
 * ADMIN_SEKOLAH dan selanjutnya hanya bekerja pada database sekolah.
 */
function getMasterSpreadsheet_() {
  const id = getMasterSpreadsheetId_();
  if (!id) throw new Error("MASTER_SPREADSHEET_ID belum dikonfigurasi.");
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      "Akun ADMIN_SEKOLAH tidak dapat membaca Spreadsheet MASTER. Pastikan akun administrator memiliki akses minimal Viewer ke MASTER. Detail: " +
        e.message,
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
  const rows = sheetRowsAsObjects_(getMasterSchoolsSheet_());
  return rows.find(function(row) { return normalizeNpsn_(row.NPSN) === target; }) || null;
}
