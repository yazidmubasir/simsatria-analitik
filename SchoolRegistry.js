/**
 * SCHOOL REGISTRY
 *
 * MASTER dibaca oleh AKUN EKSEKUSI Web App (pemilik aplikasi).
 * Identitas pengguna tetap berasal dari Session.getActiveUser().getEmail().
 *
 * PENTING:
 * - Web App wajib deployment: Execute as = Me.
 * - User accessing the web app tetap dipakai untuk identitas login.
 * - Guru/Wali/Karyawan/Siswa TIDAK perlu Viewer ke Spreadsheet MASTER.
 * - MASTER hanya menjadi sumber kebenaran ADMIN_SEKOLAH dan SCHOOLS.
 * - Modul bisnis membaca database sekolah melalui School Context.
 *
 * Tidak ada fallback registry/proxy yang menjadi sumber kebenaran login.
 * syncMasterAuthRegistry() hanya maintenance/diagnostik dan tidak diperlukan
 * agar ADMIN_SEKOLAH dapat login.
 */
function getMasterSpreadsheet_() {
  const id = getMasterSpreadsheetId_();
  if (!id) {
    throw new Error("MASTER_SPREADSHEET_ID belum dikonfigurasi.");
  }

  // Jangan membuka MASTER sebagai akun pengguna.
  // Pada deployment "Execute as: Me", pemanggilan ini berjalan sebagai
  // pemilik/akun eksekusi sehingga Guru tidak perlu diberi akses MASTER.
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      "Spreadsheet MASTER tidak dapat dibuka oleh akun eksekusi aplikasi. " +
      "Pastikan deployment Web App menggunakan Execute as: Me dan akun pemilik aplikasi memiliki akses ke MASTER. Detail: " +
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
  const rows = sheetValuesToObjects_(getMasterSchoolsSheet_());
  return rows.find(row => normalizeNpsn_(row.NPSN) === target) || null;
}
