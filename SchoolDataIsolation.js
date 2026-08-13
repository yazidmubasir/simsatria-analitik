/**
 * ============================================================
 * SIM SATRIA - SCHOOL DATA ISOLATION
 * ============================================================
 * Prinsip utama:
 * 1. SUPERADMIN adalah pemilik aplikasi.
 * 2. ADMIN_SEKOLAH ditentukan HANYA dari MASTER / registry autentikasi.
 * 3. Setelah sekolah ditemukan dari NPSN admin, seluruh data bisnis:
 *    GURU, KARYAWAN, SISWA, KELAS, USERS, CONFIG dan TRX_* wajib
 *    dibaca dari Spreadsheet sekolah tersebut.
 * 4. Binding lama ketika akun berubah dari GURU menjadi ADMIN_SEKOLAH
 *    tidak boleh lagi menentukan identitas pengguna.
 * 5. Helper di file ini server-side agar modul baru tidak perlu mengetahui
 *    ID Spreadsheet sekolah secara manual.
 * ============================================================
 */

const SCHOOL_DATA_ISOLATION_CONFIG = {
  MASTER_REFERENCE_SHEETS: ["ADMIN_SEKOLAH", "SCHOOLS"],
  SCHOOL_REFERENCE_SHEETS: [
    "CONFIG",
    "GURU",
    "KARYAWAN",
    "SISWA",
    "KELAS",
    "USERS",
    "TRX_PRESENSI",
    "TRX_PARKIR",
    "TRX_PRESTASI",
    "TRX_AGENDA_GURU",
    "TRX_SBI",
    "TRX_KEBERSIHAN",
    "TRX_KEAMANAN",
    "TRX_KERJA",
    "LOG",
  ],
};

/**
 * ADMIN_SEKOLAH/SUPERADMIN tidak boleh membawa identitas GURU lama.
 * Jika email ditemukan di registry admin, binding lokal lama dihapus.
 * Data historis pada sheet GURU tidak dihapus; yang dihapus hanya
 * hubungan autentikasi lama email -> GURU.
 */
function enforceAdminIdentityIsolation_() {
  const email = getGoogleUserEmail_();
  const admin = getAdminByEmail_(email);
  if (!admin) return null;

  const role = normalizeAuthRole_(admin.ROLE);
  const status = String(admin.STATUS || "").trim().toUpperCase();
  const isSuperAdmin = typeof isSuperAdminEmail_ === "function" && isSuperAdminEmail_(email);

  if (!isSuperAdmin && role !== "ADMIN_SEKOLAH") {
    return null;
  }

  if (status !== "ACTIVE") {
    throw new Error("Akun administrator tidak aktif.");
  }

  // Putus hubungan autentikasi lama, misalnya akun yang dahulu tercatat
  // sebagai GURU/AYA. Tidak menghapus data GURU historis.
  const bindings = getUserBindings_();
  if (Object.prototype.hasOwnProperty.call(bindings, email)) {
    delete bindings[email];
    saveUserBindings_(bindings);
  }

  clearUserContextCache_(email);

  return admin;
}

/**
 * Selalu memperoleh Spreadsheet sekolah aktif berdasarkan context server.
 * Tidak menerima spreadsheetId dari frontend.
 */
function getActiveSchoolSpreadsheet_() {
  const context = getCurrentUserContext();
  if (!context || !context.school || !context.school.spreadsheetId) {
    throw new Error("Spreadsheet sekolah aktif belum tersedia.");
  }

  const spreadsheetId = String(context.school.spreadsheetId).trim();
  if (!spreadsheetId) {
    throw new Error("SPREADSHEET_ID sekolah aktif kosong.");
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * Mengambil sheet data sekolah aktif.
 * Fungsi ini adalah jalur standar untuk GURU/KARYAWAN/SISWA/KELAS/USERS
 * dan seluruh sheet transaksi sekolah.
 */
function getActiveSchoolDataSheet_(sheetName, required) {
  const name = String(sheetName || "").trim().toUpperCase();
  if (!name) throw new Error("Nama sheet sekolah wajib diisi.");

  const ss = getActiveSchoolSpreadsheet_();
  const sheet = ss.getSheetByName(name);

  if (!sheet && required !== false) {
    throw new Error(
      'Sheet "' + name + '" tidak ditemukan pada Spreadsheet sekolah aktif untuk NPSN ' +
        getCurrentUserContext().npsn + ".",
    );
  }

  return sheet;
}

function getActiveSchoolReferenceData_(sheetName) {
  const name = String(sheetName || "").trim().toUpperCase();
  if (SCHOOL_DATA_ISOLATION_CONFIG.SCHOOL_REFERENCE_SHEETS.indexOf(name) < 0) {
    throw new Error("Sheet " + name + " bukan sheet data sekolah yang diizinkan.");
  }

  const sheet = getActiveSchoolDataSheet_(name, true);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0].map(function (header) {
    return String(header || "").trim().toUpperCase();
  });

  return values.slice(1).map(function (row) {
    const item = {};
    headers.forEach(function (header, index) {
      item[header] = row[index];
    });
    return item;
  });
}

function auditCurrentSchoolDataSource() {
  const context = getCurrentUserContext();
  const ss = getActiveSchoolSpreadsheet_();

  return {
    success: true,
    email: context.email,
    role: context.role,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    spreadsheetId: context.school.spreadsheetId,
    spreadsheetName: ss.getName(),
    sheets: SCHOOL_DATA_ISOLATION_CONFIG.SCHOOL_REFERENCE_SHEETS.map(function (name) {
      const sheet = ss.getSheetByName(name);
      return {
        sheet: name,
        exists: !!sheet,
        rows: sheet ? sheet.getLastRow() : 0,
      };
    }),
  };
}

function repairCurrentAdminIdentity() {
  const admin = enforceAdminIdentityIsolation_();
  if (!admin) {
    throw new Error("Akun Google aktif bukan ADMIN_SEKOLAH/SUPERADMIN.");
  }

  const context = getCurrentUserContext();
  return {
    success: true,
    email: context.email,
    role: context.role,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    spreadsheetId: context.school.spreadsheetId,
    message:
      "Identitas administrator sudah dipisahkan dari binding GURU lama. Data GURU historis tidak dihapus.",
  };
}
