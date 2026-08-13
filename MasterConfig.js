/**
 * ============================================================
 * SIM SATRIA - MASTER CONFIGURATION
 * ============================================================
 * MASTER hanya menyimpan data GLOBAL:
 *   - SCHOOLS
 *   - ADMIN_SEKOLAH
 *
 * GURU/WALI_KELAS/KARYAWAN/SISWA TIDAK BOLEH dimasukkan ke MASTER.
 * Semua pengguna sekolah tersebut dikelola melalui menu Manajemen
 * Pengguna pada Spreadsheet sekolah masing-masing (sheet USERS).
 * ============================================================
 */

const MASTER_AUTH_REGISTRY = {
  ADMIN_PREFIX: "SIM_SATRIA_MASTER_ADMIN_",
  SCHOOL_PREFIX: "SIM_SATRIA_MASTER_SCHOOL_",
  META_KEY: "SIM_SATRIA_MASTER_AUTH_REGISTRY_V1",
};

function getMasterSpreadsheetId_() {
  const props = PropertiesService.getScriptProperties();
  const id = String(props.getProperty("MASTER_SPREADSHEET_ID") || "").trim();
  if (!id) {
    throw new Error(
      "MASTER_SPREADSHEET_ID belum dikonfigurasi. Jalankan setupMasterSpreadsheetId() terlebih dahulu.",
    );
  }
  return id;
}

function setupMasterSpreadsheetId() {
  const MASTER_ID = "1o7l24gGB7rXsjyFJJw1plz_0ud-V74USqyLQsFbZmS0";
  if (!MASTER_ID) {
    throw new Error("Silakan isi MASTER_ID dengan ID Spreadsheet SIM SATRIA MASTER.");
  }

  const ss = SpreadsheetApp.openById(MASTER_ID);
  const sheet = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");
  if (!sheet) {
    throw new Error(
      'Spreadsheet MASTER berhasil dibuka, tetapi sheet "SCHOOLS" tidak ditemukan.',
    );
  }

  PropertiesService.getScriptProperties().setProperty(
    "MASTER_SPREADSHEET_ID",
    MASTER_ID,
  );

  return {
    success: true,
    message: "Spreadsheet SIM SATRIA MASTER berhasil dikonfigurasi.",
    spreadsheetId: MASTER_ID,
    spreadsheetName: ss.getName(),
    sheet: sheet.getName(),
  };
}

/**
 * Validasi keras struktur ADMIN_SEKOLAH pada MASTER.
 * MASTER hanya boleh berisi administrator sekolah.
 */
function validateMasterAdminSheet_(sheet) {
  if (!sheet) throw new Error("Sheet ADMIN_SEKOLAH tidak ditemukan pada MASTER.");

  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    throw new Error("Sheet ADMIN_SEKOLAH masih kosong.");
  }

  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });

  const requiredHeaders = [
    "USER_ID",
    "EMAIL",
    "NIP",
    "NAMA",
    "NPSN",
    "ROLE",
    "STATUS",
  ];

  requiredHeaders.forEach(function (header) {
    if (headers.indexOf(header) < 0) {
      throw new Error('Kolom "' + header + '" wajib ada pada sheet ADMIN_SEKOLAH.');
    }
  });

  const index = {};
  headers.forEach(function (header, i) {
    index[header] = i;
  });

  const seenEmails = {};
  const seenNpsn = {};
  let adminCount = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const email = normalizeEmail_(row[index.EMAIL]);
    const npsn = normalizeNpsn_(row[index.NPSN]);
    const role = normalizeAuthRole_(row[index.ROLE]);
    const status = String(row[index.STATUS] || "").trim().toUpperCase();

    // Baris kosong setelah data terakhir diperbolehkan.
    if (!email && !npsn && !role && !status) continue;

    if (!email) throw new Error("ADMIN_SEKOLAH baris " + (i + 1) + ": EMAIL wajib diisi.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      throw new Error("ADMIN_SEKOLAH baris " + (i + 1) + ": EMAIL tidak valid.");
    }
    if (!npsn) throw new Error("ADMIN_SEKOLAH baris " + (i + 1) + ": NPSN wajib diisi.");

    // MASTER tidak boleh menjadi database guru.
    if (role !== "ADMIN_SEKOLAH") {
      throw new Error(
        "ADMIN_SEKOLAH baris " +
          (i + 1) +
          ": ROLE harus ADMIN_SEKOLAH. GURU/WALI_KELAS/KARYAWAN/SISWA harus dikelola melalui USERS sekolah masing-masing.",
      );
    }

    if (!["ACTIVE", "INACTIVE"].includes(status)) {
      throw new Error(
        "ADMIN_SEKOLAH baris " + (i + 1) + ": STATUS harus ACTIVE atau INACTIVE.",
      );
    }

    if (seenEmails[email]) throw new Error("Email ADMIN_SEKOLAH duplikat: " + email + ".");
    if (seenNpsn[npsn]) throw new Error("NPSN ADMIN_SEKOLAH duplikat: " + npsn + ".");

    seenEmails[email] = true;
    seenNpsn[npsn] = true;
    adminCount++;
  }

  if (adminCount === 0) {
    throw new Error("Belum ada ADMIN_SEKOLAH aktif/terdaftar pada MASTER.");
  }

  return {
    success: true,
    adminCount: adminCount,
    headers: requiredHeaders,
  };
}

/**
 * Sinkronkan registry MASTER ke Script Properties.
 * Registry yang disalin hanya ADMIN_SEKOLAH dan SCHOOLS.
 * GURU/WALI_KELAS/KARYAWAN/SISWA tidak pernah menjadi registry MASTER.
 */
function syncMasterAuthRegistry() {
  const callerEmail = getGoogleUserEmail_();
  if (!isSuperAdminEmail_(callerEmail)) {
    throw new Error("Hanya SUPERADMIN yang boleh melakukan sinkronisasi MASTER.");
  }

  const masterId = getMasterSpreadsheetId_();
  const ss = SpreadsheetApp.openById(masterId);
  const props = PropertiesService.getScriptProperties();

  const adminSheet = ss.getSheetByName("ADMIN_SEKOLAH");
  const schoolSheet = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");

  if (!adminSheet) throw new Error("Sheet ADMIN_SEKOLAH tidak ditemukan pada MASTER.");
  if (!schoolSheet) throw new Error("Sheet SCHOOLS tidak ditemukan pada MASTER.");

  const adminValidation = validateMasterAdminSheet_(adminSheet);
  const admins = sheetValuesToObjects_(adminSheet);
  const schools = sheetValuesToObjects_(schoolSheet);

  // Hapus registry lama agar admin/sekolah yang dihapus dari MASTER
  // tidak tertinggal di Script Properties.
  const allProps = props.getProperties();
  const oldKeys = Object.keys(allProps).filter(function (key) {
    return (
      key.indexOf(MASTER_AUTH_REGISTRY.ADMIN_PREFIX) === 0 ||
      key.indexOf(MASTER_AUTH_REGISTRY.SCHOOL_PREFIX) === 0
    );
  });
  if (oldKeys.length) props.deleteProperties(oldKeys);

  const updates = {};
  let adminCount = 0;
  let schoolCount = 0;

  admins.forEach(function (admin) {
    const email = normalizeEmail_(admin.EMAIL);
    if (!email) return;

    const role = normalizeAuthRole_(admin.ROLE);
    if (role !== "ADMIN_SEKOLAH") {
      throw new Error(
        "Registry MASTER menolak " + email + ": hanya ROLE ADMIN_SEKOLAH yang boleh berada di ADMIN_SEKOLAH.",
      );
    }

    updates[
      MASTER_AUTH_REGISTRY.ADMIN_PREFIX + registrySafeKey_(email)
    ] = JSON.stringify(admin);
    adminCount++;
  });

  schools.forEach(function (school) {
    const npsn = normalizeNpsn_(school.NPSN);
    if (!npsn) return;
    updates[
      MASTER_AUTH_REGISTRY.SCHOOL_PREFIX + registrySafeKey_(npsn)
    ] = JSON.stringify(school);
    schoolCount++;
  });

  updates[MASTER_AUTH_REGISTRY.META_KEY] = JSON.stringify({
    syncedAt: new Date().toISOString(),
    masterSpreadsheetId: masterId,
    adminCount: adminCount,
    schoolCount: schoolCount,
    version: 2,
  });

  props.setProperties(updates, false);

  return {
    success: true,
    adminCount: adminCount,
    schoolCount: schoolCount,
    validatedAdminCount: adminValidation.adminCount,
    syncedAt: updates[MASTER_AUTH_REGISTRY.META_KEY],
    message:
      "Registry MASTER berhasil disinkronkan. MASTER hanya memuat ADMIN_SEKOLAH dan SCHOOLS; pengguna GURU/WALI_KELAS/KARYAWAN/SISWA tetap dikelola pada USERS sekolah masing-masing.",
  };
}

function getLocalMasterAdminByEmail_(email) {
  const normalized = normalizeEmail_(email);
  if (!normalized) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(
    MASTER_AUTH_REGISTRY.ADMIN_PREFIX + registrySafeKey_(normalized),
  );
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function getLocalMasterSchoolByNpsn_(npsn) {
  const normalized = normalizeNpsn_(npsn);
  if (!normalized) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(
    MASTER_AUTH_REGISTRY.SCHOOL_PREFIX + registrySafeKey_(normalized),
  );
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function getMasterAuthRegistryStatus() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(MASTER_AUTH_REGISTRY.META_KEY);
  if (!raw) {
    return {
      success: false,
      configured: false,
      message: "Registry autentikasi belum pernah disinkronkan.",
    };
  }
  try {
    const meta = JSON.parse(raw);
    return { success: true, configured: true, meta: meta };
  } catch (e) {
    return {
      success: false,
      configured: false,
      message: "Registry autentikasi rusak atau tidak valid.",
    };
  }
}

function sheetValuesToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });
  return values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function registrySafeKey_(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9._-]/g, "_");
}

function testMasterSpreadsheet() {
  const masterId = getMasterSpreadsheetId_();
  const ss = SpreadsheetApp.openById(masterId);
  const sheet = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");
  return {
    success: true,
    spreadsheetId: masterId,
    spreadsheetName: ss.getName(),
    sheet: sheet.getName(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn(),
  };
}
