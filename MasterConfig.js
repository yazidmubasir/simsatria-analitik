/**
 * ============================================================
 * SIM SATRIA - MASTER CONFIGURATION
 * ============================================================
 * MASTER_SPREADSHEET_ID tetap menjadi sumber utama.
 *
 * Untuk Web App USER_ACCESSING, akun sekolah tidak harus memiliki
 * izin membaca MASTER. Admin utama menjalankan
 * syncMasterAuthRegistry() untuk menyalin registry autentikasi
 * (ADMIN_SEKOLAH + SCHOOLS) ke Script Properties.
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
 * Sinkronkan registry MASTER ke Script Properties.
 *
 * WAJIB dijalankan oleh akun pemilik/admin utama yang mempunyai akses
 * ke Spreadsheet MASTER. Setelah sinkronisasi, akun ADMIN_SEKOLAH
 * seperti masayid11 tidak perlu membuka MASTER lagi.
 */
function syncMasterAuthRegistry() {
  const masterId = getMasterSpreadsheetId_();
  const ss = SpreadsheetApp.openById(masterId);
  const props = PropertiesService.getScriptProperties();

  const adminSheet = ss.getSheetByName("ADMIN_SEKOLAH");
  const schoolSheet = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");

  if (!adminSheet) {
    throw new Error("Sheet ADMIN_SEKOLAH tidak ditemukan pada MASTER.");
  }
  if (!schoolSheet) {
    throw new Error("Sheet SCHOOLS tidak ditemukan pada MASTER.");
  }

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
    version: 1,
  });

  props.setProperties(updates, false);

  return {
    success: true,
    adminCount: adminCount,
    schoolCount: schoolCount,
    syncedAt: updates[MASTER_AUTH_REGISTRY.META_KEY],
    message:
      "Registry autentikasi MASTER berhasil disinkronkan ke Script Properties.",
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
