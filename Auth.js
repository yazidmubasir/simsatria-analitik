/**
 * AUTH.GS
 *
 * MASTER hanya menyimpan registry sekolah dan administrator sekolah.
 * Guru/Karyawan/Wali Kelas dikelola pada Spreadsheet sekolah -> USERS.
 *
 * Routing:
 *   ADMIN: Google Account -> ADMIN_SEKOLAH -> NPSN -> SCHOOLS
 *   USER : Google Account -> scan ACTIVE SCHOOLS -> local USERS -> SCHOOLS
 */
const AUTH_CONFIG = {
  ADMIN_SHEET: "ADMIN_SEKOLAH",
  SCHOOLS_SHEET: "SCHOOLS",
  USERS_SHEET: "USERS",
  ACTIVE_STATUS: "ACTIVE",
  CACHE_SECONDS: 21600,
};

function normalizeEmail_(email) {
  return String(email || "").trim().toLowerCase();
}

function normalizeNpsn_(npsn) {
  return String(npsn || "").trim();
}

function getGoogleUserEmail_() {
  const user = Session.getActiveUser();
  const email = user ? normalizeEmail_(user.getEmail()) : "";
  if (!email) {
    throw new Error(
      'Identitas akun Google tidak dapat diperoleh. Pastikan deployment menggunakan "User accessing the web app" dan akun sudah memberikan otorisasi.',
    );
  }
  return email;
}

function getAdminSheet_() {
  const sheet = getMasterSpreadsheet_().getSheetByName(AUTH_CONFIG.ADMIN_SHEET);
  if (!sheet) throw new Error("Sheet ADMIN_SEKOLAH tidak ditemukan pada Spreadsheet Master. Jalankan setupMasterRegistry().");
  return sheet;
}

function getSchoolsSheet_() {
  const sheet = getMasterSpreadsheet_().getSheetByName(AUTH_CONFIG.SCHOOLS_SHEET);
  if (!sheet) throw new Error("Sheet SCHOOLS tidak ditemukan pada Spreadsheet Master. Jalankan setupMasterRegistry().");
  return sheet;
}

function sheetRowsAsObjects_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(function (h) { return String(h || "").trim().toUpperCase(); });
  return values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (header, index) { obj[header] = row[index]; });
    return obj;
  });
}

function getAdminByEmail_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return null;
  const cache = CacheService.getScriptCache();
  const key = "ADMIN_" + normalizedEmail.replace(/[^a-zA-Z0-9]/g, "_");
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const sheet = getAdminSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(function (h) { return String(h || "").trim().toUpperCase(); });
  const emailIndex = headers.indexOf("EMAIL");
  if (emailIndex === -1) throw new Error("Kolom EMAIL tidak ditemukan di ADMIN_SEKOLAH.");

  for (let i = 1; i < values.length; i++) {
    if (normalizeEmail_(values[i][emailIndex]) === normalizedEmail) {
      const admin = {};
      headers.forEach(function (header, index) { admin[header] = values[i][index]; });
      cache.put(key, JSON.stringify(admin), AUTH_CONFIG.CACHE_SECONDS);
      return admin;
    }
  }
  return null;
}

function getSchoolByNpsnAuth_(npsn) {
  const normalizedNpsn = normalizeNpsn_(npsn);
  if (!normalizedNpsn) return null;
  const cache = CacheService.getScriptCache();
  const key = "SCHOOL_" + normalizedNpsn;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const sheet = getSchoolsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(function (h) { return String(h || "").trim().toUpperCase(); });
  const npsnIndex = headers.indexOf("NPSN");
  if (npsnIndex === -1) throw new Error("Kolom NPSN tidak ditemukan di SCHOOLS.");

  for (let i = 1; i < values.length; i++) {
    if (normalizeNpsn_(values[i][npsnIndex]) === normalizedNpsn) {
      const school = {};
      headers.forEach(function (header, index) { school[header] = values[i][index]; });
      cache.put(key, JSON.stringify(school), AUTH_CONFIG.CACHE_SECONDS);
      return school;
    }
  }
  return null;
}

function getSchoolUserByEmail_(spreadsheetId, email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!spreadsheetId || !normalizedEmail) return null;
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
  if (!sheet) return null;
  const rows = sheetRowsAsObjects_(sheet);
  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail_(rows[i].EMAIL) === normalizedEmail) return rows[i];
  }
  return null;
}

function buildSchoolContext_(user, school, email) {
  const spreadsheetId = String(school.SPREADSHEET_ID || "").trim();
  if (!spreadsheetId) throw new Error("Spreadsheet sekolah belum dikonfigurasi.");
  return {
    authenticated: true,
    email: normalizeEmail_(email),
    userId: String(user.USER_ID || "").trim(),
    nip: String(user.NIP || "").trim(),
    nama: String(user.NAMA || "").trim(),
    role: String(user.ROLE || "GURU").trim().toUpperCase(),
    npsn: normalizeNpsn_(school.NPSN),
    school: {
      npsn: normalizeNpsn_(school.NPSN),
      namaSekolah: String(school.NAMA_SEKOLAH || "").trim(),
      spreadsheetId: spreadsheetId,
      driveFolderId: String(school.DRIVE_FOLDER_ID || "").trim(),
      alamat: String(school.ALAMAT || "").trim(),
      logoUrl: String(school.LOGO_URL || "").trim(),
      tagline: String(school.TAGLINE || "").trim(),
      warnaUtama: String(school.WARNA_UTAMA || "").trim(),
      warnaSekunder: String(school.WARNA_SEKUNDER || "").trim(),
    },
  };
}

function getLocalUserContextByEmail_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return null;
  const cache = CacheService.getScriptCache();
  const key = "LOCAL_USER_" + normalizedEmail.replace(/[^a-zA-Z0-9]/g, "_");
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(key); }
  }

  const sheet = getSchoolsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(function (h) { return String(h || "").trim().toUpperCase(); });
  const npsnIndex = headers.indexOf("NPSN");
  const statusIndex = headers.indexOf("STATUS");
  const spreadsheetIndex = headers.indexOf("SPREADSHEET_ID");
  if (npsnIndex < 0 || spreadsheetIndex < 0) return null;

  for (let i = 1; i < values.length; i++) {
    const schoolStatus = statusIndex >= 0 ? String(values[i][statusIndex] || "").trim().toUpperCase() : "ACTIVE";
    if (schoolStatus && schoolStatus !== AUTH_CONFIG.ACTIVE_STATUS) continue;
    const npsn = normalizeNpsn_(values[i][npsnIndex]);
    const spreadsheetId = String(values[i][spreadsheetIndex] || "").trim();
    if (!npsn || !spreadsheetId) continue;

    try {
      const user = getSchoolUserByEmail_(spreadsheetId, normalizedEmail);
      if (!user) continue;
      const userStatus = String(user.STATUS || "").trim().toUpperCase();
      if (userStatus !== AUTH_CONFIG.ACTIVE_STATUS) continue;
      const school = {};
      headers.forEach(function (header, index) { school[header] = values[i][index]; });
      const context = buildSchoolContext_(user, school, normalizedEmail);
      cache.put(key, JSON.stringify(context), AUTH_CONFIG.CACHE_SECONDS);
      return context;
    } catch (e) {
      console.warn("[AUTH] Gagal membaca USERS sekolah NPSN " + npsn + ": " + e.message);
    }
  }
  return null;
}

function getCurrentUserContext() {
  const email = getGoogleUserEmail_();
  const cache = CacheService.getScriptCache();
  const cacheKey = "USER_CONTEXT_V4_" + email.replace(/[^a-zA-Z0-9]/g, "_");
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { cache.remove(cacheKey); }
  }

  // Jalur utama administrator sekolah.
  const admin = getAdminByEmail_(email);
  if (admin) {
    const status = String(admin.STATUS || "").trim().toUpperCase();
    if (status !== AUTH_CONFIG.ACTIVE_STATUS) throw new Error("Akun administrator sekolah tidak aktif.");
    const npsn = normalizeNpsn_(admin.NPSN);
    if (!npsn) throw new Error("Akun administrator belum memiliki NPSN sekolah.");
    const school = getSchoolByNpsnAuth_(npsn);
    if (!school) throw new Error("Sekolah dengan NPSN " + npsn + " tidak ditemukan pada registry SIM SATRIA.");
    const schoolStatus = String(school.STATUS || "").trim().toUpperCase();
    if (schoolStatus && schoolStatus !== AUTH_CONFIG.ACTIVE_STATUS) throw new Error("Sekolah Anda tidak aktif pada SIM SATRIA.");
    const context = buildSchoolContext_(admin, school, email);
    cache.put(cacheKey, JSON.stringify(context), AUTH_CONFIG.CACHE_SECONDS);
    return context;
  }

  // Jalur pengguna sekolah: GURU/WALI_KELAS/KARYAWAN.
  const localContext = getLocalUserContextByEmail_(email);
  if (localContext) {
    cache.put(cacheKey, JSON.stringify(localContext), AUTH_CONFIG.CACHE_SECONDS);
    return localContext;
  }

  throw new Error('Akun Google "' + email + '" belum terdaftar pada SIM SATRIA. Hubungi ADMIN_SEKOLAH sekolah Anda.');
}

function checkAuthentication() {
  try {
    const c = getCurrentUserContext();
    return { success: true, authenticated: true, email: c.email, userId: c.userId, npsn: c.npsn, sekolah: c.school.namaSekolah, role: c.role };
  } catch (e) {
    return { success: false, authenticated: false, message: e.message };
  }
}

function bindMySchool(npsn) {
  const email = getGoogleUserEmail_();
  const requestedNpsn = normalizeNpsn_(npsn);
  if (!requestedNpsn) throw new Error("NPSN sekolah wajib diisi.");
  const admin = getAdminByEmail_(email);
  if (!admin) throw new Error("Akun Google belum terdaftar di ADMIN_SEKOLAH.");
  const adminNpsn = normalizeNpsn_(admin.NPSN);
  if (adminNpsn && adminNpsn !== requestedNpsn) throw new Error("NPSN yang diminta berbeda dengan NPSN akun administrator. Binding ditolak demi keamanan.");
  const school = getSchoolByNpsnAuth_(requestedNpsn);
  if (!school) throw new Error("NPSN sekolah tidak ditemukan pada SCHOOLS.");
  const spreadsheetId = String(school.SPREADSHEET_ID || "").trim();
  if (!spreadsheetId) throw new Error("SPREADSHEET_ID sekolah belum dikonfigurasi.");
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let users = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
  if (!users) users = ss.insertSheet(AUTH_CONFIG.USERS_SHEET);
  const headers = ["USER_ID", "EMAIL", "NIP", "NAMA", "ROLE", "STATUS"];
  ensureLocalHeaders_(users, headers);
  const values = users.getDataRange().getValues();
  const headerRow = values[0].map(function (h) { return String(h || "").trim().toUpperCase(); });
  const emailIndex = headerRow.indexOf("EMAIL"), userIdIndex = headerRow.indexOf("USER_ID"), nipIndex = headerRow.indexOf("NIP"), namaIndex = headerRow.indexOf("NAMA"), roleIndex = headerRow.indexOf("ROLE"), statusIndex = headerRow.indexOf("STATUS");
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) if (normalizeEmail_(values[i][emailIndex]) === email) { targetRow = i + 1; break; }
  const userId = String(admin.USER_ID || "").trim() || Utilities.getUuid();
  const row = new Array(headerRow.length).fill("");
  row[userIdIndex] = userId; row[emailIndex] = email; row[nipIndex] = String(admin.NIP || "").trim(); row[namaIndex] = String(admin.NAMA || "").trim(); row[roleIndex] = String(admin.ROLE || "ADMIN_SEKOLAH").trim().toUpperCase(); row[statusIndex] = String(admin.STATUS || "ACTIVE").trim().toUpperCase();
  if (targetRow === -1) users.getRange(users.getLastRow() + 1, 1, 1, row.length).setValues([row]); else users.getRange(targetRow, 1, 1, row.length).setValues([row]);
  clearMyAuthCache();
  return { success: true, message: "Akun berhasil di-bind ke database sekolah.", email: email, npsn: requestedNpsn, sekolah: String(school.NAMA_SEKOLAH || "").trim(), userId: userId };
}

function ensureLocalHeaders_(sheet, requiredHeaders) {
  const normalizedRequired = requiredHeaders.map(function (h) { return String(h || "").trim().toUpperCase(); });
  if (sheet.getLastColumn() === 0) { sheet.getRange(1, 1, 1, normalizedRequired.length).setValues([normalizedRequired]); sheet.setFrozenRows(1); return normalizedRequired; }
  const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h || "").trim().toUpperCase(); });
  normalizedRequired.forEach(function (header) { if (!current.includes(header)) { sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header); current.push(header); } });
  sheet.setFrozenRows(1); return current;
}

function clearMyAuthCache() {
  const email = getGoogleUserEmail_();
  clearUserContextCache_(email);
  return { success: true };
}

function refreshMySchoolContext() {
  clearMyAuthCache();
  const context = getCurrentUserContext();
  return { success: true, message: "School Context berhasil di-refresh.", email: context.email, npsn: context.npsn, sekolah: context.school.namaSekolah, spreadsheetId: context.school.spreadsheetId, driveFolderId: context.school.driveFolderId };
}
