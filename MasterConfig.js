/**
 * SIM SATRIA - MASTER CONFIG
 *
 * MASTER adalah sumber kebenaran untuk:
 *   - SCHOOLS
 *   - ADMIN_SEKOLAH
 *
 * GURU / WALI_KELAS / KARYAWAN / SISWA TIDAK disimpan di MASTER.
 * Mereka dikelola pada USERS di Spreadsheet sekolah masing-masing.
 *
 * PENTING:
 * Web App harus DEPLOY "Execute as me" (pemilik aplikasi), sehingga
 * pembacaan MASTER dilakukan oleh akun pemilik aplikasi, bukan oleh akun
 * guru/admin yang sedang login. Identitas pengguna tetap diambil dari
 * Session.getActiveUser().getEmail().
 *
 * Dengan pola ini login ADMIN_SEKOLAH tidak lagi bergantung pada
 * syncMasterAuthRegistry(). Fungsi sinkronisasi hanya dipertahankan sebagai
 * alat maintenance/diagnostik dan bukan bagian dari alur login.
 */

const MASTER_AUTH_REGISTRY = {
  ADMIN_PREFIX: "SIM_SATRIA_MASTER_ADMIN_",
  SCHOOL_PREFIX: "SIM_SATRIA_MASTER_SCHOOL_",
  META_KEY: "SIM_SATRIA_MASTER_AUTH_REGISTRY_V1",
};

function getMasterSpreadsheetId_() {
  const id = String(
    PropertiesService.getScriptProperties().getProperty("MASTER_SPREADSHEET_ID") || "",
  ).trim();
  if (!id) {
    throw new Error(
      "MASTER_SPREADSHEET_ID belum dikonfigurasi. Jalankan setupMasterSpreadsheetId() sekali sebagai pemilik aplikasi.",
    );
  }
  return id;
}

function setupMasterSpreadsheetId() {
  const MASTER_ID = "1o7l24gGB7rXsjyFJJw1plz_0ud-V74USqyLQsFbZmS0";
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const schools = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");
  const admins = ss.getSheetByName("ADMIN_SEKOLAH");
  if (!schools) throw new Error('Sheet "SCHOOLS" tidak ditemukan pada MASTER.');
  if (!admins) throw new Error('Sheet "ADMIN_SEKOLAH" tidak ditemukan pada MASTER.');

  PropertiesService.getScriptProperties().setProperty("MASTER_SPREADSHEET_ID", MASTER_ID);
  validateMasterAdminSheet_(admins);

  return {
    success: true,
    message: "Spreadsheet SIM SATRIA MASTER berhasil dikonfigurasi.",
    spreadsheetId: MASTER_ID,
    spreadsheetName: ss.getName(),
  };
}

function validateMasterAdminSheet_(sheet) {
  if (!sheet) throw new Error("Sheet ADMIN_SEKOLAH tidak ditemukan pada MASTER.");
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("Sheet ADMIN_SEKOLAH masih kosong.");

  const headers = values[0].map(h => String(h || "").trim().toUpperCase());
  ["USER_ID","EMAIL","NIP","NAMA","NPSN","ROLE","STATUS"].forEach(header => {
    if (headers.indexOf(header) < 0) {
      throw new Error('Kolom "' + header + '" wajib ada pada ADMIN_SEKOLAH.');
    }
  });

  const ix = {};
  headers.forEach((h,i) => ix[h] = i);
  const emails = {};
  const npsns = {};
  let count = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const email = normalizeEmail_(row[ix.EMAIL]);
    const npsn = normalizeNpsn_(row[ix.NPSN]);
    const role = normalizeAuthRole_(row[ix.ROLE]);
    const status = String(row[ix.STATUS] || "").trim().toUpperCase();
    if (!email && !npsn && !role && !status) continue;

    if (!email) throw new Error("ADMIN_SEKOLAH baris " + (r + 1) + ": EMAIL wajib diisi.");
    if (!npsn) throw new Error("ADMIN_SEKOLAH baris " + (r + 1) + ": NPSN wajib diisi.");
    if (role !== "ADMIN_SEKOLAH") throw new Error("ADMIN_SEKOLAH baris " + (r + 1) + ": ROLE harus ADMIN_SEKOLAH.");
    if (!["ACTIVE","INACTIVE"].includes(status)) throw new Error("ADMIN_SEKOLAH baris " + (r + 1) + ": STATUS harus ACTIVE atau INACTIVE.");
    if (emails[email]) throw new Error("Email ADMIN_SEKOLAH duplikat: " + email);
    if (npsns[npsn]) throw new Error("NPSN ADMIN_SEKOLAH duplikat: " + npsn);

    emails[email] = true;
    npsns[npsn] = true;
    count++;
  }

  if (!count) throw new Error("Belum ada ADMIN_SEKOLAH pada MASTER.");
  return { success:true, adminCount:count, headers:headers };
}

/** Maintenance only. LOGIN tidak memanggil fungsi ini. */
function syncMasterAuthRegistry() {
  const caller = getGoogleUserEmail_();
  if (!isSuperAdminEmail_(caller)) throw new Error("Hanya SUPERADMIN yang boleh melakukan sinkronisasi MASTER.");

  const ss = SpreadsheetApp.openById(getMasterSpreadsheetId_());
  const adminSheet = ss.getSheetByName("ADMIN_SEKOLAH");
  const schoolSheet = ss.getSheetByName("SCHOOLS") || ss.getSheetByName("schools");
  validateMasterAdminSheet_(adminSheet);
  if (!schoolSheet) throw new Error("Sheet SCHOOLS tidak ditemukan pada MASTER.");

  const props = PropertiesService.getScriptProperties();
  const old = props.getProperties();
  Object.keys(old).forEach(key => {
    if (key.indexOf(MASTER_AUTH_REGISTRY.ADMIN_PREFIX) === 0 || key.indexOf(MASTER_AUTH_REGISTRY.SCHOOL_PREFIX) === 0) {
      props.deleteProperty(key);
    }
  });

  const admins = sheetValuesToObjects_(adminSheet);
  const schools = sheetValuesToObjects_(schoolSheet);
  const updates = {};
  let adminCount = 0;
  let schoolCount = 0;

  admins.forEach(row => {
    const email = normalizeEmail_(row.EMAIL);
    if (!email) return;
    updates[MASTER_AUTH_REGISTRY.ADMIN_PREFIX + registrySafeKey_(email)] = JSON.stringify(row);
    adminCount++;
  });
  schools.forEach(row => {
    const npsn = normalizeNpsn_(row.NPSN);
    if (!npsn) return;
    updates[MASTER_AUTH_REGISTRY.SCHOOL_PREFIX + registrySafeKey_(npsn)] = JSON.stringify(row);
    schoolCount++;
  });
  updates[MASTER_AUTH_REGISTRY.META_KEY] = JSON.stringify({
    syncedAt:new Date().toISOString(), masterSpreadsheetId:getMasterSpreadsheetId_(), adminCount, schoolCount
  });
  props.setProperties(updates, false);
  return {success:true, adminCount, schoolCount, message:"Registry maintenance berhasil disinkronkan."};
}

function getLocalMasterAdminByEmail_(email) {
  const key = MASTER_AUTH_REGISTRY.ADMIN_PREFIX + registrySafeKey_(normalizeEmail_(email));
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function getLocalMasterSchoolByNpsn_(npsn) {
  const key = MASTER_AUTH_REGISTRY.SCHOOL_PREFIX + registrySafeKey_(normalizeNpsn_(npsn));
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function getMasterAuthRegistryStatus() {
  const raw = PropertiesService.getScriptProperties().getProperty(MASTER_AUTH_REGISTRY.META_KEY);
  if (!raw) return {success:false, configured:false, message:"Registry maintenance belum pernah disinkronkan."};
  try { return {success:true, configured:true, meta:JSON.parse(raw)}; }
  catch(e) { return {success:false, configured:false, message:"Registry tidak valid."}; }
}

function sheetValuesToObjects_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || "").trim().toUpperCase());
  return values.slice(1).filter(row => row.some(v => String(v ?? "").trim() !== "")).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}

function registrySafeKey_(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9._-]/g, "_");
}
