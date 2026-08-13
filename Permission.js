/**
 * PERMISSION SERVICE
 * Role menentukan apa yang boleh dilakukan.
 * NPSN selalu berasal dari AUTH/SCHOOL CONTEXT.
 */
const ROLE_PERMISSIONS = {
  SUPERADMIN: ["*"],
  SUPER_ADMIN: ["*"],

  ADMIN_SEKOLAH: [
    "VIEW_DASHBOARD", "MANAGE_DATABASE", "MANAGE_GURU", "MANAGE_KARYAWAN",
    "MANAGE_SISWA", "MANAGE_KELAS", "MANAGE_USERS", "MANAGE_CONFIG", "VIEW_LOG",
    "INPUT_MONITORING", "VIEW_MONITORING", "INPUT_PRESENSI", "VIEW_PRESENSI",
    "VIEW_ANALYTICS", "MANAGE_AGENDA", "MANAGE_PRESTASI", "MANAGE_SBI",
    "MANAGE_PARKIR", "MANAGE_KEBERSIHAN", "MANAGE_KEAMANAN", "MANAGE_KERJA",
  ],

  GURU: [
    "VIEW_DASHBOARD", "INPUT_PRESENSI", "VIEW_PRESENSI", "INPUT_MONITORING",
    "VIEW_MONITORING", "MANAGE_AGENDA", "MANAGE_PRESTASI", "PRINT_PDF", "UPLOAD_FILE",
  ],

  WALI_KELAS: [
    "VIEW_DASHBOARD", "INPUT_PRESENSI", "VIEW_PRESENSI", "INPUT_MONITORING",
    "VIEW_MONITORING", "MANAGE_AGENDA", "MANAGE_PRESTASI", "PRINT_PDF", "UPLOAD_FILE",
  ],

  KARYAWAN: [
    "VIEW_DASHBOARD", "INPUT_MONITORING", "VIEW_MONITORING", "MANAGE_PARKIR",
    "MANAGE_KEBERSIHAN", "MANAGE_KEAMANAN", "MANAGE_KERJA", "PRINT_PDF", "UPLOAD_FILE",
  ],

  SISWA: [
    "VIEW_DASHBOARD", "VIEW_PRESENSI", "VIEW_MONITORING", "INPUT_MONITORING",
    "PRINT_PDF", "UPLOAD_FILE",
  ],
};

function normalizePermission_(permission) { return String(permission || "").trim().toUpperCase(); }
function normalizeRole_(role) { return String(role || "").trim().toUpperCase(); }
function getRolePermissions_(role) { return (ROLE_PERMISSIONS[normalizeRole_(role)] || []).slice(); }
function hasPermission(permission, role) {
  const required = normalizePermission_(permission);
  if (!required) return false;
  const permissions = getRolePermissions_(role);
  return permissions.includes("*") || permissions.includes(required);
}
function requirePermission(permission) {
  const required = normalizePermission_(permission);
  if (!required) throw new Error("Permission wajib tidak boleh kosong.");
  const context = getCurrentUserContext();
  if (hasPermission(required, context.role)) return true;
  throw new Error("Anda tidak memiliki izin untuk melakukan tindakan ini. Role " + context.role + " tidak memiliki permission " + required + ".");
}
function hasCurrentUserPermission_(permission) { try { requirePermission(permission); return true; } catch (e) { return false; } }
function getMyPermissions() {
  const context = getCurrentUserContext();
  return { success: true, role: context.role, npsn: context.npsn, sekolah: context.school.namaSekolah, permissions: getRolePermissions_(context.role) };
}
function checkPermission(permission) {
  try {
    const context = getCurrentUserContext();
    const required = normalizePermission_(permission);
    return { success: true, role: context.role, npsn: context.npsn, permission: required, allowed: hasPermission(required, context.role) };
  } catch (e) { return { success: false, allowed: false, message: e.message }; }
}
function requireSchoolContext() {
  const context = getCurrentUserContext();
  if (!context || !context.school || !context.school.spreadsheetId) throw new Error("School Context belum tersedia.");
  return context;
}
function requireSchoolAdmin() {
  const context = requireSchoolContext();
  if (!hasPermission("MANAGE_DATABASE", context.role)) throw new Error("Akses hanya tersedia untuk ADMIN_SEKOLAH.");
  return context;
}

/**
 * Semua akses file fisik user biasa dibuat VIEWER.
 * Hak tulis diberikan hanya melalui Write Gateway yang Execute as = Me.
 */
function grantSchoolSpreadsheetViewer_(spreadsheetId, email) {
  if (!spreadsheetId || !email) throw new Error("Data akses spreadsheet tidak lengkap.");
  const file = DriveApp.getFileById(spreadsheetId);
  try { file.addViewer(email); } catch (e) {
    throw new Error("Gagal memberikan Viewer database sekolah kepada " + email + ": " + e.message);
  }
  try { file.revokePermissions(email); } catch (e) {}
  try { file.addViewer(email); } catch (e) {}
}

function grantSchoolSpreadsheetEditor_(spreadsheetId, email) {
  if (!spreadsheetId || !email) throw new Error("Data akses spreadsheet tidak lengkap.");
  const context = getCurrentUserContext();
  const role = normalizeRole_(context.role);
  if (["GURU", "WALI_KELAS", "KARYAWAN", "SISWA"].includes(role)) {
    return grantSchoolSpreadsheetViewer_(spreadsheetId, email);
  }
  const file = DriveApp.getFileById(spreadsheetId);
  file.addEditor(email);
  try { file.setShareableByEditors(false); } catch (e) {}
}

function revokeSchoolSpreadsheetAccess_(spreadsheetId, email) {
  if (!spreadsheetId || !email) return;
  try { DriveApp.getFileById(spreadsheetId).revokePermissions(email); } catch (e) { console.warn("[USER] Gagal mencabut akses spreadsheet " + email + ": " + e.message); }
}

function requireUserManager_() {
  const context = getCurrentUserContext();
  if (!["ADMIN_SEKOLAH", "SUPERADMIN"].includes(normalizeRole_(context.role))) throw new Error("Menu Manajemen Pengguna hanya dapat digunakan oleh ADMIN_SEKOLAH.");
  return context;
}
