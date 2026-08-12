/**
 * PERMISSION SERVICE
 * Role menentukan apa yang boleh dilakukan.
 * NPSN selalu berasal dari AUTH/SCHOOL CONTEXT.
 */
const ROLE_PERMISSIONS = {
  SUPERADMIN: ["*"],
  SUPER_ADMIN: ["*"],

  ADMIN_SEKOLAH: [
    "VIEW_DASHBOARD",
    "MANAGE_DATABASE",
    "MANAGE_GURU",
    "MANAGE_KARYAWAN",
    "MANAGE_SISWA",
    "MANAGE_KELAS",
    "MANAGE_USERS",
    "MANAGE_CONFIG",
    "VIEW_LOG",
    "INPUT_MONITORING",
    "VIEW_MONITORING",
    "INPUT_PRESENSI",
    "VIEW_PRESENSI",
    "VIEW_ANALYTICS",
    "MANAGE_AGENDA",
    "MANAGE_PRESTASI",
    "MANAGE_SBI",
    "MANAGE_PARKIR",
    "MANAGE_KEBERSIHAN",
    "MANAGE_KEAMANAN",
    "MANAGE_KERJA",
  ],

  GURU: [
    "VIEW_DASHBOARD",
    "INPUT_PRESENSI",
    "VIEW_PRESENSI",
    "INPUT_MONITORING",
    "VIEW_MONITORING",
    "MANAGE_AGENDA",
    "MANAGE_PRESTASI",
  ],

  WALI_KELAS: [
    "VIEW_DASHBOARD",
    "INPUT_PRESENSI",
    "VIEW_PRESENSI",
    "VIEW_MONITORING",
    "MANAGE_AGENDA",
    "MANAGE_PRESTASI",
  ],

  KARYAWAN: [
    "VIEW_DASHBOARD",
    "INPUT_MONITORING",
    "VIEW_MONITORING",
    "MANAGE_PARKIR",
    "MANAGE_KEBERSIHAN",
    "MANAGE_KEAMANAN",
    "MANAGE_KERJA",
  ],
};

function normalizePermission_(permission) {
  return String(permission || "").trim().toUpperCase();
}

function normalizeRole_(role) {
  return String(role || "").trim().toUpperCase();
}

function getRolePermissions_(role) {
  const normalizedRole = normalizeRole_(role);
  return (ROLE_PERMISSIONS[normalizedRole] || []).slice();
}

function hasPermission(permission, role) {
  const required = normalizePermission_(permission);
  if (!required) return false;
  const permissions = getRolePermissions_(role);
  return permissions.includes("*") || permissions.includes(required);
}

function requirePermission(permission) {
  const required = normalizePermission_(permission);
  if (!required) {
    throw new Error("Permission wajib tidak boleh kosong.");
  }

  const context = getCurrentUserContext();
  if (hasPermission(required, context.role)) {
    return true;
  }

  throw new Error(
    "Anda tidak memiliki izin untuk melakukan tindakan ini. Role " +
      context.role +
      " tidak memiliki permission " +
      required +
      ".",
  );
}

function hasCurrentUserPermission_(permission) {
  try {
    requirePermission(permission);
    return true;
  } catch (e) {
    return false;
  }
}

function getMyPermissions() {
  const context = getCurrentUserContext();
  return {
    success: true,
    role: context.role,
    permissions: getRolePermissions_(context.role),
  };
}

function checkPermission(permission) {
  try {
    const context = getCurrentUserContext();
    const required = normalizePermission_(permission);
    return {
      success: true,
      role: context.role,
      permission: required,
      allowed: hasPermission(required, context.role),
    };
  } catch (e) {
    return {
      success: false,
      allowed: false,
      message: e.message,
    };
  }
}

function requireSchoolContext() {
  const context = getCurrentUserContext();
  if (!context || !context.school || !context.school.spreadsheetId) {
    throw new Error("School Context belum tersedia.");
  }
  return context;
}

function requireSchoolAdmin() {
  const context = requireSchoolContext();
  if (!hasPermission("MANAGE_DATABASE", context.role)) {
    throw new Error("Akses hanya tersedia untuk ADMIN_SEKOLAH.");
  }
  return context;
}
