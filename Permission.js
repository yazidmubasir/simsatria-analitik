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

  // SISWA sudah menjadi identitas autentikasi resmi, tetapi tidak diberi
  // akses pengelolaan database sekolah.
  SISWA: [
    "VIEW_DASHBOARD",
    "VIEW_PRESENSI",
    "VIEW_MONITORING",
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
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
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
      npsn: context.npsn,
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

/**
 * SINKRONISASI PENGGUNA SEKOLAH
 *
 * Hanya ADMIN_SEKOLAH/SUPERADMIN yang dapat menjalankannya.
 */
function normalizeSchoolUserList_(input) {
  const tokens = String(input || "")
    .split(/[\n,;]+/)
    .map(function (item) { return String(item || "").trim(); })
    .filter(Boolean);

  const seen = {};
  const result = [];

  tokens.forEach(function (token) {
    const parts = token.split("|").map(function (part) { return String(part || "").trim(); });
    const email = normalizeEmail_(parts[0]);
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
    if (seen[email]) return;
    seen[email] = true;
    result.push({ email: email, nama: parts[1] || "", nip: parts[2] || "" });
  });

  return result;
}

function syncAllSchoolUsers(emailList) {
  const context = requireUserManager_();
  const entries = normalizeSchoolUserList_(emailList);
  if (!entries.length) throw new Error("Daftar email belum diisi atau tidak ada email yang valid.");

  const sheet = getOrCreateUsersSheet_(context.school.spreadsheetId);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader_);
  const index = {};
  headers.forEach(function (header, i) { index[header] = i; });

  ["USER_ID", "EMAIL", "NIP", "NAMA", "ROLE", "STATUS"].forEach(function (header) {
    if (index[header] === undefined) throw new Error("Kolom " + header + " tidak ditemukan di USERS.");
  });

  const rowByEmail = {};
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][index.EMAIL]);
    if (email) rowByEmail[email] = i + 1;
  }

  const result = {
    success: true,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    total: entries.length,
    created: 0,
    updated: 0,
    failed: 0,
    users: [],
  };

  entries.forEach(function (entry) {
    try {
      const email = entry.email;
      const rowNumber = rowByEmail[email] || -1;
      let existing = null;

      if (rowNumber > 0) {
        const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
        existing = {
          userId: String(rowValues[index.USER_ID] || "").trim(),
          nip: String(rowValues[index.NIP] || "").trim(),
          nama: String(rowValues[index.NAMA] || "").trim(),
          role: normalizeRole_(rowValues[index.ROLE]),
          status: normalizeRole_(rowValues[index.STATUS]),
        };
      }

      const userId = existing && existing.userId
        ? existing.userId
        : "USR-" + Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase();
      const nip = entry.nip || (existing ? existing.nip : "");
      const nama = entry.nama || (existing ? existing.nama : "Guru SIM SATRIA");
      const role = existing && existing.role ? existing.role : "GURU";
      const status = existing && existing.status ? existing.status : "ACTIVE";

      if (!USER_MANAGEMENT_CONFIG.ALLOWED_ROLES.includes(role)) throw new Error("Role " + role + " tidak diizinkan untuk sinkronisasi.");
      if (!["ACTIVE", "INACTIVE"].includes(status)) throw new Error("Status " + status + " tidak valid.");

      const row = new Array(headers.length).fill("");
      row[index.USER_ID] = userId;
      row[index.EMAIL] = email;
      row[index.NIP] = nip;
      row[index.NAMA] = nama;
      row[index.ROLE] = role;
      row[index.STATUS] = status;

      if (rowNumber > 0) {
        sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
        result.updated++;
      } else {
        const newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1, 1, headers.length).setValues([row]);
        rowByEmail[email] = newRow;
        result.created++;
      }

      const userRecord = { USER_ID: userId, EMAIL: email, NIP: nip, NAMA: nama, ROLE: role, STATUS: status };
      registerSchoolUserBinding_(context, userRecord);

      if (status === "ACTIVE" && role !== "SISWA") grantSchoolSpreadsheetEditor_(context.school.spreadsheetId, email);
      else if (role !== "SISWA") revokeSchoolSpreadsheetAccess_(context.school.spreadsheetId, email);

      clearUserContextCache_(email);

      result.users.push({ email: email, nama: nama, nip: nip, role: role, status: status, userId: userId, success: true });
    } catch (e) {
      result.failed++;
      result.users.push({ email: entry.email, nama: entry.nama || "", success: false, error: e.message || String(e) });
    }
  });

  result.message = "Sinkronisasi selesai: " + result.created + " dibuat, " + result.updated + " diperbarui, " + result.failed + " gagal.";
  return result;
}

function syncAllExistingSchoolUsers() {
  const context = requireUserManager_();
  const sheet = getOrCreateUsersSheet_(context.school.spreadsheetId);
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { success: true, total: 0, created: 0, updated: 0, failed: 0, users: [], message: "Belum ada pengguna pada USERS sekolah." };

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf("EMAIL");
  const nameIndex = headers.indexOf("NAMA");
  const nipIndex = headers.indexOf("NIP");
  if (emailIndex < 0) throw new Error("Kolom EMAIL tidak ditemukan di USERS.");

  const entries = [];
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    if (!email) continue;
    entries.push({
      email: email,
      nama: nameIndex >= 0 ? String(values[i][nameIndex] || "").trim() : "",
      nip: nipIndex >= 0 ? String(values[i][nipIndex] || "").trim() : "",
    });
  }

  return syncAllSchoolUsers(entries.map(function (entry) {
    return entry.email + "|" + entry.nama + "|" + entry.nip;
  }).join("\n"));
}
