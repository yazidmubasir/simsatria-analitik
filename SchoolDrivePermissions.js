/**
 * SIM SATRIA - SCHOOL DRIVE PERMISSIONS
 *
 * Sinkronisasi akses folder Drive sekolah untuk GURU/WALI_KELAS.
 * Tidak pernah membuka Spreadsheet MASTER.
 *
 * Prinsip:
 * - sumber user: USERS pada spreadsheet sekolah aktif
 * - sumber folder: context.school.driveFolderId
 * - ACTIVE GURU/WALI_KELAS -> Editor folder sekolah
 * - INACTIVE GURU/WALI_KELAS -> akses langsung folder dicabut
 * - KARYAWAN tidak ikut sinkronisasi ini
 * - Editor tidak boleh membagikan ulang folder
 */

const SCHOOL_DRIVE_PERMISSION_CONFIG = {
  TEACHER_ROLES: ["GURU", "WALI_KELAS"],
  ACTIVE_STATUS: "ACTIVE",
};

/**
 * Memberikan akses Editor ke root Drive sekolah untuk satu guru.
 * Aman dipanggil berulang kali; addEditor tidak membuat permission ganda.
 */
function grantSchoolDriveEditor_(context, email) {
  email = normalizeEmail_(email);
  if (!email) throw new Error("Email guru kosong.");

  const folder = getSchoolDriveFolder_();
  folder.addEditor(email);

  // Guru dapat mengelola file, tetapi tidak boleh mengubah permission
  // dan membagikan folder sekolah kepada pihak lain.
  try {
    folder.setShareableByEditors(false);
  } catch (e) {
    console.warn("[DRIVE PERMISSION] setShareableByEditors gagal: " + e.message);
  }

  return {
    success: true,
    email: email,
    folderId: folder.getId(),
    folderName: folder.getName(),
  };
}

/**
 * Mencabut akses langsung guru terhadap root Drive sekolah.
 */
function revokeSchoolDriveEditor_(email) {
  email = normalizeEmail_(email);
  if (!email) return { success: false, email: "", skipped: true };

  const folder = getSchoolDriveFolder_();
  try {
    folder.revokePermissions(email);
  } catch (e) {
    console.warn("[DRIVE PERMISSION] Gagal mencabut " + email + ": " + e.message);
  }

  return {
    success: true,
    email: email,
    folderId: folder.getId(),
    folderName: folder.getName(),
  };
}

/**
 * Sinkronisasi seluruh guru pada sekolah aktif.
 * Jalankan sebagai ADMIN_SEKOLAH.
 *
 * Fungsi ini membaca USERS dari spreadsheet sekolah aktif, bukan MASTER.
 */
function syncSchoolDrivePermissionsForTeachers() {
  const context = requireUserManager_();
  const spreadsheetId = String(context.school.spreadsheetId || "").trim();
  if (!spreadsheetId) {
    throw new Error("Spreadsheet sekolah aktif tidak ditemukan.");
  }

  const folder = getSchoolDriveFolder_();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || "USERS");
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      folderId: folder.getId(),
      folderName: folder.getName(),
      totalUsers: 0,
      activeTeachers: 0,
      granted: [],
      revoked: [],
      failed: [],
      message: "Belum ada pengguna pada USERS.",
    };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf("EMAIL");
  const roleIndex = headers.indexOf("ROLE");
  const statusIndex = headers.indexOf("STATUS");

  if (emailIndex < 0 || roleIndex < 0 || statusIndex < 0) {
    throw new Error("Struktur USERS tidak lengkap. Diperlukan EMAIL, ROLE, STATUS.");
  }

  const granted = [];
  const revoked = [];
  const failed = [];
  const activeTeacherEmails = [];

  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    const role = normalizeRole_(values[i][roleIndex]);
    const status = normalizeRole_(values[i][statusIndex]);

    if (!email) continue;
    if (SCHOOL_DRIVE_PERMISSION_CONFIG.TEACHER_ROLES.indexOf(role) < 0) continue;

    if (status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS) {
      activeTeacherEmails.push(email);
      try {
        folder.addEditor(email);
        granted.push(email);
      } catch (e) {
        failed.push({
          email: email,
          action: "GRANT",
          error: e.message || String(e),
        });
      }
    } else {
      try {
        folder.revokePermissions(email);
        revoked.push(email);
      } catch (e) {
        failed.push({
          email: email,
          action: "REVOKE",
          error: e.message || String(e),
        });
      }
    }
  }

  try {
    folder.setShareableByEditors(false);
  } catch (e) {
    failed.push({
      email: "",
      action: "SET_SHAREABLE_BY_EDITORS",
      error: e.message || String(e),
    });
  }

  return {
    success: failed.length === 0,
    npsn: context.npsn || "",
    sekolah: context.school.namaSekolah || "",
    spreadsheetId: spreadsheetId,
    folderId: folder.getId(),
    folderName: folder.getName(),
    totalUsers: values.length - 1,
    activeTeachers: activeTeacherEmails.length,
    granted: granted,
    revoked: revoked,
    failed: failed,
    message:
      failed.length === 0
        ? "Permission Drive sekolah berhasil disinkronkan untuk seluruh GURU/WALI_KELAS aktif."
        : "Sinkronisasi selesai tetapi ada permission yang gagal diproses. Periksa daftar failed.",
  };
}

/**
 * Dipanggil saat user GURU/WALI_KELAS dibuat atau statusnya diubah.
 * Kegagalan permission Drive tidak membatalkan transaksi USERS.
 */
function syncTeacherSchoolDrivePermission_(context, email, role, status) {
  role = normalizeRole_(role);
  status = normalizeRole_(status);
  email = normalizeEmail_(email);

  if (SCHOOL_DRIVE_PERMISSION_CONFIG.TEACHER_ROLES.indexOf(role) < 0) {
    return { success: true, skipped: true, reason: "ROLE_NOT_TEACHER", email: email };
  }

  try {
    if (status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS) {
      return grantSchoolDriveEditor_(context, email);
    }
    return revokeSchoolDriveEditor_(email);
  } catch (e) {
    return {
      success: false,
      email: email,
      error: e.message || String(e),
    };
  }
}

/**
 * Diagnostik permission Drive sekolah.
 */
function diagnoseSchoolDrivePermissions() {
  const context = requireUserManager_();
  const folder = getSchoolDriveFolder_();
  const spreadsheetId = String(context.school.spreadsheetId || "").trim();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || "USERS");
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      folderId: folder.getId(),
      folderName: folder.getName(),
      teachers: [],
    };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf("EMAIL");
  const roleIndex = headers.indexOf("ROLE");
  const statusIndex = headers.indexOf("STATUS");
  if (emailIndex < 0 || roleIndex < 0 || statusIndex < 0) {
    throw new Error("Struktur USERS tidak lengkap.");
  }

  const editors = folder.getEditors().map(function (user) {
    return normalizeEmail_(user.getEmail());
  });

  const teachers = [];
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    const role = normalizeRole_(values[i][roleIndex]);
    const status = normalizeRole_(values[i][statusIndex]);
    if (!email || SCHOOL_DRIVE_PERMISSION_CONFIG.TEACHER_ROLES.indexOf(role) < 0) continue;

    teachers.push({
      email: email,
      role: role,
      status: status,
      isEditor: editors.indexOf(email) >= 0,
    });
  }

  return {
    success: true,
    npsn: context.npsn || "",
    sekolah: context.school.namaSekolah || "",
    folderId: folder.getId(),
    folderName: folder.getName(),
    shareableByEditors: folder.isShareableByEditors(),
    teachers: teachers,
    editors: editors,
  };
}
