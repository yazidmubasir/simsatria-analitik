/**
 * SIM SATRIA - SCHOOL DRIVE PERMISSIONS
 *
 * Sinkronisasi akses Drive sekolah untuk GURU/WALI_KELAS.
 * Tidak pernah membuka Spreadsheet MASTER.
 *
 * PDF dan dokumen presensi disimpan pada:
 *   Drive Sekolah / PRESENSI
 */

const SCHOOL_DRIVE_PERMISSION_CONFIG = {
  TEACHER_ROLES: ["GURU", "WALI_KELAS"],
  ACTIVE_STATUS: "ACTIVE",
  PRESENSI_FOLDER: "PRESENSI",
};

function getSchoolPresensiFolder_() {
  const root = getSchoolDriveFolder_();
  const folders = root.getFoldersByName(SCHOOL_DRIVE_PERMISSION_CONFIG.PRESENSI_FOLDER);
  return folders.hasNext() ? folders.next() : root.createFolder(SCHOOL_DRIVE_PERMISSION_CONFIG.PRESENSI_FOLDER);
}

/** Memberikan akses Editor root sekolah dan folder PRESENSI. */
function grantSchoolDriveEditor_(context, email) {
  email = normalizeEmail_(email);
  if (!email) throw new Error("Email guru kosong.");

  const root = getSchoolDriveFolder_();
  const presensi = getSchoolPresensiFolder_();

  root.addEditor(email);
  presensi.addEditor(email);

  try { root.setShareableByEditors(false); } catch (e) {
    console.warn("[DRIVE PERMISSION] root setShareableByEditors gagal: " + e.message);
  }
  try { presensi.setShareableByEditors(false); } catch (e) {
    console.warn("[DRIVE PERMISSION] PRESENSI setShareableByEditors gagal: " + e.message);
  }

  return {
    success: true,
    email: email,
    folderId: root.getId(),
    folderName: root.getName(),
    presensiFolderId: presensi.getId(),
    presensiFolderName: presensi.getName(),
  };
}

/** Mencabut akses langsung guru terhadap root dan PRESENSI. */
function revokeSchoolDriveEditor_(email) {
  email = normalizeEmail_(email);
  if (!email) return { success: false, email: "", skipped: true };

  const root = getSchoolDriveFolder_();
  let presensi = null;
  try { presensi = getSchoolPresensiFolder_(); } catch (e) {}

  try { root.revokePermissions(email); } catch (e) {
    console.warn("[DRIVE PERMISSION] Gagal mencabut root " + email + ": " + e.message);
  }
  if (presensi) {
    try { presensi.revokePermissions(email); } catch (e) {
      console.warn("[DRIVE PERMISSION] Gagal mencabut PRESENSI " + email + ": " + e.message);
    }
  }

  return {
    success: true,
    email: email,
    folderId: root.getId(),
    folderName: root.getName(),
    presensiFolderId: presensi ? presensi.getId() : "",
  };
}

/**
 * Sinkronisasi seluruh guru pada sekolah aktif.
 * Jalankan sekali setelah deployment/update.
 * Sumber user hanya USERS pada spreadsheet sekolah aktif.
 */
function syncSchoolDrivePermissionsForTeachers() {
  const context = requireUserManager_();
  const spreadsheetId = String(context.school.spreadsheetId || "").trim();
  if (!spreadsheetId) throw new Error("Spreadsheet sekolah aktif tidak ditemukan.");

  const root = getSchoolDriveFolder_();
  const presensi = getSchoolPresensiFolder_();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || "USERS");
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      folderId: root.getId(),
      folderName: root.getName(),
      presensiFolderId: presensi.getId(),
      presensiFolderName: presensi.getName(),
      totalUsers: 0,
      activeTeachers: 0,
      granted: [], revoked: [], failed: [],
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

  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    const role = normalizeRole_(values[i][roleIndex]);
    const status = normalizeRole_(values[i][statusIndex]);
    if (!email || SCHOOL_DRIVE_PERMISSION_CONFIG.TEACHER_ROLES.indexOf(role) < 0) continue;

    if (status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS) {
      try {
        root.addEditor(email);
        presensi.addEditor(email);
        granted.push(email);
      } catch (e) {
        failed.push({ email: email, action: "GRANT", error: e.message || String(e) });
      }
    } else {
      try {
        root.revokePermissions(email);
        presensi.revokePermissions(email);
        revoked.push(email);
      } catch (e) {
        failed.push({ email: email, action: "REVOKE", error: e.message || String(e) });
      }
    }
  }

  try { root.setShareableByEditors(false); } catch (e) {
    failed.push({ email: "", action: "ROOT_SHARE", error: e.message || String(e) });
  }
  try { presensi.setShareableByEditors(false); } catch (e) {
    failed.push({ email: "", action: "PRESENSI_SHARE", error: e.message || String(e) });
  }

  return {
    success: failed.length === 0,
    npsn: context.npsn || "",
    sekolah: context.school.namaSekolah || "",
    spreadsheetId: spreadsheetId,
    folderId: root.getId(),
    folderName: root.getName(),
    presensiFolderId: presensi.getId(),
    presensiFolderName: presensi.getName(),
    totalUsers: values.length - 1,
    activeTeachers: granted.length,
    granted: granted,
    revoked: revoked,
    failed: failed,
    message: failed.length === 0
      ? "Permission root dan folder PRESENSI sekolah berhasil disinkronkan untuk seluruh GURU/WALI_KELAS aktif."
      : "Sinkronisasi selesai tetapi ada permission yang gagal diproses. Periksa failed.",
  };
}

/** Dipanggil saat user dibuat/diubah. */
function syncTeacherSchoolDrivePermission_(context, email, role, status) {
  role = normalizeRole_(role);
  status = normalizeRole_(status);
  email = normalizeEmail_(email);

  if (SCHOOL_DRIVE_PERMISSION_CONFIG.TEACHER_ROLES.indexOf(role) < 0) {
    return { success: true, skipped: true, reason: "ROLE_NOT_TEACHER", email: email };
  }

  try {
    return status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS
      ? grantSchoolDriveEditor_(context, email)
      : revokeSchoolDriveEditor_(email);
  } catch (e) {
    return { success: false, email: email, error: e.message || String(e) };
  }
}

/**
 * Diagnostik. Menunjukkan permission root dan PRESENSI sekaligus.
 */
function diagnoseSchoolDrivePermissions() {
  const context = requireUserManager_();
  const root = getSchoolDriveFolder_();
  const presensi = getSchoolPresensiFolder_();
  const spreadsheetId = String(context.school.spreadsheetId || "").trim();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || "USERS");
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      folderId: root.getId(),
      folderName: root.getName(),
      presensiFolderId: presensi.getId(),
      presensiFolderName: presensi.getName(),
      teachers: [],
    };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf("EMAIL");
  const roleIndex = headers.indexOf("ROLE");
  const statusIndex = headers.indexOf("STATUS");
  if (emailIndex < 0 || roleIndex < 0 || statusIndex < 0) throw new Error("Struktur USERS tidak lengkap.");

  const rootEditors = root.getEditors().map(function(u) { return normalizeEmail_(u.getEmail()); });
  const presensiEditors = presensi.getEditors().map(function(u) { return normalizeEmail_(u.getEmail()); });
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
      rootEditor: rootEditors.indexOf(email) >= 0,
      presensiEditor: presensiEditors.indexOf(email) >= 0,
    });
  }

  return {
    success: true,
    npsn: context.npsn || "",
    sekolah: context.school.namaSekolah || "",
    folderId: root.getId(),
    folderName: root.getName(),
    presensiFolderId: presensi.getId(),
    presensiFolderName: presensi.getName(),
    rootEditors: rootEditors,
    presensiEditors: presensiEditors,
    teachers: teachers,
  };
}
