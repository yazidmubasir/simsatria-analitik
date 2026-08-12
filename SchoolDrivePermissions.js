/**
 * SIM SATRIA - SCHOOL DRIVE PERMISSIONS
 *
 * Sinkronisasi akses Drive sekolah untuk GURU/WALI_KELAS.
 * Tidak pernah membuka Spreadsheet MASTER.
 *
 * Root sekolah dan folder modul penting diberikan akses Editor
 * secara eksplisit. Ini menghindari masalah inheritance permission
 * pada folder Drive yang digunakan oleh Web App user-accessing.
 */

const SCHOOL_DRIVE_PERMISSION_CONFIG = {
  TEACHER_ROLES: ['GURU', 'WALI_KELAS'],
  ACTIVE_STATUS: 'ACTIVE',
  MODULE_FOLDERS: ['PRESENSI', 'AGENDA'],
};

function getSchoolPresensiFolder_() {
  return getSchoolModuleFolder_('PRESENSI');
}

function getSchoolAgendaFolder_() {
  return getSchoolModuleFolder_('AGENDA');
}

function getSchoolDrivePermissionFolders_() {
  const root = getSchoolDriveFolder_();
  const folders = {
    root: root,
    PRESENSI: getSchoolPresensiFolder_(),
    AGENDA: getSchoolAgendaFolder_(),
  };
  return folders;
}

function addEditorVerified_(folder, email, label) {
  email = normalizeEmail_(email);
  if (!email) throw new Error('Email guru kosong.');

  folder.addEditor(email);

  // Verifikasi langsung. Jangan menganggap addEditor berhasil
  // apabila Google Drive menolak permission karena policy/domain.
  const editors = folder.getEditors().map(function(user) {
    return normalizeEmail_(user.getEmail());
  });

  if (editors.indexOf(email) < 0) {
    throw new Error(
      'Permission Editor untuk ' + email + ' belum terdeteksi pada folder ' +
      (label || folder.getName()) + '.'
    );
  }

  return true;
}

/** Memberikan akses Editor root sekolah + modul PRESENSI + AGENDA. */
function grantSchoolDriveEditor_(context, email) {
  email = normalizeEmail_(email);
  if (!email) throw new Error('Email guru kosong.');

  const folders = getSchoolDrivePermissionFolders_();

  addEditorVerified_(folders.root, email, 'ROOT SEKOLAH');
  addEditorVerified_(folders.PRESENSI, email, 'PRESENSI');
  addEditorVerified_(folders.AGENDA, email, 'AGENDA');

  try { folders.root.setShareableByEditors(false); } catch (e) {}
  try { folders.PRESENSI.setShareableByEditors(false); } catch (e) {}
  try { folders.AGENDA.setShareableByEditors(false); } catch (e) {}

  return {
    success: true,
    email: email,
    folderId: folders.root.getId(),
    folderName: folders.root.getName(),
    presensiFolderId: folders.PRESENSI.getId(),
    agendaFolderId: folders.AGENDA.getId(),
  };
}

/** Mencabut akses langsung guru terhadap root, PRESENSI, dan AGENDA. */
function revokeSchoolDriveEditor_(email) {
  email = normalizeEmail_(email);
  if (!email) return { success: false, email: '', skipped: true };

  const folders = getSchoolDrivePermissionFolders_();

  [folders.root, folders.PRESENSI, folders.AGENDA].forEach(function(folder) {
    try {
      folder.revokePermissions(email);
    } catch (e) {
      console.warn('[DRIVE PERMISSION] Gagal mencabut ' + folder.getName() + ' ' + email + ': ' + e.message);
    }
  });

  return {
    success: true,
    email: email,
    folderId: folders.root.getId(),
    presensiFolderId: folders.PRESENSI.getId(),
    agendaFolderId: folders.AGENDA.getId(),
  };
}

/**
 * Sinkronisasi seluruh guru pada sekolah aktif.
 * Sumber user hanya USERS pada spreadsheet sekolah aktif.
 */
function syncSchoolDrivePermissionsForTeachers() {
  const context = requireUserManager_();
  const spreadsheetId = String(context.school.spreadsheetId || '').trim();
  if (!spreadsheetId) throw new Error('Spreadsheet sekolah aktif tidak ditemukan.');

  const folders = getSchoolDrivePermissionFolders_();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || 'USERS');
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      folderId: folders.root.getId(),
      folderName: folders.root.getName(),
      presensiFolderId: folders.PRESENSI.getId(),
      agendaFolderId: folders.AGENDA.getId(),
      totalUsers: 0,
      activeTeachers: 0,
      granted: [],
      revoked: [],
      failed: [],
      message: 'Belum ada pengguna pada USERS.',
    };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf('EMAIL');
  const roleIndex = headers.indexOf('ROLE');
  const statusIndex = headers.indexOf('STATUS');
  if (emailIndex < 0 || roleIndex < 0 || statusIndex < 0) {
    throw new Error('Struktur USERS tidak lengkap. Diperlukan EMAIL, ROLE, STATUS.');
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
        addEditorVerified_(folders.root, email, 'ROOT SEKOLAH');
        addEditorVerified_(folders.PRESENSI, email, 'PRESENSI');
        addEditorVerified_(folders.AGENDA, email, 'AGENDA');
        granted.push(email);
      } catch (e) {
        failed.push({
          email: email,
          action: 'GRANT',
          error: e && e.message ? e.message : String(e),
        });
      }
    } else {
      try {
        [folders.root, folders.PRESENSI, folders.AGENDA].forEach(function(folder) {
          try { folder.revokePermissions(email); } catch (ignore) {}
        });
        revoked.push(email);
      } catch (e) {
        failed.push({
          email: email,
          action: 'REVOKE',
          error: e && e.message ? e.message : String(e),
        });
      }
    }
  }

  try { folders.root.setShareableByEditors(false); } catch (e) {}
  try { folders.PRESENSI.setShareableByEditors(false); } catch (e) {}
  try { folders.AGENDA.setShareableByEditors(false); } catch (e) {}

  return {
    success: failed.length === 0,
    npsn: context.npsn || '',
    sekolah: context.school.namaSekolah || '',
    spreadsheetId: spreadsheetId,
    folderId: folders.root.getId(),
    folderName: folders.root.getName(),
    presensiFolderId: folders.PRESENSI.getId(),
    presensiFolderName: folders.PRESENSI.getName(),
    agendaFolderId: folders.AGENDA.getId(),
    agendaFolderName: folders.AGENDA.getName(),
    totalUsers: values.length - 1,
    activeTeachers: granted.length,
    granted: granted,
    revoked: revoked,
    failed: failed,
    message: failed.length === 0
      ? 'Permission Editor root, PRESENSI, dan AGENDA berhasil disinkronkan untuk seluruh GURU/WALI_KELAS aktif.'
      : 'Sinkronisasi selesai tetapi ada permission yang gagal diproses. Periksa failed.',
  };
}

/** Dipanggil saat user dibuat/diubah. */
function syncTeacherSchoolDrivePermission_(context, email, role, status) {
  role = normalizeRole_(role);
  status = normalizeRole_(status);
  email = normalizeEmail_(email);

  if (SCHOOL_DRIVE_PERMISSION_CONFIG.TEACHER_ROLES.indexOf(role) < 0) {
    return { success: true, skipped: true, reason: 'ROLE_NOT_TEACHER', email: email };
  }

  try {
    return status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS
      ? grantSchoolDriveEditor_(context, email)
      : revokeSchoolDriveEditor_(email);
  } catch (e) {
    return { success: false, email: email, error: e && e.message ? e.message : String(e) };
  }
}

/**
 * Diagnostik permission root, PRESENSI, dan AGENDA.
 */
function diagnoseSchoolDrivePermissions() {
  const context = requireUserManager_();
  const folders = getSchoolDrivePermissionFolders_();
  const spreadsheetId = String(context.school.spreadsheetId || '').trim();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || 'USERS');
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      folderId: folders.root.getId(),
      folderName: folders.root.getName(),
      presensiFolderId: folders.PRESENSI.getId(),
      agendaFolderId: folders.AGENDA.getId(),
      teachers: [],
    };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf('EMAIL');
  const roleIndex = headers.indexOf('ROLE');
  const statusIndex = headers.indexOf('STATUS');
  if (emailIndex < 0 || roleIndex < 0 || statusIndex < 0) throw new Error('Struktur USERS tidak lengkap.');

  const editorEmails_ = function(folder) {
    return folder.getEditors().map(function(u) {
      return normalizeEmail_(u.getEmail());
    });
  };

  const rootEditors = editorEmails_(folders.root);
  const presensiEditors = editorEmails_(folders.PRESENSI);
  const agendaEditors = editorEmails_(folders.AGENDA);
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
      agendaEditor: agendaEditors.indexOf(email) >= 0,
    });
  }

  return {
    success: true,
    npsn: context.npsn || '',
    sekolah: context.school.namaSekolah || '',
    folderId: folders.root.getId(),
    folderName: folders.root.getName(),
    presensiFolderId: folders.PRESENSI.getId(),
    presensiFolderName: folders.PRESENSI.getName(),
    agendaFolderId: folders.AGENDA.getId(),
    agendaFolderName: folders.AGENDA.getName(),
    rootEditors: rootEditors,
    presensiEditors: presensiEditors,
    agendaEditors: agendaEditors,
    teachers: teachers,
  };
}
