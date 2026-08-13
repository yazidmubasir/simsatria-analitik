/**
 * SIM SATRIA - SCHOOL DRIVE PERMISSIONS
 *
 * User sekolah aktif diberikan Editor langsung pada folder Drive sekolah
 * agar upload/penyimpanan dari Web App USER_ACCESSING berjalan langsung.
 * Tidak menggunakan Write Gateway.
 * ADMIN_SEKOLAH tetap menjadi pengelola/editor folder sekolah.
 */

const SCHOOL_DRIVE_PERMISSION_CONFIG = {
  EDITOR_ROLES: ['GURU', 'WALI_KELAS', 'KARYAWAN', 'SISWA'],
  ACTIVE_STATUS: 'ACTIVE',
  MODULE_FOLDERS: ['PRESENSI', 'AGENDA'],
};

function getSchoolPresensiFolder_() { return getSchoolModuleFolder_('PRESENSI'); }
function getSchoolAgendaFolder_() { return getSchoolModuleFolder_('AGENDA'); }
function getSchoolDrivePermissionFolders_() {
  const root = getSchoolDriveFolder_();
  return { root: root, PRESENSI: getSchoolPresensiFolder_(), AGENDA: getSchoolAgendaFolder_() };
}

function addEditorVerified_(folder, email, label) {
  email = normalizeEmail_(email);
  if (!email) throw new Error('Email pengguna kosong.');
  try { folder.addEditor(email); } catch (e) {
    throw new Error('Gagal memberikan Editor pada ' + (label || folder.getName()) + ': ' + e.message);
  }
  const editors = folder.getEditors().map(function(user) { return normalizeEmail_(user.getEmail()); });
  if (editors.indexOf(email) < 0) {
    throw new Error('Permission Editor untuk ' + email + ' belum terdeteksi pada folder ' + (label || folder.getName()) + '.');
  }
  return true;
}

function grantSchoolDriveEditor_(context, email, role) {
  email = normalizeEmail_(email);
  role = normalizeRole_(role);
  if (SCHOOL_DRIVE_PERMISSION_CONFIG.EDITOR_ROLES.indexOf(role) < 0) {
    return { success: true, skipped: true, reason: 'ROLE_NOT_EDITOR_USER' };
  }
  const folders = getSchoolDrivePermissionFolders_();
  addEditorVerified_(folders.root, email, 'ROOT SEKOLAH');
  addEditorVerified_(folders.PRESENSI, email, 'PRESENSI');
  addEditorVerified_(folders.AGENDA, email, 'AGENDA');
  return {
    success: true,
    email: email,
    role: role,
    folderId: folders.root.getId(),
    folderName: folders.root.getName(),
    presensiFolderId: folders.PRESENSI.getId(),
    agendaFolderId: folders.AGENDA.getId(),
  };
}

function revokeSchoolDriveEditor_(email) {
  email = normalizeEmail_(email);
  if (!email) return { success: false, email: '', skipped: true };
  const folders = getSchoolDrivePermissionFolders_();
  [folders.root, folders.PRESENSI, folders.AGENDA].forEach(function(folder) {
    try { folder.revokePermissions(email); } catch (e) {}
  });
  return { success: true, email: email };
}

function syncSchoolDrivePermissionsForTeachers() {
  const context = requireUserManager_();
  const spreadsheetId = String(context.school.spreadsheetId || '').trim();
  if (!spreadsheetId) throw new Error('Spreadsheet sekolah aktif tidak ditemukan.');
  const folders = getSchoolDrivePermissionFolders_();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || 'USERS');
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { success: true, totalUsers: 0, granted: [], revoked: [], failed: [], message: 'Belum ada pengguna pada USERS.' };
  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf('EMAIL'), roleIndex = headers.indexOf('ROLE'), statusIndex = headers.indexOf('STATUS');
  if (emailIndex < 0 || roleIndex < 0 || statusIndex < 0) throw new Error('Struktur USERS tidak lengkap.');
  const granted = [], revoked = [], failed = [];
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    const role = normalizeRole_(values[i][roleIndex]);
    const status = normalizeRole_(values[i][statusIndex]);
    if (!email || SCHOOL_DRIVE_PERMISSION_CONFIG.EDITOR_ROLES.indexOf(role) < 0) continue;
    try {
      if (status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS) {
        grantSchoolDriveEditor_(context, email, role);
        granted.push(email);
      } else {
        revokeSchoolDriveEditor_(email);
        revoked.push(email);
      }
    } catch (e) {
      failed.push({ email: email, action: status === 'ACTIVE' ? 'GRANT_EDITOR' : 'REVOKE', error: e.message || String(e) });
    }
  }
  return {
    success: failed.length === 0,
    npsn: context.npsn || '',
    sekolah: context.school.namaSekolah || '',
    spreadsheetId: spreadsheetId,
    folderId: folders.root.getId(),
    folderName: folders.root.getName(),
    presensiFolderId: folders.PRESENSI.getId(),
    agendaFolderId: folders.AGENDA.getId(),
    granted: granted,
    revoked: revoked,
    failed: failed,
    message: failed.length === 0 ? 'Permission Editor berhasil disinkronkan untuk pengguna sekolah.' : 'Sinkronisasi selesai dengan beberapa kegagalan.',
  };
}

function syncTeacherSchoolDrivePermission_(context, email, role, status) {
  role = normalizeRole_(role); status = normalizeRole_(status); email = normalizeEmail_(email);
  if (SCHOOL_DRIVE_PERMISSION_CONFIG.EDITOR_ROLES.indexOf(role) < 0) return { success: true, skipped: true, reason: 'ROLE_NOT_EDITOR_USER', email: email };
  try {
    return status === SCHOOL_DRIVE_PERMISSION_CONFIG.ACTIVE_STATUS
      ? grantSchoolDriveEditor_(context, email, role)
      : revokeSchoolDriveEditor_(email);
  } catch (e) {
    return { success: false, email: email, error: e.message || String(e) };
  }
}

function diagnoseSchoolDrivePermissions() {
  const context = requireUserManager_();
  const folders = getSchoolDrivePermissionFolders_();
  const spreadsheetId = String(context.school.spreadsheetId || '').trim();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(USER_MANAGEMENT_CONFIG.SHEET || 'USERS');
  if (!sheet) throw new Error('Sheet "USERS" tidak ditemukan.');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { success: true, teachers: [] };
  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf('EMAIL'), roleIndex = headers.indexOf('ROLE'), statusIndex = headers.indexOf('STATUS');
  const editors_ = function(folder) { return folder.getEditors().map(function(u) { return normalizeEmail_(u.getEmail()); }); };
  const viewers_ = function(folder) { return folder.getViewers().map(function(u) { return normalizeEmail_(u.getEmail()); }); };
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]), role = normalizeRole_(values[i][roleIndex]), status = normalizeRole_(values[i][statusIndex]);
    if (!email || SCHOOL_DRIVE_PERMISSION_CONFIG.EDITOR_ROLES.indexOf(role) < 0) continue;
    result.push({
      email: email, role: role, status: status,
      rootEditor: editors_(folders.root).indexOf(email) >= 0, rootViewer: viewers_(folders.root).indexOf(email) >= 0,
      presensiEditor: editors_(folders.PRESENSI).indexOf(email) >= 0, presensiViewer: viewers_(folders.PRESENSI).indexOf(email) >= 0,
      agendaEditor: editors_(folders.AGENDA).indexOf(email) >= 0, agendaViewer: viewers_(folders.AGENDA).indexOf(email) >= 0,
    });
  }
  return { success: true, npsn: context.npsn, sekolah: context.school.namaSekolah, teachers: result };
}
