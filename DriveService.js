/**
 * DRIVE SERVICE
 * Semua operasi Drive mengikuti School Context sekolah aktif.
 */

function getSchoolDriveFolder_() {
  const c = getCurrentUserContext() || {};
  const school = c.school || {};

  const candidates = [
    school.driveFolderId,
    school.driveRootFolderId,
    c.driveFolderId,
    c.driveRootFolderId,
  ];

  const ids = [];
  candidates.forEach(function(value) {
    if (value && typeof value === 'object') {
      value = value.id || value.folderId || value.ID || '';
    }
    value = String(value || '').trim();
    if (value && ids.indexOf(value) === -1) ids.push(value);
  });

  if (!ids.length) {
    throw new Error(
      'Folder Drive sekolah belum dikonfigurasi pada School Context. ' +
      'Diperlukan driveFolderId atau driveRootFolderId.'
    );
  }

  let lastError = null;
  for (let i = 0; i < ids.length; i++) {
    try {
      return DriveApp.getFolderById(ids[i]);
    } catch (e) {
      lastError = e;
    }
  }

  throw new Error(
    'Folder Drive sekolah tidak dapat diakses oleh akun yang sedang menjalankan aplikasi. ' +
    'ID yang dicoba: ' + ids.join(', ') + '. ' +
    'Pastikan folder masih ada dan akun guru memiliki akses Editor. Detail: ' +
    (lastError && lastError.message ? lastError.message : String(lastError || 'Unknown error'))
  );
}

function getSchoolModuleFolder_(moduleName) {
  const name = String(moduleName || '').trim().toUpperCase();
  if (!name) throw new Error('Nama modul wajib diisi.');

  const root = getSchoolDriveFolder_();
  const folders = root.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();

  return root.createFolder(name);
}

function testSchoolDriveAccess() {
  const c = getCurrentUserContext() || {};
  const folder = getSchoolDriveFolder_();
  const school = c.school || {};

  return {
    success: true,
    email: c.email || '',
    npsn: c.npsn || school.npsn || '',
    sekolah: school.namaSekolah || '',
    folderId: folder.getId(),
    folderName: folder.getName(),
  };
}

/** Diagnostik akses Drive khusus Agenda Mengajar. */
function diagnoseSchoolAgendaDriveAccess() {
  const c = getCurrentUserContext() || {};
  const school = c.school || {};
  const result = {
    success: false,
    email: c.email || '',
    npsn: c.npsn || school.npsn || '',
    sekolah: school.namaSekolah || '',
    root: null,
    agenda: null,
    error: '',
  };

  try {
    const root = getSchoolDriveFolder_();
    result.root = {
      id: root.getId(),
      name: root.getName(),
    };

    const folders = root.getFoldersByName('AGENDA');
    if (folders.hasNext()) {
      const folder = folders.next();
      result.agenda = {
        exists: true,
        id: folder.getId(),
        name: folder.getName(),
      };
    } else {
      result.agenda = {
        exists: false,
        message: 'Folder AGENDA belum ada dan akan dibuat saat upload pertama.'
      };
    }

    result.success = true;
    return result;
  } catch (e) {
    result.error = e && e.message ? e.message : String(e);
    return result;
  }
}

/**
 * Upload base64 dari frontend.
 * dataUrl: data:<mime>;base64,...
 */
function uploadFileToSchoolDrive(dataUrl, fileName, moduleName) {
  const c = getCurrentUserContext() || {};
  const school = c.school || {};

  // Agenda Mengajar menggunakan MANAGE_AGENDA; modul lain tetap memakai INPUT_MONITORING.
  const module = String(moduleName || 'GENERAL').trim().toUpperCase();
  if (module === 'AGENDA' && typeof requirePermission === 'function') {
    requirePermission('MANAGE_AGENDA');
  } else if (typeof requirePermission === 'function') {
    requirePermission('INPUT_MONITORING');
  }

  if (!dataUrl || !fileName) {
    throw new Error('Data file dan nama file wajib diisi.');
  }

  const match = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Format data file tidak valid.');

  const mimeType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const folder = getSchoolModuleFolder_(module);

  let file;
  try {
    file = folder.createFile(blob);
  } catch (e) {
    throw new Error(
      'File gagal dibuat di folder ' + folder.getName() + ' sekolah. ' +
      'Pastikan akun guru memiliki akses Editor. Detail: ' +
      (e && e.message ? e.message : String(e))
    );
  }

  return {
    success: true,
    npsn: c.npsn || school.npsn || '',
    school: school.namaSekolah || '',
    fileId: file.getId(),
    fileName: file.getName(),
    folderId: folder.getId(),
    folder: folder.getName(),
    url: file.getUrl(),
  };
}

function setupDriveSekolahSaya() {
  const context = getCurrentUserContext();
  requirePermission('MANAGE_KELAS');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const rootFolder = getSchoolDriveFolder_();
    const moduleFolders = [
      'PRESENSI','AGENDA','PRESTASI','SBI','PARKIR',
      'KEBERSIHAN','KEAMANAN','KERJA','UMUM'
    ];
    const existing = [];
    const created = [];
    const duplicates = [];

    moduleFolders.forEach(function(folderName) {
      const folders = rootFolder.getFoldersByName(folderName);
      const found = [];
      while (folders.hasNext()) {
        const folder = folders.next();
        found.push({ id: folder.getId(), name: folder.getName(), url: folder.getUrl() });
      }

      if (found.length === 0) {
        const newFolder = rootFolder.createFolder(folderName);
        created.push({ name: folderName, id: newFolder.getId(), url: newFolder.getUrl() });
      } else if (found.length === 1) {
        existing.push({ name: folderName, id: found[0].id, url: found[0].url });
      } else {
        duplicates.push({ name: folderName, count: found.length, folders: found });
      }
    });

    let status = 'ALREADY_COMPLETE';
    if (duplicates.length > 0) status = 'DUPLICATE_FOUND';
    else if (created.length > 0) status = 'COMPLETED_MISSING';

    return {
      success: status !== 'DUPLICATE_FOUND',
      status: status,
      email: context.email,
      npsn: context.npsn,
      sekolah: context.school.namaSekolah,
      rootFolderId: rootFolder.getId(),
      rootFolderName: rootFolder.getName(),
      created: created,
      existing: existing,
      duplicates: duplicates,
      totalRequired: moduleFolders.length,
      totalCreated: created.length,
      totalExisting: existing.length,
      totalDuplicateTypes: duplicates.length,
      message: duplicates.length
        ? 'Ditemukan folder dengan nama ganda. Tidak ada folder yang dihapus.'
        : created.length
          ? created.length + ' folder baru dibuat. ' + existing.length + ' folder dipertahankan.'
          : 'Semua folder SIM SATRIA sudah tersedia.'
    };
  } finally {
    lock.releaseLock();
  }
}

function getDriveSetupMessage_(status, existing, created) {
  if (status === 'CREATED_ALL') return 'Semua folder SIM SATRIA berhasil dibuat.';
  if (status === 'COMPLETED_MISSING') {
    return 'Setup selesai. ' + created.length + ' folder baru dibuat dan ' + existing.length + ' folder dipertahankan.';
  }
  return 'Semua folder SIM SATRIA sudah tersedia. Tidak ada folder baru yang dibuat.';
}

function auditDriveSekolahSaya() {
  const context = getCurrentUserContext();
  requirePermission('MANAGE_KELAS');
  const rootFolder = getSchoolDriveFolder_();
  const moduleFolders = ['PRESENSI','AGENDA','PRESTASI','SBI','PARKIR','KEBERSIHAN','KEAMANAN','KERJA','UMUM'];
  const result = [];

  moduleFolders.forEach(function(folderName) {
    const folders = rootFolder.getFoldersByName(folderName);
    const list = [];
    while (folders.hasNext()) {
      const folder = folders.next();
      list.push({ id: folder.getId(), name: folder.getName(), url: folder.getUrl(), files: countFilesInFolder_(folder) });
    }
    result.push({ name: folderName, count: list.length, folders: list });
  });

  return {
    success: true,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    rootFolderId: rootFolder.getId(),
    rootFolderName: rootFolder.getName(),
    folders: result,
  };
}

function countFilesInFolder_(folder) {
  let count = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    files.next();
    count++;
  }
  return count;
}
