/**
 * Central policy for write operations.
 *
 * NOTE: This does not grant physical Editor permission to user accounts.
 * Non-admin mutations must be sent through Write Gateway.
 */
function requireAppWritePermission_(permission) {
  requirePermission(permission || 'INPUT_MONITORING');
  const context = getCurrentUserContext();
  const role = normalizeRole_(context.role);
  if (['ADMIN_SEKOLAH', 'SUPERADMIN', 'SUPER_ADMIN'].indexOf(role) >= 0) {
    return { mode: 'DIRECT', context: context };
  }
  if (['GURU', 'WALI_KELAS', 'KARYAWAN', 'SISWA'].indexOf(role) >= 0) {
    return { mode: 'GATEWAY', context: context };
  }
  throw new Error('Role tidak diizinkan melakukan operasi tulis.');
}

function appendSchoolRowByPolicy_(permission, sheetName, row) {
  const policy = requireAppWritePermission_(permission);
  if (policy.mode === 'GATEWAY') return gatewayAppendRow_(sheetName, row);
  const sheet = getSchoolSheet_(sheetName);
  sheet.appendRow(row);
  return { success: true, mode: 'DIRECT', sheet: sheet.getName(), rowNumber: sheet.getLastRow() };
}

function setSchoolValuesByPolicy_(permission, sheetName, startRow, startCol, values) {
  const policy = requireAppWritePermission_(permission);
  if (policy.mode === 'GATEWAY') return gatewaySetValues_(sheetName, startRow, startCol, values);
  const sheet = getSchoolSheet_(sheetName);
  sheet.getRange(startRow, startCol, values.length, values[0].length).setValues(values);
  return { success: true, mode: 'DIRECT', sheet: sheet.getName() };
}

function uploadSchoolFileByPolicy_(permission, moduleName, file) {
  const policy = requireAppWritePermission_(permission || 'UPLOAD_FILE');
  if (policy.mode === 'GATEWAY') return gatewayUploadFile_(moduleName, file);
  const folder = getSchoolModuleFolder_(moduleName);
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType || 'application/octet-stream', file.name);
  const created = folder.createFile(blob);
  return { success: true, mode: 'DIRECT', fileId: created.getId(), fileName: created.getName(), url: created.getUrl() };
}
