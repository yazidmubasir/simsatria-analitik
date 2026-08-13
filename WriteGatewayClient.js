/**
 * SIM SATRIA - WRITE GATEWAY CLIENT
 *
 * Project utama tetap USER_ACCESSING. Semua operasi mutasi database/Drive
 * yang dipanggil oleh user non-admin diarahkan ke project Write Gateway
 * yang Execute as = Me.
 *
 * Konfigurasi pada Script Properties project utama:
 * SIM_SATRIA_WRITE_GATEWAY_URL
 * SIM_SATRIA_WRITE_GATEWAY_SECRET
 */
const WRITE_GATEWAY_CLIENT_CONFIG = {
  URL_PROPERTY: 'SIM_SATRIA_WRITE_GATEWAY_URL',
  SECRET_PROPERTY: 'SIM_SATRIA_WRITE_GATEWAY_SECRET',
  TIMEOUT_MS: 30000,
};

function isViewerWriteRole_() {
  try {
    const role = normalizeRole_(getCurrentUserContext().role);
    return ['GURU', 'WALI_KELAS', 'KARYAWAN', 'SISWA'].indexOf(role) >= 0;
  } catch (e) {
    return false;
  }
}

function callWriteGateway_(action, payload) {
  const url = String(PropertiesService.getScriptProperties().getProperty(WRITE_GATEWAY_CLIENT_CONFIG.URL_PROPERTY) || '').trim();
  const secret = String(PropertiesService.getScriptProperties().getProperty(WRITE_GATEWAY_CLIENT_CONFIG.SECRET_PROPERTY) || '');
  if (!url) throw new Error('Write Gateway URL belum dikonfigurasi.');
  if (!secret) throw new Error('Write Gateway secret belum dikonfigurasi.');

  const context = getCurrentUserContext();
  const body = Object.assign({}, payload || {}, {
    action: String(action || '').trim().toUpperCase(),
    email: context.email,
    npsn: context.npsn,
    role: context.role,
    timestamp: Date.now(),
    secret: secret,
  });

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  let result;
  try {
    result = JSON.parse(response.getContentText() || '{}');
  } catch (e) {
    throw new Error('Respons Write Gateway tidak valid. HTTP ' + response.getResponseCode());
  }
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300 || !result.success) {
    throw new Error(result.error || 'Write Gateway gagal memproses permintaan.');
  }
  return result;
}

function gatewayAppendRow_(sheetName, row) {
  return callWriteGateway_('APPEND_ROW', { sheet: sheetName, row: row });
}

function gatewaySetValues_(sheetName, startRow, startCol, values) {
  return callWriteGateway_('SET_VALUES', {
    sheet: sheetName,
    startRow: startRow,
    startCol: startCol,
    values: values,
  });
}

function gatewayUploadFile_(moduleName, file) {
  return callWriteGateway_('UPLOAD_FILE', {
    module: moduleName,
    file: file,
  });
}

function gatewayCreatePdf_(fileName, html, folderId, moduleName) {
  return callWriteGateway_('CREATE_PDF', {
    fileName: fileName,
    html: html,
    folderId: folderId,
    module: moduleName,
  });
}
