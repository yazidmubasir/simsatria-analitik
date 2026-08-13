/**
 * WRITE GATEWAY SERVICES
 *
 * The gateway runs as its deploying/admin account. Caller identity is still
 * taken from Session.getActiveUser() when the deployment requires sign-in.
 * The caller is validated against the USERS sheet in the configured school
 * spreadsheet before any write is performed.
 */
function handleGatewayRequest_(request, requestId) {
  validateRequestSize_(request);

  const action = String(request.action || "").trim().toUpperCase();
  if (!action) throw new Error("Action gateway wajib diisi.");

  const identity = getGatewayCallerIdentity_();
  const user = authorizeGatewayUser_(identity.email, request);
  const context = gatewayContext_();

  if (normalizeGatewayNpsn_(request.npsn) && normalizeGatewayNpsn_(request.npsn) !== context.npsn) {
    throw new Error("NPSN request tidak sesuai dengan gateway sekolah.");
  }

  switch (action) {
    case "PING":
      return gatewaySuccess_(requestId, { action: action, email: user.email, npsn: context.npsn });
    case "SAVE_ROW":
      requireGatewayRole_(user, ["ADMIN_SEKOLAH", "GURU", "WALI_KELAS", "KARYAWAN", "SISWA"]);
      return saveRow_(request, user, requestId);
    case "UPDATE_ROW":
      requireGatewayRole_(user, ["ADMIN_SEKOLAH", "GURU", "WALI_KELAS", "KARYAWAN", "SISWA"]);
      return updateRow_(request, user, requestId);
    case "READ_ROWS":
      requireGatewayRole_(user, GATEWAY_CONFIG.ALLOWED_ROLES);
      return readRows_(request, user, requestId);
    case "CREATE_PDF":
      requireGatewayRole_(user, GATEWAY_CONFIG.ALLOWED_ROLES);
      return createPdf_(request, user, requestId);
    case "UPLOAD_FILE":
      requireGatewayRole_(user, GATEWAY_CONFIG.ALLOWED_ROLES);
      return uploadFile_(request, user, requestId);
    default:
      throw new Error("Action gateway tidak diizinkan: " + action);
  }
}

function getGatewayCallerIdentity_() {
  const active = Session.getActiveUser();
  const email = active ? normalizeGatewayEmail_(active.getEmail()) : "";
  if (!email) {
    throw new Error("Identitas Google pemanggil tidak tersedia. Gateway harus dideploy dengan akses yang mewajibkan pengguna login.");
  }
  return { email: email };
}

function authorizeGatewayUser_(email, request) {
  const ss = SpreadsheetApp.openById(gatewayContext_().spreadsheetId);
  const usersSheet = ss.getSheetByName("USERS");
  if (!usersSheet) throw new Error("Sheet USERS belum tersedia pada Spreadsheet sekolah.");

  const rows = sheetObjects_(usersSheet);
  const user = rows.find(function(row) {
    return normalizeGatewayEmail_(row.EMAIL) === email;
  });
  if (!user) throw new Error("Akun " + email + " tidak ditemukan pada USERS sekolah.");

  const status = String(user.STATUS || "").trim().toUpperCase();
  const role = String(user.ROLE || "").trim().toUpperCase();
  if (status !== GATEWAY_CONFIG.ACTIVE_STATUS) throw new Error("Akun pengguna tidak aktif.");
  if (!GATEWAY_CONFIG.ALLOWED_ROLES.includes(role)) throw new Error("Role pengguna tidak diizinkan.");

  const requestedUserId = String(request.userId || "").trim();
  if (requestedUserId && String(user.USER_ID || "").trim() !== requestedUserId) {
    throw new Error("USER_ID request tidak sesuai dengan akun Google.");
  }

  return {
    userId: String(user.USER_ID || "").trim(),
    email: email,
    nip: String(user.NIP || "").trim(),
    nama: String(user.NAMA || "").trim(),
    role: role,
    status: status,
  };
}

function saveRow_(request, user, requestId) {
  const sheetName = normalizeGatewaySheet_(request.sheet);
  const allowed = GATEWAY_WRITE_ALLOWLIST[sheetName];
  if (!allowed) throw new Error("Sheet tidak diizinkan untuk SAVE_ROW: " + sheetName);

  const values = request.values && typeof request.values === "object" ? request.values : null;
  if (!values) throw new Error("values wajib berupa object.");

  const ss = SpreadsheetApp.openById(gatewayContext_().spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " belum tersedia.");

  const headers = ensureHeadersFromGateway_(sheet, allowed);
  const row = new Array(headers.length).fill("");
  const now = new Date();
  const txId = String(request.transactionId || Utilities.getUuid()).trim();
  const system = {
    TRANSACTION_ID: txId,
    TIMESTAMP: now,
    NPSN: gatewayContext_().npsn,
    USER_ID: user.userId,
    EMAIL: user.email,
    NIP: user.nip,
    NAMA_USER: user.nama,
    ROLE: user.role,
  };

  headers.forEach(function(header, index) {
    if (Object.prototype.hasOwnProperty.call(system, header)) row[index] = system[header];
    else if (Object.prototype.hasOwnProperty.call(values, header)) row[index] = sanitizeCell_(values[header]);
  });

  validateRequiredBusinessData_(sheetName, values);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  appendGatewayLog_(user, "SAVE", sheetName, "Data disimpan melalui Write Gateway", txId);

  return gatewaySuccess_(requestId, {
    action: "SAVE_ROW",
    sheet: sheetName,
    transactionId: txId,
    rowNumber: sheet.getLastRow(),
  });
}

function updateRow_(request, user, requestId) {
  const sheetName = normalizeGatewaySheet_(request.sheet);
  const allowed = GATEWAY_WRITE_ALLOWLIST[sheetName];
  if (!allowed) throw new Error("Sheet tidak diizinkan untuk UPDATE_ROW: " + sheetName);

  const rowNumber = Number(request.rowNumber);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) throw new Error("rowNumber tidak valid.");

  const values = request.values && typeof request.values === "object" ? request.values : null;
  if (!values) throw new Error("values wajib berupa object.");

  const ss = SpreadsheetApp.openById(gatewayContext_().spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " belum tersedia.");
  if (rowNumber > sheet.getLastRow()) throw new Error("Baris target tidak ditemukan.");

  const headers = ensureHeadersFromGateway_(sheet, allowed);
  const current = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  headers.forEach(function(header, index) {
    if (Object.prototype.hasOwnProperty.call(values, header)) current[index] = sanitizeCell_(values[header]);
  });
  sheet.getRange(rowNumber, 1, 1, current.length).setValues([current]);
  appendGatewayLog_(user, "UPDATE", sheetName, "Data diperbarui melalui Write Gateway", "");

  return gatewaySuccess_(requestId, { action: "UPDATE_ROW", sheet: sheetName, rowNumber: rowNumber });
}

function readRows_(request, user, requestId) {
  const sheetName = normalizeGatewaySheet_(request.sheet);
  if (!GATEWAY_READ_ALLOWLIST.includes(sheetName)) throw new Error("Sheet tidak diizinkan untuk READ_ROWS.");
  const ss = SpreadsheetApp.openById(gatewayContext_().spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " belum tersedia.");
  const values = sheet.getDataRange().getDisplayValues();
  const maxRows = Math.min(Number(request.limit) || 500, 1000);
  return gatewaySuccess_(requestId, {
    action: "READ_ROWS",
    sheet: sheetName,
    headers: values.length ? values[0] : [],
    rows: values.length > 1 ? values.slice(1, maxRows + 1) : [],
  });
}

function uploadFile_(request, user, requestId) {
  const name = String(request.name || "").trim();
  const mimeType = String(request.mimeType || "application/octet-stream").trim();
  const base64 = String(request.base64 || "").trim();
  if (!name || !base64) throw new Error("name dan base64 wajib diisi.");
  if (base64.length > GATEWAY_CONFIG.MAX_BODY_BYTES) throw new Error("File terlalu besar untuk Gateway.");

  const folderId = String(request.folderId || gatewayContext_().driveFolderId || "").trim();
  if (!folderId) throw new Error("Folder Drive Gateway belum dikonfigurasi.");
  const folder = DriveApp.getFolderById(folderId);
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, name);
  const file = folder.createFile(blob);
  appendGatewayLog_(user, "UPLOAD", "DRIVE", "File diunggah melalui Write Gateway: " + file.getName(), "");

  return gatewaySuccess_(requestId, {
    action: "UPLOAD_FILE",
    fileId: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
  });
}

function createPdf_(request, user, requestId) {
  const name = String(request.name || "SIM SATRIA.pdf").trim();
  const html = String(request.html || "").trim();
  if (!html) throw new Error("html PDF wajib diisi.");
  if (html.length > GATEWAY_CONFIG.MAX_BODY_BYTES) throw new Error("Konten PDF terlalu besar.");

  const folderId = String(request.folderId || gatewayContext_().driveFolderId || "").trim();
  if (!folderId) throw new Error("Folder Drive Gateway belum dikonfigurasi.");

  const blob = Utilities.newBlob(html, "text/html", name.replace(/\.pdf$/i, "") + ".html");
  const temp = DriveApp.createFile(blob);
  try {
    const pdf = temp.getBlob().getAs(MimeType.PDF).setName(name);
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(pdf);
    appendGatewayLog_(user, "CREATE_PDF", "DRIVE", "PDF dibuat melalui Write Gateway: " + file.getName(), "");
    return gatewaySuccess_(requestId, { action: "CREATE_PDF", fileId: file.getId(), name: file.getName(), url: file.getUrl() });
  } finally {
    try { temp.setTrashed(true); } catch (ignore) {}
  }
}

function requireGatewayRole_(user, allowedRoles) {
  if (!allowedRoles.includes(user.role)) throw new Error("Role " + user.role + " tidak memiliki izin operasi ini.");
}

function validateRequiredBusinessData_(sheetName, values) {
  if (sheetName === "TRX_PRESENSI" && !String(values.TANGGAL || "").trim()) throw new Error("TANGGAL presensi wajib diisi.");
}

function ensureHeadersFromGateway_(sheet, allowed) {
  const required = ["TRANSACTION_ID","TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE"].concat(allowed);
  const existing = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v){return String(v||"").trim().toUpperCase();}) : [];
  const headers = existing.slice();
  required.forEach(function(h) { if (!headers.includes(h)) headers.push(h); });
  if (!existing.length) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  else if (headers.length !== existing.length) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return headers;
}

function appendGatewayLog_(user, action, module, description, transactionId) {
  const ss = SpreadsheetApp.openById(gatewayContext_().spreadsheetId);
  let sheet = ss.getSheetByName("LOG");
  if (!sheet) sheet = ss.insertSheet("LOG");
  const required = ["TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE","ACTION","MODULE","DESCRIPTION","TRANSACTION_ID"];
  const headers = ensureHeadersFromGateway_(sheet, required.slice(7));
  const row = headers.map(function(h) {
    return ({
      TIMESTAMP: new Date(), NPSN: gatewayContext_().npsn, USER_ID: user.userId, EMAIL: user.email,
      NIP: user.nip, NAMA_USER: user.nama, ROLE: user.role, ACTION: action, MODULE: module,
      DESCRIPTION: description, TRANSACTION_ID: transactionId,
    })[h] || "";
  });
  sheet.getRange(sheet.getLastRow()+1,1,1,row.length).setValues([row]);
}

function sheetObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(function(h){return String(h||"").trim().toUpperCase();});
  return data.slice(1).map(function(row){
    const obj = {};
    headers.forEach(function(h,i){obj[h]=row[i];});
    return obj;
  }).filter(function(obj){return Object.keys(obj).some(function(k){return String(obj[k]||"").trim() !== "";});});
}

function normalizeGatewayEmail_(email) { return String(email || "").trim().toLowerCase(); }
function normalizeGatewayNpsn_(npsn) { return String(npsn || "").trim(); }
function normalizeGatewaySheet_(name) { return String(name || "").trim().toUpperCase(); }
function sanitizeCell_(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" && value.length > GATEWAY_CONFIG.MAX_CELL_CHARS) throw new Error("Nilai cell terlalu panjang.");
  return value;
}
function validateRequestSize_(request) {
  const json = JSON.stringify(request || {});
  if (json.length > GATEWAY_CONFIG.MAX_BODY_BYTES) throw new Error("Request gateway terlalu besar.");
}
function gatewaySuccess_(requestId, data) {
  return Object.assign({ success: true, requestId: requestId }, data || {});
}
