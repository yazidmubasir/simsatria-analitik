function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("SIM SATRIA")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupMasterRegistry() {
  const ss = getMasterSpreadsheet_();
  ensureSheetHeaders_(ss, "SCHOOLS", ["NPSN","NAMA_SEKOLAH","STATUS","SPREADSHEET_ID","DRIVE_FOLDER_ID","ALAMAT","LOGO_URL","TAGLINE","WARNA_UTAMA","WARNA_SEKUNDER"]);
  ensureSheetHeaders_(ss, "ADMIN_SEKOLAH", ["USER_ID","EMAIL","NIP","NAMA","NPSN","ROLE","STATUS"]);
  return { success: true, spreadsheetId: ss.getId(), spreadsheetName: ss.getName(), message: "Registry Master siap. Sheet SCHOOLS dan ADMIN_SEKOLAH tersedia." };
}

function ensureSheetHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const current = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (!current.length || current.every((v) => String(v).trim() === "")) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }
  const normalized = current.map((v) => String(v || "").trim().toUpperCase());
  headers.forEach((h) => {
    if (!normalized.includes(h.toUpperCase())) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      normalized.push(h.toUpperCase());
    }
  });
  sh.setFrozenRows(1);
}

function getLoginInfo() {
  // Selalu bersihkan cache autentikasi saat membaca identitas login.
  // Jika email sudah menjadi ADMIN_SEKOLAH, binding GURU lama diputus
  // sebelum context dibangun. Data historis GURU tetap aman di sheet sekolah.
  const email = getGoogleUserEmail_();
  clearUserContextCache_(email);
  if (typeof enforceAdminIdentityIsolation_ === "function") {
    enforceAdminIdentityIsolation_();
    clearUserContextCache_(email);
  }

  const context = getCurrentUserContext();
  return { success: true, user: { userId: context.userId, email: context.email, nip: context.nip, nama: context.nama, role: context.role }, school: { npsn: context.npsn, namaSekolah: context.school.namaSekolah, alamat: context.school.alamat, logoUrl: context.school.logoUrl, tagline: context.school.tagline, warnaUtama: context.school.warnaUtama, warnaSekunder: context.school.warnaSekunder } };
}

function getSchoolContextInfo() {
  const email = getGoogleUserEmail_();
  clearUserContextCache_(email);
  if (typeof enforceAdminIdentityIsolation_ === "function") enforceAdminIdentityIsolation_();
  clearUserContextCache_(email);
  const c = getCurrentUserContext();
  return { success: true, email: c.email, userId: c.userId, nip: c.nip, nama: c.nama, role: c.role, npsn: c.npsn, sekolah: c.school.namaSekolah, spreadsheetId: c.school.spreadsheetId, driveFolderId: c.school.driveFolderId };
}

function testSchoolResources() {
  const c = getCurrentUserContext();
  const result = { email: c.email, npsn: c.npsn, sekolah: c.school.namaSekolah, spreadsheet: { success: false, id: c.school.spreadsheetId, name: "", error: "" }, drive: { success: false, id: c.school.driveFolderId, name: "", error: "" } };
  try { const ss = SpreadsheetApp.openById(c.school.spreadsheetId); result.spreadsheet.success = true; result.spreadsheet.name = ss.getName(); } catch (e) { result.spreadsheet.error = e.message; }
  try { const folder = DriveApp.getFolderById(c.school.driveFolderId); result.drive.success = true; result.drive.name = folder.getName(); } catch (e) { result.drive.error = e.message; }
  return result;
}

function testSchoolContextSpeed() { const start = Date.now(); return { elapsedMs: Date.now() - start, context: getSchoolContextInfo() }; }

function getKoneksiView() {
  const html = HtmlService.createHtmlOutputFromFile("koneksi").getContent();
  let js = HtmlService.createHtmlOutputFromFile("koneksi_js").getContent();
  js = js.replace(/^\s*<script[^>]*>/i, "").replace(/<\/script>\s*$/i, "");
  return { success: true, html: html, js: js };
}

function getPresensiPerkelasView() {
  const html = HtmlService.createHtmlOutputFromFile("presensiPerkelas").getContent();
  let js = HtmlService.createHtmlOutputFromFile("presensiPerkelas_js").getContent();

  // Buang wrapper <script> karena JS akan diinjeksi sebagai elemen script.
  js = js.replace(/^\s*<script[^>]*>/i, "").replace(/<\/script>\s*$/i, "");

  // Gunakan endpoint security baru tanpa mengubah frontend lama.
  js = js
    .replace(/\.getKelasPresensiPerkelas\(\)/g, ".securePresensiLoad()")
    .replace(/\.getPresensiPerkelasData\(/g, ".securePresensiGetData(")
    .replace(/\.simpanPresensiPerkelas\(/g, ".securePresensiSave(");

  // KRITIS: PP_init() harus dipanggil DI DALAM script yang diinjeksi.
  // Function declaration dari script dinamis tidak selalu tersedia pada
  // lexical scope fungsi menu, sehingga memanggil PP_init() dari menu
  // menyebabkan "PP_init() tidak ditemukan".
  js += '\n;setTimeout(function(){try{if(typeof PP_init === "function"){PP_init();}else{console.error("PP_init tidak ditemukan pada modul Presensi Per Kelas.");}}catch(e){console.error("PP_init ERROR:",e);}},0);\n';

  return { success: true, html: html, js: js };
}

function getManajemenPenggunaView() {
  return getUserManagementView();
}
