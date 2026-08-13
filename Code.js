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
  // Identitas ditentukan oleh Auth.js.
  // Guru/Wali/Karyawan/Siswa yang sudah memiliki binding sekolah tidak pernah
  // membuka MASTER. ADMIN_SEKOLAH baru membaca MASTER pada jalur admin.
  clearUserContextCache_(getGoogleUserEmail_());
  const context = getCurrentUserContext();
  return {
    success: true,
    user: {
      userId: context.userId,
      email: context.email,
      nip: context.nip,
      nama: context.nama,
      role: context.role,
    },
    school: {
      npsn: context.npsn,
      namaSekolah: context.school.namaSekolah,
      alamat: context.school.alamat,
      logoUrl: context.school.logoUrl,
      tagline: context.school.tagline,
      warnaUtama: context.school.warnaUtama,
      warnaSekunder: context.school.warnaSekunder,
    },
  };
}

function getSchoolContextInfo() {
  clearUserContextCache_(getGoogleUserEmail_());
  const c = getCurrentUserContext();
  return {
    success: true,
    email: c.email,
    userId: c.userId,
    nip: c.nip,
    nama: c.nama,
    role: c.role,
    npsn: c.npsn,
    sekolah: c.school.namaSekolah,
    spreadsheetId: c.school.spreadsheetId,
    driveFolderId: c.school.driveFolderId,
  };
}

function getKoneksiView() {
  const html = HtmlService.createHtmlOutputFromFile("koneksi").getContent();
  let js = HtmlService.createHtmlOutputFromFile("koneksi_js").getContent();
  js = js.replace(/^\s*<script[^>]*>/i, "").replace(/<\/script>\s*$/i, "");
  return { success: true, html: html, js: js };
}

function getPresensiPerkelasView() {
  const html = HtmlService.createHtmlOutputFromFile("presensiPerkelas").getContent();
  let js = HtmlService.createHtmlOutputFromFile("presensiPerkelas_js").getContent();
  js = js.replace(/^\s*<script[^>]*>/i, "").replace(/<\/script>\s*$/i, "");
  js = js
    .replace(/\.getKelasPresensiPerkelas\(\)/g, ".securePresensiLoad()")
    .replace(/\.getPresensiPerkelasData\(/g, ".securePresensiGetData(")
    .replace(/\.simpanPresensiPerkelas\(/g, ".securePresensiSave(");
  js += '\n;setTimeout(function(){try{if(typeof PP_init === "function"){PP_init();}else{console.error("PP_init tidak ditemukan pada modul Presensi Per Kelas.");}}catch(e){console.error("PP_init ERROR:",e);}},0);\n';
  return { success: true, html: html, js: js };
}

function getManajemenPenggunaView() {
  const view = getUserManagementView();
  let html = view.html || "";
  let js = view.js || "";

  // Tombol ini sengaja disisipkan dari server agar modul manajemen pengguna
  // tetap satu sumber dengan file Database.js. Tidak ada duplikasi struktur
  // database di frontend.
  const anchor = '<button class="btn btn-success btn-sm px-3" type="button" onclick="mpBukaForm()">';
  const databaseButton = '<button class="btn btn-outline-primary btn-sm px-3" type="button" id="mpDatabaseSetupBtn" onclick="mpSiapkanDatabaseSekolah(this)"><i class="bi bi-database-check me-1"></i> Siapkan Database Sekolah</button>';
  if (html.indexOf("mpDatabaseSetupBtn") < 0 && html.indexOf(anchor) >= 0) {
    html = html.replace(anchor, databaseButton + anchor);
  }

  // Backend tetap memakai setupDatabaseSekolahSaya() dari Database.js.
  // Fungsi tersebut bersifat non-destructive: sheet yang sudah ada tidak
  // dihapus, data tidak dipindahkan, dan header yang sudah ada dipertahankan.
  js += '\nwindow.mpSiapkanDatabaseSekolah=function(btn){\n' +
    '  if(!confirm("Siapkan struktur database untuk Spreadsheet sekolah ini?\\n\\nSheet yang belum ada akan dibuat. Sheet dan data yang sudah ada akan dipertahankan."))return;\n' +
    '  btn=btn||document.getElementById("mpDatabaseSetupBtn");\n' +
    '  var oldHtml=btn?btn.innerHTML:"";\n' +
    '  if(btn){btn.disabled=true;btn.innerHTML="<span class=\\"spinner-border spinner-border-sm me-1\\"></span> Menyiapkan...";}\n' +
    '  google.script.run.withSuccessHandler(function(r){\n' +
    '    if(btn){btn.disabled=false;btn.innerHTML=oldHtml;}\n' +
    '    if(!r||!r.success){if(typeof mpAlert==="function")mpAlert((r&&r.message)||"Gagal menyiapkan database sekolah.","danger");return;}\n' +
    '    var dibuat=(r.createdSheets||[]).map(function(x){return x.name;});\n' +
    '    var sudahAda=(r.existingSheets||[]).map(function(x){return x.name;});\n' +
    '    var tambahan=(r.addedHeaders||[]).map(function(x){return x.sheet+": "+(x.headers||[]).join(", ");});\n' +
    '    var detail=[];\n' +
    '    if(dibuat.length)detail.push("Dibuat: "+dibuat.join(", "));\n' +
    '    if(sudahAda.length)detail.push("Dipertahankan: "+sudahAda.join(", "));\n' +
    '    if(tambahan.length)detail.push("Header dilengkapi: "+tambahan.join(" | "));\n' +
    '    if(typeof mpAlert==="function")mpAlert((r.message||"Setup database selesai.")+(detail.length?" "+detail.join(" "):""),r.status==="ALREADY_COMPLETE"?"info":"success");\n' +
    '    if(typeof mpMuat==="function")mpMuat();\n' +
    '  }).withFailureHandler(function(e){\n' +
    '    if(btn){btn.disabled=false;btn.innerHTML=oldHtml;}\n' +
    '    if(typeof mpAlert==="function")mpAlert(e.message||String(e),"danger");\n' +
    '  }).setupDatabaseSekolahSaya();\n' +
    '};\n';

  return { success: true, html: html, js: js };
}
