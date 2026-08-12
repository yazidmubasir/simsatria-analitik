/**
 * DATABASE SERVICES
 * Selalu bekerja pada Spreadsheet sekolah dari School Context.
 */
function normalizeHeader_(v) {
  return String(v || "").trim().toUpperCase();
}

function getSheetHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(normalizeHeader_);
}

function setupHeaders_(sheet, headers) {
  if (!sheet) throw new Error("Sheet tidak tersedia.");
  headers = headers.map(normalizeHeader_);
  const current = getSheetHeaders_(sheet);
  if (!current.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return headers;
  }
  const merged = current.slice();
  headers.forEach((h) => {
    if (!merged.includes(h)) merged.push(h);
  });
  if (merged.length !== current.length) {
    sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
  }
  sheet.setFrozenRows(1);
  return merged;
}

function ensureTransactionSheet_(sheetName, businessHeaders) {
  const ss = getSchoolSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  const systemHeaders = [
    "TRANSACTION_ID","TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE",
  ];
  return setupHeaders_(sheet, systemHeaders.concat(businessHeaders || []));
}

function initializeSchoolTransactionSheets() {
  requirePermission("MANAGE_KELAS");
  const modules = {
    TRX_PRESENSI: ["TANGGAL","KELAS","NISN","NAMA_SISWA","STATUS","KETERANGAN"],
    TRX_PARKIR: ["TANGGAL","KENDALA","SOLUSI","UPLOAD_FOTO_PARKIR"],
    TRX_PRESTASI: ["TANGGAL","NAMA_SISWA","JENIS","TINGKAT","KETERANGAN"],
    TRX_AGENDA_GURU: [
      "TANGGAL","SESI","KELAS","TUJUAN_PEMBELAJARAN","MATERI_PEMBELAJARAN",
      "DPL","PENGALAMAN_BELAJAR","PRINSIP_PEMBELAJARAN","REKAP_MURID_TIDAK_IKUT",
      "BUKTI_FISIK","NAMA_GURU","MAPEL","KETERANGAN",
    ],
    TRX_SBI: ["INDIKATOR","SUBINDIKATOR","URAIAN_KEGIATAN","HAMBATAN","SOLUSI","KARAKTER","BUKTI_FISIK"],
    TRX_KEBERSIHAN: ["TANGGAL","KENDALA","SOLUSI","BUKTI_FISIK"],
    TRX_KEAMANAN: ["TANGGAL","KENDALA","SOLUSI","BUKTI_FISIK"],
    TRX_KERJA: [
      "TANGGAL_PELAKSANAAN","SESI","BIDANG_TUGAS","TARGET_PEKERJAAN","URAIAN_PEKERJAAN",
      "KENDALA","TINDAK_LANJUT","REFLEKSI","BUKTI_FISIK",
    ],
  };
  Object.keys(modules).forEach((name) => ensureTransactionSheet_(name, modules[name]));
  ensureLogSheet_();
  return { success: true, school: getSchoolContext().namaSekolah, npsn: getSchoolContext().npsn };
}

function ensureLogSheet_() {
  const ss = getSchoolSpreadsheet_();
  let sh = ss.getSheetByName("LOG");
  if (!sh) sh = ss.insertSheet("LOG");
  setupHeaders_(sh, [
    "TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE","ACTION","MODULE","DESCRIPTION","TRANSACTION_ID",
  ]);
  return sh;
}

function setupDatabaseSekolahSaya() {
  const context = requireSchoolAdmin();
  const permission = "MANAGE_DATABASE";
  if (!hasPermission(permission, context.role)) {
    throw new Error("Akses setup database hanya untuk ADMIN_SEKOLAH.");
  }
  const ss = SpreadsheetApp.openById(context.school.spreadsheetId);

  const masterSheets = {
    CONFIG: ["KEY","VALUE","KETERANGAN"],
    GURU: ["NIP","NAMA","EMAIL","NO_HP","STATUS"],
    SISWA: ["NISN","NIS","NAMA","JK","KELAS","STATUS"],
    KARYAWAN: ["NIP","NAMA","JABATAN","EMAIL","NO_HP","STATUS"],
    KELAS: ["KELAS","TINGKAT","JURUSAN","WALI_KELAS","STATUS"],
  };

  const transactionSheets = {
    TRX_PRESENSI: ["TANGGAL","KELAS","NISN","NAMA_SISWA","STATUS","KETERANGAN"],
    TRX_PARKIR: ["TANGGAL","KENDALA","SOLUSI","UPLOAD_FOTO_PARKIR"],
    TRX_PRESTASI: ["TANGGAL","NAMA_SISWA","JENIS","TINGKAT","KETERANGAN"],
    TRX_AGENDA_GURU: [
      "TANGGAL","SESI","KELAS","TUJUAN_PEMBELAJARAN","MATERI_PEMBELAJARAN","DPL",
      "PENGALAMAN_BELAJAR","PRINSIP_PEMBELAJARAN","REKAP_MURID_TIDAK_IKUT","BUKTI_FISIK",
      "NAMA_GURU","MAPEL","KETERANGAN",
    ],
    TRX_SBI: ["INDIKATOR","SUBINDIKATOR","URAIAN_KEGIATAN","HAMBATAN","SOLUSI","KARAKTER","BUKTI_FISIK"],
    TRX_KEBERSIHAN: ["TANGGAL","KENDALA","SOLUSI","BUKTI_FISIK"],
    TRX_KEAMANAN: ["TANGGAL","KENDALA","SOLUSI","BUKTI_FISIK"],
    TRX_KERJA: [
      "TANGGAL_PELAKSANAAN","SESI","BIDANG_TUGAS","TARGET_PEKERJAAN","URAIAN_PEKERJAAN",
      "KENDALA","TINDAK_LANJUT","REFLEKSI","BUKTI_FISIK",
    ],
  };

  const transactionSystemHeaders = [
    "TRANSACTION_ID","TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE",
  ];
  const createdSheets = [];
  const existingSheets = [];
  const addedHeaders = [];

  function ensureSheet_(sheetName, headers, category) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      createdSheets.push({ name: sheetName, category: category });
    } else {
      existingSheets.push({ name: sheetName, category: category });
    }
    const result = ensureHeadersNonDestructive_(sheet, headers);
    if (result.added && result.added.length > 0) {
      addedHeaders.push({ sheet: sheetName, category: category, headers: result.added });
    }
  }

  Object.keys(masterSheets).forEach((sheetName) => ensureSheet_(sheetName, masterSheets[sheetName], "MASTER"));
  Object.keys(transactionSheets).forEach((sheetName) =>
    ensureSheet_(sheetName, transactionSystemHeaders.concat(transactionSheets[sheetName]), "TRANSACTION"),
  );
  ensureSheet_(
    "LOG",
    ["TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE","ACTION","MODULE","DESCRIPTION","TRANSACTION_ID"],
    "SYSTEM",
  );

  const config = ss.getSheetByName("CONFIG");
  if (!config) throw new Error("CONFIG gagal dibuat.");

  const configKeys = {
    NPSN: context.npsn,
    NAMA_SEKOLAH: context.school.namaSekolah,
    SPREADSHEET_ID: context.school.spreadsheetId,
    DRIVE_FOLDER_ID: context.school.driveFolderId,
    STATUS: "ACTIVE",
    VERSI_DATABASE: "1.0",
  };

  const configLastRow = config.getLastRow();
  const existingConfig = {};
  if (configLastRow >= 2) {
    config.getRange(2, 1, configLastRow - 1, 3).getValues().forEach((row) => {
      const key = String(row[0] || "").trim().toUpperCase();
      if (key) existingConfig[key] = { row: row, value: row[1] };
    });
  }

  const configToAdd = [];
  Object.keys(configKeys).forEach((key) => {
    if (!existingConfig[key]) configToAdd.push([key, configKeys[key], getConfigDescription_(key)]);
  });
  if (configToAdd.length > 0) {
    config.getRange(config.getLastRow() + 1, 1, configToAdd.length, 3).setValues(configToAdd);
  }

  const hasChanges = createdSheets.length > 0 || addedHeaders.length > 0 || configToAdd.length > 0;
  return {
    success: true,
    status: hasChanges ? "COMPLETED_MISSING" : "ALREADY_COMPLETE",
    email: context.email,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    spreadsheet: ss.getName(),
    spreadsheetId: ss.getId(),
    createdSheets: createdSheets,
    existingSheets: existingSheets,
    addedHeaders: addedHeaders,
    addedConfig: configToAdd,
    message: hasChanges
      ? "Database sekolah berhasil dilengkapi tanpa menghapus data yang sudah ada."
      : "Database sekolah sudah lengkap. Tidak ada perubahan.",
  };
}

function getDatabaseSetupMessage_(createdSheets, addedHeaders) {
  if (createdSheets.length === 0 && addedHeaders.length === 0) {
    return "Database sudah lengkap. Tidak ada sheet atau header yang diubah.";
  }
  return "Setup database selesai. " + createdSheets.length + " sheet dibuat dan " +
    addedHeaders.length + " sheet memperoleh tambahan header. Data yang sudah ada dipertahankan.";
}

function ensureHeadersNonDestructive_(sheet, requiredHeaders) {
  if (!sheet) throw new Error("Sheet tidak ditemukan.");
  const headers = requiredHeaders.map((h) => String(h || "").trim().toUpperCase()).filter((h) => h !== "");
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    if (headers.length > 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return { existing: [], added: headers.slice() };
  }

  const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((h) =>
    String(h || "").trim().toUpperCase(),
  );
  const added = [];
  headers.forEach((header) => {
    if (!existingHeaders.includes(header)) {
      const newColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColumn).setValue(header);
      existingHeaders.push(header);
      added.push(header);
    }
  });
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  return { existing: existingHeaders, added: added };
}

function getConfigDescription_(key) {
  const descriptions = {
    NPSN: "NPSN sekolah",
    NAMA_SEKOLAH: "Nama sekolah",
    SPREADSHEET_ID: "ID Spreadsheet sekolah",
    DRIVE_FOLDER_ID: "ID folder Drive sekolah",
    STATUS: "Status sekolah",
    VERSI_DATABASE: "Versi struktur database",
  };
  return descriptions[key] || "";
}
