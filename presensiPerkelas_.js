/**
 * SIM SATRIA - PRESENSI PER KELAS
 *
 * Canonical compatibility layer.
 * Business logic and permission checks live in PresensiSecurity.js.
 * This file intentionally contains only one public implementation
 * for each legacy endpoint so there are no duplicate function names.
 */

const PRESENSI_CONFIG = {
  SHEET_SISWA: "SISWA",
  SHEET_KELAS: "KELAS",
  SHEET_TRANSAKSI: "TRX_PRESENSI",
  SHEET_LOG: "LOG",
  STATUS_PRESENSI: ["HADIR", "IZIN", "SAKIT", "ALPA"],
};

/**
 * Legacy endpoint: load menu/context.
 * Uses the secured Presensi service.
 */
function loadPresensiPerkelas() {
  return securePresensiLoad();
}

/**
 * Legacy endpoint used by older frontend code.
 * Permission is enforced by securePresensiLoad().
 */
function getKelasPresensiPerkelas() {
  return securePresensiLoad();
}

/**
 * Legacy endpoint used by older frontend code.
 * Permission and school/kelas validation are enforced server-side.
 */
function getPresensiPerkelasData(tanggal, kelas) {
  return securePresensiGetData(tanggal, kelas);
}

/**
 * Legacy endpoint used by the current frontend.
 * The authoritative student identity/name comes from SISWA on the
 * active school spreadsheet; INPUT_PRESENSI is checked server-side.
 */
function simpanPresensiPerkelas(tanggal, kelas, data) {
  const result = securePresensiSave(tanggal, kelas, data);

  // Preserve the existing audit trail without making logging failure
  // invalidate a successful transaction.
  if (result && result.success) {
    try {
      const context = getCurrentUserContext();
      writePresensiLog_(context.school.spreadsheetId, {
        npsn: context.npsn || "",
        userId: context.userId || "",
        email: context.email || "",
        nip: context.nip || "",
        namaUser: context.nama || "",
        role: context.role || "",
        action: "SIMPAN",
        module: "PRESENSI_PERKELAS",
        description:
          "Presensi kelas " +
          result.kelas +
          " tanggal " +
          result.tanggal +
          ". INSERT: " +
          (result.inserted || 0) +
          ", UPDATE: " +
          (result.updated || 0),
        transactionId: result.transactionId || "",
      });
    } catch (error) {
      console.error("[PRESENSI LOG]", error);
    }
  }

  return result;
}

/**
 * Mengambil status presensi yang sudah tersimpan.
 * VIEW permission wajib dipenuhi sebelum membaca transaksi.
 */
function cekPresensiPerkelas(tanggal, kelas) {
  requirePermission("VIEW_PRESENSI");

  tanggal = String(tanggal || "").trim();
  kelas = String(kelas || "").trim();
  if (!tanggal || !kelas) return [];

  const context = getCurrentUserContext();
  const ss = getSchoolSpreadsheet_();
  const sheet = ss.getSheetByName(PRESENSI_CONFIG.SHEET_TRANSAKSI);
  if (!sheet) {
    throw new Error("Sheet TRX_PRESENSI tidak ditemukan.");
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = normalizeHeaders_(values[0]);
  const tanggalIndex = findHeaderIndex_(headers, ["TANGGAL"]);
  const kelasIndex = findHeaderIndex_(headers, ["KELAS"]);
  const nisnIndex = findHeaderIndex_(headers, ["NISN"]);
  const statusIndex = findHeaderIndex_(headers, ["STATUS"]);

  if (
    tanggalIndex < 0 ||
    kelasIndex < 0 ||
    nisnIndex < 0 ||
    statusIndex < 0
  ) {
    return [];
  }

  const hasil = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowTanggal = normalizeDateString_(row[tanggalIndex]);
    const rowKelas = String(row[kelasIndex] || "").trim();

    if (rowTanggal !== tanggal) continue;
    if (rowKelas.toUpperCase() !== kelas.toUpperCase()) continue;

    hasil.push({
      nisn: String(row[nisnIndex] || "").trim(),
      status: String(row[statusIndex] || "").trim().toUpperCase(),
    });
  }

  return hasil;
}

/**
 * Backward-compatible helper for code that previously called the
 * private class loader directly.
 */
function getKelasPresensi_(ss) {
  if (!ss) throw new Error("Spreadsheet sekolah tidak tersedia.");
  const sheet = ss.getSheetByName(PRESENSI_CONFIG.SHEET_KELAS);
  if (!sheet) throw new Error("Sheet KELAS tidak ditemukan.");

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = normalizeHeaders_(values[0]);
  const kelasIndex = findHeaderIndex_(headers, [
    "KELAS",
    "NAMA_KELAS",
    "ROMBEL",
  ]);
  if (kelasIndex < 0) throw new Error("Kolom KELAS tidak ditemukan.");

  const statusIndex = findHeaderIndex_(headers, ["STATUS"]);
  const seen = {};
  const kelas = [];

  for (let i = 1; i < values.length; i++) {
    const namaKelas = String(values[i][kelasIndex] || "").trim();
    if (!namaKelas) continue;

    if (statusIndex >= 0) {
      const status = String(values[i][statusIndex] || "")
        .trim()
        .toUpperCase();
      if (status && status !== "ACTIVE" && status !== "AKTIF") continue;
    }

    const key = namaKelas.toUpperCase();
    if (!seen[key]) {
      seen[key] = true;
      kelas.push(namaKelas);
    }
  }

  return kelas.sort(function (a, b) {
    return String(a).localeCompare(String(b), "id", {
      numeric: true,
      sensitivity: "base",
    });
  });
}

function normalizeHeaders_(headers) {
  return headers.map(function (header) {
    return String(header || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "_");
  });
}

function findHeaderIndex_(headers, names) {
  for (let i = 0; i < names.length; i++) {
    const target = String(names[i] || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "_");
    const index = headers.indexOf(target);
    if (index >= 0) return index;
  }
  return -1;
}

function normalizeDateString_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || "Asia/Jakarta",
      "yyyy-MM-dd",
    );
  }

  const text = String(value || "").trim();
  if (!text) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const date = new Date(text);
  if (isNaN(date.getTime())) return text;

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone() || "Asia/Jakarta",
    "yyyy-MM-dd",
  );
}

/**
 * Audit logger for Presensi.
 * Logging failure must never roll back a valid attendance transaction.
 */
function writePresensiLog_(spreadsheetId, logData) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(PRESENSI_CONFIG.SHEET_LOG);
    const headers = [
      "TIMESTAMP",
      "NPSN",
      "USER_ID",
      "EMAIL",
      "NIP",
      "NAMA_USER",
      "ROLE",
      "ACTION",
      "MODULE",
      "DESCRIPTION",
      "TRANSACTION_ID",
    ];

    if (!sheet) {
      sheet = ss.insertSheet(PRESENSI_CONFIG.SHEET_LOG);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }

    const row = [
      new Date(),
      logData.npsn || "",
      logData.userId || "",
      logData.email || "",
      logData.nip || "",
      logData.namaUser || "",
      logData.role || "",
      logData.action || "SIMPAN",
      logData.module || "PRESENSI_PERKELAS",
      logData.description || "",
      logData.transactionId || "",
    ];

    const targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    return {
      success: true,
      sheet: PRESENSI_CONFIG.SHEET_LOG,
      row: targetRow,
      transactionId: logData.transactionId || "",
    };
  } catch (error) {
    console.error("[LOG PRESENSI]", error);
    return {
      success: false,
      error: error && error.message ? error.message : String(error),
    };
  }
}

function testLoadPresensiPerkelas() {
  return loadPresensiPerkelas();
}

function testKelasPresensiPerkelas() {
  return getKelasPresensiPerkelas();
}

function testCekPresensiPerkelas() {
  const tanggal = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || "Asia/Jakarta",
    "yyyy-MM-dd",
  );
  return cekPresensiPerkelas(tanggal, "XI-A");
}
