/**
 * GATEWAY CONFIGURATION
 *
 * Configure these values in Script Properties of the separate Gateway project.
 * Do NOT put credentials/secrets in the main SIM SATRIA project.
 *
 * Required properties:
 *   GATEWAY_NPSN
 *   GATEWAY_SPREADSHEET_ID
 *
 * Optional:
 *   GATEWAY_DRIVE_FOLDER_ID
 */
const GATEWAY_CONFIG = Object.freeze({
  VERSION: "1.0.0",
  ALLOWED_ROLES: ["ADMIN_SEKOLAH", "GURU", "WALI_KELAS", "KARYAWAN", "SISWA"],
  ACTIVE_STATUS: "ACTIVE",
  MAX_BODY_BYTES: 900000,
  MAX_CELL_CHARS: 45000,
});

const GATEWAY_WRITE_ALLOWLIST = Object.freeze({
  TRX_PRESENSI: ["TANGGAL", "KELAS", "NISN", "NAMA_SISWA", "STATUS", "KETERANGAN"],
  TRX_PARKIR: ["TANGGAL", "KENDALA", "SOLUSI", "UPLOAD_FOTO_PARKIR"],
  TRX_PRESTASI: ["TANGGAL", "NAMA_SISWA", "JENIS", "TINGKAT", "KETERANGAN"],
  TRX_AGENDA_GURU: [
    "TANGGAL", "SESI", "KELAS", "TUJUAN_PEMBELAJARAN", "MATERI_PEMBELAJARAN",
    "DPL", "PENGALAMAN_BELAJAR", "PRINSIP_PEMBELAJARAN", "REKAP_MURID_TIDAK_IKUT",
    "BUKTI_FISIK", "NAMA_GURU", "MAPEL", "KETERANGAN",
  ],
  TRX_SBI: ["INDIKATOR", "SUBINDIKATOR", "URAIAN_KEGIATAN", "HAMBATAN", "SOLUSI", "KARAKTER", "BUKTI_FISIK"],
  TRX_KEBERSIHAN: ["TANGGAL", "KENDALA", "SOLUSI", "BUKTI_FISIK"],
  TRX_KEAMANAN: ["TANGGAL", "KENDALA", "SOLUSI", "BUKTI_FISIK"],
  TRX_KERJA: [
    "TANGGAL_PELAKSANAAN", "SESI", "BIDANG_TUGAS", "TARGET_PEKERJAAN", "URAIAN_PEKERJAAN",
    "KENDALA", "TINDAK_LANJUT", "REFLEKSI", "BUKTI_FISIK",
  ],
  LOG: ["ACTION", "MODULE", "DESCRIPTION", "TRANSACTION_ID"],
});

const GATEWAY_READ_ALLOWLIST = Object.freeze([
  "CONFIG", "GURU", "SISWA", "KARYAWAN", "KELAS",
  "TRX_PRESENSI", "TRX_PARKIR", "TRX_PRESTASI", "TRX_AGENDA_GURU",
  "TRX_SBI", "TRX_KEBERSIHAN", "TRX_KEAMANAN", "TRX_KERJA", "LOG",
]);

function getGatewayProperty_(name, required) {
  const value = String(PropertiesService.getScriptProperties().getProperty(name) || "").trim();
  if (required && !value) throw new Error("Script Property " + name + " belum dikonfigurasi pada Write Gateway.");
  return value;
}

function gatewayContext_() {
  return {
    npsn: getGatewayProperty_("GATEWAY_NPSN", true),
    spreadsheetId: getGatewayProperty_("GATEWAY_SPREADSHEET_ID", true),
    driveFolderId: getGatewayProperty_("GATEWAY_DRIVE_FOLDER_ID", false),
  };
}
