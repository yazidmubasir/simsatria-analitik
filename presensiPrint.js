/**
 * SIM SATRIA - PRESENSI PRINT ENGINE
 *
 * Semua PDF presensi WAJIB disimpan ke folder PRESENSI milik sekolah aktif.
 * Tidak ada fallback ke My Drive guru.
 * Guru tidak membaca Spreadsheet MASTER.
 */

function getKelasUntukCetakPresensi(tanggalAwal, tanggalAkhir) {
  requirePermission("VIEW_PRESENSI");
  tanggalAwal = String(tanggalAwal || "").trim();
  tanggalAkhir = String(tanggalAkhir || "").trim();
  if (!tanggalAwal) throw new Error("Tanggal awal belum dipilih.");
  if (!tanggalAkhir) throw new Error("Tanggal akhir belum dipilih.");
  if (tanggalAwal > tanggalAkhir) throw new Error("Tanggal awal tidak boleh lebih besar dari tanggal akhir.");

  const context = getCurrentUserContext();
  const spreadsheetId = String(context.school && context.school.spreadsheetId ? context.school.spreadsheetId : context.spreadsheetId || "").trim();
  if (!spreadsheetId) throw new Error("Spreadsheet sekolah aktif tidak ditemukan.");

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("TRX_PRESENSI");
  if (!sheet) throw new Error("TRX_PRESENSI tidak ditemukan.");
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
  const colTanggal = headers.indexOf("TANGGAL");
  const colKelas = headers.indexOf("KELAS");
  if (colTanggal < 0 || colKelas < 0) throw new Error("Kolom TANGGAL atau KELAS tidak ditemukan.");

  const kelasSet = {};
  for (let i = 1; i < values.length; i++) {
    const tanggal = String(values[i][colTanggal] || "").trim();
    const kelas = String(values[i][colKelas] || "").trim();
    if (tanggal >= tanggalAwal && tanggal <= tanggalAkhir && kelas) kelasSet[kelas] = true;
  }

  return Object.keys(kelasSet).sort(function(a, b) {
    return String(a).localeCompare(String(b), "id", { numeric: true, sensitivity: "base" });
  });
}

function cetakPresensiPerkelas(tanggalAwal, tanggalAkhir, kelas) {
  requirePermission("VIEW_PRESENSI");

  tanggalAwal = String(tanggalAwal || "").trim();
  tanggalAkhir = String(tanggalAkhir || "").trim();
  kelas = String(kelas || "").trim();
  if (!tanggalAwal) throw new Error("Tanggal awal belum dipilih.");
  if (!tanggalAkhir) throw new Error("Tanggal akhir belum dipilih.");
  if (!kelas) throw new Error("Kelas belum dipilih.");
  if (tanggalAwal > tanggalAkhir) throw new Error("Tanggal awal tidak boleh lebih besar dari tanggal akhir.");

  const context = getCurrentUserContext();
  const school = context.school || {};
  const spreadsheetId = String(school.spreadsheetId || context.spreadsheetId || "").trim();
  const npsn = String(school.npsn || context.npsn || "").trim();
  const namaSekolah = String(school.namaSekolah || school.nama || context.sekolah || "").trim();
  if (!spreadsheetId) throw new Error("Spreadsheet sekolah aktif tidak ditemukan.");

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("TRX_PRESENSI");
  if (!sheet) throw new Error("TRX_PRESENSI tidak ditemukan.");
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error("Belum ada data presensi.");

  const headers = values[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
  const col = {
    tanggal: headers.indexOf("TANGGAL"),
    kelas: headers.indexOf("KELAS"),
    nisn: headers.indexOf("NISN"),
    nama: headers.indexOf("NAMA_SISWA"),
    status: headers.indexOf("STATUS"),
    keterangan: headers.indexOf("KETERANGAN"),
  };
  if (col.tanggal < 0 || col.kelas < 0 || col.nisn < 0 || col.nama < 0 || col.status < 0) {
    throw new Error("Struktur TRX_PRESENSI tidak lengkap.");
  }

  const siswaMap = new Map();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const tanggal = String(row[col.tanggal] || "").trim();
    const rowKelas = String(row[col.kelas] || "").trim();
    if (tanggal < tanggalAwal || tanggal > tanggalAkhir) continue;
    if (rowKelas.toUpperCase() !== kelas.toUpperCase()) continue;

    const nisn = String(row[col.nisn] || "").trim();
    const nama = String(row[col.nama] || "").trim();
    const status = String(row[col.status] || "").trim().toUpperCase();
    if (!nisn) continue;

    if (!siswaMap.has(nisn)) {
      siswaMap.set(nisn, { nisn: nisn, nama: nama, hadir: 0, sakit: 0, izin: 0, alpa: 0, lainnya: 0 });
    }
    const siswa = siswaMap.get(nisn);
    if (status === "HADIR") siswa.hadir++;
    else if (status === "SAKIT") siswa.sakit++;
    else if (status === "IZIN") siswa.izin++;
    else if (status === "ALPA") siswa.alpa++;
    else siswa.lainnya++;
  }

  const siswa = Array.from(siswaMap.values()).sort(function(a, b) {
    return String(a.nama).localeCompare(String(b.nama), "id", { sensitivity: "base" });
  });
  if (!siswa.length) throw new Error("Tidak ada data presensi kelas " + kelas + " pada periode tersebut.");

  const total = { hadir: 0, sakit: 0, izin: 0, alpa: 0, lainnya: 0 };
  siswa.forEach(function(item) {
    total.hadir += item.hadir;
    total.sakit += item.sakit;
    total.izin += item.izin;
    total.alpa += item.alpa;
    total.lainnya += item.lainnya;
  });

  const kepalaSekolah = getKepalaSekolahForPrint_(context);
  const template = HtmlService.createTemplateFromFile("presensiPerkelasPrintTemplate");
  template.data = {
    namaSekolah: namaSekolah,
    npsn: npsn,
    kelas: kelas,
    tanggalAwal: tanggalAwal,
    tanggalAkhir: tanggalAkhir,
    siswa: siswa,
    total: total,
    kepalaSekolah: kepalaSekolah,
  };

  const html = template.evaluate().setTitle("Presensi " + kelas);
  const pdfBlob = html.getBlob().getAs(MimeType.PDF);
  const uniqueId = Utilities.getUuid().replace(/-/g, "").substring(0, 8).toUpperCase();
  const safeKelas = kelas.replace(/[\\/:*?"<>|]/g, "_");
  const fileName = "Presensi_" + safeKelas + "_" + tanggalAwal + "_" + tanggalAkhir + "_" + uniqueId + ".pdf";
  pdfBlob.setName(fileName);

  // WAJIB: simpan pada folder PRESENSI milik sekolah aktif.
  // Tidak boleh fallback ke My Drive guru karena akan membuat dokumen
  // sekolah tersebar pada Drive pribadi masing-masing guru.
  const fileResult = createPresensiPdfFile_(context, pdfBlob);
  const file = fileResult.file;

  try {
    writePresensiLog_(spreadsheetId, {
      npsp: npsn,
      npsn: npsn,
      userId: context.userId || "",
      email: context.email || "",
      nip: context.nip || "",
      namaUser: context.nama || "",
      role: context.role || "",
      action: "CETAK",
      module: "PRESENSI_PERKELAS",
      description: "Cetak PDF presensi kelas " + kelas + ", periode " + tanggalAwal + " s.d. " + tanggalAkhir + ", file: " + fileName,
      transactionId: "PRINT-" + uniqueId,
    });
  } catch (error) {
    console.error("[PRESENSI PRINT LOG]", error);
  }

  return {
    success: true,
    fileId: file.getId(),
    fileName: file.getName(),
    folderName: fileResult.folderName,
    folderId: fileResult.folderId,
    url: file.getUrl(),
    storage: "SCHOOL_DRIVE",
    kelas: kelas,
    tanggalAwal: tanggalAwal,
    tanggalAkhir: tanggalAkhir,
    jumlahSiswa: siswa.length,
    total: total,
  };
}

/**
 * Simpan PDF presensi secara deterministik ke:
 *
 *   Drive Sekolah
 *      └── PRESENSI
 *
 * Tidak ada fallback ke My Drive.
 */
function createPresensiPdfFile_(context, pdfBlob) {
  const school = context && context.school ? context.school : {};
  const rootFolderId = String(
    school.driveRootFolderId ||
    school.driveFolderId ||
    context.driveRootFolderId ||
    context.driveFolderId ||
    ""
  ).trim();

  if (!rootFolderId) {
    throw new Error(
      "Folder Drive sekolah belum dikonfigurasi. PDF tidak dibuat agar tidak tersimpan di My Drive guru."
    );
  }

  let root;
  try {
    root = DriveApp.getFolderById(rootFolderId);
  } catch (e) {
    throw new Error(
      "Folder Drive sekolah tidak dapat diakses oleh akun ini. Pastikan guru sudah mendapat permission ke folder sekolah. Detail: " +
      (e.message || String(e))
    );
  }

  let folder;
  try {
    const folders = root.getFoldersByName("PRESENSI");
    folder = folders.hasNext() ? folders.next() : root.createFolder("PRESENSI");
  } catch (e) {
    throw new Error(
      "Folder PRESENSI sekolah tidak dapat diakses/dibuat. Pastikan guru memiliki akses Editor ke folder sekolah. Detail: " +
      (e.message || String(e))
    );
  }

  let file;
  try {
    file = folder.createFile(pdfBlob);
  } catch (e) {
    throw new Error(
      "PDF gagal disimpan ke folder PRESENSI sekolah. Tidak ada fallback ke My Drive. Detail: " +
      (e.message || String(e))
    );
  }

  return {
    file: file,
    folderId: folder.getId(),
    folderName: folder.getName(),
    storage: "SCHOOL_DRIVE",
  };
}

function getKepalaSekolahForPrint_(context) {
  const school = context && context.school ? context.school : {};
  return {
    nama: String(school.namaKepalaSekolah || school.kepalaSekolah || context.namaKepalaSekolah || "").trim(),
    nip: String(school.nipKepalaSekolah || school.nipKepala || context.nipKepalaSekolah || "").trim(),
  };
}

// Kompatibilitas fungsi lama. Tidak membaca MASTER lagi.
function getKepalaSekolahFromMaster_(npsn) {
  return { nama: "", nip: "" };
}
