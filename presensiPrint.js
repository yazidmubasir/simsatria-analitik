/**
 * SIM SATRIA - PRESENSI PRINT ENGINE
 *
 * All public print/read endpoints are protected by VIEW_PRESENSI.
 */

function getKelasUntukCetakPresensi(tanggalAwal, tanggalAkhir) {
  requirePermission("VIEW_PRESENSI");

  tanggalAwal = String(tanggalAwal || "").trim();
  tanggalAkhir = String(tanggalAkhir || "").trim();

  if (!tanggalAwal) throw new Error("Tanggal awal belum dipilih.");
  if (!tanggalAkhir) throw new Error("Tanggal akhir belum dipilih.");
  if (tanggalAwal > tanggalAkhir) {
    throw new Error("Tanggal awal tidak boleh lebih besar dari tanggal akhir.");
  }

  const context = getCurrentUserContext();
  const spreadsheetId = String(
    context.school && context.school.spreadsheetId
      ? context.school.spreadsheetId
      : context.spreadsheetId || "",
  ).trim();

  if (!spreadsheetId) {
    throw new Error("Spreadsheet sekolah aktif tidak ditemukan.");
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("TRX_PRESENSI");
  if (!sheet) throw new Error("TRX_PRESENSI tidak ditemukan.");

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });
  const colTanggal = headers.indexOf("TANGGAL");
  const colKelas = headers.indexOf("KELAS");

  if (colTanggal < 0 || colKelas < 0) {
    throw new Error("Kolom TANGGAL atau KELAS tidak ditemukan.");
  }

  const kelasSet = {};
  for (let i = 1; i < values.length; i++) {
    const tanggal = String(values[i][colTanggal] || "").trim();
    const kelas = String(values[i][colKelas] || "").trim();

    if (
      tanggal >= tanggalAwal &&
      tanggal <= tanggalAkhir &&
      kelas
    ) {
      kelasSet[kelas] = true;
    }
  }

  return Object.keys(kelasSet).sort(function (a, b) {
    return String(a).localeCompare(String(b), "id", {
      numeric: true,
      sensitivity: "base",
    });
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
  if (tanggalAwal > tanggalAkhir) {
    throw new Error("Tanggal awal tidak boleh lebih besar dari tanggal akhir.");
  }

  const context = getCurrentUserContext();
  const school = context.school || {};
  const spreadsheetId = String(
    school.spreadsheetId || context.spreadsheetId || "",
  ).trim();
  const npsn = String(school.npsn || context.npsn || "").trim();
  const namaSekolah = String(
    school.namaSekolah || school.nama || context.sekolah || "",
  ).trim();

  if (!spreadsheetId) {
    throw new Error("Spreadsheet sekolah aktif tidak ditemukan.");
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("TRX_PRESENSI");
  if (!sheet) throw new Error("TRX_PRESENSI tidak ditemukan.");

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error("Belum ada data presensi.");

  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });
  const col = {
    tanggal: headers.indexOf("TANGGAL"),
    kelas: headers.indexOf("KELAS"),
    nisn: headers.indexOf("NISN"),
    nama: headers.indexOf("NAMA_SISWA"),
    status: headers.indexOf("STATUS"),
    keterangan: headers.indexOf("KETERANGAN"),
  };

  if (
    col.tanggal < 0 ||
    col.kelas < 0 ||
    col.nisn < 0 ||
    col.nama < 0 ||
    col.status < 0
  ) {
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
      siswaMap.set(nisn, {
        nisn: nisn,
        nama: nama,
        hadir: 0,
        sakit: 0,
        izin: 0,
        alpa: 0,
        lainnya: 0,
      });
    }

    const siswa = siswaMap.get(nisn);
    if (status === "HADIR") siswa.hadir++;
    else if (status === "SAKIT") siswa.sakit++;
    else if (status === "IZIN") siswa.izin++;
    else if (status === "ALPA") siswa.alpa++;
    else siswa.lainnya++;
  }

  const siswa = Array.from(siswaMap.values()).sort(function (a, b) {
    return String(a.nama).localeCompare(String(b.nama), "id", {
      sensitivity: "base",
    });
  });

  if (!siswa.length) {
    throw new Error(
      "Tidak ada data presensi kelas " +
        kelas +
        " pada periode tersebut.",
    );
  }

  const total = {
    hadir: 0,
    sakit: 0,
    izin: 0,
    alpa: 0,
    lainnya: 0,
  };

  siswa.forEach(function (item) {
    total.hadir += item.hadir;
    total.sakit += item.sakit;
    total.izin += item.izin;
    total.alpa += item.alpa;
    total.lainnya += item.lainnya;
  });

  const kepalaSekolah = getKepalaSekolahFromMaster_(npsn);
  const template = HtmlService.createTemplateFromFile(
    "presensiPerkelasPrintTemplate",
  );

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
  const fileName =
    "Presensi_" +
    safeKelas +
    "_" +
    tanggalAwal +
    "_" +
    tanggalAkhir +
    "_" +
    uniqueId +
    ".pdf";

  pdfBlob.setName(fileName);

  const folder = getPresensiFolder_(context);
  const file = folder.createFile(pdfBlob);

  try {
    writePresensiLog_(spreadsheetId, {
      npsn: npsn,
      userId: context.userId || "",
      email: context.email || "",
      nip: context.nip || "",
      namaUser: context.nama || "",
      role: context.role || "",
      action: "CETAK",
      module: "PRESENSI_PERKELAS",
      description:
        "Cetak PDF presensi kelas " +
        kelas +
        ", periode " +
        tanggalAwal +
        " s.d. " +
        tanggalAkhir +
        ", file: " +
        fileName,
      transactionId: "PRINT-" + uniqueId,
    });
  } catch (error) {
    console.error("[PRESENSI PRINT LOG]", error);
  }

  return {
    success: true,
    fileId: file.getId(),
    fileName: file.getName(),
    folderName: folder.getName(),
    url: file.getUrl(),
    kelas: kelas,
    tanggalAwal: tanggalAwal,
    tanggalAkhir: tanggalAkhir,
    jumlahSiswa: siswa.length,
    total: total,
  };
}

function getPresensiFolder_(context) {
  const school = context.school || {};
  const rootFolderId =
    school.driveRootFolderId ||
    school.driveFolderId ||
    context.driveRootFolderId ||
    context.driveFolderId ||
    "";

  if (rootFolderId) {
    const root = DriveApp.getFolderById(rootFolderId);
    const folders = root.getFoldersByName("PRESENSI");
    if (folders.hasNext()) return folders.next();
    return root.createFolder("PRESENSI");
  }

  const folders = DriveApp.getFoldersByName("PRESENSI");
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder("PRESENSI");
}

function getKepalaSekolahFromMaster_(npsn) {
  npsn = String(npsn || "").trim();
  if (!npsn) return { nama: "", nip: "" };

  const masterId = getMasterSpreadsheetId_();
  if (!masterId) {
    throw new Error("Spreadsheet SIM SATRIA MASTER belum dikonfigurasi.");
  }

  const ss = SpreadsheetApp.openById(masterId);
  const sheet = ss.getSheetByName("schools");
  if (!sheet) {
    throw new Error('Sheet "schools" pada SIM SATRIA MASTER tidak ditemukan.');
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { nama: "", nip: "" };

  const headers = values[0].map(function (h) {
    return String(h || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "_");
  });

  function findColumn(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const index = headers.indexOf(candidates[i]);
      if (index >= 0) return index;
    }
    return -1;
  }

  const colNpsn = findColumn(["NPSN"]);
  const colNama = findColumn([
    "NAMA_KEPALA_SEKOLAH",
    "KEPALA_SEKOLAH",
    "NAMA_KEPSEK",
    "KEPALA",
  ]);
  const colNip = findColumn([
    "NIP_KEPALA_SEKOLAH",
    "NIP_KEPSEK",
    "NIP_KEPALA",
    "NIP",
  ]);

  if (colNpsn < 0) {
    throw new Error("Kolom NPSN pada sheet schools tidak ditemukan.");
  }

  for (let i = 1; i < values.length; i++) {
    const rowNpsn = String(values[i][colNpsn] || "").trim();
    if (rowNpsn === npsn) {
      return {
        nama: colNama >= 0 ? String(values[i][colNama] || "").trim() : "",
        nip: colNip >= 0 ? String(values[i][colNip] || "").trim() : "",
      };
    }
  }

  return { nama: "", nip: "" };
}
