/**
 * SCHOOL REGISTRY
 *
 * MASTER_SPREADSHEET_ID disimpan di Script Properties.
 *
 * Penting:
 * Web App memakai USER_ACCESSING. Guru/admin sekolah tidak harus
 * mempunyai izin membaca MASTER. Jika MASTER tidak dapat dibuka oleh
 * akun aktif, getMasterSpreadsheet_() menyediakan read-only proxy dari
 * registry autentikasi yang sebelumnya disinkronkan oleh admin utama.
 */
function getMasterSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(
    "MASTER_SPREADSHEET_ID",
  );
  if (!id) {
    throw new Error(
      "MASTER_SPREADSHEET_ID belum dikonfigurasi pada Script Properties.",
    );
  }

  try {
    // Jika akun aktif memang mempunyai akses MASTER, gunakan spreadsheet asli.
    return SpreadsheetApp.openById(id);
  } catch (masterError) {
    // Jika tidak mempunyai akses MASTER, JANGAN menghentikan autentikasi
    // pengguna sekolah. getAdminByEmail_() akan mendapatkan null dari proxy
    // lalu getCurrentUserContext() dapat melanjutkan ke binding lokal GURU,
    // WALI_KELAS, atau KARYAWAN.
    return createLocalMasterRegistryProxy_();
  }
}

function createLocalMasterRegistryProxy_() {
  return {
    getSheetByName: function (sheetName) {
      const normalized = String(sheetName || "").trim().toUpperCase();

      if (normalized === "ADMIN_SEKOLAH") {
        return createLocalRegistrySheet_("ADMIN");
      }

      if (normalized === "SCHOOLS") {
        return createLocalRegistrySheet_("SCHOOLS");
      }

      return null;
    },
  };
}

function createLocalRegistrySheet_(type) {
  const props = PropertiesService.getScriptProperties();
  const prefix =
    type === "ADMIN"
      ? MASTER_AUTH_REGISTRY.ADMIN_PREFIX
      : MASTER_AUTH_REGISTRY.SCHOOL_PREFIX;

  const all = props.getProperties();
  const rows = [];

  Object.keys(all).forEach(function (key) {
    if (key.indexOf(prefix) !== 0) return;
    try {
      rows.push(JSON.parse(all[key]));
    } catch (e) {
      // Abaikan entry registry yang rusak.
    }
  });

  let headers = [];
  if (type === "ADMIN") {
    headers = ["USER_ID", "EMAIL", "NIP", "NAMA", "NPSN", "ROLE", "STATUS"];
  } else {
    headers = [
      "NPSN",
      "NAMA_SEKOLAH",
      "STATUS",
      "SPREADSHEET_ID",
      "DRIVE_FOLDER_ID",
      "ALAMAT",
      "LOGO_URL",
      "TAGLINE",
      "WARNA_UTAMA",
      "WARNA_SEKUNDER",
    ];
  }

  const values = [headers];
  rows.forEach(function (row) {
    values.push(
      headers.map(function (header) {
        return row[header] === undefined ? "" : row[header];
      }),
    );
  });

  return {
    getDataRange: function () {
      return {
        getValues: function () {
          return values;
        },
      };
    },
    getLastRow: function () {
      return values.length;
    },
    getLastColumn: function () {
      return headers.length;
    },
    getName: function () {
      return type === "ADMIN" ? "ADMIN_SEKOLAH" : "SCHOOLS";
    },
  };
}

function setMasterSpreadsheetId(id) {
  id = String(id || "").trim();
  if (!id) {
    throw new Error("ID Spreadsheet Master wajib diisi.");
  }
  SpreadsheetApp.openById(id);
  PropertiesService.getScriptProperties().setProperty(
    "MASTER_SPREADSHEET_ID",
    id,
  );
  return {
    success: true,
    spreadsheetId: id,
  };
}

function getSchoolByNpsn(npsn) {
  const school = getSchoolByNpsnAuth_(npsn);
  if (!school) {
    throw new Error("Sekolah tidak ditemukan.");
  }
  return school;
}
