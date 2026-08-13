/**
 * SIM SATRIA - AUTHENTICATION
 *
 * SKEMA FINAL:
 * 1. SUPERADMIN hanya satu akun pemilik aplikasi.
 * 2. ADMIN_SEKOLAH ditentukan dari MASTER / registry autentikasi.
 * 3. GURU/WALI_KELAS/KARYAWAN/SISWA ditentukan dari USERS pada
 *    Spreadsheet sekolah masing-masing.
 * 4. Binding lokal hanya menjadi locator sekolah untuk user non-admin.
 * 5. Data bisnis tidak pernah memilih Spreadsheet dari frontend.
 *
 * Web App tetap USER_ACCESSING agar Session.getActiveUser() mengenali
 * akun Google pengguna.
 */
const AUTH_CONFIG = {
  ADMIN_SHEET: "ADMIN_SEKOLAH",
  SCHOOLS_SHEET: "SCHOOLS",
  USERS_SHEET: "USERS",
  USER_BINDINGS_PROPERTY: "SIM_SATRIA_USER_BINDINGS_V2",
  ACTIVE_STATUS: "ACTIVE",
  SUPERADMIN_EMAIL: "yazid.mubasir12@admin.sma.belajar.id",
  CACHE_SECONDS: 21600,
  CONTEXT_CACHE_VERSION: "V7",
  ADMIN_CACHE_VERSION: "V3",
};

function normalizeEmail_(email) {
  return String(email || "").trim().toLowerCase();
}

function normalizeNpsn_(npsn) {
  return String(npsn || "").trim();
}

function normalizeAuthRole_(role) {
  return String(role || "").trim().toUpperCase();
}

function getGoogleUserEmail_() {
  const user = Session.getActiveUser();
  const email = user ? normalizeEmail_(user.getEmail()) : "";
  if (!email) {
    throw new Error(
      'Identitas akun Google tidak dapat diperoleh. Pastikan deployment menggunakan "User accessing the web app" dan akun sudah memberikan otorisasi.',
    );
  }
  return email;
}

function isSuperAdminEmail_(email) {
  return normalizeEmail_(email) === normalizeEmail_(AUTH_CONFIG.SUPERADMIN_EMAIL);
}

function getAdminSheet_() {
  const sheet = getMasterSpreadsheet_().getSheetByName(AUTH_CONFIG.ADMIN_SHEET);
  if (!sheet) {
    throw new Error(
      "Sheet ADMIN_SEKOLAH tidak ditemukan pada Spreadsheet Master. Jalankan syncMasterAuthRegistry().",
    );
  }
  return sheet;
}

function getSchoolsSheet_() {
  const sheet = getMasterSpreadsheet_().getSheetByName(AUTH_CONFIG.SCHOOLS_SHEET);
  if (!sheet) {
    throw new Error(
      "Sheet SCHOOLS tidak ditemukan pada Spreadsheet Master. Jalankan syncMasterAuthRegistry().",
    );
  }
  return sheet;
}

function sheetRowsAsObjects_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });
  return values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function getAdminByEmail_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return null;

  const cache = CacheService.getScriptCache();
  const safe = normalizedEmail.replace(/[^a-zA-Z0-9]/g, "_");
  const key = "ADMIN_" + AUTH_CONFIG.ADMIN_CACHE_VERSION + "_" + safe;
  const cached = cache.get(key);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      cache.remove(key);
    }
  }

  const sheet = getAdminSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });
  const emailIndex = headers.indexOf("EMAIL");
  if (emailIndex === -1) {
    throw new Error("Kolom EMAIL tidak ditemukan di ADMIN_SEKOLAH.");
  }

  for (let i = 1; i < values.length; i++) {
    if (normalizeEmail_(values[i][emailIndex]) === normalizedEmail) {
      const admin = {};
      headers.forEach(function (header, index) {
        admin[header] = values[i][index];
      });
      cache.put(key, JSON.stringify(admin), AUTH_CONFIG.CACHE_SECONDS);
      return admin;
    }
  }

  // SUPERADMIN adalah identitas global aplikasi. Jika emailnya belum
  // tercantum di ADMIN_SEKOLAH, tetap dikenali sebagai SUPERADMIN, tetapi
  // konteks sekolah tetap harus tersedia jika modul sekolah akan digunakan.
  if (isSuperAdminEmail_(normalizedEmail)) {
    return {
      EMAIL: normalizedEmail,
      NAMA: "Pemilik Aplikasi",
      ROLE: "SUPERADMIN",
      STATUS: "ACTIVE",
      NPSN: "",
    };
  }

  return null;
}

function getSchoolByNpsnAuth_(npsn) {
  const normalizedNpsn = normalizeNpsn_(npsn);
  if (!normalizedNpsn) return null;

  const cache = CacheService.getScriptCache();
  const key = "SCHOOL_" + normalizedNpsn;
  const cached = cache.get(key);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      cache.remove(key);
    }
  }

  const sheet = getSchoolsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0].map(function (h) {
    return String(h || "").trim().toUpperCase();
  });
  const npsnIndex = headers.indexOf("NPSN");
  if (npsnIndex === -1) {
    throw new Error("Kolom NPSN tidak ditemukan di SCHOOLS.");
  }

  for (let i = 1; i < values.length; i++) {
    if (normalizeNpsn_(values[i][npsnIndex]) === normalizedNpsn) {
      const school = {};
      headers.forEach(function (header, index) {
        school[header] = values[i][index];
      });
      cache.put(key, JSON.stringify(school), AUTH_CONFIG.CACHE_SECONDS);
      return school;
    }
  }
  return null;
}

function getSchoolUserByEmail_(spreadsheetId, email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!spreadsheetId || !normalizedEmail) return null;

  let ss;
  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    throw new Error(
      "Akun " + normalizedEmail + " belum dapat membaca Spreadsheet sekolah. " +
      "Pastikan akun diberi akses ke database sekolah. Detail: " + e.message,
    );
  }

  const sheet = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
  if (!sheet) return null;

  const rows = sheetRowsAsObjects_(sheet);
  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail_(rows[i].EMAIL) === normalizedEmail) {
      return rows[i];
    }
  }
  return null;
}

function buildSchoolContext_(user, school, email) {
  const spreadsheetId = String(school.SPREADSHEET_ID || "").trim();
  if (!spreadsheetId) {
    throw new Error("Spreadsheet sekolah belum dikonfigurasi.");
  }

  return {
    authenticated: true,
    email: normalizeEmail_(email),
    userId: String(user.USER_ID || "").trim(),
    nip: String(user.NIP || "").trim(),
    nama: String(user.NAMA || "").trim(),
    role: normalizeAuthRole_(user.ROLE || "GURU"),
    npsn: normalizeNpsn_(school.NPSN),
    school: {
      npsn: normalizeNpsn_(school.NPSN),
      namaSekolah: String(school.NAMA_SEKOLAH || "").trim(),
      spreadsheetId: spreadsheetId,
      driveFolderId: String(school.DRIVE_FOLDER_ID || "").trim(),
      alamat: String(school.ALAMAT || "").trim(),
      logoUrl: String(school.LOGO_URL || "").trim(),
      tagline: String(school.TAGLINE || "").trim(),
      warnaUtama: String(school.WARNA_UTAMA || "").trim(),
      warnaSekunder: String(school.WARNA_SEKUNDER || "").trim(),
    },
  };
}

function getUserBindings_() {
  const raw = PropertiesService.getScriptProperties().getProperty(
    AUTH_CONFIG.USER_BINDINGS_PROPERTY,
  );
  if (!raw) return {};
  try {
    const data = JSON.parse(raw);
    return data && typeof data === "object" ? data : {};
  } catch (e) {
    return {};
  }
}

function saveUserBindings_(bindings) {
  PropertiesService.getScriptProperties().setProperty(
    AUTH_CONFIG.USER_BINDINGS_PROPERTY,
    JSON.stringify(bindings || {}),
  );
}

function registerSchoolUserBinding_(context, user) {
  if (!context || !context.school || !user) {
    throw new Error("Data binding pengguna tidak lengkap.");
  }

  const email = normalizeEmail_(user.EMAIL || user.email);
  if (!email) throw new Error("Email pengguna wajib diisi.");

  const role = normalizeAuthRole_(user.ROLE || user.role);
  if (!["GURU", "WALI_KELAS", "KARYAWAN", "SISWA"].includes(role)) {
    throw new Error("Role pengguna sekolah tidak diizinkan.");
  }

  const status = String(user.STATUS || user.status || "ACTIVE")
    .trim()
    .toUpperCase();

  const bindings = getUserBindings_();
  bindings[email] = {
    userId: String(user.USER_ID || user.userId || "").trim(),
    nip: String(user.NIP || user.nip || "").trim(),
    nama: String(user.NAMA || user.nama || "").trim(),
    role: role,
    status: status,
    npsn: normalizeNpsn_(context.npsn),
    spreadsheetId: String(context.school.spreadsheetId || "").trim(),
    namaSekolah: String(context.school.namaSekolah || "").trim(),
    driveFolderId: String(context.school.driveFolderId || "").trim(),
    alamat: String(context.school.alamat || "").trim(),
    logoUrl: String(context.school.logoUrl || "").trim(),
    tagline: String(context.school.tagline || "").trim(),
    warnaUtama: String(context.school.warnaUtama || "").trim(),
    warnaSekunder: String(context.school.warnaSekunder || "").trim(),
    updatedAt: new Date().toISOString(),
  };

  saveUserBindings_(bindings);
  clearUserContextCache_(email);
  return bindings[email];
}

function removeSchoolUserBinding_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return;
  const bindings = getUserBindings_();
  if (Object.prototype.hasOwnProperty.call(bindings, normalizedEmail)) {
    delete bindings[normalizedEmail];
    saveUserBindings_(bindings);
  }
  clearUserContextCache_(normalizedEmail);
}

function getBoundSchoolUserContext_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return null;

  const bindings = getUserBindings_();
  const binding = bindings[normalizedEmail];
  if (!binding) return null;

  if (String(binding.status || "").trim().toUpperCase() !== AUTH_CONFIG.ACTIVE_STATUS) {
    throw new Error("Akun pengguna sekolah tidak aktif. Hubungi ADMIN_SEKOLAH.");
  }

  if (!binding.npsn || !binding.spreadsheetId) {
    throw new Error("Binding akun sekolah belum lengkap. Hubungi ADMIN_SEKOLAH.");
  }

  const school = {
    NPSN: binding.npsn,
    NAMA_SEKOLAH: binding.namaSekolah,
    STATUS: AUTH_CONFIG.ACTIVE_STATUS,
    SPREADSHEET_ID: binding.spreadsheetId,
    DRIVE_FOLDER_ID: binding.driveFolderId,
    ALAMAT: binding.alamat,
    LOGO_URL: binding.logoUrl,
    TAGLINE: binding.tagline,
    WARNA_UTAMA: binding.warnaUtama,
    WARNA_SEKUNDER: binding.warnaSekunder,
  };

  // Binding hanya menentukan SEKOLAH. Identitas user selalu dibaca ulang
  // dari USERS sekolah agar perubahan nama/role/status langsung berlaku.
  const user = getSchoolUserByEmail_(binding.spreadsheetId, normalizedEmail);
  if (!user) {
    throw new Error(
      'Akun "' + normalizedEmail + '" tidak ditemukan pada USERS sekolah NPSN ' +
      binding.npsn + ". Hubungi ADMIN_SEKOLAH.",
    );
  }

  const role = normalizeAuthRole_(user.ROLE);
  const status = String(user.STATUS || "").trim().toUpperCase();
  const allowed = ["GURU", "WALI_KELAS", "KARYAWAN", "SISWA"];
  if (!allowed.includes(role)) {
    throw new Error(
      "Role " + role + " pada USERS tidak valid untuk akun pengguna sekolah.",
    );
  }
  if (status !== AUTH_CONFIG.ACTIVE_STATUS) {
    throw new Error("Akun pengguna sekolah tidak aktif. Hubungi ADMIN_SEKOLAH.");
  }

  // Segarkan binding dari USERS agar binding lama tidak pernah menjadi sumber
  // kebenaran untuk nama, NIP, role, atau status.
  const refreshedBinding = {
    USER_ID: user.USER_ID,
    EMAIL: normalizedEmail,
    NIP: user.NIP,
    NAMA: user.NAMA,
    ROLE: role,
    STATUS: status,
  };
  registerSchoolUserBinding_({ npsn: binding.npsn, school: school }, refreshedBinding);

  return buildSchoolContext_(user, school, normalizedEmail);
}

function getCurrentUserContext() {
  const email = getGoogleUserEmail_();

  // Jangan cache context user. Role/Nama/Status pada USERS harus selalu
  // menjadi sumber kebenaran sehingga perubahan ADMIN_SEKOLAH tidak tertahan
  // oleh cache lama dan akun GURU seperti AYA tidak terbawa lagi.
  const admin = getAdminByEmail_(email);

  if (admin) {
    const status = String(admin.STATUS || "").trim().toUpperCase();
    if (status !== AUTH_CONFIG.ACTIVE_STATUS) {
      throw new Error("Akun administrator sekolah tidak aktif.");
    }

    const isSuperAdmin = isSuperAdminEmail_(email);
    const role = isSuperAdmin ? "SUPERADMIN" : normalizeAuthRole_(admin.ROLE);

    if (!isSuperAdmin && role !== "ADMIN_SEKOLAH") {
      // ADMIN_SEKOLAH tetap harus berasal dari MASTER. Jangan menerima role
      // admin yang hanya ditulis di USERS sekolah.
      throw new Error("Akun administrator tidak memiliki role ADMIN_SEKOLAH yang sah pada MASTER.");
    }

    const npsn = normalizeNpsn_(admin.NPSN);
    if (!npsn) {
      throw new Error(
        isSuperAdmin
          ? "SUPERADMIN terautentikasi secara global. Untuk membuka modul sekolah, akun pemilik harus memiliki NPSN pada ADMIN_SEKOLAH atau context sekolah yang sah."
          : "Akun administrator belum memiliki NPSN sekolah.",
      );
    }

    const school = getSchoolByNpsnAuth_(npsn);
    if (!school) {
      throw new Error(
        "Sekolah dengan NPSN " + npsn + " tidak ditemukan pada registry SIM SATRIA.",
      );
    }

    const schoolStatus = String(school.STATUS || "").trim().toUpperCase();
    if (schoolStatus && schoolStatus !== AUTH_CONFIG.ACTIVE_STATUS) {
      throw new Error("Sekolah Anda tidak aktif pada SIM SATRIA.");
    }

    const adminUser = Object.assign({}, admin, { ROLE: role });
    // SUPERADMIN tetap boleh menggunakan context sekolah yang tercantum pada
    // ADMIN_SEKOLAH, tetapi role globalnya tetap SUPERADMIN.
    return buildSchoolContext_(adminUser, school, email);
  }

  // User sekolah sama sekali tidak perlu membuka MASTER. Binding lokal hanya
  // menyimpan locator sekolah, sedangkan USERS sekolah menjadi sumber data user.
  const boundUser = getBoundSchoolUserContext_(email);
  if (boundUser) return boundUser;

  throw new Error(
    'Akun Google "' +
      email +
      '" belum terdaftar sebagai pengguna sekolah. Hubungi ADMIN_SEKOLAH sekolah Anda.',
  );
}

function checkAuthentication() {
  try {
    const c = getCurrentUserContext();
    return {
      success: true,
      authenticated: true,
      email: c.email,
      userId: c.userId,
      npsn: c.npsn,
      sekolah: c.school.namaSekolah,
      role: c.role,
    };
  } catch (e) {
    return {
      success: false,
      authenticated: false,
      message: e.message,
    };
  }
}

function bindMySchool(npsn) {
  const email = getGoogleUserEmail_();
  const requestedNpsn = normalizeNpsn_(npsn);
  if (!requestedNpsn) throw new Error("NPSN sekolah wajib diisi.");

  const admin = getAdminByEmail_(email);
  if (!admin) {
    throw new Error("Hanya ADMIN_SEKOLAH yang dapat melakukan binding sekolah.");
  }

  const adminNpsn = normalizeNpsn_(admin.NPSN);
  if (adminNpsn && adminNpsn !== requestedNpsn) {
    throw new Error(
      "NPSN yang diminta berbeda dengan NPSN akun administrator. Binding ditolak demi keamanan.",
    );
  }

  const school = getSchoolByNpsnAuth_(requestedNpsn);
  if (!school) throw new Error("NPSN sekolah tidak ditemukan pada SCHOOLS.");

  const spreadsheetId = String(school.SPREADSHEET_ID || "").trim();
  if (!spreadsheetId) throw new Error("SPREADSHEET_ID sekolah belum dikonfigurasi.");

  const user = getSchoolUserByEmail_(spreadsheetId, email);
  if (!user) throw new Error("Akun belum terdaftar pada USERS sekolah.");

  const context = buildSchoolContext_(Object.assign({}, admin, { ROLE: isSuperAdminEmail_(email) ? "SUPERADMIN" : "ADMIN_SEKOLAH" }), school, email);
  registerSchoolUserBinding_(context, user);
  clearMyAuthCache();

  return {
    success: true,
    message: "Binding pengguna berhasil dibuat.",
    email: email,
    npsn: requestedNpsn,
    sekolah: String(school.NAMA_SEKOLAH || "").trim(),
  };
}

function ensureLocalHeaders_(sheet, requiredHeaders) {
  const normalizedRequired = requiredHeaders.map(function (h) {
    return String(h || "").trim().toUpperCase();
  });

  if (sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, normalizedRequired.length).setValues([normalizedRequired]);
    sheet.setFrozenRows(1);
    return normalizedRequired;
  }

  const current = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(function (h) {
      return String(h || "").trim().toUpperCase();
    });

  normalizedRequired.forEach(function (header) {
    if (!current.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      current.push(header);
    }
  });

  sheet.setFrozenRows(1);
  return current;
}

function clearUserContextCache_(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) return;
  const safe = normalizedEmail.replace(/[^a-zA-Z0-9]/g, "_");
  const cache = CacheService.getScriptCache();
  cache.remove("USER_CONTEXT_" + AUTH_CONFIG.CONTEXT_CACHE_VERSION + "_" + safe);
  cache.remove("USER_CONTEXT_V6_" + safe);
  cache.remove("USER_CONTEXT_V5_" + safe);
  cache.remove("USER_CONTEXT_V4_" + safe);
  cache.remove("LOCAL_USER_" + safe);
  cache.remove("ADMIN_" + AUTH_CONFIG.ADMIN_CACHE_VERSION + "_" + safe);
  cache.remove("ADMIN_V2_" + safe);
  cache.remove("ADMIN_" + safe);
}

function clearMyAuthCache() {
  const email = getGoogleUserEmail_();
  clearUserContextCache_(email);
  return { success: true };
}

function refreshMySchoolContext() {
  clearMyAuthCache();
  const context = getCurrentUserContext();
  return {
    success: true,
    message: "School Context berhasil di-refresh.",
    email: context.email,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    spreadsheetId: context.school.spreadsheetId,
    driveFolderId: context.school.driveFolderId,
    role: context.role,
  };
}

function syncSchoolUserBinding(email) {
  const context = getCurrentUserContext();
  if (!["ADMIN_SEKOLAH", "SUPERADMIN"].includes(normalizeAuthRole_(context.role))) {
    throw new Error("Hanya ADMIN_SEKOLAH yang dapat melakukan sinkronisasi binding pengguna.");
  }

  const normalizedEmail = normalizeEmail_(email);
  if (!normalizedEmail) throw new Error("Email pengguna wajib diisi.");

  const user = getSchoolUserByEmail_(context.school.spreadsheetId, normalizedEmail);
  if (!user) throw new Error("Pengguna tidak ditemukan pada USERS sekolah.");

  const binding = registerSchoolUserBinding_(context, user);
  return {
    success: true,
    email: normalizedEmail,
    userId: binding.userId,
    role: binding.role,
    status: binding.status,
    npsn: binding.npsn,
    sekolah: binding.namaSekolah,
  };
}
