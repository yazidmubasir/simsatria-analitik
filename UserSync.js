/**
 * SIM SATRIA - SINKRONISASI SEMUA PENGGUNA SEKOLAH
 *
 * Admin Sekolah cukup memasukkan daftar email guru sekali saja.
 * Sistem akan:
 * 1. memastikan setiap email masuk USERS sekolah;
 * 2. membuat USER_ID bila belum ada;
 * 3. menetapkan role/status default GURU/ACTIVE;
 * 4. membuat school binding untuk autentikasi;
 * 5. memberikan akses spreadsheet sekolah untuk akun ACTIVE;
 * 6. membersihkan cache user.
 *
 * Input dapat berupa string dipisahkan baris baru, koma, atau titik koma.
 * Baris juga boleh memakai format:
 *   email|Nama Guru|NIP
 * Jika hanya email, nama/NIP yang sudah ada tidak akan ditimpa.
 */

function normalizeUserEmailList_(input) {
  const raw = String(input || "");
  const tokens = raw
    .split(/[\n,;]+/)
    .map(function (item) { return String(item || "").trim(); })
    .filter(Boolean);

  const seen = {};
  const result = [];

  tokens.forEach(function (token) {
    const parts = token.split("|").map(function (part) {
      return String(part || "").trim();
    });
    const email = normalizeEmail_(parts[0]);
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
    if (seen[email]) return;
    seen[email] = true;
    result.push({
      email: email,
      nama: parts[1] || "",
      nip: parts[2] || "",
    });
  });

  return result;
}

function syncAllSchoolUsers(emailList) {
  const context = requireUserManager_();
  const entries = normalizeUserEmailList_(emailList);
  if (!entries.length) {
    throw new Error("Daftar email guru belum diisi atau tidak ada email yang valid.");
  }

  const sheet = getOrCreateUsersSheet_(context.school.spreadsheetId);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader_);
  const index = {};
  headers.forEach(function (header, i) { index[header] = i; });

  ["USER_ID", "EMAIL", "NIP", "NAMA", "ROLE", "STATUS"].forEach(function (header) {
    if (index[header] === undefined) {
      throw new Error("Kolom " + header + " tidak ditemukan di USERS.");
    }
  });

  const rowByEmail = {};
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][index.EMAIL]);
    if (email) rowByEmail[email] = i + 1;
  }

  const result = {
    success: true,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    total: entries.length,
    created: 0,
    updated: 0,
    alreadyActive: 0,
    failed: 0,
    users: [],
  };

  entries.forEach(function (entry) {
    try {
      const email = entry.email;
      const rowNumber = rowByEmail[email] || -1;
      let existing = null;

      if (rowNumber > 0) {
        const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
        existing = {
          userId: String(rowValues[index.USER_ID] || "").trim(),
          nip: String(rowValues[index.NIP] || "").trim(),
          nama: String(rowValues[index.NAMA] || "").trim(),
          role: normalizeRole_(rowValues[index.ROLE]),
          status: normalizeRole_(rowValues[index.STATUS]),
        };
      }

      const userId = existing && existing.userId
        ? existing.userId
        : "USR-" + Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase();
      const nip = entry.nip || (existing ? existing.nip : "");
      const nama = entry.nama || (existing ? existing.nama : "Guru SIM SATRIA");
      const role = existing && existing.role ? existing.role : "GURU";
      const status = existing && existing.status ? existing.status : "ACTIVE";

      if (!USER_MANAGEMENT_CONFIG.ALLOWED_ROLES.includes(role)) {
        throw new Error("Role " + role + " tidak diizinkan untuk sinkronisasi.");
      }
      if (!["ACTIVE", "INACTIVE"].includes(status)) {
        throw new Error("Status " + status + " tidak valid.");
      }

      const row = new Array(headers.length).fill("");
      row[index.USER_ID] = userId;
      row[index.EMAIL] = email;
      row[index.NIP] = nip;
      row[index.NAMA] = nama;
      row[index.ROLE] = role;
      row[index.STATUS] = status;

      if (rowNumber > 0) {
        sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
        result.updated++;
      } else {
        const newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1, 1, headers.length).setValues([row]);
        rowByEmail[email] = newRow;
        result.created++;
      }

      const userRecord = {
        USER_ID: userId,
        EMAIL: email,
        NIP: nip,
        NAMA: nama,
        ROLE: role,
        STATUS: status,
      };

      const binding = registerSchoolUserBinding_(context, userRecord);

      if (status === "ACTIVE") {
        grantSchoolSpreadsheetEditor_(context.school.spreadsheetId, email);
        result.alreadyActive++;
      } else {
        revokeSchoolSpreadsheetAccess_(context.school.spreadsheetId, email);
      }

      clearUserContextCache_(email);

      result.users.push({
        email: email,
        nama: nama,
        nip: nip,
        role: role,
        status: status,
        userId: userId,
        npsn: binding.npsn,
        sekolah: binding.namaSekolah,
        success: true,
      });
    } catch (e) {
      result.failed++;
      result.users.push({
        email: entry.email,
        nama: entry.nama || "",
        success: false,
        error: e.message || String(e),
      });
    }
  });

  result.message =
    "Sinkronisasi selesai: " +
    result.created + " dibuat, " +
    result.updated + " diperbarui, " +
    result.failed + " gagal.";

  return result;
}

/**
 * Sinkronisasi ulang semua user yang SUDAH ada pada sheet USERS.
 * Tidak memerlukan input daftar email.
 */
function syncAllExistingSchoolUsers() {
  const context = requireUserManager_();
  const sheet = getOrCreateUsersSheet_(context.school.spreadsheetId);
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      success: true,
      total: 0,
      created: 0,
      updated: 0,
      failed: 0,
      users: [],
      message: "Belum ada pengguna pada USERS sekolah.",
    };
  }

  const headers = values[0].map(normalizeHeader_);
  const emailIndex = headers.indexOf("EMAIL");
  const nameIndex = headers.indexOf("NAMA");
  const nipIndex = headers.indexOf("NIP");
  if (emailIndex < 0) throw new Error("Kolom EMAIL tidak ditemukan di USERS.");

  const entries = [];
  for (let i = 1; i < values.length; i++) {
    const email = normalizeEmail_(values[i][emailIndex]);
    if (!email) continue;
    entries.push({
      email: email,
      nama: nameIndex >= 0 ? String(values[i][nameIndex] || "").trim() : "",
      nip: nipIndex >= 0 ? String(values[i][nipIndex] || "").trim() : "",
    });
  }

  return syncAllSchoolUsers(entries.map(function (entry) {
    return entry.email + "|" + entry.nama + "|" + entry.nip;
  }).join("\n"));
}

/**
 * Memasukkan daftar email guru ke USERS tanpa perlu membuka Spreadsheet.
 * Hanya ADMIN_SEKOLAH yang boleh memanggil fungsi ini.
 */
function importTeacherEmails(emailList) {
  return syncAllSchoolUsers(emailList);
}
