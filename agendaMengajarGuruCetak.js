/**
 * ============================================================
 * SIM SATRIA MULTI SCHOOL
 * MODUL : TABEL & CETAK AGENDA GURU
 * ============================================================
 *
 * File ini menggantikan versi lama yang:
 * - menggunakan getActiveSpreadsheet()
 * - menggunakan ID folder Drive sekolah lama
 * - membaca kolom berdasarkan posisi tetap
 * - masih mengacu NAMA_GURU / MAPEL / KETERANGAN
 *
 * Prinsip:
 * 1. Spreadsheet selalu berasal dari School Context sekolah aktif.
 * 2. Semua pembacaan TRX_AGENDA_GURU berdasarkan NAMA HEADER.
 * 3. Tidak membuat NAMA_GURU, MAPEL, KETERANGAN.
 * 4. PDF disimpan pada folder sekolah aktif.
 * 5. Template PDF berada di dalam file GS ini sehingga modul
 *    hanya membutuhkan 2 file: .gs dan .html.
 * ============================================================
 */

/* ============================================================
   SCHOOL CONTEXT
   ============================================================ */

function agendaCetakGetSchoolContext_() {
  var candidates = [
    'getSchoolContextInfo',
    'getCurrentSchoolContext',
    'getMySchoolContext',
    'getSchoolContext'
  ];

  for (var i = 0; i < candidates.length; i++) {
    var fn = candidates[i];

    try {
      if (typeof this[fn] === 'function') {
        var ctx = this[fn]();

        if (ctx && typeof ctx === 'object') {
          return ctx;
        }
      }
    } catch (err) {
      // lanjut ke resolver berikutnya
    }
  }

  throw new Error(
    'School Context sekolah aktif tidak tersedia. ' +
    'Pastikan fungsi School Context utama SIM SATRIA sudah aktif.'
  );
}


/* ============================================================
   AMBIL SPREADSHEET SEKOLAH AKTIF
   ============================================================ */

function agendaCetakGetActiveSchoolSpreadsheet_() {

  var ctx = agendaCetakGetSchoolContext_();

  var id =
    ctx.spreadsheetId ||
    ctx.spreadsheet_id ||
    ctx.SPREADSHEET_ID ||
    (ctx.school && (
      ctx.school.spreadsheetId ||
      ctx.school.spreadsheet_id
    )) ||
    (ctx.resource && (
      ctx.resource.spreadsheetId ||
      ctx.resource.spreadsheet_id
    ));

  if (!id) {
    throw new Error(
      'Spreadsheet sekolah aktif belum tersedia pada School Context.'
    );
  }

  return SpreadsheetApp.openById(String(id));
}


/* ============================================================
   NORMALISASI HEADER
   ============================================================ */

function agendaCetakNormalizeHeader_(value) {

  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '_')
    .replace(/[^A-Z0-9_]/g, '');
}


/* ============================================================
   HEADER MAP
   ============================================================ */

function agendaCetakHeaderMap_(sheet) {

  if (!sheet) {
    throw new Error('Sheet tidak ditemukan.');
  }

  var lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return {};
  }

  var headers =
    sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0];

  var map = {};

  headers.forEach(function(header, index) {

    var key =
      agendaCetakNormalizeHeader_(header);

    if (key) {
      map[key] = index;
    }

  });

  return map;
}


/* ============================================================
   NILAI KOLOM DENGAN ALIAS
   ============================================================ */

function agendaCetakGetValue_(row, map, aliases) {

  for (var i = 0; i < aliases.length; i++) {

    var key =
      agendaCetakNormalizeHeader_(aliases[i]);

    if (
      Object.prototype.hasOwnProperty.call(map, key)
    ) {
      return row[map[key]] == null
        ? ''
        : row[map[key]];
    }

  }

  return '';
}


/* ============================================================
   SHEET TRX AGENDA
   ============================================================ */

function agendaCetakGetTransactionSheet_() {

  var ss =
    agendaCetakGetActiveSchoolSpreadsheet_();

  var sh =
    ss.getSheetByName('TRX_AGENDA_GURU');

  if (!sh) {
    throw new Error(
      'Sheet TRX_AGENDA_GURU tidak ditemukan pada Spreadsheet sekolah aktif.'
    );
  }

  return sh;
}


/* ============================================================
   GET ALL DATA
   ============================================================ */

function getAgendaGuruTable() {

  var sh =
    agendaCetakGetTransactionSheet_();

  var values =
    sh.getDataRange().getDisplayValues();

  if (values.length <= 1) {
    return [];
  }

  var map =
    agendaCetakHeaderMap_(sh);

  var hasil = [];

  for (var i = 1; i < values.length; i++) {

    var row =
      values[i];

    hasil.push({

      row: i + 1,

      transactionId:
        agendaCetakGetValue_(row, map, [
          'TRANSACTION_ID'
        ]),

      timestamp:
        agendaCetakGetValue_(row, map, [
          'TIMESTAMP'
        ]),

      nip:
        agendaCetakGetValue_(row, map, [
          'NIP'
        ]),

      namaUser:
        agendaCetakGetValue_(row, map, [
          'NAMA_USER'
        ]),

      tanggal:
        agendaCetakGetValue_(row, map, [
          'TANGGAL'
        ]),

      sesi:
        agendaCetakGetValue_(row, map, [
          'SESI'
        ]),

      kelas:
        agendaCetakGetValue_(row, map, [
          'KELAS'
        ]),

      tujuanPembelajaran:
        agendaCetakGetValue_(row, map, [
          'TUJUAN_PEMBELAJARAN',
          'TUJUAN'
        ]),

      dpl:
        agendaCetakGetValue_(row, map, [
          'DPL'
        ]),

      pengalamanBelajar:
        agendaCetakGetValue_(row, map, [
          'PENGALAMAN_BELAJAR',
          'PM'
        ]),

      prinsipPembelajaran:
        agendaCetakGetValue_(row, map, [
          'PRINSIP_PEMBELAJARAN',
          'PRINSIP'
        ]),

      rekapMuridTidakIkut:
        agendaCetakGetValue_(row, map, [
          'REKAP_MURID_TIDAK_IKUT',
          'SISWA_TIDAK_MASUK'
        ]),

      materiPembelajaran:
        agendaCetakGetValue_(row, map, [
          'MATERI_PEMBELAJARAN',
          'MATERI'
        ]),

      buktiFisik:
        agendaCetakGetValue_(row, map, [
          'BUKTI_FISIK',
          'FOTO'
        ])

    });

  }

  return hasil.reverse();
}


/* ============================================================
   GET NAMA GURU
   ============================================================ */

function getAgendaGuruListMultiSchool() {

  var ss =
    agendaCetakGetActiveSchoolSpreadsheet_();

  var sh =
    ss.getSheetByName('Guru');

  if (!sh) {
    throw new Error(
      'Sheet Guru tidak ditemukan pada Spreadsheet sekolah aktif.'
    );
  }

  var values =
    sh.getDataRange().getDisplayValues();

  if (values.length <= 1) {
    return [];
  }

  var lastColumn =
    sh.getLastColumn();

  var headers =
    values[0];

  var map = {};

  headers.forEach(function(header, index) {
    map[
      agendaCetakNormalizeHeader_(header)
    ] = index;
  });

  var idxNama =
    map.NAMA != null
      ? map.NAMA
      : (
        map.NAMA_GURU != null
          ? map.NAMA_GURU
          : 2
      );

  var idxNip =
    map.NIP != null
      ? map.NIP
      : 1;

  var hasil = [];
  var seen = {};

  for (var i = 1; i < values.length; i++) {

    var nama =
      String(values[i][idxNama] || '').trim();

    var nip =
      String(values[i][idxNip] || '').trim();

    if (!nama) continue;

    var key =
      nama.toUpperCase();

    if (!seen[key]) {

      seen[key] = true;

      hasil.push({
        nama: nama,
        nip: nip
      });

    }

  }

  hasil.sort(function(a, b) {
    return a.nama.localeCompare(
      b.nama,
      'id',
      { sensitivity: 'base' }
    );
  });

  return hasil;
}


/* ============================================================
   USER INFO / ROLE
   ============================================================ */

function getAgendaCurrentUserInfoMultiSchool() {

  var ctx =
    agendaCetakGetSchoolContext_();

  var user =
    ctx.user ||
    ctx.currentUser ||
    {};

  return {

    userId:
      user.userId ||
      user.user_id ||
      ctx.userId ||
      ctx.user_id ||
      '',

    email:
      user.email ||
      ctx.email ||
      '',

    nama:
      user.nama ||
      user.namaUser ||
      user.name ||
      ctx.namaUser ||
      ctx.nama ||
      '',

    role:
      user.role ||
      ctx.role ||
      '',

    nip:
      user.nip ||
      user.NIP ||
      ctx.nip ||
      ctx.NIP ||
      ''

  };

}


/* ============================================================
   VALIDASI AKSES GURU
   ============================================================ */

function agendaCetakCanAccessGuru_(nip, nama) {

  var info =
    getAgendaCurrentUserInfoMultiSchool();

  var role =
    String(info.role || '').toUpperCase();

  if (
    role !== 'GURU' &&
    role !== 'GURU_MAPEL'
  ) {
    return true;
  }

  var currentNip =
    String(info.nip || '').trim();

  var selectedNip =
    String(nip || '').trim();

  if (currentNip && selectedNip) {
    return currentNip === selectedNip;
  }

  return (
    String(nama || '').trim().toUpperCase() ===
    String(info.nama || '').trim().toUpperCase()
  );
}

/* ============================================================
   PARSE TANGGAL
   ============================================================ */

function parseTanggalAgenda(val) {

  if (!val) return null;

  if (val instanceof Date) {
    return new Date(val.getTime());
  }

  var str =
    String(val).trim();

  if (!str) return null;

  var p;

  if (str.indexOf('/') !== -1) {

    p = str.split('/');

    if (p.length === 3) {

      return new Date(
        parseInt(p[2], 10),
        parseInt(p[1], 10) - 1,
        parseInt(p[0], 10)
      );

    }

  }

  if (str.indexOf('-') !== -1) {

    p = str.split('-');

    if (p.length === 3) {

      if (p[0].length === 4) {

        return new Date(
          parseInt(p[0], 10),
          parseInt(p[1], 10) - 1,
          parseInt(p[2], 10)
        );

      }

      return new Date(
        parseInt(p[2], 10),
        parseInt(p[1], 10) - 1,
        parseInt(p[0], 10)
      );

    }

  }

  var d =
    new Date(str);

  return isNaN(d.getTime())
    ? null
    : d;
}


/* ============================================================
   FORMAT TANGGAL
   ============================================================ */

function formatTanggalIndo(tglStr) {

  if (!tglStr) return '-';

  var d =
    parseTanggalAgenda(tglStr);

  if (!d) return tglStr;

  var bulan = [
    'Januari',
    'Februari',
    'Maret',
    'April',
    'Mei',
    'Juni',
    'Juli',
    'Agustus',
    'September',
    'Oktober',
    'November',
    'Desember'
  ];

  return (
    String(d.getDate()).padStart(2, '0') +
    ' ' +
    bulan[d.getMonth()] +
    ' ' +
    d.getFullYear()
  );
}


/* ============================================================
   GET DATA BY NAMA + PERIODE
   ============================================================ */

function getAgendaGuruByFilterMultiSchool(
  nama,
  tglAwal,
  tglAkhir,
  nip
) {

  if (!nama) {
    throw new Error(
      'Nama guru wajib dipilih.'
    );
  }

  if (!nip) {
    throw new Error(
      'NIP guru wajib tersedia.'
    );
  }

  if (!tglAwal || !tglAkhir) {
    throw new Error(
      'Tanggal awal dan tanggal akhir wajib diisi.'
    );
  }


  const sh =
    agendaCetakGetTransactionSheet_();

  if (!sh) {
    throw new Error(
      'Sheet TRX_AGENDA_GURU tidak ditemukan.'
    );
  }


  const values =
    sh
      .getDataRange()
      .getDisplayValues();


  if (values.length <= 1) {
    return [];
  }


  const map =
    agendaCetakHeaderMap_(sh);


  const awal =
    parseTanggalAgenda(
      tglAwal
    );

  const akhir =
    parseTanggalAgenda(
      tglAkhir
    );


  if (!awal || !akhir) {
    throw new Error(
      'Tanggal tidak valid.'
    );
  }


  awal.setHours(
    0,
    0,
    0,
    0
  );

  akhir.setHours(
    23,
    59,
    59,
    999
  );


  /*
   * NIP GURU YANG DIPILIH
   */
  const selectedNip =
    String(nip)
      .trim()
      .replace(/\s+/g, '');


  const hasil = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];


    /* =====================================================
       NIP
       ===================================================== */

    const rowNip =
      String(
        agendaCetakGetValue_(
          row,
          map,
          ['NIP']
        ) || ''
      )
        .trim()
        .replace(/\s+/g, '');


    /*
     * IDENTITAS GURU HARUS NIP
     */
    if (
      rowNip !== selectedNip
    ) {
      continue;
    }


    /* =====================================================
       NAMA GURU
       ===================================================== */

    const namaGuru =
      agendaCetakGetValue_(
        row,
        map,
        ['NAMA_GURU']
      );


    /* =====================================================
       TANGGAL
       ===================================================== */

    const tanggal =
      agendaCetakGetValue_(
        row,
        map,
        ['TANGGAL']
      );


    const tanggalData =
      parseTanggalAgenda(
        tanggal
      );


    if (!tanggalData) {
      continue;
    }


    tanggalData.setHours(
      0,
      0,
      0,
      0
    );


    if (
      tanggalData < awal ||
      tanggalData > akhir
    ) {
      continue;
    }


    /* =====================================================
       DATA AGENDA
       ===================================================== */

    hasil.push({

      row:
        i + 1,

      transactionId:
        agendaCetakGetValue_(
          row,
          map,
          ['TRANSACTION_ID']
        ),

      timestamp:
        agendaCetakGetValue_(
          row,
          map,
          ['TIMESTAMP']
        ),

      npsn:
        agendaCetakGetValue_(
          row,
          map,
          ['NPSN']
        ),

      userId:
        agendaCetakGetValue_(
          row,
          map,
          ['USER_ID']
        ),

      email:
        agendaCetakGetValue_(
          row,
          map,
          ['EMAIL']
        ),

      nip:
        rowNip,

      namaUser:
        agendaCetakGetValue_(
          row,
          map,
          ['NAMA_USER']
        ),

      role:
        agendaCetakGetValue_(
          row,
          map,
          ['ROLE']
        ),

      namaGuru:
        namaGuru,

      mapel:
        agendaCetakGetValue_(
          row,
          map,
          ['MAPEL']
        ),

      keterangan:
        agendaCetakGetValue_(
          row,
          map,
          ['KETERANGAN']
        ),

      tanggal:
        tanggal,

      sesi:
        agendaCetakGetValue_(
          row,
          map,
          ['SESI']
        ),

      kelas:
        agendaCetakGetValue_(
          row,
          map,
          ['KELAS']
        ),

      tujuanPembelajaran:
        agendaCetakGetValue_(
          row,
          map,
          ['TUJUAN_PEMBELAJARAN']
        ),

      materiPembelajaran:
        agendaCetakGetValue_(
          row,
          map,
          ['MATERI_PEMBELAJARAN']
        ),

      dpl:
        agendaCetakGetValue_(
          row,
          map,
          ['DPL']
        ),

      pengalamanBelajar:
        agendaCetakGetValue_(
          row,
          map,
          ['PENGALAMAN_BELAJAR']
        ),

      prinsipPembelajaran:
        agendaCetakGetValue_(
          row,
          map,
          ['PRINSIP_PEMBELAJARAN']
        ),

      rekapMuridTidakIkut:
        agendaCetakGetValue_(
          row,
          map,
          ['REKAP_MURID_TIDAK_IKUT']
        ),

      buktiFisik:
        agendaCetakGetValue_(
          row,
          map,
          ['BUKTI_FISIK']
        )

    });

  }


  /*
   * URUTKAN BERDASARKAN TANGGAL
   */

  hasil.sort(
    function(a, b) {

      const da =
        parseTanggalAgenda(
          a.tanggal
        );

      const db =
        parseTanggalAgenda(
          b.tanggal
        );


      if (!da && !db) {
        return 0;
      }

      if (!da) {
        return 1;
      }

      if (!db) {
        return -1;
      }


      return (
        da.getTime() -
        db.getTime()
      );

    }
  );


  return hasil;

}


/* ============================================================
   HAPUS DATA
   ============================================================ */

function hapusAgendaGuruTable(row) {

  var sh =
    agendaCetakGetTransactionSheet_();

  var rowNumber =
    parseInt(row, 10);

  if (
    !rowNumber ||
    rowNumber < 2 ||
    rowNumber > sh.getLastRow()
  ) {
    throw new Error(
      'Nomor baris data tidak valid.'
    );
  }

  /*
   * Untuk keamanan Multi School, row hanya boleh dihapus
   * jika memang berada pada Spreadsheet sekolah aktif.
   */
  sh.deleteRow(rowNumber);

  return 'Data agenda berhasil dihapus.';
}


/* ============================================================
   MASTER SEKOLAH / KEPALA SEKOLAH
   ============================================================ */

function agendaCetakGetKepalaSekolah_() {

  var ctx =
    agendaCetakGetSchoolContext_();

  var school =
    ctx.school ||
    ctx.sekolah ||
    {};

  var nama =
    school.namaKepalaSekolah ||
    school.kepalaSekolah ||
    school.namaKepsek ||
    school.kepsekNama ||
    ctx.namaKepalaSekolah ||
    '';

  var nip =
    school.nipKepalaSekolah ||
    school.nipKepsek ||
    ctx.nipKepalaSekolah ||
    '';

  /*
   * Jika School Context sudah membawa data kepala sekolah,
   * tidak perlu membuka master.
   */
  if (nama || nip) {
    return {
      nama: nama,
      nip: nip
    };
  }

  /*
   * Fallback ke Spreadsheet MASTER bila fungsi/ID tersedia.
   */
  var masterId = '';

  try {

    if (typeof getMasterSpreadsheetId_ === 'function') {
      masterId =
        getMasterSpreadsheetId_() || '';
    }

  } catch (e) {}

  masterId =
    masterId ||
    ctx.masterSpreadsheetId ||
    ctx.master_spreadsheet_id ||
    PropertiesService
      .getScriptProperties()
      .getProperty('MASTER_SPREADSHEET_ID') ||
    '';

  if (!masterId) {
    return {
      nama: '',
      nip: ''
    };
  }

  try {

    var ss =
      SpreadsheetApp.openById(masterId);

    var sh =
      ss.getSheetByName('schools');

    if (!sh) {
      return {
        nama: '',
        nip: ''
      };
    }

    var values =
      sh.getDataRange().getDisplayValues();

    if (values.length <= 1) {
      return {
        nama: '',
        nip: ''
      };
    }

    var map =
      agendaCetakHeaderMap_(sh);

    var npsn =
      school.npsn ||
      ctx.npsn ||
      '';

    for (var i = 1; i < values.length; i++) {

      var row =
        values[i];

      var rowNpsn =
        agendaCetakGetValue_(row, map, [
          'NPSN'
        ]);

      if (
        String(rowNpsn).trim() !==
        String(npsn).trim()
      ) {
        continue;
      }

      nama =
        agendaCetakGetValue_(row, map, [
          'NAMA_KEPALA_SEKOLAH',
          'KEPALA_SEKOLAH',
          'NAMA_KEPSEK'
        ]);

      nip =
        agendaCetakGetValue_(row, map, [
          'NIP_KEPALA_SEKOLAH',
          'NIP_KEPSEK',
          'NIP'
        ]);

      break;
    }

  } catch (e) {
    // fallback tetap kosong, tidak menggagalkan laporan
  }

  return {
    nama: nama || '',
    nip: nip || ''
  };
}


/* ============================================================
   ESCAPE HTML
   ============================================================ */

function escapeHtml(str) {

  if (str == null) {
    return '';
  }

  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

}


/* ============================================================
   CETAK PDF
   ============================================================ */

function cetakPDFAgendaGuru(
  nama,
  nip,
  tglAwal,
  tglAkhir
) {

  if (!nama) {
    throw new Error(
      'Nama guru wajib dipilih.'
    );
  }

  if (!tglAwal || !tglAkhir) {
    throw new Error(
      'Tanggal awal dan tanggal akhir wajib diisi.'
    );
  }

  var awal =
    parseTanggalAgenda(tglAwal);

  var akhir =
    parseTanggalAgenda(tglAkhir);

  if (!awal || !akhir) {
    throw new Error(
      'Format tanggal tidak valid.'
    );
  }

  if (awal > akhir) {
    throw new Error(
      'Tanggal awal tidak boleh lebih besar dari tanggal akhir.'
    );
  }

  var data =
    getAgendaGuruByFilterMultiSchool(
      nama,
      tglAwal,
      tglAkhir,
      nip
    );

  if (!data.length) {
    throw new Error(
      'Data agenda mengajar tidak ditemukan untuk periode ini.'
    );
  }

  var ctx =
    agendaCetakGetSchoolContext_();

  var school =
    ctx.school ||
    ctx.sekolah ||
    {};

  var infoGuru =
    data[0];

  var kepala =
    agendaCetakGetKepalaSekolah_();

  var timeZone =
    Session.getScriptTimeZone();

  var tanggalCetak =
    Utilities.formatDate(
      new Date(),
      timeZone,
      'dd MMMM yyyy'
    );

  var rows = '';

  data.forEach(function(item, index) {

    var dpl =
      item.dpl || '-';

    var pengalaman =
      item.pengalamanBelajar || '-';

    var prinsip =
      item.prinsipPembelajaran || '-';

    var rekap =
      item.rekapMuridTidakIkut || '-';

    var materi =
      item.materiPembelajaran || '-';

    rows +=
      '<tr>' +

      '<td class="center">' +
      (index + 1) +
      '</td>' +

      '<td class="center">' +
      escapeHtml(item.tanggal || '-') +
      '</td>' +

      '<td class="center">' +
      escapeHtml(item.sesi || '-') +
      '</td>' +

      '<td class="center">' +
      escapeHtml(item.kelas || '-') +
      '</td>' +

      '<td>' +
      escapeHtml(item.tujuanPembelajaran || '-') +
      '</td>' +

      '<td>' +
      '<div><b>DPL:</b> ' +
      escapeHtml(dpl) +
      '</div>' +

      '<div><b>Pengalaman:</b> ' +
      escapeHtml(pengalaman) +
      '</div>' +

      '<div><b>Prinsip:</b> ' +
      escapeHtml(prinsip) +
      '</div>' +

      '</td>' +

      '<td>' +
      escapeHtml(rekap) +
      '</td>' +

      '<td>' +
      '<strong>' +
      escapeHtml(materi) +
      '</strong>' +

      '</td>' +

      '</tr>';

  });


  /*
   * Template PDF disimpan langsung di GS.
   * Dengan demikian hanya perlu 2 file:
   * Code GS + HTML modul.
   */
  var html =
    agendaCetakPdfTemplate_();

  html =
    html
      .replace(/\{\{NAMA_SEKOLAH\}\}/g,
        escapeHtml(
          school.namaSekolah ||
          school.nama ||
          ctx.namaSekolah ||
          ctx.sekolah ||
          '-'
        )
      )

      .replace(/\{\{NPSN\}\}/g,
        escapeHtml(
          school.npsn ||
          ctx.npsn ||
          '-'
        )
      )

      .replace(/\{\{NAMA_GURU\}\}/g,
        escapeHtml(nama)
      )

      .replace(/\{\{NIP_GURU\}\}/g,
        escapeHtml(infoGuru.nip || '-')
      )

      .replace(/\{\{PERIODE_AWAL\}\}/g,
        escapeHtml(formatTanggalIndo(tglAwal))
      )

      .replace(/\{\{PERIODE_AKHIR\}\}/g,
        escapeHtml(formatTanggalIndo(tglAkhir))
      )

      .replace(/\{\{TANGGAL_CETAK\}\}/g,
        escapeHtml(tanggalCetak)
      )

      .replace(/\{\{NAMA_KEPSEK\}\}/g,
        escapeHtml(kepala.nama || '-')
      )

      .replace(/\{\{NIP_KEPSEK\}\}/g,
        escapeHtml(kepala.nip || '-')
      )

      .replace(/\{\{TABEL_ROWS\}\}/g,
        rows);


  var blob =
    HtmlService
      .createHtmlOutput(html)
      .getBlob()
      .getAs(MimeType.PDF);


  /*
   * Folder PDF mengikuti folder sekolah aktif.
   */
  var folder =
    agendaCetakGetPdfFolder_();


  var safeName =
    String(nama)
      .replace(/[^a-zA-Z0-9._-]/g, '_')
      .replace(/_+/g, '_');


  var stamp =
    Utilities.formatDate(
      new Date(),
      timeZone,
      'yyyyMMdd_HHmmss'
    );


  var fileName =
    'AgendaGuru_' +
    safeName +
    '_' +
    stamp +
    '.pdf';


  var file =
    folder
      .createFile(blob)
      .setName(fileName);


  return file.getUrl();
}


/* ============================================================
   FOLDER PDF SEKOLAH AKTIF
   ============================================================ */

/* ============================================================
   FOLDER PDF AGENDA
   MULTI SCHOOL
   ============================================================ */

function agendaCetakGetPdfFolder_() {

  /*
   * ==========================================================
   * SCHOOL CONTEXT
   * ==========================================================
   */

  var ctx =
    agendaCetakGetSchoolContext_();


  if (!ctx) {

    throw new Error(
      'School Context sekolah aktif tidak tersedia.'
    );

  }


  /*
   * ==========================================================
   * 1. JIKA SCHOOL CONTEXT SUDAH MEMBERIKAN
   *    FOLDER AGENDA
   * ==========================================================
   */

  var resource =
    ctx.resource ||
    ctx.resources ||
    {};


  var school =
    ctx.school ||
    ctx.sekolah ||
    {};


  var agendaFolderId =
    resource.agendaPdfFolderId ||
    resource.agenda_pdf_folder_id ||
    resource.agendaFolderId ||
    resource.agenda_folder_id ||

    school.agendaPdfFolderId ||
    school.agenda_pdf_folder_id ||
    school.agendaFolderId ||
    school.agenda_folder_id ||

    ctx.agendaPdfFolderId ||
    ctx.agendaFolderId ||
    '';


  if (agendaFolderId) {

    try {

      return DriveApp.getFolderById(
        String(agendaFolderId)
      );

    } catch (e) {

      console.warn(
        'Folder AGENDA dari context tidak dapat dibuka: ' +
        e.message
      );

    }

  }


  /*
   * ==========================================================
   * 2. ROOT FOLDER SEKOLAH
   * ==========================================================
   */

  var rootId =
    resource.rootFolderId ||
    resource.root_folder_id ||
    resource.driveFolderId ||
    resource.drive_folder_id ||
    resource.folderId ||
    resource.folder_id ||

    school.rootFolderId ||
    school.root_folder_id ||
    school.driveFolderId ||
    school.drive_folder_id ||
    school.folderId ||
    school.folder_id ||

    ctx.rootFolderId ||
    ctx.driveFolderId ||
    ctx.folderId ||
    '';


  if (rootId) {

    try {

      var root =
        DriveApp.getFolderById(
          String(rootId)
        );


      return agendaCetakEnsureSubfolder_(
        root,
        'AGENDA'
      );

    } catch (e) {

      console.warn(
        'Root folder sekolah tidak dapat digunakan: ' +
        e.message
      );

    }

  }


  /*
   * ==========================================================
   * 3. FALLBACK DARI SPREADSHEET SEKOLAH AKTIF
   *
   * INI YANG MENJADI SOLUSI UTAMA.
   * ==========================================================
   */

  var spreadsheetId =
    ctx.spreadsheetId ||
    ctx.spreadsheet_id ||
    ctx.SPREADSHEET_ID ||

    (
      ctx.school &&
      (
        ctx.school.spreadsheetId ||
        ctx.school.spreadsheet_id
      )
    ) ||

    (
      ctx.resource &&
      (
        ctx.resource.spreadsheetId ||
        ctx.resource.spreadsheet_id
      )
    ) ||

    (
      ctx.resources &&
      (
        ctx.resources.spreadsheetId ||
        ctx.resources.spreadsheet_id
      )
    ) ||

    '';


  if (!spreadsheetId) {

    /*
     * Coba ambil dari fungsi spreadsheet aktif
     * modul agenda.
     */

    try {

      var ss =
        agendaCetakGetActiveSchoolSpreadsheet_();

      spreadsheetId =
        ss.getId();

    } catch (e) {

      console.warn(
        'Spreadsheet aktif tidak berhasil diperoleh: ' +
        e.message
      );

    }

  }


  /*
   * Jika Spreadsheet ID tersedia,
   * cari parent folder Spreadsheet.
   */

  if (spreadsheetId) {

    try {

      var spreadsheetFile =
        DriveApp.getFileById(
          String(spreadsheetId)
        );


      var parents =
        spreadsheetFile.getParents();


      if (parents.hasNext()) {

        var parentFolder =
          parents.next();


        /*
         * Cari folder AGENDA.
         */

        return agendaCetakEnsureSubfolder_(
          parentFolder,
          'AGENDA'
        );

      }

    } catch (e) {

      console.warn(
        'Parent Spreadsheet sekolah tidak dapat digunakan: ' +
        e.message
      );

    }

  }


  /*
   * ==========================================================
   * 4. COBA FUNGSI ROOT DRIVE DARI SISTEM UTAMA
   * ==========================================================
   */

  var rootCandidates = [

    'getSchoolDriveRootFolder_',

    'getSchoolRootFolder_',

    'getMySchoolRootFolder_',

    'getSchoolDriveFolder_',

    'getMySchoolDriveFolder_'

  ];


  for (
    var i = 0;
    i < rootCandidates.length;
    i++
  ) {

    try {

      var fn =
        rootCandidates[i];


      if (
        typeof this[fn] ===
        'function'
      ) {

        var folder =
          this[fn]();


        if (folder) {

          return agendaCetakEnsureSubfolder_(
            folder,
            'AGENDA'
          );

        }

      }

    } catch (e) {

      console.warn(
        fn +
        ' gagal: ' +
        e.message
      );

    }

  }


  /*
   * ==========================================================
   * 5. ERROR TERAKHIR
   * ==========================================================
   */

  throw new Error(
    'Folder AGENDA sekolah aktif tidak dapat ditemukan. ' +
    'School Context sudah terbaca tetapi tidak menyediakan ' +
    'folderId dan Spreadsheet sekolah aktif tidak memiliki parent folder.'
  );

}


/* ============================================================
   SUBFOLDER AGENDA
   ============================================================ */

function agendaCetakEnsureSubfolder_(
  parent,
  name
) {

  var folders =
    parent.getFoldersByName(name);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parent.createFolder(name);
}


/* ============================================================
   TEMPLATE PDF
   A4 PORTRAIT
   ============================================================ */

function agendaCetakPdfTemplate_() {

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +

    '<meta charset="UTF-8">' +

    '<style>' +

    '@page {' +
      'size:A4 portrait;' +
      'margin:9mm 8mm 10mm 8mm;' +
    '}' +

    'body {' +
      'font-family:Arial,sans-serif;' +
      'font-size:8px;' +
      'color:#222;' +
      'margin:0;' +
    '}' +

    '.school {' +
      'font-size:14px;' +
      'font-weight:bold;' +
      'text-align:center;' +
      'text-transform:uppercase;' +
      'margin-bottom:2px;' +
    '}' +

    '.npsn {' +
      'font-size:8px;' +
      'text-align:center;' +
      'margin-bottom:5px;' +
    '}' +

    '.title {' +
      'font-size:12px;' +
      'font-weight:bold;' +
      'text-align:center;' +
      'margin-bottom:3px;' +
    '}' +

    '.periode {' +
      'font-size:8px;' +
      'text-align:center;' +
      'margin-bottom:8px;' +
    '}' +

    '.info {' +
      'width:100%;' +
      'border-collapse:collapse;' +
      'margin-bottom:7px;' +
    '}' +

    '.info td {' +
      'padding:2px 3px;' +
      'vertical-align:top;' +
    '}' +

    '.label {' +
      'font-weight:bold;' +
      'width:80px;' +
    '}' +

    'table.data {' +
      'width:100%;' +
      'border-collapse:collapse;' +
      'table-layout:fixed;' +
      'font-size:7px;' +
    '}' +

    'table.data th {' +
      'background:#3C3332;' +
      'color:#fff;' +
      'border:0.5px solid #3C3332;' +
      'padding:4px 3px;' +
      'text-align:center;' +
      'font-weight:bold;' +
    '}' +

    'table.data td {' +
      'border:0.5px solid #B9B0AF;' +
      'padding:3px;' +
      'vertical-align:top;' +
      'word-wrap:break-word;' +
    '}' +

    '.center {' +
      'text-align:center;' +
      'vertical-align:middle !important;' +
    '}' +

    '.signature {' +
      'width:100%;' +
      'border-collapse:collapse;' +
      'margin-top:16px;' +
      'font-size:8px;' +
    '}' +

    '.signature td {' +
      'width:50%;' +
      'text-align:center;' +
      'vertical-align:top;' +
    '}' +

    '.space {' +
      'height:42px;' +
    '}' +

    '.footer {' +
      'margin-top:8px;' +
      'font-size:6.5px;' +
      'text-align:center;' +
      'color:#666;' +
    '}' +

    '</style>' +

    '</head>' +

    '<body>' +

    '<div class="school">' +
      '{{NAMA_SEKOLAH}}' +
    '</div>' +

    '<div class="npsn">' +
      'NPSN: {{NPSN}}' +
    '</div>' +

    '<div class="title">' +
      'LAPORAN AGENDA MENGAJAR GURU' +
    '</div>' +

    '<div class="periode">' +
      'Periode {{PERIODE_AWAL}} s.d. {{PERIODE_AKHIR}}' +
    '</div>' +

    '<table class="info">' +
      '<tr>' +
        '<td class="label">Nama Guru</td>' +
        '<td>: {{NAMA_GURU}}</td>' +
      '</tr>' +
      '<tr>' +
        '<td class="label">NIP</td>' +
        '<td>: {{NIP_GURU}}</td>' +
      '</tr>' +
    '</table>' +

    '<table class="data">' +

      '<thead>' +
        '<tr>' +
          '<th style="width:4%;">No</th>' +
          '<th style="width:8%;">Tanggal</th>' +
          '<th style="width:6%;">Sesi</th>' +
          '<th style="width:7%;">Kelas</th>' +
          '<th style="width:21%;">Tujuan Pembelajaran</th>' +
          '<th style="width:22%;">DPL / Pengalaman / Prinsip</th>' +
          '<th style="width:14%;">Murid Tidak Ikut</th>' +
          '<th style="width:18%;">Materi Pembelajaran</th>' +
        '</tr>' +
      '</thead>' +

      '<tbody>' +
        '{{TABEL_ROWS}}' +
      '</tbody>' +

    '</table>' +

    '<table class="signature">' +

      '<tr>' +

        '<td>' +
          'Mengetahui,<br>' +
          'Kepala Sekolah' +
          '<div class="space"></div>' +
          '<b>{{NAMA_KEPSEK}}</b><br>' +
          'NIP. {{NIP_KEPSEK}}' +
        '</td>' +

        '<td>' +
          'Guru' +
          '<div class="space"></div>' +
          '<b>{{NAMA_GURU}}</b><br>' +
          'NIP. {{NIP_GURU}}' +
        '</td>' +

      '</tr>' +

    '</table>' +

    '<div class="footer">' +
      'Dicetak melalui SIM SATRIA — {{TANGGAL_CETAK}}' +
    '</div>' +

    '</body>' +
    '</html>';
}


/* ============================================================
   TEST MODUL
   ============================================================ */

function testAgendaCetakMultiSchool() {

  var ctx =
    agendaCetakGetSchoolContext_();

  var ss =
    agendaCetakGetActiveSchoolSpreadsheet_();

  var sh =
    ss.getSheetByName('TRX_AGENDA_GURU');

  return {

    success: true,

    npsn:
      ctx.npsn ||
      (ctx.school && ctx.school.npsn) ||
      '',

    sekolah:
      ctx.namaSekolah ||
      (ctx.school && (
        ctx.school.namaSekolah ||
        ctx.school.nama
      )) ||
      '',

    spreadsheetId:
      ss.getId(),

    spreadsheetName:
      ss.getName(),

    trxAgendaGuru:
      !!sh,

    jumlahBaris:
      sh
        ? Math.max(0, sh.getLastRow() - 1)
        : 0

  };

}

function agendaGetSchoolContext_() {

  var result = null;


  /* =====================================================
     1. COBA SCHOOL CONTEXT UTAMA
     ===================================================== */

  try {

    if (
      typeof getSchoolContext === 'function'
    ) {

      result =
        getSchoolContext();

    }

  } catch (e) {

    console.warn(
      'getSchoolContext gagal: ' +
      e.message
    );

  }


  if (result) {
    return agendaNormalizeSchoolContext_(
      result
    );
  }


  /* =====================================================
     2. COBA CURRENT SCHOOL CONTEXT
     ===================================================== */

  try {

    if (
      typeof getCurrentSchoolContext === 'function'
    ) {

      result =
        getCurrentSchoolContext();

    }

  } catch (e) {

    console.warn(
      'getCurrentSchoolContext gagal: ' +
      e.message
    );

  }


  if (result) {
    return agendaNormalizeSchoolContext_(
      result
    );
  }


  /* =====================================================
     3. COBA SCHOOL RESOURCES
     ===================================================== */

  try {

    if (
      typeof getSchoolResources === 'function'
    ) {

      result =
        getSchoolResources();

    }

  } catch (e) {

    console.warn(
      'getSchoolResources gagal: ' +
      e.message
    );

  }


  if (result) {
    return agendaNormalizeSchoolContext_(
      result
    );
  }


  /* =====================================================
     4. COBA LOGIN INFO
     ===================================================== */

  try {

    if (
      typeof getLoginInfo === 'function'
    ) {

      result =
        getLoginInfo();

    }

  } catch (e) {

    console.warn(
      'getLoginInfo gagal: ' +
      e.message
    );

  }


  if (result) {

    return agendaNormalizeSchoolContext_(
      result
    );

  }


  throw new Error(
    'School Context SIM SATRIA tidak dapat diperoleh.'
  );

}

function agendaNormalizeSchoolContext_(
  info
) {

  info =
    info || {};


  var school =
    info.school ||
    info.sekolahData ||
    info.sekolah ||
    {};


  var resource =
    info.resource ||
    info.resources ||
    info.schoolResources ||
    {};


  /*
   * Jika sekolah berupa string,
   * tetap simpan nama sekolah.
   */

  var namaSekolah =
    typeof school === 'string'
      ? school
      : (
          school.namaSekolah ||
          school.nama ||
          info.namaSekolah ||
          info.sekolah ||
          ''
        );


  /*
   * NPSN
   */

  var npsn =
    (
      typeof school === 'object'
        ? (
            school.npsn ||
            school.NPSN ||
            ''
          )
        : ''
    ) ||
    info.npsn ||
    '';


  /*
   * Spreadsheet ID
   */

  var spreadsheetId =
    (
      typeof school === 'object'
        ? (
            school.spreadsheetId ||
            school.spreadsheet_id ||
            school.databaseSpreadsheetId ||
            school.database_spreadsheet_id ||
            ''
          )
        : ''
    ) ||

    resource.spreadsheetId ||
    resource.spreadsheet_id ||
    resource.databaseSpreadsheetId ||
    resource.database_spreadsheet_id ||

    info.spreadsheetId ||
    info.spreadsheet_id ||
    info.databaseSpreadsheetId ||
    info.database_spreadsheet_id ||

    '';


  /*
   * Root folder
   */

  var rootFolderId =
    (
      typeof school === 'object'
        ? (
            school.rootFolderId ||
            school.root_folder_id ||
            school.driveFolderId ||
            school.drive_folder_id ||
            ''
          )
        : ''
    ) ||

    resource.rootFolderId ||
    resource.root_folder_id ||
    resource.driveFolderId ||
    resource.drive_folder_id ||

    info.rootFolderId ||
    info.root_folder_id ||
    info.driveFolderId ||
    info.drive_folder_id ||

    '';


  return {

    success:
      info.success !== false,

    school: {

      namaSekolah:
        namaSekolah,

      npsn:
        npsn,

      spreadsheetId:
        spreadsheetId,

      rootFolderId:
        rootFolderId,

      driveFolderId:
        rootFolderId

    },

    resource:
      resource,

    resources:
      resource,

    spreadsheetId:
      spreadsheetId,

    rootFolderId:
      rootFolderId,

    driveFolderId:
      rootFolderId

  };

}



/* =========================================================
   FOLDER PDF AGENDA
   MULTI SCHOOL
   ========================================================= */

function agendaGetPdfFolder_() {

  /*
   * =======================================================
   * 1. AMBIL SCHOOL CONTEXT YANG SUDAH ADA
   * =======================================================
   *
   * Gunakan fungsi:
   *
   *     getAgendaSchoolContext_()
   *
   * BUKAN:
   *
   *     agendaGetSchoolContext_()
   */

  var ctx =
    getAgendaSchoolContext_();


  if (!ctx) {

    throw new Error(
      'School Context sekolah aktif tidak tersedia.'
    );

  }


  /* =======================================================
     2. PRIORITAS PERTAMA:
        FOLDER ID DARI SCHOOL CONTEXT
     ======================================================= */

  var folderId =
    String(
      ctx.folderId ||
      ''
    ).trim();


  if (folderId) {

    try {

      var rootFolder =
        DriveApp.getFolderById(
          folderId
        );


      /*
       * Cari folder AGENDA
       */

      var folders =
        rootFolder.getFoldersByName(
          AGENDA_GURU_CONFIG.FOLDER_AGENDA
        );


      if (folders.hasNext()) {

        return folders.next();

      }


      /*
       * Jika belum ada,
       * buat folder AGENDA.
       */

      return rootFolder.createFolder(
        AGENDA_GURU_CONFIG.FOLDER_AGENDA
      );

    } catch (err) {

      console.warn(
        'Folder Drive dari School Context tidak dapat digunakan: ' +
        err.message
      );

    }

  }


  /* =======================================================
     3. PRIORITAS KEDUA:
        SPREADSHEET SEKOLAH AKTIF
     ======================================================= */

  var spreadsheetId =
    String(
      ctx.spreadsheetId ||
      ''
    ).trim();


  if (!spreadsheetId) {

    throw new Error(
      'Spreadsheet sekolah aktif belum tersedia pada School Context.'
    );

  }


  var spreadsheetFile;

  try {

    spreadsheetFile =
      DriveApp.getFileById(
        spreadsheetId
      );

  } catch (err) {

    throw new Error(
      'Spreadsheet sekolah aktif tidak dapat diakses melalui Drive. ' +
      err.message
    );

  }


  /* =======================================================
     4. CARI PARENT FOLDER SPREADSHEET
     ======================================================= */

  var parents =
    spreadsheetFile.getParents();


  if (parents.hasNext()) {

    var parentFolder =
      parents.next();


    /*
     * Cari AGENDA di folder yang sama
     * dengan Spreadsheet sekolah.
     */

    var agendaFolders =
      parentFolder.getFoldersByName(
        AGENDA_GURU_CONFIG.FOLDER_AGENDA
      );


    if (agendaFolders.hasNext()) {

      return agendaFolders.next();

    }


    /*
     * Jika belum ada → buat.
     */

    return parentFolder.createFolder(
      AGENDA_GURU_CONFIG.FOLDER_AGENDA
    );

  }


  /* =======================================================
     5. SPREADSHEET TIDAK MEMPUNYAI PARENT
     ======================================================= */

  throw new Error(
    'Folder Drive sekolah aktif belum tersedia. ' +
    'Spreadsheet sekolah aktif tidak berada di dalam folder Drive.'
  );

}
function testAgendaPdfFolderMultiSchool() {

  var ctx =
    getAgendaSchoolContext_();


  var folder =
    agendaGetPdfFolder_();


  return {

    success: true,

    sekolah:
      ctx.namaSekolah,

    npsn:
      ctx.npsn,

    spreadsheetId:
      ctx.spreadsheetId,

    spreadsheetName:
      ctx.spreadsheetName,

    schoolFolderId:
      ctx.folderId || '',

    pdfFolderId:
      folder.getId(),

    pdfFolderName:
      folder.getName(),

    pdfFolderUrl:
      folder.getUrl()

  };

}