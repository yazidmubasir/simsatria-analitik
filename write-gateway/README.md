# SIM SATRIA Write Gateway

Gateway ini adalah Apps Script Web App TERPISAH dari Web App SIM SATRIA utama.

## Tujuan

- Web App utama tetap `Execute as: User accessing the web app`.
- User sekolah (GURU, WALI_KELAS, KARYAWAN, SISWA) tetap dapat diberi `Viewer` pada Spreadsheet/Drive secara fisik.
- Operasi tulis dilakukan oleh Gateway yang dideploy `Execute as: Me` oleh akun yang mempunyai hak Editor pada database/Drive sekolah.

## File

- `Code.gs` - doGet/doPost endpoint.
- `GatewayConfig.gs` - allowlist sheet/action dan Script Properties.
- `GatewayService.gs` - autentikasi caller, otorisasi role, write Spreadsheet, upload Drive, PDF.
- `appsscript.json` - manifest Gateway.

## Script Properties wajib

Pada project Gateway, buka **Project Settings > Script Properties** dan isi:

- `GATEWAY_NPSN` = NPSN sekolah, contoh `20306175`.
- `GATEWAY_SPREADSHEET_ID` = ID Spreadsheet sekolah yang harus ditulis Gateway.
- `GATEWAY_DRIVE_FOLDER_ID` = ID folder Drive sekolah (wajib untuk upload/PDF).

Jangan menyimpan password atau OAuth token di repository.

## Deployment

1. Buka project Apps Script Gateway.
2. Masukkan tiga file `.gs` dan `appsscript.json` dari folder ini.
3. Pastikan akun yang melakukan deployment mempunyai **Editor** terhadap Spreadsheet dan folder Drive sekolah.
4. Deploy > New deployment > Web app.
5. Execute as: **Me**.
6. Who has access: **Anyone** (pengguna tetap harus dapat teridentifikasi; gateway menolak request jika `Session.getActiveUser().getEmail()` kosong).
7. Gunakan URL `/exec`.

## Endpoint

`GET /exec` -> health check.

`POST /exec` menerima JSON:

```json
{
  "action": "SAVE_ROW",
  "npsn": "20306175",
  "userId": "USER_ID_DARI_USERS",
  "sheet": "TRX_PRESENSI",
  "values": {
    "TANGGAL": "2026-08-13",
    "KELAS": "X-1",
    "NISN": "...",
    "NAMA_SISWA": "...",
    "STATUS": "HADIR",
    "KETERANGAN": ""
  }
}
```

Gateway tidak mempercayai `email`, `role`, atau `userId` yang dikirim client. Email caller diambil dari `Session.getActiveUser()`, lalu dicocokkan dengan `USERS` pada Spreadsheet sekolah yang dikonfigurasi.

## Catatan keamanan

- Gateway hanya dapat menulis sheet yang ada pada allowlist.
- Gateway tidak menerima spreadsheetId dari client.
- NPSN gateway ditetapkan di Script Properties.
- User harus `ACTIVE` dan role harus terdaftar.
- `TRANSACTION_ID`, `TIMESTAMP`, NPSN, userId, email, NIP, nama user, dan role ditulis oleh gateway.
- Web App utama TIDAK perlu diubah menjadi `Execute as: Me`.

## Status integrasi

Gateway ini harus diuji dan dideploy terlebih dahulu. Setelah endpoint health dan `SAVE_ROW` berhasil, baru fungsi-fungsi simpan/upload pada Web App utama dialihkan ke `WriteGatewayClient`.
