# SIM SATRIA — Viewer + Write Gateway

## Target architecture

- ADMIN_SEKOLAH: Editor pada Spreadsheet dan Drive sekolah.
- GURU/WALI_KELAS/KARYAWAN/SISWA: Viewer secara fisik pada Spreadsheet/Drive.
- User biasa tetap dapat input/simpan/cetak/upload melalui aplikasi.
- Mutasi database dan Drive dilakukan oleh Write Gateway yang dideploy sebagai `Execute as: Me` pada akun pengelola sekolah.

## Deployment Write Gateway

Buat project Apps Script terpisah dari web app utama dan salin `WriteGateway.js`.

Set Script Properties gateway:

- `SIM_SATRIA_MASTER_SPREADSHEET_ID` = ID MASTER.
- `SIM_SATRIA_WRITE_GATEWAY_SECRET` = secret acak panjang.

Deploy gateway sebagai Web App:

- Execute as: **Me** (akun pengelola yang memiliki Editor pada database/Drive sekolah).
- Who has access: sesuai kebijakan domain; endpoint harus dapat dipanggil oleh web app utama.

Pada project web utama set Script Properties:

- `SIM_SATRIA_WRITE_GATEWAY_URL` = URL Web App gateway.
- `SIM_SATRIA_WRITE_GATEWAY_SECRET` = secret yang sama.

## Penting

`WriteGateway.js` sengaja berada sebagai project/deployment terpisah. Jangan mengganti `appsscript.json` web utama dari `USER_ACCESSING` menjadi `ME`, karena identitas login GURU/SISWA digunakan oleh Auth.js untuk school binding.

## Migrasi modul

Setiap operasi mutasi yang saat ini langsung memakai `SpreadsheetApp.openById(...).appendRow`, `setValues`, `setValue`, `insertRow`, `deleteRow`, atau operasi Drive `createFile` untuk user biasa harus dialihkan ke `gatewayAppendRow_`, `gatewaySetValues_`, `gatewayUploadFile_`, atau `gatewayCreatePdf_`.

Operasi administrasi yang memang hanya dilakukan ADMIN_SEKOLAH tetap boleh menggunakan Spreadsheet/Drive secara langsung.
