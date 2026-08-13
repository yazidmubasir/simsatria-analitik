# Viewer + Write Gateway

The main web app remains `USER_ACCESSING`. GURU, WALI_KELAS, KARYAWAN and SISWA are physical Viewers of the school Spreadsheet/Drive. They can still use application forms because mutating operations are routed to a separate Apps Script Write Gateway deployed `Execute as: Me` by the school administrator/service account.

The gateway validates NPSN, email, role, ACTIVE status and request timestamp before writing to the school spreadsheet or Drive. The main project never exposes the gateway secret to the UI.

Deployment is required before non-admin write operations can succeed: configure `SIM_SATRIA_WRITE_GATEWAY_URL` and `SIM_SATRIA_WRITE_GATEWAY_SECRET` in the main project; configure `SIM_SATRIA_MASTER_SPREADSHEET_ID` and the same secret in the gateway project.
