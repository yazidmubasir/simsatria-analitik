/**
 * PRESENSI SECURITY SERVICE
 * Server-side security boundary for Presensi Per Kelas.
 */
function securePresensiLoad() {
  requirePermission("VIEW_PRESENSI");
  const context = getCurrentUserContext();
  const ss = getSchoolSpreadsheet_();
  const sheet = ss.getSheetByName("KELAS");
  if (!sheet) throw new Error("Sheet KELAS tidak ditemukan pada Spreadsheet sekolah.");
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return {success:true,npsn:context.npsn,sekolah:context.school.namaSekolah,kelas:[]};
  const headers = values[0].map(normalizeHeader_);
  let kelasIndex=-1; ["KELAS","NAMA_KELAS","ROMBEL"].some(function(h){kelasIndex=headers.indexOf(h);return kelasIndex>-1;});
  if (kelasIndex<0) throw new Error("Kolom KELAS tidak ditemukan pada sheet KELAS.");
  const statusIndex=headers.indexOf("STATUS"),seen={},kelas=[];
  for(let i=1;i<values.length;i++){const nama=String(values[i][kelasIndex]||"").trim();if(!nama)continue;if(statusIndex>-1){const s=String(values[i][statusIndex]||"").trim().toUpperCase();if(s&&s!=="ACTIVE"&&s!=="AKTIF")continue;}const key=nama.toUpperCase();if(!seen[key]){seen[key]=true;kelas.push(nama);}}
  kelas.sort(function(a,b){return String(a).localeCompare(String(b),"id",{numeric:true,sensitivity:"base"});});
  return {success:true,email:context.email,userId:context.userId,nama:context.nama,role:context.role,nip:context.nip,npsn:context.npsn,sekolah:context.school.namaSekolah,kelas:kelas};
}
function securePresensiGetData(tanggal,kelas){
  requirePermission("VIEW_PRESENSI");
  tanggal=String(tanggal||"").trim();kelas=String(kelas||"").trim();if(!tanggal)throw new Error("Tanggal belum dipilih.");if(!kelas)throw new Error("Kelas belum dipilih.");
  const context=getCurrentUserContext(),ss=getSchoolSpreadsheet_(),sheet=ss.getSheetByName("SISWA");if(!sheet)throw new Error("Sheet SISWA tidak ditemukan pada Spreadsheet sekolah.");
  const values=sheet.getDataRange().getDisplayValues();if(values.length<2)return {success:true,tanggal:tanggal,kelas:kelas,npsn:context.npsn,sekolah:context.school.namaSekolah,data:[]};
  const headers=values[0].map(normalizeHeader_),nisnIndex=headers.indexOf("NISN"),namaIndex=headers.indexOf("NAMA")>-1?headers.indexOf("NAMA"):headers.indexOf("NAMA_SISWA"),kelasIndex=headers.indexOf("KELAS")>-1?headers.indexOf("KELAS"):headers.indexOf("ROMBEL"),statusIndex=headers.indexOf("STATUS");
  if(namaIndex<0||kelasIndex<0)throw new Error("Kolom NAMA/KELAS pada sheet SISWA tidak ditemukan.");
  const data=[];for(let i=1;i<values.length;i++){const rowKelas=String(values[i][kelasIndex]||"").trim();if(rowKelas.toUpperCase()!==kelas.toUpperCase())continue;if(statusIndex>-1){const s=String(values[i][statusIndex]||"").trim().toUpperCase();if(s&&s!=="ACTIVE"&&s!=="AKTIF")continue;}data.push({nisn:nisnIndex>-1?String(values[i][nisnIndex]||"").trim():"",nama:String(values[i][namaIndex]||"").trim(),kelas:rowKelas,status:"HADIR"});}
  data.sort(function(a,b){return String(a.nama).localeCompare(String(b.nama),"id",{sensitivity:"base"});});
  return {success:true,tanggal:tanggal,kelas:kelas,npsn:context.npsn,sekolah:context.school.namaSekolah,data:data,jumlah:data.length};
}
function securePresensiSave(tanggal,kelas,data){
  requirePermission("INPUT_PRESENSI");
  tanggal=String(tanggal||"").trim();kelas=String(kelas||"").trim();if(!tanggal)throw new Error("Tanggal belum dipilih.");if(!kelas)throw new Error("Kelas belum dipilih.");if(!Array.isArray(data)||!data.length)throw new Error("Data presensi kosong.");
  const context=getCurrentUserContext(),ss=getSchoolSpreadsheet_(),siswaSheet=ss.getSheetByName("SISWA"),trx=ss.getSheetByName("TRX_PRESENSI");if(!siswaSheet)throw new Error("Sheet SISWA tidak ditemukan pada Spreadsheet sekolah.");if(!trx)throw new Error("Sheet TRX_PRESENSI tidak ditemukan pada Spreadsheet sekolah.");
  const sv=siswaSheet.getDataRange().getDisplayValues();if(sv.length<2)throw new Error("Data SISWA sekolah kosong.");const sh=sv[0].map(normalizeHeader_),sNisn=sh.indexOf("NISN"),sNama=sh.indexOf("NAMA")>-1?sh.indexOf("NAMA"):sh.indexOf("NAMA_SISWA"),sKelas=sh.indexOf("KELAS")>-1?sh.indexOf("KELAS"):sh.indexOf("ROMBEL"),sStatus=sh.indexOf("STATUS");if(sNisn<0||sNama<0||sKelas<0)throw new Error("Struktur SISWA tidak lengkap.");
  const allowed={};for(let i=1;i<sv.length;i++){const row=sv[i],rk=String(row[sKelas]||"").trim();if(rk.toUpperCase()!==kelas.toUpperCase())continue;if(sStatus>-1){const s=String(row[sStatus]||"").trim().toUpperCase();if(s&&s!=="ACTIVE"&&s!=="AKTIF")continue;}const nisn=String(row[sNisn]||"").trim();if(nisn)allowed[nisn]={nama:String(row[sNama]||"").trim(),kelas:rk};}
  const headers=trx.getRange(1,1,1,trx.getLastColumn()).getDisplayValues()[0].map(normalizeHeader_),required=["TRANSACTION_ID","TIMESTAMP","NPSN","USER_ID","EMAIL","NIP","NAMA_USER","ROLE","TANGGAL","KELAS","NISN","NAMA_SISWA","STATUS","KETERANGAN"],missing=required.filter(function(h){return headers.indexOf(h)<0;});if(missing.length)throw new Error("Header TRX_PRESENSI belum lengkap: "+missing.join(", "));const col={};headers.forEach(function(h,i){col[h]=i;});
  const lock=LockService.getScriptLock();lock.waitLock(30000);try{const now=new Date(),tx=createSecurePresensiTransactionId_(context.npsn),existingValues=trx.getLastRow()>=2?trx.getRange(2,1,trx.getLastRow()-1,headers.length).getDisplayValues():[],existingMap={};existingValues.forEach(function(row,i){const nisn=String(row[col.NISN]||"").trim(),d=String(row[col.TANGGAL]||"").trim(),k=String(row[col.KELAS]||"").trim();if(nisn&&d&&k)existingMap[[d,k,nisn].join("|")]=i+2;});let inserted=0,updated=0;
    data.forEach(function(item){const nisn=String(item&&item.nisn||"").trim();if(!nisn)throw new Error("NISN siswa tidak boleh kosong.");const student=allowed[nisn];if(!student)throw new Error("Siswa dengan NISN "+nisn+" tidak terdaftar pada kelas "+kelas+".");const status=String(item.status||"HADIR").trim().toUpperCase();if(["HADIR","IZIN","SAKIT","ALPA"].indexOf(status)<0)throw new Error("Status presensi tidak valid: "+status);const row=new Array(headers.length).fill("");row[col.TRANSACTION_ID]=tx;row[col.TIMESTAMP]=now;row[col.NPSN]=context.npsn;row[col.USER_ID]=context.userId;row[col.EMAIL]=context.email;row[col.NIP]=context.nip;row[col.NAMA_USER]=context.nama;row[col.ROLE]=context.role;row[col.TANGGAL]=tanggal;row[col.KELAS]=student.kelas;row[col.NISN]=nisn;row[col.NAMA_SISWA]=student.nama;row[col.STATUS]=status;row[col.KETERANGAN]=String(item.keterangan||"").trim();const key=[tanggal,kelas,nisn].join("|");const target=existingMap[key];if(target){trx.getRange(target,1,1,headers.length).setValues([row]);updated++;}else{const nr=trx.getLastRow()+1;trx.getRange(nr,1,1,headers.length).setValues([row]);existingMap[key]=nr;inserted++;}});
    return {success:true,npsn:context.npsn,sekolah:context.school.namaSekolah,tanggal:tanggal,kelas:kelas,inserted:inserted,updated:updated,transactionId:tx};
  }finally{lock.releaseLock();}
}
function createSecurePresensiTransactionId_(npsn){const tz=Session.getScriptTimeZone()||"Asia/Jakarta";return "PRS-"+String(npsn||"UNKNOWN")+"-"+Utilities.formatDate(new Date(),tz,"yyyyMMddHHmmss")+"-"+Utilities.getUuid().replace(/-/g,"").substring(0,8).toUpperCase();}
