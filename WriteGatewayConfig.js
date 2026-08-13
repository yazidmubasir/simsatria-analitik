/**
 * SIM SATRIA - WRITE GATEWAY CONFIG
 *
 * Web App utama tetap Execute as: User accessing the web app.
 * URL berikut adalah deployment Web App WRITE GATEWAY yang dijalankan
 * menggunakan akun berotoritas Editor terhadap database/Drive sekolah.
 */
const WRITE_GATEWAY_CONFIG = Object.freeze({
  URL: "https://script.google.com/macros/s/AKfycbxihmOttlgS5kAcMgcY9wz3_O-j0LeohajZ54tnU7zS3bFm3Tel5KsS936_MPitw--s/exec",
  ENABLED: true,
  TIMEOUT_MS: 30000,
});

function getWriteGatewayUrl_() {
  const url = String(WRITE_GATEWAY_CONFIG.URL || "").trim();
  if (!url) throw new Error("URL Write Gateway belum dikonfigurasi.");
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(url)) {
    throw new Error("URL Write Gateway tidak valid. Gunakan URL deployment /exec.");
  }
  return url;
}

function getWriteGatewayConfig() {
  return {
    enabled: WRITE_GATEWAY_CONFIG.ENABLED === true,
    urlConfigured: !!String(WRITE_GATEWAY_CONFIG.URL || "").trim(),
    timeoutMs: WRITE_GATEWAY_CONFIG.TIMEOUT_MS,
  };
}
