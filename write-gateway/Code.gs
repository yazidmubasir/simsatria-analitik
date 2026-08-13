/**
 * SIM SATRIA WRITE GATEWAY
 * Separate Apps Script Web App.
 * DEPLOYMENT: Execute as = Me (gateway owner/admin account).
 * The main SIM SATRIA Web App remains Execute as = User accessing the web app.
 */

function doGet() {
  return jsonResponse_({
    success: true,
    service: "SIM_SATRIA_WRITE_GATEWAY",
    version: GATEWAY_CONFIG.VERSION,
    status: "ONLINE",
  });
}

function doPost(e) {
  const requestId = Utilities.getUuid();
  try {
    const body = parseRequestBody_(e);
    const result = handleGatewayRequest_(body, requestId);
    return jsonResponse_(result);
  } catch (err) {
    console.error("WRITE_GATEWAY_ERROR", requestId, err && err.stack ? err.stack : err);
    return jsonResponse_({
      success: false,
      requestId: requestId,
      error: safeErrorMessage_(err),
    });
  }
}

function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Request body kosong.");
  }
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error("Request JSON tidak valid.");
  }
  if (!body || typeof body !== "object") throw new Error("Payload gateway tidak valid.");
  return body;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeErrorMessage_(err) {
  const message = err && err.message ? err.message : String(err || "Unknown error");
  return message.length > 500 ? message.slice(0, 500) : message;
}
