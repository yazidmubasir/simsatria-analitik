function getWriteGatewayStatus() {
  const context = getCurrentUserContext();
  const url = String(PropertiesService.getScriptProperties().getProperty('SIM_SATRIA_WRITE_GATEWAY_URL') || '').trim();
  const configured = !!url && !!PropertiesService.getScriptProperties().getProperty('SIM_SATRIA_WRITE_GATEWAY_SECRET');
  return {
    success: true,
    configured: configured,
    role: context.role,
    npsn: context.npsn,
    sekolah: context.school.namaSekolah,
    mode: ['ADMIN_SEKOLAH', 'SUPERADMIN', 'SUPER_ADMIN'].indexOf(normalizeRole_(context.role)) >= 0 ? 'DIRECT_ADMIN' : 'WRITE_GATEWAY',
  };
}
