/**
 * ============================================================================
 * CODE.GS - PUNTO DE ENTRADA WEB APP Y API DEL SERVIDOR
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appName = APP_CONFIG.APP_NAME;
  template.subtitle = APP_CONFIG.SUBTITLE;

  return template.evaluate()
    .setTitle(APP_CONFIG.APP_NAME + ' - ' + APP_CONFIG.SUBTITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// FUNCIONES EXPUESTAS AL FRONTEND (google.script.run)
// ============================================================================

function apiGetBootstrap() {
  return DataService.getBootstrapData();
}

function apiGetTableRows(tableName) {
  return DataService.getTableRows(tableName);
}

function apiSaveRecord(tableName, recordData, mediaUploads) {
  return DataService.saveRecord(tableName, recordData, mediaUploads);
}

function apiDeleteRecord(tableName, rowNumber) {
  return DataService.deleteRecord(tableName, rowNumber);
}

function apiRegeneratePdf(tableName, rowNumber) {
  return DataService.regeneratePdf(tableName, rowNumber);
}

function apiGetDashboardStats() {
  return DataService.getDashboardStats();
}

function apiGetMapPoints(tableName) {
  return DataService.getMapPoints(tableName);
}

function apiRunExcelBackup() {
  return BackupExcel.runFullExcelBackup();
}

function apiListExcelBackups() {
  return DriveManager.listExcelBackups(25);
}

function apiToggleDailyBackup(enable) {
  return enable ? BackupExcel.setupDailyBackupTrigger() : BackupExcel.removeDailyBackupTrigger();
}

function apiEnsureDriveFolders() {
  return DriveManager.ensureBaseStructure();
}
