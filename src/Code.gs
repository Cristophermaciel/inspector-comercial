/**
 * ============================================================================
 * CODE.GS - PUNTO DE ENTRADA WEB APP + API REST JSON PARA PWA GITHUB PAGES
 * Permite usar la app tanto desde el enlace directo de Google Apps Script
 * como desde la PWA instalable en móvil y escritorio (GitHub Pages).
 * ============================================================================
 */

function doGet(e) {
  // Si recibe ?action=... responde como API JSON para la PWA (móvil/escritorio)
  if (e && e.parameter && e.parameter.action) {
    return _handleApiRequest(e.parameter.action, e.parameter);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.appName = APP_CONFIG.APP_NAME;
  template.subtitle = APP_CONFIG.SUBTITLE;

  return template.evaluate()
    .setTitle(APP_CONFIG.APP_NAME + ' - ' + APP_CONFIG.SUBTITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Endpoint POST JSON para guardar registros, subir fotos/firmas, generar PDFs
 * y ejecutar respaldos Excel desde la PWA instalable en GitHub Pages.
 */
function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    return _handleApiRequest(payload.action, payload);
  } catch (err) {
    return _jsonResponse({ ok: false, error: err.message });
  }
}

function _handleApiRequest(action, params) {
  try {
    let result;
    switch (action) {
      case 'bootstrap':
        result = DataService.getBootstrapData();
        break;
      case 'rows':
        result = DataService.getTableRows(params.table);
        break;
      case 'saveRecord':
        result = DataService.saveRecord(params.table, params.record || {}, params.mediaUploads || {});
        break;
      case 'deleteRecord':
        result = DataService.deleteRecord(params.table, params.rowNumber);
        break;
      case 'regeneratePdf':
        result = DataService.regeneratePdf(params.table, params.rowNumber);
        break;
      case 'dashboard':
        result = DataService.getDashboardStats();
        break;
      case 'map':
        result = DataService.getMapPoints(params.table);
        break;
      case 'runExcelBackup':
        result = BackupExcel.runFullExcelBackup();
        break;
      case 'listExcelBackups':
        result = DriveManager.listExcelBackups(25);
        break;
      case 'toggleDailyBackup':
        result = params.enable ? BackupExcel.setupDailyBackupTrigger() : BackupExcel.removeDailyBackupTrigger();
        break;
      case 'feriaAttendance':
        result = DataService.getFeriaAttendance(params.feria, params.fecha);
        break;
      case 'saveFeriaAttendance':
        result = DataService.saveFeriaAttendance(params.payload || {});
        break;
      default:
        result = { ok: false, error: 'Acción no reconocida: ' + action };
    }
    return _jsonResponse(result);
  } catch (err) {
    return _jsonResponse({ ok: false, error: err.message });
  }
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// FUNCIONES EXPUESTAS AL FRONTEND (google.script.run)
// ============================================================================

function apiGetBootstrap() { return DataService.getBootstrapData(); }
function apiGetTableRows(tableName) { return DataService.getTableRows(tableName); }
function apiSaveRecord(tableName, recordData, mediaUploads) { return DataService.saveRecord(tableName, recordData, mediaUploads); }
function apiDeleteRecord(tableName, rowNumber) { return DataService.deleteRecord(tableName, rowNumber); }
function apiRegeneratePdf(tableName, rowNumber) { return DataService.regeneratePdf(tableName, rowNumber); }
function apiGetDashboardStats() { return DataService.getDashboardStats(); }
function apiGetMapPoints(tableName) { return DataService.getMapPoints(tableName); }
function apiRunExcelBackup() { return BackupExcel.runFullExcelBackup(); }
function apiListExcelBackups() { return DriveManager.listExcelBackups(25); }
function apiToggleDailyBackup(enable) { return enable ? BackupExcel.setupDailyBackupTrigger() : BackupExcel.removeDailyBackupTrigger(); }
function apiEnsureDriveFolders() { return DriveManager.ensureBaseStructure(); }
function apiGetFeriaAttendance(feria, fecha) { return DataService.getFeriaAttendance(feria, fecha); }
function apiSaveFeriaAttendance(payload) { return DataService.saveFeriaAttendance(payload); }
