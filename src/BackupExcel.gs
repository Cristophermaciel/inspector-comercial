/**
 * ============================================================================
 * BACKUP EXCEL SERVICE - RESPALDO EN FORMATO .XLSX EN GOOGLE DRIVE
 * Exporta ambas hojas de cálculo (INSPECTOR COMERCIAL y NOTIFICACIONES DEUDAS)
 * a archivos Excel (.xlsx) con fecha y hora dentro de:
 * INSPECTOR_COMERCIAL_DRIVE / RESPALDOS_EXCEL
 * ============================================================================
 */

const BackupExcel = {

  /**
   * Ejecuta un respaldo completo en Excel (.xlsx) de ambas planillas.
   */
  runFullExcelBackup() {
    const folder = DriveManager.getExcelBackupFolder();
    const stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd_HHmm');
    const results = [];

    const targets = [
      { key: 'MAIN', id: APP_CONFIG.SPREADSHEETS.MAIN, prefix: 'INSPECTOR_COMERCIAL' },
      { key: 'DEUDAS', id: APP_CONFIG.SPREADSHEETS.DEUDAS, prefix: 'NOTIFICACIONES_DEUDAS' }
    ];

    targets.forEach(t => {
      try {
        const exportUrl = 'https://docs.google.com/spreadsheets/d/' + t.id + '/export?format=xlsx';
        const response = UrlFetchApp.fetch(exportUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
          throw new Error('HTTP ' + response.getResponseCode() + ' al exportar ' + t.prefix);
        }

        const fileName = t.prefix + '_' + stamp + '.xlsx';
        const blob = response.getBlob().setName(fileName);
        const file = folder.createFile(blob);

        results.push({
          ok: true,
          key: t.key,
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
          sizeKb: Math.round(file.getSize() / 1024)
        });
      } catch (e) {
        results.push({
          ok: false,
          key: t.key,
          error: e.message
        });
      }
    });

    return {
      timestamp: Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'),
      folderUrl: folder.getUrl(),
      files: results
    };
  },

  /**
   * Activa un disparador automático diario (23:00 hrs Chile) para generar respaldos Excel.
   */
  setupDailyBackupTrigger() {
    const fnName = 'scheduledDailyExcelBackup';
    const existing = ScriptApp.getProjectTriggers();
    existing.forEach(tr => {
      if (tr.getHandlerFunction() === fnName) {
        ScriptApp.deleteTrigger(tr);
      }
    });

    ScriptApp.newTrigger(fnName)
      .timeBased()
      .everyDays(1)
      .atHour(23)
      .create();

    return { ok: true, message: 'Respaldo automático en Excel programado todos los días a las 23:00 hrs.' };
  },

  /**
   * Elimina el disparador automático de respaldo Excel.
   */
  removeDailyBackupTrigger() {
    const fnName = 'scheduledDailyExcelBackup';
    const existing = ScriptApp.getProjectTriggers();
    existing.forEach(tr => {
      if (tr.getHandlerFunction() === fnName) {
        ScriptApp.deleteTrigger(tr);
      }
    });
    return { ok: true, message: 'Respaldo automático desactivado.' };
  },

  hasDailyBackupTrigger() {
    return ScriptApp.getProjectTriggers().some(tr => tr.getHandlerFunction() === 'scheduledDailyExcelBackup');
  }
};

/**
 * Función global invocada por el Trigger diario de Apps Script.
 */
function scheduledDailyExcelBackup() {
  BackupExcel.runFullExcelBackup();
}
