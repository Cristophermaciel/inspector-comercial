/**
 * ============================================================================
 * PDF SERVICE - REPLICA LOS 5 BOTS DE AUTOMATIZACIÓN DE APPSHEET
 * 1. COMERCIO          -> REPORTE DE FISCALIZACION COMERCIAL DIARIA
 * 2. FUERA DE HORA     -> REPORTES FUERA DE HORARIO ESTABLECIDO
 * 3. GIGANTOGRAFIA     -> INFORME DE PUBLICIDAD
 * 4. PATENTES EN MORA  -> FISCALIZACIONES MOROSAS
 * 5. REPORTE           -> REPORTES A JEFATURAS
 * ============================================================================
 */

const PdfService = {

  /**
   * Genera el PDF para un registro usando la plantilla original de Google Docs
   * (reemplazando etiquetas AppSheet <<[COLUMNA]>>). Si la plantilla no es accesible
   * por permisos, genera automáticamente un informe PDF municipal estructurado en HTML.
   */
  generatePdfForRecord(tableName, record) {
    const bot = APP_CONFIG.PDF_BOTS[tableName];
    if (!bot) {
      throw new Error('La tabla ' + tableName + ' no tiene un bot de PDF configurado.');
    }

    const pdfFolder = DriveManager.getPdfFolder(bot.folderName);
    const prefixVal = record[bot.prefixColumn] || record['ID'] || 'REG';
    const stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
    const safePrefix = String(prefixVal).replace(/[\\/:*?"<>|]/g, '-').substring(0, 50);
    const fileName = safePrefix + '_' + stamp + '.pdf';

    let pdfFile = null;

    // Intento 1: usar la plantilla original de Google Docs de AppSheet
    if (bot.templateDocId) {
      try {
        pdfFile = this._generateFromGoogleDocTemplate(bot.templateDocId, record, fileName, pdfFolder);
      } catch (e) {
        Logger.log('No se pudo usar plantilla Google Docs (' + bot.templateDocId + '), usando generador HTML: ' + e.message);
      }
    }

    // Intento 2 (Respaldo automático): generar PDF municipal limpio desde HTML
    if (!pdfFile) {
      pdfFile = this._generateFromHtmlFallback(tableName, record, fileName, pdfFolder);
    }

    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      try {
        pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e2) {}
    }

    return {
      id: pdfFile.getId(),
      name: pdfFile.getName(),
      url: pdfFile.getUrl(),
      folderUrl: pdfFolder.getUrl(),
      folderName: bot.folderName
    };
  },

  /**
   * Copia el Google Doc plantilla, reemplaza etiquetas <<[COLUMNA]>> y convierte a PDF.
   */
  _generateFromGoogleDocTemplate(templateDocId, record, fileName, pdfFolder) {
    const templateFile = DriveApp.getFileById(templateDocId);
    const tempCopy = templateFile.makeCopy('TEMP_' + fileName, pdfFolder);
    try {
      const doc = DocumentApp.openById(tempCopy.getId());
      const body = doc.getBody();
      const header = doc.getHeader();
      const footer = doc.getFooter();

      Object.keys(record).forEach(col => {
        const val = record[col] !== undefined && record[col] !== null ? String(record[col]) : '';
        const escapedCol = col.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const pattern = '<<\\[' + escapedCol + '\\]>>';
        body.replaceText(pattern, val);
        if (header) header.replaceText(pattern, val);
        if (footer) footer.replaceText(pattern, val);
      });

      // Limpiar cualquier etiqueta no encontrada
      body.replaceText('<<\\[[^\\]]+\\]>>', '');
      doc.saveAndClose();

      const pdfBlob = tempCopy.getAs(MimeType.PDF).setName(fileName);
      return pdfFolder.createFile(pdfBlob);
    } finally {
      tempCopy.setTrashed(true);
    }
  },

  /**
   * Genera un informe PDF formal en HTML (respaldo si la plantilla Doc no está compartida).
   */
  _generateFromHtmlFallback(tableName, record, fileName, pdfFolder) {
    const schema = TABLE_SCHEMAS[tableName] || { label: tableName, columns: [] };
    const nowStr = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');

    let rowsHtml = '';
    const imageFields = [];

    schema.columns.forEach(col => {
      if (col.name === 'PDF') return;
      const val = record[col.name];
      if (val === undefined || val === null || val === '') return;

      if (col.type === 'Image' || col.type === 'Signature') {
        imageFields.push({ label: col.displayName || col.name, url: String(val) });
      } else {
        rowsHtml += '<tr>' +
          '<td style="width:36%;font-weight:bold;background:#f4f7fb;border:1px solid #d7e2f0;padding:7px 10px;">' +
          this._esc(col.displayName || col.name) + '</td>' +
          '<td style="border:1px solid #d7e2f0;padding:7px 10px;">' + this._esc(val) + '</td>' +
          '</tr>';
      }
    });

    let mediaHtml = '';
    if (imageFields.length > 0) {
      mediaHtml = '<h3 style="color:#0b7bde;margin-top:18px;border-bottom:1px solid #0b7bde;padding-bottom:4px;">Registro Fotográfico y Firma</h3><ul>';
      imageFields.forEach(m => {
        mediaHtml += '<li style="margin:6px 0;"><strong>' + this._esc(m.label) + ':</strong> ' +
          '<a href="' + this._esc(m.url) + '">' + this._esc(m.url) + '</a></li>';
      });
      mediaHtml += '</ul>';
    }

    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
      '<body style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222;margin:24px;">' +
      '<div style="border-bottom:3px solid #0b7bde;padding-bottom:10px;margin-bottom:16px;">' +
      '<div style="font-size:18px;font-weight:bold;color:#0b7bde;">' + this._esc(APP_CONFIG.APP_NAME) + '</div>' +
      '<div style="font-size:13px;color:#555;">' + this._esc(APP_CONFIG.SUBTITLE) + ' — ' + this._esc(schema.label) + '</div>' +
      '<div style="font-size:11px;color:#777;margin-top:4px;">Generado: ' + nowStr + ' | ID: ' + this._esc(record.ID || '') + '</div>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;">' + rowsHtml + '</table>' +
      mediaHtml +
      '<div style="margin-top:26px;font-size:10px;color:#888;text-align:center;border-top:1px solid #e0e0e0;padding-top:8px;">' +
      this._esc(APP_CONFIG.CREATOR) + ' | ' + this._esc(APP_CONFIG.CONTACT) +
      '</div></body></html>';

    const blob = Utilities.newBlob(html, MimeType.HTML, fileName + '.html')
      .getAs(MimeType.PDF)
      .setName(fileName);
    return pdfFolder.createFile(blob);
  },

  _esc(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
};
