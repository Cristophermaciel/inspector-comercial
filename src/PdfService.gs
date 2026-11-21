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

  /**
   * Genera un Reporte PDF Diario Individual de Asistencia para una Feria y Fecha específica.
   * Se guarda en: INSPECTOR_COMERCIAL_DRIVE / DOCUMENTOS_PDF / ASISTENCIA FERIAS LIBRES / FERIA <FERIA>/
   */
  generateFeriaDailyPdf(feriaName, fechaStr, reportData) {
    const cleanFeria = String(feriaName || 'GENERAL').toUpperCase().replace(/^FERIA\s+/, '').trim();
    const safeFeriaFile = cleanFeria.replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');
    const safeFecha = String(fechaStr || Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyy-MM-dd')).trim();
    const fileName = 'ASISTENCIA_FERIA_' + safeFeriaFile + '_' + safeFecha + '.pdf';
    const pdfFolder = DriveManager.getFeriaPdfFolder(cleanFeria);

    // Si ya existía un PDF con el mismo nombre para esa feria y fecha, mover el anterior a papelera para evitar duplicados
    try {
      const existingFiles = pdfFolder.getFilesByName(fileName);
      while (existingFiles.hasNext()) {
        existingFiles.next().setTrashed(true);
      }
    } catch (e) {}

    const nowStr = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    const summary = reportData.summary || {};
    const rows = reportData.rows || [];
    const inspectores = reportData.inspectores || 'No especificado';
    const sheetName = reportData.sheetName || ('CONTROL_FERIA_' + safeFeriaFile);
    const tipoCierre = reportData.tipoCierre || 'CIERRE DIARIO (17:00 HRS)';

    // Formatear fecha dd/MM/yyyy para visualización
    const fParts = safeFecha.split('-');
    const fechaDisplay = fParts.length === 3 ? (fParts[2] + '/' + fParts[1] + '/' + fParts[0]) : safeFecha;

    let tableRowsHtml = '';
    rows.forEach(item => {
      const est = String(item.estado || 'PENDIENTE').toUpperCase();
      let badgeBg = '#f2f4f7';
      let badgeColor = '#475467';
      if (est === 'ASISTE') {
        badgeBg = '#e6f6ec';
        badgeColor = '#12805c';
      } else if (est === 'NO ASISTE') {
        badgeBg = '#fdecec';
        badgeColor = '#c92a2a';
      }

      const medio = String(item.medio || '').toUpperCase();
      const medioBg = medio === 'QR' ? '#e6f2fc' : (medio ? '#fff4e5' : '#f8fafc');
      const medioColor = medio === 'QR' ? '#0b7bde' : (medio ? '#b54708' : '#667085');

      tableRowsHtml += '<tr>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;text-align:center;font-weight:bold;color:#0c347d;">#' + this._esc(item.puesto) + '</td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;"><strong>' + this._esc(item.nombre || '-') + '</strong></td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;white-space:nowrap;">' + this._esc(item.rut || '-') + '</td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;">' + this._esc(item.rol || '-') + '</td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;">' + this._esc(item.patente || '-') + '</td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;">' + this._esc(item.giro || '-') + '</td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;text-align:center;"><span style="display:inline-block;padding:2px 7px;border-radius:10px;font-weight:bold;font-size:10px;background:' + badgeBg + ';color:' + badgeColor + ';">' + this._esc(est) + '</span></td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;text-align:center;"><span style="display:inline-block;padding:2px 6px;border-radius:8px;font-weight:bold;font-size:10px;background:' + medioBg + ';color:' + medioColor + ';">' + this._esc(medio || '-') + '</span></td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;text-align:center;white-space:nowrap;">' + this._esc(item.hora || '-') + '</td>' +
        '<td style="border:1px solid #d7e2f0;padding:5px 6px;">' + this._esc(item.inspector || '-') + '</td>' +
        '</tr>';
    });

    if (!tableRowsHtml) {
      tableRowsHtml = '<tr><td colspan="10" style="border:1px solid #d7e2f0;padding:14px;text-align:center;color:#667085;">Sin registros de asistencia para esta jornada.</td></tr>';
    }

    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
      '@page { size: letter landscape; margin: 14mm; }' +
      'body { font-family: Arial, Helvetica, sans-serif; font-size: 10.5px; color: #1d2939; margin: 0; }' +
      '.header { border-bottom: 3px solid #0c347d; padding-bottom: 8px; margin-bottom: 12px; }' +
      '.title { font-size: 16px; font-weight: bold; color: #0c347d; }' +
      '.subtitle { font-size: 12px; color: #475467; margin-top: 2px; }' +
      '.meta { background: #f4f7fb; border: 1px solid #d7e2f0; border-radius: 6px; padding: 8px 10px; margin-bottom: 12px; }' +
      '.meta table { width: 100%; border-collapse: collapse; }' +
      '.meta td { padding: 2px 6px; font-size: 10.5px; }' +
      '.kpis { width: 100%; border-collapse: separate; border-spacing: 6px 0; margin: 0 -6px 12px -6px; }' +
      '.kpi { border: 1px solid #d7e2f0; border-top: 3px solid #0c347d; border-radius: 6px; padding: 7px 8px; text-align: center; background: #fff; }' +
      '.kpi-lbl { font-size: 9px; color: #667085; text-transform: uppercase; font-weight: bold; }' +
      '.kpi-val { font-size: 16px; font-weight: bold; margin-top: 2px; }' +
      'table.data { width: 100%; border-collapse: collapse; font-size: 9.8px; }' +
      'table.data th { background: #0c347d; color: #ffffff; padding: 6px 5px; text-align: left; border: 1px solid #0c347d; font-size: 9.5px; }' +
      '.footer { margin-top: 16px; font-size: 9.5px; color: #667085; text-align: center; border-top: 1px solid #e0e0e0; padding-top: 6px; }' +
      '</style></head><body>' +
      '<div class="header">' +
        '<div class="title">' + this._esc(APP_CONFIG.APP_NAME) + '</div>' +
        '<div class="subtitle">' + this._esc(APP_CONFIG.SUBTITLE) + ' — <strong>REPORTE DIARIO INDIVIDUAL DE ASISTENCIA</strong></div>' +
      '</div>' +
      '<div class="meta"><table>' +
        '<tr>' +
          '<td><strong>Feria Libre:</strong> FERIA ' + this._esc(cleanFeria) + '</td>' +
          '<td><strong>Fecha Jornada:</strong> ' + this._esc(fechaDisplay) + '</td>' +
          '<td><strong>Tope de Toma:</strong> 17:00 hrs</td>' +
        '</tr>' +
        '<tr>' +
          '<td><strong>Tipo de Cierre:</strong> ' + this._esc(tipoCierre) + '</td>' +
          '<td><strong>Hoja Excel:</strong> ' + this._esc(sheetName) + '</td>' +
          '<td><strong>Generado:</strong> ' + this._esc(nowStr) + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td colspan="3"><strong>Inspector(es) a Cargo:</strong> ' + this._esc(inspectores) + '</td>' +
        '</tr>' +
      '</table></div>' +
      '<table class="kpis"><tr>' +
        '<td class="kpi"><div class="kpi-lbl">Puestos Ocupados</div><div class="kpi-val" style="color:#0c347d;">' + this._esc(summary.ocupados || 0) + '</div></td>' +
        '<td class="kpi" style="border-top-color:#12805c;"><div class="kpi-lbl">Asisten</div><div class="kpi-val" style="color:#12805c;">' + this._esc(summary.asisten || 0) + '</div></td>' +
        '<td class="kpi" style="border-top-color:#c92a2a;"><div class="kpi-lbl">No Asisten</div><div class="kpi-val" style="color:#c92a2a;">' + this._esc(summary.noAsisten || 0) + '</div></td>' +
        '<td class="kpi" style="border-top-color:#d97706;"><div class="kpi-lbl">Pendientes</div><div class="kpi-val" style="color:#d97706;">' + this._esc(summary.pendientes || 0) + '</div></td>' +
        '<td class="kpi" style="border-top-color:#0b7bde;"><div class="kpi-lbl">Ingreso por QR</div><div class="kpi-val" style="color:#0b7bde;">' + this._esc(summary.porQr || 0) + '</div></td>' +
        '<td class="kpi" style="border-top-color:#475467;"><div class="kpi-lbl">Ingreso Manual</div><div class="kpi-val" style="color:#475467;">' + this._esc(summary.manual || 0) + '</div></td>' +
      '</tr></table>' +
      '<table class="data"><thead><tr>' +
        '<th style="text-align:center;width:54px;">PUESTO</th>' +
        '<th>TITULAR / FERIANTE</th>' +
        '<th style="width:82px;">RUT</th>' +
        '<th style="width:65px;">ROL</th>' +
        '<th style="width:60px;">PATENTE</th>' +
        '<th>GIRO</th>' +
        '<th style="text-align:center;width:78px;">ESTADO</th>' +
        '<th style="text-align:center;width:62px;">MEDIO</th>' +
        '<th style="text-align:center;width:64px;">HORA</th>' +
        '<th>INSPECTOR</th>' +
      '</tr></thead><tbody>' + tableRowsHtml + '</tbody></table>' +
      '<div class="footer">' +
        'Documento oficial generado al cierre diario de asistencia (tope 17:00 hrs) — ' +
        this._esc(APP_CONFIG.CREATOR) + ' | ' + this._esc(APP_CONFIG.CONTACT) +
      '</div></body></html>';

    const blob = Utilities.newBlob(html, MimeType.HTML, fileName + '.html')
      .getAs(MimeType.PDF)
      .setName(fileName);
    const pdfFile = pdfFolder.createFile(blob);

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
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
      folderUrl: pdfFolder.getUrl(),
      folderName: 'FERIA ' + cleanFeria,
      feria: cleanFeria,
      fecha: safeFecha,
      generatedAt: nowStr
    };
  },

  _esc(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
};
