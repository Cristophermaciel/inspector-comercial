/**
 * ============================================================================
 * DRIVE MANAGER - CARPETAS INDEPENDIENTES EN GOOGLE DRIVE
 * Organiza automáticamente fotos, firmas, documentos PDF y respaldos Excel
 * en carpetas independientes por módulo y por registro.
 * ============================================================================
 */

const DriveManager = {

  /**
   * Obtiene o crea una subcarpeta dentro de un padre (o en la raíz de Drive).
   */
  getOrCreateFolder(name, parentFolder) {
    const safeName = String(name || 'SIN_NOMBRE')
      .replace(/[\\/:*?"<>|]/g, '_')
      .trim()
      .substring(0, 100);
    const parent = parentFolder || DriveApp.getRootFolder();
    const it = parent.getFoldersByName(safeName);
    if (it.hasNext()) {
      return it.next();
    }
    return parent.createFolder(safeName);
  },

  /**
   * Retorna la carpeta raíz independiente de la aplicación.
   */
  getAppRootFolder() {
    return this.getOrCreateFolder(APP_CONFIG.DRIVE_ROOT_FOLDER_NAME, null);
  },

  /**
   * Garantiza que existan las 4 carpetas base independientes:
   * FOTOS, FIRMAS, DOCUMENTOS_PDF y RESPALDOS_EXCEL.
   */
  ensureBaseStructure() {
    const root = this.getAppRootFolder();
    const fotos = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.FOTOS, root);
    const firmas = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.FIRMAS, root);
    const pdfs = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.DOCUMENTOS_PDF, root);
    const excel = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.RESPALDOS_EXCEL, root);
    const cierres = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.CIERRES_ANUALES || 'CIERRES_ANUALES', root);

    // Crear carpetas independientes de PDF por cada Bot
    Object.keys(APP_CONFIG.PDF_BOTS).forEach(tableName => {
      this.getOrCreateFolder(APP_CONFIG.PDF_BOTS[tableName].folderName, pdfs);
    });
    // Carpeta de reportes PDF diarios individuales de Asistencia de Ferias Libres
    this.getOrCreateFolder(APP_CONFIG.FERIA_PDF_FOLDER_NAME || 'ASISTENCIA FERIAS LIBRES', pdfs);
    // Carpeta de Balances Anuales PDF
    this.getOrCreateFolder(APP_CONFIG.ANNUAL_PDF_FOLDER_NAME || 'BALANCES ANUALES', pdfs);

    return {
      root: { id: root.getId(), name: root.getName(), url: root.getUrl() },
      fotos: { id: fotos.getId(), name: fotos.getName(), url: fotos.getUrl() },
      firmas: { id: firmas.getId(), name: firmas.getName(), url: firmas.getUrl() },
      pdfs: { id: pdfs.getId(), name: pdfs.getName(), url: pdfs.getUrl() },
      excel: { id: excel.getId(), name: excel.getName(), url: excel.getUrl() },
      cierres: { id: cierres.getId(), name: cierres.getName(), url: cierres.getUrl() }
    };
  },

  /**
   * Obtiene/crea la carpeta específica para guardar una foto o firma de un registro:
   * INSPECTOR_COMERCIAL_DRIVE / FOTOS (o FIRMAS) / <TABLA> / <ID_REGISTRO>
   */
  getRecordMediaFolder(categoryKey, tableName, recordKey, recordLabel) {
    const root = this.getAppRootFolder();
    const catFolder = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS[categoryKey], root);
    const tableFolder = this.getOrCreateFolder(tableName, catFolder);
    const labelPart = recordLabel ? '_' + String(recordLabel).replace(/[^a-zA-Z0-9áéíóúñÁÉÍÓÚÑ\s-]/g, '').trim().substring(0, 35) : '';
    const recordFolderName = String(recordKey || 'REG') + labelPart;
    return this.getOrCreateFolder(recordFolderName, tableFolder);
  },

  /**
   * Obtiene/crea la carpeta de PDFs de un módulo:
   * INSPECTOR_COMERCIAL_DRIVE / DOCUMENTOS_PDF / <NOMBRE_CARPETA_BOT>
   */
  getPdfFolder(folderName) {
    const root = this.getAppRootFolder();
    const pdfsRoot = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.DOCUMENTOS_PDF, root);
    return this.getOrCreateFolder(folderName || 'GENERAL', pdfsRoot);
  },

  /**
   * Obtiene/crea la carpeta individual de reportes PDF diarios de una feria:
   * INSPECTOR_COMERCIAL_DRIVE / DOCUMENTOS_PDF / ASISTENCIA FERIAS LIBRES / FERIA <NOMBRE_FERIA>
   */
  getFeriaPdfFolder(feriaName) {
    const baseFolder = this.getPdfFolder(APP_CONFIG.FERIA_PDF_FOLDER_NAME || 'ASISTENCIA FERIAS LIBRES');
    const clean = String(feriaName || 'GENERAL').toUpperCase().replace(/^FERIA\s+/, '').trim();
    return this.getOrCreateFolder('FERIA ' + clean, baseFolder);
  },

  /**
   * Obtiene/crea la carpeta de respaldos Excel:
   * INSPECTOR_COMERCIAL_DRIVE / RESPALDOS_EXCEL
   */
  getExcelBackupFolder() {
    const root = this.getAppRootFolder();
    return this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.RESPALDOS_EXCEL, root);
  },

  /**
   * Obtiene/crea la carpeta de archivo de un Cierre de Año:
   * INSPECTOR_COMERCIAL_DRIVE / CIERRES_ANUALES / AÑO_<YYYY>
   */
  getAnnualCloseFolder(year) {
    const root = this.getAppRootFolder();
    const base = this.getOrCreateFolder(APP_CONFIG.DRIVE_SUBFOLDERS.CIERRES_ANUALES || 'CIERRES_ANUALES', root);
    return this.getOrCreateFolder('AÑO_' + String(year || new Date().getFullYear()), base);
  },

  /**
   * Obtiene/crea la carpeta de Balances Anuales PDF:
   * INSPECTOR_COMERCIAL_DRIVE / DOCUMENTOS_PDF / BALANCES ANUALES
   */
  getAnnualPdfFolder() {
    return this.getPdfFolder(APP_CONFIG.ANNUAL_PDF_FOLDER_NAME || 'BALANCES ANUALES');
  },

  /**
   * Guarda un archivo codificado en Data URL (base64) dentro de Drive
   * y retorna su URL pública de visualización directa y el ID.
   */
  saveBase64File(dataUrl, fileName, folder) {
    if (!dataUrl || typeof dataUrl !== 'string' || !dataUrl.startsWith('data:')) {
      return null;
    }
    const parts = dataUrl.split(',');
    const meta = parts[0];
    const base64Data = parts[1];
    const mimeMatch = meta.match(/data:([^;]+);base64/);
    const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const ext = mimeType.indexOf('png') !== -1 ? '.png'
      : mimeType.indexOf('pdf') !== -1 ? '.pdf'
      : '.jpg';

    const cleanFileName = String(fileName || 'archivo')
      .replace(/[\\/:*?"<>|]/g, '_')
      .trim() + ext;

    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, cleanFileName);
    const file = folder.createFile(blob);

    // Permitir visualización a quien tenga el enlace para previsualizar en la Web App
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // Si el dominio Workspace restringe compartir externo, mantener permiso de dominio/usuario
      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e2) {
        Logger.log('Aviso permisos Drive: ' + e2.message);
      }
    }

    return {
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      thumbnailUrl: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800'
    };
  },

  /**
   * Lista los archivos más recientes dentro de una carpeta de respaldo Excel.
   */
  listExcelBackups(limit) {
    const folder = this.getExcelBackupFolder();
    const files = folder.getFiles();
    const list = [];
    while (files.hasNext()) {
      const f = files.next();
      list.push({
        id: f.getId(),
        name: f.getName(),
        url: f.getUrl(),
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + f.getId(),
        date: Utilities.formatDate(f.getDateCreated(), APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'),
        timestamp: f.getDateCreated().getTime(),
        sizeKb: Math.round(f.getSize() / 1024)
      });
    }
    list.sort((a, b) => b.timestamp - a.timestamp);
    return list.slice(0, limit || 20);
  }
};
