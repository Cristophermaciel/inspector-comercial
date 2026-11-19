/**
 * ============================================================================
 * DATA SERVICE - MOTOR CRUD PARA LAS 16 TABLAS, FÓRMULAS VIRTUALES Y DASHBOARD
 * ============================================================================
 */

const DataService = {

  _getSpreadsheet(spreadsheetKey) {
    const id = APP_CONFIG.SPREADSHEETS[spreadsheetKey || 'MAIN'];
    return SpreadsheetApp.openById(id);
  },

  _getSheet(tableName) {
    const schema = TABLE_SCHEMAS[tableName];
    if (!schema) throw new Error('Tabla no encontrada en configuración: ' + tableName);
    const ss = this._getSpreadsheet(schema.spreadsheetKey);
    const sheet = ss.getSheetByName(schema.sheetName);
    if (!sheet) throw new Error('Hoja no encontrada en Google Sheets: ' + schema.sheetName);
    return { sheet, schema };
  },

  /**
   * Verifica identidad del usuario actual.
   */
  getCurrentUser() {
    let email = '';
    try {
      email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
    } catch (e) {
      email = '';
    }
    const lower = email.toLowerCase();
    const isAuthorized = !APP_CONFIG.REQUIRE_AUTH_LIST ||
      !lower ||
      APP_CONFIG.AUTHORIZED_USERS.indexOf(lower) !== -1;

    return { email: email || 'Usuario Municipal', isAuthorized };
  },

  /**
   * Carga inicial de la Web App: esquemas, menú dinámico, funcionarios y carpetas Drive.
   */
  getBootstrapData() {
    const user = this.getCurrentUser();
    if (!user.isAuthorized) {
      return { authorized: false, user };
    }

    const driveFolders = DriveManager.ensureBaseStructure();
    const funcionarios = this._safeReadRows('FUNCIONARIO');
    const navTables = {};
    ['MENU', 'SUBMENU', 'SUBMENU 2', 'SUBMENU 3', 'LISTADOS', 'FERIAS LIBRES', 'DOCUMENTOS'].forEach(t => {
      navTables[t] = this._safeReadRows(t);
    });

    return {
      authorized: true,
      user,
      config: {
        appName: APP_CONFIG.APP_NAME,
        subtitle: APP_CONFIG.SUBTITLE,
        creator: APP_CONFIG.CREATOR,
        contact: APP_CONFIG.CONTACT,
        primaryColor: APP_CONFIG.PRIMARY_COLOR
      },
      schemas: TABLE_SCHEMAS,
      viewMap: APPSHEET_VIEW_MAP,
      funcionarios: funcionarios.map(f => ({
        id: f['ID'] || '',
        nombre: f['NOMBRE'] || '',
        rol: f['ROL'] || '',
        imagen: f['IMAGEN'] || ''
      })).filter(f => f.nombre),
      navTables,
      driveFolders,
      dailyBackupEnabled: BackupExcel.hasDailyBackupTrigger()
    };
  },

  _safeReadRows(tableName) {
    try {
      return this.getTableRows(tableName).rows;
    } catch (e) {
      Logger.log('Aviso leyendo ' + tableName + ': ' + e.message);
      return [];
    }
  },

  /**
   * Lee todas las filas de una tabla y calcula columnas virtuales.
   */
  getTableRows(tableName) {
    const { sheet, schema } = this._getSheet(tableName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return { tableName, headers: [], rows: [], summary: {} };
    }

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const rows = [];

    for (let r = 1; r < values.length; r++) {
      const rowArr = values[r];
      const hasAnyValue = rowArr.some(v => v !== '' && v !== null && v !== undefined);
      if (!hasAnyValue) continue;

      const obj = { _RowNumber: r + 1 };
      headers.forEach((h, idx) => {
        if (!h) return;
        let val = rowArr[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
        }
        obj[h] = val !== undefined && val !== null ? val : '';
      });

      // Columna virtual DIAS RESTANTES para Hoja 1 (Notificaciones por Deudas)
      if (tableName === 'Hoja 1' && obj['FECHA DE TERMINO']) {
        obj['DIAS RESTANTES'] = this._computeDaysRemaining(rowArr[headers.indexOf('FECHA DE TERMINO')]);
      }

      rows.push(obj);
    }

    // Orden por defecto (ej. LISTADO DE FERIANTES por N° Puesto)
    if (schema.defaultSort) {
      const col = schema.defaultSort.column;
      const dir = schema.defaultSort.asc ? 1 : -1;
      rows.sort((a, b) => {
        const na = Number(a[col]);
        const nb = Number(b[col]);
        if (!isNaN(na) && !isNaN(nb)) return (na - nb) * dir;
        return String(a[col] || '').localeCompare(String(b[col] || '')) * dir;
      });
    } else {
      // Más recientes primero
      rows.reverse();
    }

    // Columna virtual TOTAL RECAUDADO en PATENTES EN MORA
    const summary = {};
    if (tableName === 'PATENTES EN MORA') {
      let totalMora = 0;
      rows.forEach(row => {
        if (String(row['ESTADO PATENTE'] || '').toUpperCase() === 'EN MORA') {
          const num = this._parseMoney(row['VALOR PATENTE']);
          totalMora += num;
        }
      });
      summary.totalPatentesMorosas = totalMora;
    }

    return { tableName, headers, rows, summary };
  },

  /**
   * Crea o actualiza un registro, sube fotos/firmas a carpetas independientes en Drive
   * y dispara el Bot de PDF automático cuando es un registro nuevo.
   */
  saveRecord(tableName, recordData, mediaUploads) {
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      const { sheet, schema } = this._getSheet(tableName);
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

      const isEdit = Boolean(recordData._RowNumber && Number(recordData._RowNumber) >= 2);
      const keyCol = schema.keyCol;

      // Generar ID único si es registro nuevo
      if (!isEdit && keyCol && keyCol !== '_RowNumber' && !recordData[keyCol]) {
        recordData[keyCol] = Utilities.getUuid().substring(0, 8).toUpperCase();
      }

      // Inicializar FECHA HORA / FECHA si viene vacío
      schema.columns.forEach(col => {
        if (!isEdit && col.initial === 'NOW' && !recordData[col.name]) {
          recordData[col.name] = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
        }
        if (!isEdit && col.initial === 'TODAY' && !recordData[col.name]) {
          recordData[col.name] = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'dd/MM/yyyy');
        }
      });

      const recordKey = recordData[keyCol] || ('ROW_' + (sheet.getLastRow() + 1));
      const recordLabel = schema.labelCol ? recordData[schema.labelCol] : '';

      // 1. Subir fotos y firmas a carpetas independientes en Drive
      const uploadedMedia = {};
      if (mediaUploads && typeof mediaUploads === 'object') {
        Object.keys(mediaUploads).forEach(colName => {
          const dataUrl = mediaUploads[colName];
          if (!dataUrl || !String(dataUrl).startsWith('data:')) return;

          const colDef = schema.columns.find(c => c.name === colName);
          const isSig = colDef && colDef.type === 'Signature';
          const categoryKey = isSig ? 'FIRMAS' : 'FOTOS';
          const folder = DriveManager.getRecordMediaFolder(categoryKey, tableName, recordKey, recordLabel);
          const stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
          const saved = DriveManager.saveBase64File(dataUrl, colName + '_' + stamp, folder);
          if (saved) {
            recordData[colName] = saved.url;
            uploadedMedia[colName] = saved;
          }
        });
      }

      // 2. Generar PDF automático (réplica de los 5 Bots de AppSheet al agregar registro)
      let pdfResult = null;
      if (!isEdit && schema.hasPdfBot && APP_CONFIG.PDF_BOTS[tableName]) {
        try {
          pdfResult = PdfService.generatePdfForRecord(tableName, recordData);
          const pdfCol = APP_CONFIG.PDF_BOTS[tableName].pdfColumn;
          if (pdfCol) {
            recordData[pdfCol] = pdfResult.url;
          }
        } catch (e) {
          Logger.log('Error generando PDF automático: ' + e.message);
        }
      }

      // 3. Escribir fila en Google Sheets respetando el orden exacto de encabezados
      const rowValues = headers.map(h => {
        if (!h) return '';
        const val = recordData[h];
        if (Array.isArray(val)) return val.join(' , ');
        return val !== undefined && val !== null ? val : '';
      });

      let targetRowNum;
      if (isEdit) {
        targetRowNum = Number(recordData._RowNumber);
        sheet.getRange(targetRowNum, 1, 1, headers.length).setValues([rowValues]);
      } else {
        targetRowNum = sheet.getLastRow() + 1;
        sheet.getRange(targetRowNum, 1, 1, headers.length).setValues([rowValues]);
      }

      recordData._RowNumber = targetRowNum;

      return {
        ok: true,
        isEdit,
        record: recordData,
        uploadedMedia,
        pdf: pdfResult
      };
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Elimina una fila por su _RowNumber.
   */
  deleteRecord(tableName, rowNumber) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const num = Number(rowNumber);
      if (!num || num < 2) throw new Error('Número de fila inválido.');
      const { sheet } = this._getSheet(tableName);
      sheet.deleteRow(num);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Genera (o regenera) manualmente el PDF de un registro existente y actualiza la columna PDF.
   */
  regeneratePdf(tableName, rowNumber) {
    const { sheet } = this._getSheet(tableName);
    const data = this.getTableRows(tableName);
    const record = data.rows.find(r => Number(r._RowNumber) === Number(rowNumber));
    if (!record) throw new Error('Registro no encontrado.');

    const pdf = PdfService.generatePdfForRecord(tableName, record);
    const bot = APP_CONFIG.PDF_BOTS[tableName];
    if (bot && bot.pdfColumn) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
      const colIdx = headers.indexOf(bot.pdfColumn);
      if (colIdx !== -1) {
        sheet.getRange(Number(rowNumber), colIdx + 1).setValue(pdf.url);
      }
    }
    return { ok: true, pdf };
  },

  /**
   * Calcula todos los indicadores y los 6 gráficos exactos del DASHBOARD de AppSheet.
   */
  getDashboardStats() {
    const comercio = this._safeReadRows('COMERCIO');
    const fueraHora = this._safeReadRows('FUERA DE HORA');
    const publicidad = this._safeReadRows('GIGANTOGRAFIA');
    const mora = this._safeReadRows('PATENTES EN MORA');
    const reportes = this._safeReadRows('REPORTE');
    const feriantes = this._safeReadRows('ASISTENCIA');

    const countBy = (arr, col) => {
      const map = {};
      arr.forEach(r => {
        const k = String(r[col] || 'Sin especificar').trim() || 'Sin especificar';
        map[k] = (map[k] || 0) + 1;
      });
      return Object.keys(map)
        .map(k => ({ label: k, value: map[k] }))
        .sort((a, b) => b.value - a.value);
    };

    // Total valor patentes en mora y agrupación por estado
    let totalMora = 0;
    let totalAlDia = 0;
    const moraPorEstado = {};
    mora.forEach(r => {
      const estado = String(r['ESTADO PATENTE'] || 'Sin estado').trim();
      const val = this._parseMoney(r['VALOR PATENTE']);
      moraPorEstado[estado] = (moraPorEstado[estado] || 0) + val;
      if (estado.toUpperCase() === 'EN MORA') totalMora += val;
      if (estado.toUpperCase() === 'AL DIA') totalAlDia += val;
    });

    return {
      kpis: {
        totalComercio: comercio.length,
        totalFueraHora: fueraHora.length,
        totalPublicidad: publicidad.length,
        totalMoraCount: mora.length,
        totalMoraMonto: totalMora,
        totalAlDiaMonto: totalAlDia,
        totalReportes: reportes.length,
        totalFeriantes: feriantes.length
      },
      charts: {
        topPublicidad: countBy(publicidad, 'INSPECTOR'),
        topFiscalizador: countBy(comercio, 'INSPECTOR'),
        topFueraHora: countBy(fueraHora, 'INSPECTOR'),
        permisoAlDia: countBy(comercio, 'PERMISO AL DIA'),
        valorPatentesMora: Object.keys(moraPorEstado).map(k => ({ label: k, value: moraPorEstado[k] })),
        totalesPatentesMora: countBy(mora, 'ESTADO PATENTE')
      }
    };
  },

  /**
   * Extrae los puntos con coordenadas GPS válidas para el mapa FISCALIZACIONES COMERCIALES.
   */
  getMapPoints(tableName) {
    const target = tableName || 'COMERCIO';
    const rows = this._safeReadRows(target);
    const points = [];

    rows.forEach(r => {
      const loc = String(r['UBICACION'] || r['DIRECCION'] || '').trim();
      const match = loc.match(/(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/);
      if (match) {
        points.push({
          rowNumber: r._RowNumber,
          id: r['ID'] || r._RowNumber,
          lat: parseFloat(match[1]),
          lng: parseFloat(match[2]),
          title: r['NOMBRE LOCAL'] || r['NOMBRE LOCAL COMERCIAL'] || r['NOMBRE DUEÑO'] || r['FUNCIONARIO'] || 'Registro #' + r._RowNumber,
          subtitle: r['DIRECCION'] || '',
          inspector: r['INSPECTOR'] || r['FISCALIZADOR'] || r['FUNCIONARIO'] || '',
          fecha: r['FECHA HORA'] || r['FECHA'] || '',
          patente: r['CUENTA CON PATENTE'] || r['ESTADO PATENTE'] || ''
        });
      }
    });

    return points;
  },

  _parseMoney(val) {
    if (typeof val === 'number') return val;
    const cleaned = String(val || '0').replace(/[^0-9.-]/g, '');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  },

  _computeDaysRemaining(dateVal) {
    if (!dateVal) return '';
    const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return '';
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    d.setHours(0, 0, 0, 0);
    const diff = Math.round((d.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    return String(diff);
  }
};
