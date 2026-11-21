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
    let sheet = ss.getSheetByName(schema.sheetName);
    if (!sheet) {
      // Crear la hoja automáticamente con encabezados en orden exacto (ej. FERIAS NAVIDEÑAS)
      sheet = ss.insertSheet(schema.sheetName);
      const cols = schema.columns.map(c => c.name);
      if (cols.length > 0) {
        sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
        sheet.setFrozenRows(1);
      }
    }
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

  _normalizeFuncionarios(rawRows) {
    const out = [];
    const seen = {};
    const norm = s => String(s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/\s+/g, ' ').trim();

    const addIfValid = (item) => {
      const nombre = String(item.nombre || item.NOMBRE || '').trim();
      const rol = String(item.rol || item.ROL || '').trim();
      if (!nombre) return;
      if (/EQUIPO/i.test(rol) || /EQUIPO/i.test(nombre)) return;
      const key = norm(nombre);
      if (seen[key]) return;
      seen[key] = true;
      out.push({
        id: item.id || item.ID || '',
        nombre: nombre,
        rol: rol || 'INSPECTOR',
        imagen: item.imagen || item.IMAGEN || ''
      });
    };

    (rawRows || []).forEach(addIfValid);
    (APP_CONFIG.DEFAULT_FUNCIONARIOS || []).forEach(addIfValid);
    return out;
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
    const funcionarios = this._normalizeFuncionarios(this._safeReadRows('FUNCIONARIO'));
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
      funcionarios,
      navTables,
      driveFolders,
      dailyBackupEnabled: BackupExcel.hasDailyBackupTrigger(),
      dailyFeriaCloseEnabled: this.hasDailyFeriaCloseTrigger(),
      feriaCutoffHour: APP_CONFIG.FERIA_ATTENDANCE_CUTOFF_HOUR || 17
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
      if (tableName === 'Hoja 1') {
        const termIdx = headers.indexOf('FECHA DE TERMINO');
        const computed = termIdx !== -1 ? this._computeDaysRemaining(rowArr[termIdx]) : '';
        if (computed !== '') {
          obj['DIAS RESTANTES'] = computed;
        } else if (obj['DIAS RESTANTES'] !== '' && obj['DIAS RESTANTES'] !== undefined && !isNaN(Number(obj['DIAS RESTANTES']))) {
          obj['DIAS RESTANTES'] = String(Math.round(Number(obj['DIAS RESTANTES'])));
        }
      }

      rows.push(obj);
    }

    // Orden por defecto (ej. LISTADO DE FERIANTES por N° Puesto, Hoja 1 por DIAS RESTANTES)
    if (schema.defaultSort) {
      const col = schema.defaultSort.column;
      const dir = schema.defaultSort.asc ? 1 : -1;
      rows.sort((a, b) => {
        const va = a[col];
        const vb = b[col];
        const emptyA = va === '' || va === null || va === undefined;
        const emptyB = vb === '' || vb === null || vb === undefined;
        if (emptyA && emptyB) return 0;
        if (emptyA) return 1;
        if (emptyB) return -1;
        const na = Number(va);
        const nb = Number(vb);
        if (!isNaN(na) && !isNaN(nb)) {
          if (na !== nb) return (na - nb) * dir;
          if (tableName === 'Hoja 1') {
            const numA = Number(a['N° NOTIFICACION']) || 0;
            const numB = Number(b['N° NOTIFICACION']) || 0;
            return numA - numB;
          }
          return 0;
        }
        return String(va || '').localeCompare(String(vb || '')) * dir;
      });
    } else {
      // Más recientes primero
      rows.reverse();
    }

    // Columna virtual TOTAL RECAUDADO en PATENTES EN MORA y resumen de vencidos en Hoja 1
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
    } else if (tableName === 'Hoja 1') {
      let vencidos = 0, vencenHoy = 0, vigentes = 0;
      rows.forEach(row => {
        const raw = row['DIAS RESTANTES'];
        if (raw === '' || raw === null || raw === undefined) return;
        const n = Number(raw);
        if (isNaN(n)) return;
        if (n < 0) vencidos++;
        else if (n === 0) vencenHoy++;
        else vigentes++;
      });
      summary.vencidosNoPago = vencidos;
      summary.vencenHoy = vencenHoy;
      summary.vigentes = vigentes;
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

      // Garantizar CODIGO_QR único para cada feriante en su feria correspondiente
      if (tableName === 'ASISTENCIA') {
        if (!recordData['ID']) {
          recordData['ID'] = Utilities.getUuid().substring(0, 8).toUpperCase();
        }
        recordData['CODIGO_QR'] = this._buildUniqueFerianteQrCode(
          recordData['FERIA'],
          recordData['N° Puesto'],
          recordData['ID'],
          recordData._RowNumber || (sheet.getLastRow() + 1)
        );
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

      // 3. Garantizar encabezados en orden exacto si la hoja estuviera vacía,
      //    y agregar automáticamente columnas nuevas del esquema (ej. INSPECTOR ACOMPAÑANTE) al final sin desplazar columnas existentes.
      const hasHeaders = headers.some(h => h !== '');
      let finalHeaders = headers.slice();
      if (!hasHeaders) {
        finalHeaders = schema.columns.map(c => c.name);
        sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
      } else {
        const existingNorm = {};
        finalHeaders.forEach(h => {
          if (h) existingNorm[h.toUpperCase().trim()] = true;
        });
        const missingHeaders = [];
        schema.columns.forEach(c => {
          const normName = c.name.toUpperCase().trim();
          if (!existingNorm[normName] && recordData[c.name] !== undefined) {
            missingHeaders.push(c.name);
            existingNorm[normName] = true;
          }
        });
        if (missingHeaders.length > 0) {
          sheet.getRange(1, finalHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
          finalHeaders = finalHeaders.concat(missingHeaders);
        }
      }

      // Mapa rápido de definición de columnas (insensible a mayúsculas/espacios extra)
      const colDefByNorm = {};
      schema.columns.forEach(c => {
        colDefByNorm[c.name.toUpperCase().trim()] = c;
      });

      // Mapa de valores ingresados normalizados para garantizar orden exacto en Excel/Sheets
      const recordByNorm = {};
      Object.keys(recordData).forEach(k => {
        recordByNorm[k.toUpperCase().trim()] = recordData[k];
      });

      // Escribir cada celda en el orden exacto de la hoja respetando tipos numéricos/fórmulas para estadísticas
      const targetRowNum = isEdit ? Number(recordData._RowNumber) : (sheet.getLastRow() + 1);
      const rowValues = finalHeaders.map(h => {
        if (!h) return '';
        const norm = h.toUpperCase().trim();
        const colDef = colDefByNorm[norm];

        // Mantener fórmula de días restantes en Hoja 1 (Notificaciones por Deudas)
        if (tableName === 'Hoja 1' && norm === 'DIAS RESTANTES') {
          return '=INDIRECT("E"&ROW())-TODAY()';
        }

        let val = recordData[h] !== undefined ? recordData[h] : recordByNorm[norm];
        if (Array.isArray(val)) return val.join(' , ');
        if (val === undefined || val === null || val === '') return '';

        // Guardar precios y números como valores numéricos reales (clave para estadísticas y tablas dinámicas Excel)
        if (colDef && (colDef.type === 'Number' || colDef.type === 'Price')) {
          const parsed = this._parseMoney(val);
          return isNaN(parsed) ? val : parsed;
        }

        return val;
      });

      sheet.getRange(targetRowNum, 1, 1, finalHeaders.length).setValues([rowValues]);

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
    const feriasNavidenas = this._safeReadRows('FERIAS NAVIDEÑAS');
    const notificaciones = this._safeReadRows('Hoja 1');

    let totalVencidosNoPago = 0;
    let totalVencenHoy = 0;
    let totalNotifVigentes = 0;
    notificaciones.forEach(r => {
      const raw = r['DIAS RESTANTES'];
      if (raw === '' || raw === null || raw === undefined) return;
      const n = Number(raw);
      if (isNaN(n)) return;
      if (n < 0) totalVencidosNoPago++;
      else if (n === 0) totalVencenHoy++;
      else totalNotifVigentes++;
    });

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
        totalFeriantes: feriantes.length,
        totalFeriasNavidenas: feriasNavidenas.length,
        totalNotificaciones: notificaciones.length,
        totalVencidosNoPago: totalVencidosNoPago,
        totalNotifVigentes: totalNotifVigentes
      },
      charts: {
        topPublicidad: countBy(publicidad, 'INSPECTOR'),
        topFiscalizador: countBy(comercio, 'INSPECTOR'),
        topFueraHora: countBy(fueraHora, 'INSPECTOR'),
        permisoAlDia: countBy(comercio, 'PERMISO AL DIA'),
        valorPatentesMora: Object.keys(moraPorEstado).map(k => ({ label: k, value: moraPorEstado[k] })),
        totalesPatentesMora: countBy(mora, 'ESTADO PATENTE'),
        estadoNotificacionesDeudas: [
          { label: 'Vencidos por No Pago', value: totalVencidosNoPago },
          { label: 'Vencen Hoy', value: totalVencenHoy },
          { label: 'En Plazo / Vigentes', value: totalNotifVigentes }
        ]
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
          inspector: r['INSPECTOR'] || r['FISCALIZADOR'] || r['FUNCIONARIO'] || r['CONTROLADO POR'] || '',
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
    let d = null;
    if (dateVal instanceof Date) {
      d = new Date(dateVal.getTime());
    } else {
      const s = String(dateVal).trim();
      const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (dmy) {
        d = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
      } else if (ymd) {
        d = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
      } else {
        d = new Date(s);
      }
    }
    if (!d || isNaN(d.getTime())) return '';
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    d.setHours(0, 0, 0, 0);
    const diff = Math.round((d.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    return String(diff);
  },

  // ==========================================================================
  // CONTROL DE ASISTENCIA INDEPENDIENTE POR FERIA (PUESTOS 1 AL 100 + QR)
  // Crea automáticamente una hoja independiente por cada feria en Google Sheets:
  // CONTROL_FERIA_MODELO, CONTROL_FERIA_LAS_PRADERAS, etc.
  // Cierre diario automático a las 17:00 hrs con generación de PDF individual.
  // ==========================================================================
  _normalizeFeriaName(feriaName) {
    return String(feriaName || 'GENERAL')
      .toUpperCase()
      .replace(/^FERIA\s+/, '')
      .replace(/\s+/g, ' ')
      .trim();
  },

  _buildUniqueFerianteQrCode(feriaName, puestoRaw, idVal, rowNum) {
    const feria = this._normalizeFeriaName(feriaName || 'SIN_FERIA');
    const cleanFeria = feria
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '') || 'SIN_FERIA';
    const puesto = this._extractPuestoNumber(puestoRaw);
    const puestoPart = puesto ? ('P' + String(puesto).padStart(3, '0')) : 'SP';
    const cleanId = String(idVal || ('R' + (rowNum || '0')))
      .toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 6);
    return cleanFeria + '-' + puestoPart + '-' + (cleanId || '0000');
  },

  /**
   * Audita TODOS los registros de la hoja ASISTENCIA y garantiza que cada feriante
   * tenga un ID y un CODIGO_QR 100% único en su feria correspondiente.
   */
  ensureAllFeriantesQrCodes() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const { sheet } = this._getSheet('ASISTENCIA');
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return { ok: true, totalRows: 0, updatedCount: 0, byFeria: {} };
      }

      let lastCol = Math.max(sheet.getLastColumn(), 1);
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

      let idColIdx = headers.findIndex(h => h.toUpperCase() === 'ID');
      let feriaColIdx = headers.findIndex(h => h.toUpperCase() === 'FERIA');
      let puestoColIdx = headers.findIndex(h => h.toUpperCase().indexOf('PUESTO') !== -1);
      let qrColIdx = headers.findIndex(h => h.toUpperCase() === 'CODIGO_QR');

      if (qrColIdx === -1) {
        qrColIdx = lastCol;
        sheet.getRange(1, qrColIdx + 1).setValue('CODIGO_QR');
        lastCol++;
      }

      const numRows = lastRow - 1;
      const values = sheet.getRange(2, 1, numRows, lastCol).getValues();
      const usedCodes = {};
      const byFeria = {};
      let updatedCount = 0;
      let missingFeriaCount = 0;
      let missingPuestoCount = 0;

      const idUpdates = [];
      const qrUpdates = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const hasAny = row.some(v => v !== '' && v !== null && v !== undefined);
        const curId = idColIdx !== -1 ? String(row[idColIdx] || '').trim() : '';
        const curQr = String(row[qrColIdx] || '').trim();

        if (!hasAny) {
          idUpdates.push([curId]);
          qrUpdates.push([curQr]);
          continue;
        }

        let finalId = curId;
        if (!finalId && idColIdx !== -1) {
          finalId = Utilities.getUuid().substring(0, 8).toUpperCase();
          updatedCount++;
        }

        const rawFeria = feriaColIdx !== -1 ? String(row[feriaColIdx] || '').trim() : '';
        const normFeria = rawFeria ? this._normalizeFeriaName(rawFeria) : 'SIN_FERIA';
        if (!rawFeria) missingFeriaCount++;
        byFeria[normFeria] = (byFeria[normFeria] || 0) + 1;

        const rawPuesto = puestoColIdx !== -1 ? row[puestoColIdx] : '';
        if (!this._extractPuestoNumber(rawPuesto)) missingPuestoCount++;

        let code = this._buildUniqueFerianteQrCode(normFeria, rawPuesto, finalId, i + 2);
        if (usedCodes[code]) {
          let seq = 2;
          while (usedCodes[code + '-' + seq]) seq++;
          code = code + '-' + seq;
        }
        usedCodes[code] = true;

        if (code !== curQr) updatedCount++;
        idUpdates.push([finalId]);
        qrUpdates.push([code]);
      }

      if (idColIdx !== -1) {
        sheet.getRange(2, idColIdx + 1, numRows, 1).setNumberFormat('@').setValues(idUpdates);
      }
      sheet.getRange(2, qrColIdx + 1, numRows, 1).setNumberFormat('@').setValues(qrUpdates);
      SpreadsheetApp.flush();

      return {
        ok: true,
        totalRows: numRows,
        updatedCount,
        missingFeriaCount,
        missingPuestoCount,
        byFeria
      };
    } finally {
      lock.releaseLock();
    }
  },

  _formatSheetDateYmd(val, ssTz) {
    if (!val) return '';
    if (val instanceof Date) {
      return Utilities.formatDate(val, ssTz || APP_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = String(val).replace(/^'/, '').trim();
    const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymd) {
      return ymd[1] + '-' + String(ymd[2]).padStart(2, '0') + '-' + String(ymd[3]).padStart(2, '0');
    }
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      return dmy[3] + '-' + String(dmy[2]).padStart(2, '0') + '-' + String(dmy[1]).padStart(2, '0');
    }
    return s;
  },

  _formatSheetTimeHms(val, ssTz) {
    if (!val) return '';
    if (val instanceof Date) {
      return Utilities.formatDate(val, ssTz || APP_CONFIG.TIMEZONE, 'HH:mm:ss');
    }
    return String(val).replace(/^'/, '').trim();
  },

  _getChileNowParts() {
    const now = new Date();
    const tz = APP_CONFIG.TIMEZONE || 'America/Santiago';
    return {
      now,
      ymd: Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
      hour: parseInt(Utilities.formatDate(now, tz, 'H'), 10),
      minute: parseInt(Utilities.formatDate(now, tz, 'm'), 10),
      timeStr: Utilities.formatDate(now, tz, 'HH:mm:ss'),
      stampStr: Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss')
    };
  },

  _extractPuestoNumber(raw) {
    if (raw === undefined || raw === null) return null;
    const m = String(raw).match(/\d+/);
    if (!m) return null;
    const n = parseInt(m[0], 10);
    return (n >= 1 && n <= (APP_CONFIG.FERIA_MAX_PUESTOS || 100)) ? n : null;
  },

  _getOrCreateFeriaControlSheet(feriaName) {
    const ss = this._getSpreadsheet('MAIN');
    const normalized = this._normalizeFeriaName(feriaName);
    const clean = normalized.replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');
    const sheetName = ('CONTROL_FERIA_' + clean).substring(0, 31);
    let sheet = ss.getSheetByName(sheetName);
    const headers = [
      'ID_REGISTRO', 'FECHA_JORNADA', 'HORA_MARCA', 'FERIA', 'N_PUESTO',
      'ID_FERIANTE', 'NOMBRE_DUENO', 'RUT_DUENO', 'ROL', 'PATENTE',
      'GIRO', 'ESTADO_ASISTENCIA', 'MEDIO_INGRESO', 'INSPECTOR',
      'UBICACION_PUESTO', 'OBSERVACIONES', 'INSPECTOR_ACOMPANANTE'
    ];
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    } else if (sheet.getLastColumn() < headers.length) {
      const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => String(h || '').trim());
      if (existing.indexOf('INSPECTOR_ACOMPANANTE') === -1) {
        sheet.getRange(1, headers.length).setValue('INSPECTOR_ACOMPANANTE');
      }
    }
    return { sheet, headers, sheetName, ssTz: ss.getSpreadsheetTimeZone() };
  },

  _getOrCreateFeriaClosuresSheet() {
    const ss = this._getSpreadsheet('MAIN');
    const sheetName = 'CIERRES_ASISTENCIA_FERIAS';
    const headers = [
      'ID_CIERRE', 'FECHA_JORNADA', 'FERIA', 'HORA_CIERRE', 'TIPO_CIERRE',
      'PUESTOS_OCUPADOS', 'ASISTEN', 'NO_ASISTEN', 'PENDIENTES',
      'POR_QR', 'MANUAL', 'INSPECTORES', 'HOJA_EXCEL',
      'PDF_NOMBRE', 'PDF_URL', 'PDF_ID'
    ];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    return { sheet, headers, sheetName, ssTz: ss.getSpreadsheetTimeZone() };
  },

  _getFeriaClosureRecord(feriaName, fechaStr) {
    const feria = this._normalizeFeriaName(feriaName);
    const { sheet, headers, ssTz } = this._getOrCreateFeriaClosuresSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const values = sheet.getRange(2, 1, lastRow - 1, Math.min(sheet.getLastColumn(), headers.length)).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      const r = values[i];
      const rFecha = this._formatSheetDateYmd(r[1], ssTz);
      const rFeria = this._normalizeFeriaName(r[2]);
      if (rFeria === feria && rFecha === fechaStr) {
        return {
          _RowNumber: i + 2,
          idCierre: r[0] || '',
          fecha: rFecha,
          feria: rFeria,
          horaCierre: this._formatSheetTimeHms(r[3], ssTz),
          tipoCierre: String(r[4] || ''),
          ocupados: Number(r[5]) || 0,
          asisten: Number(r[6]) || 0,
          noAsisten: Number(r[7]) || 0,
          pendientes: Number(r[8]) || 0,
          porQr: Number(r[9]) || 0,
          manual: Number(r[10]) || 0,
          inspectores: String(r[11] || ''),
          sheetName: String(r[12] || ''),
          pdfName: String(r[13] || ''),
          pdfUrl: String(r[14] || ''),
          pdfId: String(r[15] || ''),
          downloadUrl: r[15] ? ('https://drive.google.com/uc?export=download&id=' + r[15]) : ''
        };
      }
    }
    return null;
  },

  _getFeriaClosureStatus(feriaName, fechaStr) {
    const feria = this._normalizeFeriaName(feriaName);
    const chile = this._getChileNowParts();
    const cutoffHour = Number(APP_CONFIG.FERIA_ATTENDANCE_CUTOFF_HOUR);
    const cutoffH = isNaN(cutoffHour) ? 17 : cutoffHour;
    const fecha = String(fechaStr || chile.ymd).trim();

    const record = this._getFeriaClosureRecord(feria, fecha);
    const isPastDate = fecha < chile.ymd;
    const isPastCutoffToday = (fecha === chile.ymd) && (chile.hour >= cutoffH);
    const isClosed = Boolean(record || isPastDate || isPastCutoffToday);

    let reason = 'ABIERTA';
    if (record) {
      reason = record.tipoCierre || 'CERRADA';
    } else if (isPastDate) {
      reason = 'JORNADA_ANTERIOR';
    } else if (isPastCutoffToday) {
      reason = 'TOPE_17H';
    }

    return {
      isClosed,
      reason,
      cutoffHour: cutoffH,
      chileDate: chile.ymd,
      chileTime: chile.timeStr,
      closedAt: record ? record.horaCierre : (isPastCutoffToday || isPastDate ? (String(cutoffH).padStart(2, '0') + ':00:00') : ''),
      pdfName: record ? record.pdfName : '',
      pdfUrl: record ? record.pdfUrl : '',
      pdfId: record ? record.pdfId : '',
      downloadUrl: record ? record.downloadUrl : '',
      autoTriggerEnabled: this.hasDailyFeriaCloseTrigger()
    };
  },

  getFeriaAttendance(feriaName, fechaStr) {
    const feria = this._normalizeFeriaName(feriaName);
    const chile = this._getChileNowParts();
    const targetFecha = String(fechaStr || chile.ymd).trim();
    const { sheet, headers, sheetName, ssTz } = this._getOrCreateFeriaControlSheet(feria);
    const lastRow = sheet.getLastRow();
    const marks = {};

    if (lastRow >= 2) {
      const readCols = Math.min(sheet.getLastColumn(), headers.length);
      const values = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();
      values.forEach((row, idx) => {
        const fVal = this._formatSheetDateYmd(row[1], ssTz);
        if (targetFecha && fVal !== targetFecha) return;
        const puesto = Number(row[4]);
        if (!puesto) return;
        marks[puesto] = {
          _RowNumber: idx + 2,
          idRegistro: row[0],
          fecha: fVal,
          hora: this._formatSheetTimeHms(row[2], ssTz),
          feria: row[3] || feria,
          puesto: puesto,
          idFeriante: row[5] || '',
          nombre: row[6] || '',
          rut: row[7] || '',
          rol: row[8] || '',
          patente: row[9] || '',
          giro: row[10] || '',
          estado: row[11] || '',
          medio: row[12] || '',
          inspector: row[13] || '',
          ubicacion: row[14] || '',
          observaciones: row[15] || '',
          inspectorAcompanante: row[16] || ''
        };
      });
    }

    const closure = this._getFeriaClosureStatus(feria, targetFecha);
    return {
      ok: true,
      feria,
      fecha: targetFecha,
      sheetName,
      cutoffHour: closure.cutoffHour,
      closure,
      marks
    };
  },

  saveFeriaAttendance(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const feria = this._normalizeFeriaName(payload.feria);
      const chile = this._getChileNowParts();
      const fecha = String(payload.fecha || chile.ymd).trim();
      const puesto = this._extractPuestoNumber(payload.puesto) || 0;
      const idFeriante = String(payload.idFeriante || '').trim();
      if (!feria || (!puesto && !idFeriante)) throw new Error('Feria y N° de puesto (o ID feriante) son obligatorios.');

      // Validar cierre diario (tope 17:00 hrs o jornada ya cerrada)
      const closure = this._getFeriaClosureStatus(feria, fecha);
      if (closure.isClosed) {
        const cutoffStr = String(closure.cutoffHour).padStart(2, '0') + ':00 hrs';
        throw new Error(
          'La toma de asistencia para FERIA ' + feria + ' del día ' + fecha +
          ' se encuentra CERRADA (tope diario ' + cutoffStr + '). Ya no se pueden registrar ni modificar marcas.'
        );
      }

      const { sheet, headers, sheetName, ssTz } = this._getOrCreateFeriaControlSheet(feria);
      const hora = chile.timeStr;

      // Buscar si ya existe marca para esta feria + fecha + puesto para actualizarla
      const lastRow = sheet.getLastRow();
      let existingRow = null;
      let idReg = Utilities.getUuid().substring(0, 8).toUpperCase();
      if (lastRow >= 2) {
        const existingVals = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
        for (let i = 0; i < existingVals.length; i++) {
          const rFecha = this._formatSheetDateYmd(existingVals[i][1], ssTz);
          const rPuesto = Number(existingVals[i][4]) || 0;
          const rId = String(existingVals[i][5] || '').trim();
          const sameStall = puesto > 0
            ? (rPuesto === puesto && (!idFeriante || !rId || rId === idFeriante))
            : (idFeriante && rId === idFeriante);
          if (rFecha === fecha && sameStall) {
            existingRow = i + 2;
            idReg = existingVals[i][0] || idReg;
            break;
          }
        }
      }

      const acompStr = Array.isArray(payload.inspectorAcompanante)
        ? payload.inspectorAcompanante.join(' , ')
        : (payload.inspectorAcompanante || '');

      const rowArr = [
        idReg,
        fecha,
        hora,
        feria,
        puesto,
        payload.idFeriante || '',
        payload.nombre || '',
        payload.rut || '',
        payload.rol || '',
        payload.patente || '',
        payload.giro || '',
        payload.estado || 'ASISTE',
        payload.medio || 'MANUAL',
        payload.inspector || '',
        payload.ubicacion || '',
        payload.observaciones || '',
        acompStr
      ];

      const targetRow = existingRow || (sheet.getLastRow() + 1);
      // Formatear FECHA_JORNADA y HORA_MARCA como texto plano para evitar desfase por zona horaria
      sheet.getRange(targetRow, 2, 1, 2).setNumberFormat('@');
      sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowArr]);
      SpreadsheetApp.flush();

      return {
        ok: true,
        sheetName,
        rowNumber: targetRow,
        mark: {
          _RowNumber: targetRow,
          idRegistro: idReg,
          fecha,
          hora,
          feria,
          puesto,
          idFeriante: rowArr[5],
          nombre: rowArr[6],
          rut: rowArr[7],
          rol: rowArr[8],
          patente: rowArr[9],
          giro: rowArr[10],
          estado: rowArr[11],
          medio: rowArr[12],
          inspector: rowArr[13],
          ubicacion: rowArr[14],
          observaciones: rowArr[15],
          inspectorAcompanante: rowArr[16]
        }
      };
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Cierra la jornada de una feria para una fecha específica y crea su documento PDF diario individual.
   */
  closeFeriaAttendanceAndGeneratePdf(feriaName, fechaStr, options) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const feria = this._normalizeFeriaName(feriaName);
      const chile = this._getChileNowParts();
      const fecha = String(fechaStr || chile.ymd).trim();
      const opts = options || {};
      const tipoCierre = opts.tipoCierre || 'CIERRE_MANUAL';

      const att = this.getFeriaAttendance(feria, fecha);
      const marks = att.marks || {};
      const padronRows = this._safeReadRows('ASISTENCIA').filter(
        r => this._normalizeFeriaName(r['FERIA']) === feria
      );

      const byPuesto = {};
      padronRows.forEach(r => {
        const p = this._extractPuestoNumber(r['N° Puesto']);
        if (p && !byPuesto[p]) byPuesto[p] = r;
      });

      const reportRows = [];
      const inspectoresSet = {};
      let ocupados = 0, asisten = 0, noAsisten = 0, porQr = 0, manual = 0;

      const maxPuestos = APP_CONFIG.FERIA_MAX_PUESTOS || 100;
      for (let p = 1; p <= maxPuestos; p++) {
        const r = byPuesto[p] || null;
        const m = marks[p] || null;
        if (r) ocupados++;
        if (!r && !m) continue;

        if (m) {
          if (m.estado === 'ASISTE') asisten++;
          else if (m.estado === 'NO ASISTE') noAsisten++;
          if (String(m.medio || '').toUpperCase() === 'QR') porQr++;
          else manual++;
          if (m.inspector) inspectoresSet[String(m.inspector).trim()] = true;
          if (m.inspectorAcompanante) {
            String(m.inspectorAcompanante).split(',').forEach(a => {
              const t = a.trim();
              if (t) inspectoresSet[t] = true;
            });
          }
        }

        reportRows.push({
          puesto: p,
          nombre: (m && m.nombre) || (r && r['NOMBRE DUEÑO']) || ('PUESTO #' + p),
          rut: (m && m.rut) || (r && r['RUT DUEÑO']) || '',
          rol: (m && m.rol) || (r && r['ROL']) || '',
          patente: (m && m.patente) || (r && r['PATENTE']) || '',
          giro: (m && m.giro) || (r && r['GIRO']) || '',
          estado: m ? m.estado : 'PENDIENTE',
          medio: m ? m.medio : '',
          hora: m ? m.hora : '',
          inspector: m ? (m.inspector + (m.inspectorAcompanante ? ' / ' + m.inspectorAcompanante : '')) : ''
        });
      }

      const pendientes = Math.max(ocupados - (asisten + noAsisten), 0);
      const inspectoresStr = Object.keys(inspectoresSet).join(', ') || 'No especificado';

      const pdf = PdfService.generateFeriaDailyPdf(feria, fecha, {
        sheetName: att.sheetName,
        tipoCierre: tipoCierre === 'AUTOMATICO_17H' ? 'CIERRE AUTOMÁTICO DIARIO (17:00 HRS)' : 'CIERRE DE JORNADA (TOPE 17:00 HRS)',
        inspectores: inspectoresStr,
        summary: { ocupados, asisten, noAsisten, pendientes, porQr, manual },
        rows: reportRows
      });

      // Guardar / actualizar el registro del cierre y enlace PDF en CIERRES_ASISTENCIA_FERIAS
      const { sheet, headers } = this._getOrCreateFeriaClosuresSheet();
      const existing = this._getFeriaClosureRecord(feria, fecha);
      const idCierre = existing ? existing.idCierre : Utilities.getUuid().substring(0, 8).toUpperCase();
      const targetRow = existing ? existing._RowNumber : (sheet.getLastRow() + 1);

      const closureRow = [
        idCierre,
        fecha,
        feria,
        chile.timeStr,
        tipoCierre,
        ocupados,
        asisten,
        noAsisten,
        pendientes,
        porQr,
        manual,
        inspectoresStr,
        att.sheetName,
        pdf.name,
        pdf.url,
        pdf.id
      ];

      sheet.getRange(targetRow, 2, 1, 1).setNumberFormat('@');
      sheet.getRange(targetRow, 4, 1, 1).setNumberFormat('@');
      sheet.getRange(targetRow, 1, 1, headers.length).setValues([closureRow]);
      SpreadsheetApp.flush();

      return {
        ok: true,
        feria,
        fecha,
        sheetName: att.sheetName,
        pdf,
        closure: this._getFeriaClosureStatus(feria, fecha)
      };
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Lista los reportes PDF diarios individuales generados (opcionalmente filtrados por feria).
   */
  listFeriaDailyPdfs(feriaName, limit) {
    const filterFeria = feriaName ? this._normalizeFeriaName(feriaName) : '';
    const { sheet, headers, ssTz } = this._getOrCreateFeriaClosuresSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(2, 1, lastRow - 1, Math.min(sheet.getLastColumn(), headers.length)).getValues();
    const list = [];
    values.forEach(r => {
      const f = this._normalizeFeriaName(r[2]);
      if (filterFeria && f !== filterFeria) return;
      if (!r[14]) return;
      list.push({
        idCierre: r[0] || '',
        fecha: this._formatSheetDateYmd(r[1], ssTz),
        feria: f,
        horaCierre: this._formatSheetTimeHms(r[3], ssTz),
        tipoCierre: String(r[4] || ''),
        ocupados: Number(r[5]) || 0,
        asisten: Number(r[6]) || 0,
        noAsisten: Number(r[7]) || 0,
        pendientes: Number(r[8]) || 0,
        porQr: Number(r[9]) || 0,
        manual: Number(r[10]) || 0,
        inspectores: String(r[11] || ''),
        sheetName: String(r[12] || ''),
        pdfName: String(r[13] || ''),
        pdfUrl: String(r[14] || ''),
        pdfId: String(r[15] || ''),
        downloadUrl: r[15] ? ('https://drive.google.com/uc?export=download&id=' + r[15]) : String(r[14] || '')
      });
    });

    list.sort((a, b) => (b.fecha + ' ' + b.horaCierre).localeCompare(a.fecha + ' ' + a.horaCierre));
    return list.slice(0, limit || 40);
  },

  /**
   * Proceso automático de las 17:00 hrs Chile:
   * Recorre todas las ferias con asistencias registradas hoy, cierra cada jornada
   * y genera su reporte PDF diario individual en Google Drive.
   */
  runDailyAutoCloseAt17(fechaOverride) {
    const chile = this._getChileNowParts();
    const fecha = String(fechaOverride || chile.ymd).trim();
    const ss = this._getSpreadsheet('MAIN');

    // Reunir todas las ferias del catálogo + cualquier hoja CONTROL_FERIA_* existente
    const feriasMap = {};
    (APP_CONFIG.FERIAS_LIST || []).forEach(f => {
      feriasMap[this._normalizeFeriaName(f)] = true;
    });
    ss.getSheets().forEach(sh => {
      const name = sh.getName();
      if (name.indexOf('CONTROL_FERIA_') === 0) {
        const lastRow = sh.getLastRow();
        if (lastRow >= 2) {
          const sampleFeria = sh.getRange(2, 4).getValue();
          if (sampleFeria) {
            feriasMap[this._normalizeFeriaName(sampleFeria)] = true;
          } else {
            feriasMap[name.replace(/^CONTROL_FERIA_/, '').replace(/_/g, ' ')] = true;
          }
        }
      }
    });

    const results = [];
    Object.keys(feriasMap).forEach(feria => {
      try {
        const att = this.getFeriaAttendance(feria, fecha);
        const marksCount = Object.keys(att.marks || {}).length;
        if (marksCount === 0) return;

        // Evitar duplicar si ya existe PDF generado para esta feria y fecha
        const existing = this._getFeriaClosureRecord(feria, fecha);
        if (existing && existing.pdfUrl) {
          results.push({ feria, fecha, skipped: true, reason: 'PDF ya generado', pdfUrl: existing.pdfUrl });
          return;
        }

        const res = this.closeFeriaAttendanceAndGeneratePdf(feria, fecha, { tipoCierre: 'AUTOMATICO_17H' });
        results.push({ feria, fecha, ok: true, marksCount, pdfUrl: res.pdf.url, pdfName: res.pdf.name });
      } catch (e) {
        Logger.log('Error cerrando feria ' + feria + ' (' + fecha + '): ' + e.message);
        results.push({ feria, fecha, ok: false, error: e.message });
      }
    });

    return { ok: true, fecha, closedCount: results.filter(r => r.ok).length, results };
  },

  setupDailyFeriaCloseTrigger() {
    const fnName = 'scheduledDailyFeriaAttendanceClose';
    const existing = ScriptApp.getProjectTriggers();
    existing.forEach(tr => {
      if (tr.getHandlerFunction() === fnName) {
        ScriptApp.deleteTrigger(tr);
      }
    });

    const cutoffHour = Number(APP_CONFIG.FERIA_ATTENDANCE_CUTOFF_HOUR);
    const h = isNaN(cutoffHour) ? 17 : cutoffHour;
    ScriptApp.newTrigger(fnName)
      .timeBased()
      .everyDays(1)
      .atHour(h)
      .nearMinute(5)
      .create();

    return { ok: true, enabled: true, message: 'Cierre automático diario y creación de PDF individual por feria programado a las ' + String(h).padStart(2, '0') + ':00 hrs.' };
  },

  removeDailyFeriaCloseTrigger() {
    const fnName = 'scheduledDailyFeriaAttendanceClose';
    const existing = ScriptApp.getProjectTriggers();
    existing.forEach(tr => {
      if (tr.getHandlerFunction() === fnName) {
        ScriptApp.deleteTrigger(tr);
      }
    });
    return { ok: true, enabled: false, message: 'Cierre automático diario de ferias desactivado.' };
  },

  hasDailyFeriaCloseTrigger() {
    try {
      return ScriptApp.getProjectTriggers().some(tr => tr.getHandlerFunction() === 'scheduledDailyFeriaAttendanceClose');
    } catch (e) {
      return false;
    }
  },

  // ==========================================================================
  // CIERRE DE AÑO Y BALANCE ANUAL (ARCHIVO POR AÑO + PDF ESTADÍSTICO + RESET)
  // ==========================================================================
  _extractYearMonth(val) {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) {
      return {
        year: parseInt(Utilities.formatDate(val, APP_CONFIG.TIMEZONE, 'yyyy'), 10),
        month: parseInt(Utilities.formatDate(val, APP_CONFIG.TIMEZONE, 'M'), 10)
      };
    }
    const s = String(val).replace(/^'/, '').trim();
    const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymd) return { year: parseInt(ymd[1], 10), month: parseInt(ymd[2], 10) };
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) return { year: parseInt(dmy[3], 10), month: parseInt(dmy[2], 10) };
    return null;
  },

  _getOrCreateAnnualClosuresSheet() {
    const ss = this._getSpreadsheet('MAIN');
    const sheetName = 'CIERRES_ANUALES';
    const headers = [
      'ID_CIERRE', 'ANIO', 'FECHA_CIERRE', 'RESPONSABLE', 'TIPO_CIERRE',
      'TOTAL_PROCEDIMIENTOS', 'COMERCIO', 'FUERA_HORA', 'PUBLICIDAD',
      'PATENTES_MORA', 'MONTO_MORA', 'REPORTES', 'NOTIFICACIONES',
      'ASISTENCIAS_FERIAS', 'CARPETA_ANIO_URL', 'EXCEL_MAIN_URL',
      'EXCEL_DEUDAS_URL', 'PDF_NOMBRE', 'PDF_URL', 'PDF_ID'
    ];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    return { sheet, headers, sheetName, ssTz: ss.getSpreadsheetTimeZone() };
  },

  listAnnualClosures() {
    const { sheet, headers } = this._getOrCreateAnnualClosuresSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, Math.min(sheet.getLastColumn(), headers.length)).getValues();
    const list = [];
    values.forEach(r => {
      if (!r[1]) return;
      list.push({
        idCierre: String(r[0] || ''),
        year: String(r[1] || ''),
        fechaCierre: r[2] instanceof Date ? Utilities.formatDate(r[2], APP_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm') : String(r[2] || ''),
        responsable: String(r[3] || ''),
        tipoCierre: String(r[4] || ''),
        totalProcedimientos: Number(r[5]) || 0,
        comercio: Number(r[6]) || 0,
        fueraHora: Number(r[7]) || 0,
        publicidad: Number(r[8]) || 0,
        patentesMora: Number(r[9]) || 0,
        montoMora: Number(r[10]) || 0,
        reportes: Number(r[11]) || 0,
        notificaciones: Number(r[12]) || 0,
        asistenciasFerias: Number(r[13]) || 0,
        folderUrl: String(r[14] || ''),
        excelMainUrl: String(r[15] || ''),
        excelDeudasUrl: String(r[16] || ''),
        pdfName: String(r[17] || ''),
        pdfUrl: String(r[18] || ''),
        pdfId: String(r[19] || ''),
        downloadPdfUrl: r[19] ? ('https://drive.google.com/uc?export=download&id=' + r[19]) : String(r[18] || '')
      });
    });
    list.sort((a, b) => String(b.year).localeCompare(String(a.year)) || String(b.fechaCierre).localeCompare(String(a.fechaCierre)));
    return list;
  },

  getAnnualBalanceStats(yearParam) {
    const chile = this._getChileNowParts();
    const currentYear = parseInt(chile.ymd.substring(0, 4), 10);
    const targetYear = Number(yearParam) || currentYear;

    const filterByYear = (rows, dateCols) => {
      return (rows || []).filter(r => {
        let ym = null;
        for (let i = 0; i < dateCols.length; i++) {
          ym = this._extractYearMonth(r[dateCols[i]]);
          if (ym) break;
        }
        // Si el registro tiene fecha, filtrar por año; si no tiene fecha, incluirlo en el año en curso
        if (!ym) return targetYear === currentYear;
        return ym.year === targetYear;
      });
    };

    const comercio = filterByYear(this._safeReadRows('COMERCIO'), ['FECHA HORA', 'FECHA']);
    const fueraHora = filterByYear(this._safeReadRows('FUERA DE HORA'), ['FECHA HORA', 'FECHA']);
    const publicidad = filterByYear(this._safeReadRows('GIGANTOGRAFIA'), ['FECHA HORA', 'FECHA']);
    const mora = filterByYear(this._safeReadRows('PATENTES EN MORA'), ['FECHA FISCALIZACION', 'FECHA HORA', 'FECHA']);
    const reportes = filterByYear(this._safeReadRows('REPORTE'), ['FECHA HORA', 'FECHA']);
    const feriasNavidenas = filterByYear(this._safeReadRows('FERIAS NAVIDEÑAS'), ['FECHA HORA', 'FECHA']);
    const semanaPenaflorina = filterByYear(this._safeReadRows('SEMANA PEÑAFLORINA'), ['FECHA HORA', 'FECHA']);
    const notificaciones = filterByYear(this._safeReadRows('Hoja 1'), ['FECHA DE NOTIFICACION', 'FECHA DE TERMINO', 'FECHA']);
    const padronFeriantes = this._safeReadRows('ASISTENCIA');

    // Leer asistencias de todas las hojas CONTROL_FERIA_*
    const ss = this._getSpreadsheet('MAIN');
    const ssTz = ss.getSpreadsheetTimeZone();
    const feriaStatsMap = {};
    (APP_CONFIG.FERIAS_LIST || []).forEach(f => {
      const nf = this._normalizeFeriaName(f);
      feriaStatsMap[nf] = { feria: nf, padron: 0, asisten: 0, noAsisten: 0, porQr: 0, manual: 0 };
    });
    padronFeriantes.forEach(r => {
      const nf = this._normalizeFeriaName(r['FERIA']);
      if (!nf || nf === 'GENERAL') return;
      if (!feriaStatsMap[nf]) feriaStatsMap[nf] = { feria: nf, padron: 0, asisten: 0, noAsisten: 0, porQr: 0, manual: 0 };
      feriaStatsMap[nf].padron++;
    });

    const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    const monthlyCounts = new Array(12).fill(0);
    const addMonth = (val) => {
      const ym = this._extractYearMonth(val);
      if (ym && ym.year === targetYear && ym.month >= 1 && ym.month <= 12) {
        monthlyCounts[ym.month - 1]++;
      }
    };

    const inspMap = {};
    const addInsp = (name, field) => {
      const k = String(name || '').trim();
      if (!k) return;
      if (!inspMap[k]) inspMap[k] = { inspector: k, comercio: 0, fueraHora: 0, publicidad: 0, reportes: 0, feriasAsist: 0, total: 0 };
      inspMap[k][field]++;
      inspMap[k].total++;
    };

    comercio.forEach(r => { addMonth(r['FECHA HORA'] || r['FECHA']); addInsp(r['INSPECTOR'], 'comercio'); });
    fueraHora.forEach(r => { addMonth(r['FECHA HORA'] || r['FECHA']); addInsp(r['INSPECTOR'], 'fueraHora'); });
    publicidad.forEach(r => { addMonth(r['FECHA HORA'] || r['FECHA']); addInsp(r['INSPECTOR'], 'publicidad'); });
    mora.forEach(r => { addMonth(r['FECHA FISCALIZACION'] || r['FECHA HORA']); });
    reportes.forEach(r => { addMonth(r['FECHA HORA'] || r['FECHA']); addInsp(r['FUNCIONARIO'] || r['INSPECTOR'], 'reportes'); });
    feriasNavidenas.forEach(r => { addMonth(r['FECHA HORA'] || r['FECHA']); });
    semanaPenaflorina.forEach(r => { addMonth(r['FECHA HORA'] || r['FECHA']); });
    notificaciones.forEach(r => { addMonth(r['FECHA DE NOTIFICACION'] || r['FECHA DE TERMINO']); });

    let totalAsistenciasFerias = 0, totalAsistenciasAsiste = 0, totalAsistenciasNoAsiste = 0, totalAsistenciasQr = 0, totalAsistenciasManual = 0;
    ss.getSheets().forEach(sh => {
      const shName = sh.getName();
      if (shName.indexOf('CONTROL_FERIA_') !== 0) return;
      const lr = sh.getLastRow();
      if (lr < 2) return;
      const lc = Math.min(sh.getLastColumn(), 17);
      const vals = sh.getRange(2, 1, lr - 1, lc).getValues();
      vals.forEach(row => {
        const fStr = this._formatSheetDateYmd(row[1], ssTz);
        const ym = this._extractYearMonth(fStr);
        if (ym && ym.year !== targetYear) return;
        const nf = this._normalizeFeriaName(row[3] || shName.replace(/^CONTROL_FERIA_/, '').replace(/_/g, ' '));
        if (!feriaStatsMap[nf]) feriaStatsMap[nf] = { feria: nf, padron: 0, asisten: 0, noAsisten: 0, porQr: 0, manual: 0 };
        const est = String(row[11] || '').toUpperCase().trim();
        const med = String(row[12] || '').toUpperCase().trim();
        totalAsistenciasFerias++;
        if (est === 'ASISTE') { totalAsistenciasAsiste++; feriaStatsMap[nf].asisten++; }
        else if (est === 'NO ASISTE') { totalAsistenciasNoAsiste++; feriaStatsMap[nf].noAsisten++; }
        if (med === 'QR') { totalAsistenciasQr++; feriaStatsMap[nf].porQr++; }
        else { totalAsistenciasManual++; feriaStatsMap[nf].manual++; }
        addMonth(fStr);
        addInsp(row[13], 'feriasAsist');
      });
    });

    const countBy = (arr, col) => {
      const m = {};
      arr.forEach(r => {
        const k = String(r[col] || 'Sin especificar').trim() || 'Sin especificar';
        m[k] = (m[k] || 0) + 1;
      });
      return Object.keys(m).map(k => ({ label: k, value: m[k] })).sort((a, b) => b.value - a.value);
    };

    let totalMoraMonto = 0, totalAlDiaMonto = 0;
    const moraPorEstado = {};
    mora.forEach(r => {
      const est = String(r['ESTADO PATENTE'] || 'Sin estado').trim().toUpperCase() || 'SIN ESTADO';
      const val = this._parseMoney(r['VALOR PATENTE']);
      moraPorEstado[est] = (moraPorEstado[est] || 0) + val;
      if (est === 'EN MORA') totalMoraMonto += val;
      if (est === 'AL DIA') totalAlDiaMonto += val;
    });

    let totalVencidosNoPago = 0, totalVencenHoy = 0, totalNotifVigentes = 0;
    notificaciones.forEach(r => {
      const raw = r['DIAS RESTANTES'];
      if (raw === '' || raw === null || raw === undefined) return;
      const n = Number(raw);
      if (isNaN(n)) return;
      if (n < 0) totalVencidosNoPago++;
      else if (n === 0) totalVencenHoy++;
      else totalNotifVigentes++;
    });

    const totalProcedimientos = comercio.length + fueraHora.length + publicidad.length +
      mora.length + reportes.length + feriasNavidenas.length + semanaPenaflorina.length +
      notificaciones.length + totalAsistenciasFerias;

    return {
      ok: true,
      year: targetYear,
      generatedAt: chile.stampStr,
      kpis: {
        totalProcedimientos,
        totalComercio: comercio.length,
        totalFueraHora: fueraHora.length,
        totalPublicidad: publicidad.length,
        totalMoraCount: mora.length,
        totalMoraMonto: Math.round(totalMoraMonto),
        totalAlDiaMonto: Math.round(totalAlDiaMonto),
        totalReportes: reportes.length,
        totalFeriasNavidenas: feriasNavidenas.length,
        totalSemanaPenaflorina: semanaPenaflorina.length,
        totalNotificaciones: notificaciones.length,
        totalVencidosNoPago,
        totalVencenHoy,
        totalNotifVigentes,
        totalAsistenciasFerias,
        totalAsistenciasAsiste,
        totalAsistenciasNoAsiste,
        totalAsistenciasQr,
        totalAsistenciasManual,
        totalPadronFeriantes: padronFeriantes.length
      },
      monthly: monthNames.map((label, idx) => ({ month: idx + 1, label, total: monthlyCounts[idx] })),
      byInspector: Object.keys(inspMap).map(k => inspMap[k]).sort((a, b) => b.total - a.total),
      commercePatente: countBy(comercio, 'CUENTA CON PATENTE'),
      commercePermiso: countBy(comercio, 'PERMISO AL DIA'),
      moraByEstado: Object.keys(moraPorEstado).map(k => ({ label: k, value: Math.round(moraPorEstado[k]) })).sort((a, b) => b.value - a.value),
      deudasEstado: [
        { label: 'Vencidos por No Pago', value: totalVencidosNoPago },
        { label: 'Vencen Hoy', value: totalVencenHoy },
        { label: 'En Plazo / Vigentes', value: totalNotifVigentes }
      ],
      feriasBreakdown: Object.keys(feriaStatsMap).map(k => feriaStatsMap[k]).sort((a, b) => (b.asisten + b.padron) - (a.asisten + a.padron)),
      history: this.listAnnualClosures()
    };
  },

  /**
   * Genera el documento PDF de Balance Anual SIN borrar datos (para consultas o reportes parciales).
   */
  generateAnnualBalancePdf(yearParam) {
    const stats = this.getAnnualBalanceStats(yearParam);
    const user = this.getCurrentUser();
    const pdf = PdfService.generateAnnualBalancePdf(stats.year, stats, {
      isClosure: false,
      userEmail: user.email
    });
    return { ok: true, year: stats.year, pdf, stats };
  },

  /**
   * Ejecuta el Cierre de Año completo:
   * 1. Exporta respaldos Excel (.xlsx) y copias históricas de Google Sheets en CIERRES_ANUALES/AÑO_<YYYY>.
   * 2. Genera el PDF oficial de Balance Anual con todas las estadísticas.
   * 3. Registra el cierre en CIERRES_ANUALES.
   * 4. Reinicia en cero únicamente las tablas operativas (conserva padrón ASISTENCIA + QRs, FUNCIONARIO y menús).
   */
  executeYearClose(yearParam, confirmCode) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const chile = this._getChileNowParts();
      const targetYear = Number(yearParam) || parseInt(chile.ymd.substring(0, 4), 10);
      const expected = 'CERRAR ' + targetYear;
      if (String(confirmCode || '').trim().toUpperCase() !== expected) {
        throw new Error('Código de confirmación inválido. Debes escribir exactamente: ' + expected);
      }

      const user = this.getCurrentUser();
      const stats = this.getAnnualBalanceStats(targetYear);
      const yearFolder = DriveManager.getAnnualCloseFolder(targetYear);
      const stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd_HHmm');

      // 1. Exportar respaldos .xlsx y copias históricas Google Sheets ANTES de borrar nada
      const targets = [
        { key: 'MAIN', id: APP_CONFIG.SPREADSHEETS.MAIN, prefix: 'INSPECTOR_COMERCIAL' },
        { key: 'DEUDAS', id: APP_CONFIG.SPREADSHEETS.DEUDAS, prefix: 'NOTIFICACIONES_DEUDAS' }
      ];
      const excelFiles = [];
      let excelMainUrl = '', excelDeudasUrl = '';

      targets.forEach(t => {
        const exportUrl = 'https://docs.google.com/spreadsheets/d/' + t.id + '/export?format=xlsx';
        const response = UrlFetchApp.fetch(exportUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        });
        if (response.getResponseCode() !== 200) {
          throw new Error('Falló el respaldo Excel de ' + t.prefix + ' (HTTP ' + response.getResponseCode() + '). Cierre abortado: NO se borró ningún dato.');
        }
        const xlsxName = t.prefix + '_AÑO_' + targetYear + '_CIERRE_' + stamp + '.xlsx';
        const xlsxFile = yearFolder.createFile(response.getBlob().setName(xlsxName));
        if (t.key === 'MAIN') excelMainUrl = xlsxFile.getUrl();
        if (t.key === 'DEUDAS') excelDeudasUrl = xlsxFile.getUrl();

        excelFiles.push({
          key: t.key,
          id: xlsxFile.getId(),
          name: xlsxFile.getName(),
          url: xlsxFile.getUrl(),
          downloadUrl: 'https://drive.google.com/uc?export=download&id=' + xlsxFile.getId(),
          sizeKb: Math.round(xlsxFile.getSize() / 1024)
        });

        // Copia nativa en Google Sheets dentro de la carpeta del año
        try {
          DriveApp.getFileById(t.id).makeCopy('HISTORICO_' + t.prefix + '_AÑO_' + targetYear + '_' + stamp, yearFolder);
        } catch (e) {
          Logger.log('Aviso copia nativa Sheets ' + t.prefix + ': ' + e.message);
        }
      });

      // 2. Generar PDF oficial de Balance Anual
      const pdf = PdfService.generateAnnualBalancePdf(targetYear, stats, {
        isClosure: true,
        userEmail: user.email,
        excelFiles
      });

      // 3. Registrar cierre en CIERRES_ANUALES
      const { sheet: closeSheet, headers: closeHeaders } = this._getOrCreateAnnualClosuresSheet();
      const idCierre = 'ANUAL_' + targetYear + '_' + Utilities.getUuid().substring(0, 6).toUpperCase();
      const k = stats.kpis;
      const closeRow = [
        idCierre,
        String(targetYear),
        chile.stampStr,
        user.email,
        'CIERRE_ANUAL_COMPLETO',
        k.totalProcedimientos,
        k.totalComercio,
        k.totalFueraHora,
        k.totalPublicidad,
        k.totalMoraCount,
        k.totalMoraMonto,
        k.totalReportes,
        k.totalNotificaciones,
        k.totalAsistenciasFerias,
        yearFolder.getUrl(),
        excelMainUrl,
        excelDeudasUrl,
        pdf.name,
        pdf.url,
        pdf.id
      ];
      const targetRow = closeSheet.getLastRow() + 1;
      closeSheet.getRange(targetRow, 1, 1, closeHeaders.length).setValues([closeRow]);

      // 4. Limpiar (dejar en cero manteniendo fila 1 de encabezados) SOLO las tablas operativas
      const clearedTables = [];
      const clearSheetData = (sh, label) => {
        if (!sh) return;
        const lr = sh.getLastRow();
        if (lr >= 2) {
          sh.deleteRows(2, lr - 1);
          clearedTables.push({ name: label, clearedRows: lr - 1 });
        } else {
          clearedTables.push({ name: label, clearedRows: 0 });
        }
      };

      const resetList = APP_CONFIG.ANNUAL_RESET_TABLES || [
        'COMERCIO', 'FUERA DE HORA', 'GIGANTOGRAFIA', 'PATENTES EN MORA',
        'REPORTE', 'FERIAS NAVIDEÑAS', 'SEMANA PEÑAFLORINA', 'Hoja 1'
      ];
      resetList.forEach(tName => {
        try {
          const { sheet } = this._getSheet(tName);
          clearSheetData(sheet, tName);
        } catch (e) {
          Logger.log('Aviso limpiando ' + tName + ': ' + e.message);
        }
      });

      // Limpiar hojas CONTROL_FERIA_* y CIERRES_ASISTENCIA_FERIAS en MAIN
      const mainSs = this._getSpreadsheet('MAIN');
      mainSs.getSheets().forEach(sh => {
        const n = sh.getName();
        if (n.indexOf('CONTROL_FERIA_') === 0 || n === 'CIERRES_ASISTENCIA_FERIAS') {
          clearSheetData(sh, n);
        }
      });

      SpreadsheetApp.flush();

      return {
        ok: true,
        year: targetYear,
        idCierre,
        folderUrl: yearFolder.getUrl(),
        excelFiles,
        pdf,
        clearedTables,
        preservedTables: ['ASISTENCIA (Padrón Feriantes + QRs)', 'FUNCIONARIO', 'CIERRES_ANUALES', 'Menús y Catálogos'],
        stats
      };
    } finally {
      lock.releaseLock();
    }
  }
};
