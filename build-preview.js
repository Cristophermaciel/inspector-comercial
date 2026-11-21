const fs = require('fs');
const path = require('path');

const srcDir = path.join(__dirname, 'src');
const styles = fs.readFileSync(path.join(srcDir, 'Styles.html'), 'utf8');
const indexTpl = fs.readFileSync(path.join(srcDir, 'Index.html'), 'utf8');
const appScript = fs.readFileSync(path.join(srcDir, 'AppScript.html'), 'utf8');
const configGs = fs.readFileSync(path.join(srcDir, 'Config.gs'), 'utf8');

const previewShim = `
<script>
// ============================================================================
// ADAPTADOR PWA (GITHUB PAGES)
// - Si hay URL de Google Apps Script configurada: conecta EN VIVO a Sheets/Drive
// - Si no: carga TODOS los 1.715 registros reales desde ./data/initial-data.json
//   y guarda nuevos ingresos/ediciones localmente hasta conectar Apps Script.
// ============================================================================
if (typeof google === 'undefined' || !google.script) {
  window.__IS_PREVIEW_SHIM__ = true;
  (function() {
    ${configGs.replace(/const /g, 'var ')}

    var DEFAULT_GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwZo8HgRFojCedsb0_zP7tXjMKxqdSY_5YIX-Auz-WJuQYWMIPjkW3tygRAndrvrupZpg/exec';
    var LIVE_API_URL = (localStorage.getItem('INSPECTOR_GAS_API_URL') || DEFAULT_GAS_API_URL).trim();
    window.__GAS_LIVE_CONNECTED__ = Boolean(LIVE_API_URL);
    window.__DEFAULT_GAS_API_URL__ = DEFAULT_GAS_API_URL;

    // ========================================================================
    // MOTOR DE CACHÉ ULTRA-RÁPIDO (STALE-WHILE-REVALIDATE + PREFETCH + DEDUP)
    // ========================================================================
    var SWR_PREFIX = 'INSP_SWR_V3:';
    var memCache = {};
    var inFlight = {};
    var prefetchedOnce = false;

    function normalizeFuncionarios(rawRows) {
      var out = [];
      var seen = {};
      var norm = function(s) {
        return String(s || '')
          .normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')
          .toUpperCase().replace(/\\s+/g, ' ').trim();
      };
      var add = function(f) {
        var nombre = String((f && (f.nombre || f.NOMBRE)) || '').trim();
        var rol = String((f && (f.rol || f.ROL)) || '').trim();
        if (!nombre) return;
        if (/EQUIPO/i.test(rol) || /EQUIPO/i.test(nombre)) return;
        var key = norm(nombre);
        if (seen[key]) return;
        seen[key] = true;
        out.push({
          id: (f && (f.id || f.ID)) || '',
          nombre: nombre,
          rol: rol || 'INSPECTOR',
          imagen: (f && (f.imagen || f.IMAGEN)) || ''
        });
      };
      (rawRows || []).forEach(add);
      (APP_CONFIG.DEFAULT_FUNCIONARIOS || []).forEach(add);
      return out;
    }

    function normFeriaKey(f) {
      return String(f || '').toUpperCase().replace(/^FERIA\\s+/, '').replace(/\\s+/g, ' ').trim();
    }

    function getCacheKey(action, payload) {
      if (action === 'rows' || action === 'map') return action + ':' + ((payload && payload.table) || '');
      if (action === 'feriaAttendance') return action + ':' + normFeriaKey(payload && payload.feria) + ':' + ((payload && payload.fecha) || '');
      if (action === 'listFeriaPdfs') return action + ':' + normFeriaKey(payload && payload.feria);
      return action;
    }

    function readCache(key) {
      if (memCache[key]) return memCache[key];
      try {
        var raw = localStorage.getItem(SWR_PREFIX + key);
        if (raw) {
          var parsed = JSON.parse(raw);
          if (parsed && parsed.data) {
            memCache[key] = parsed;
            return parsed;
          }
        }
      } catch (e) {}
      return null;
    }

    function writeCache(key, data) {
      if (!data || data.ok === false || data.error) return;
      var entry = { ts: Date.now(), data: data };
      memCache[key] = entry;
      try {
        var str = JSON.stringify(entry);
        if (str.length < 450000) {
          localStorage.setItem(SWR_PREFIX + key, str);
        }
      } catch (e) {}
    }

    function invalidateCacheForWrite(action, payload) {
      var keysToClear = ['dashboard'];
      if (payload && payload.table) {
        keysToClear.push('rows:' + payload.table);
        keysToClear.push('map:' + payload.table);
      }
      if (action === 'saveFeriaAttendance' && payload && payload.payload) {
        keysToClear.push('feriaAttendance:' + normFeriaKey(payload.payload.feria) + ':' + (payload.payload.fecha || ''));
      }
      if (action === 'closeFeriaAttendance' && payload) {
        keysToClear.push('feriaAttendance:' + normFeriaKey(payload.feria) + ':' + (payload.fecha || ''));
        keysToClear.push('listFeriaPdfs:' + normFeriaKey(payload.feria));
        keysToClear.push('listFeriaPdfs:');
      }
      if (action === 'toggleFeriaCloseTrigger') {
        keysToClear.push('bootstrap');
      }
      if (action === 'runExcelBackup' || action === 'toggleDailyBackup') {
        keysToClear.push('listExcelBackups');
        keysToClear.push('bootstrap');
      }
      keysToClear.forEach(function(k) {
        delete memCache[k];
        try { localStorage.removeItem(SWR_PREFIX + k); } catch (e) {}
      });
    }

    window.__clearInspectorLiveCache = function() {
      memCache = {};
      try {
        Object.keys(localStorage).forEach(function(k) {
          if (k.indexOf(SWR_PREFIX) === 0) localStorage.removeItem(k);
        });
      } catch (e) {}
    };

    function fetchLiveGasRaw(action, payload) {
      var isRead = (action === 'bootstrap' || action === 'rows' || action === 'dashboard' || action === 'map' || action === 'listExcelBackups' || action === 'feriaAttendance' || action === 'listFeriaPdfs');
      var req;
      if (isRead) {
        var qs = '?action=' + encodeURIComponent(action);
        if (payload && payload.table) qs += '&table=' + encodeURIComponent(payload.table);
        if (payload && payload.feria) qs += '&feria=' + encodeURIComponent(payload.feria);
        if (payload && payload.fecha) qs += '&fecha=' + encodeURIComponent(payload.fecha);
        if (payload && payload.limit) qs += '&limit=' + encodeURIComponent(payload.limit);
        req = fetch(LIVE_API_URL + qs, { method: 'GET' });
      } else {
        req = fetch(LIVE_API_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(Object.assign({ action: action }, payload || {}))
        });
      }
      return req.then(function(r) {
        return r.text().then(function(txt) {
          var data;
          try { data = JSON.parse(txt); }
          catch (e) { throw new Error('Respuesta no válida del servidor Apps Script (HTTP ' + r.status + ').'); }
          if (data && data.ok === false && data.error) {
            throw new Error(data.error);
          }
          return data;
        });
      });
    }

    function scheduleBackgroundPrefetch() {
      if (prefetchedOnce || !LIVE_API_URL) return;
      prefetchedOnce = true;
      setTimeout(function() {
        var targets = [
          { action: 'dashboard', payload: null },
          { action: 'rows', payload: { table: 'ASISTENCIA' } }
        ];
        targets.forEach(function(t) {
          var ck = getCacheKey(t.action, t.payload);
          var existing = readCache(ck);
          if (!existing || (Date.now() - existing.ts > 120000)) {
            if (!inFlight[ck]) {
              inFlight[ck] = fetchLiveGasRaw(t.action, t.payload)
                .then(function(res) { writeCache(ck, res); return res; })
                .catch(function() { return null; })
                .finally(function() { delete inFlight[ck]; });
            }
          }
        });
      }, 1200);
    }

    function buildInstantBootstrapShell() {
      return {
        authorized: true,
        isLiveBackend: true,
        user: { email: 'cristopher.maciel@penaflor.cl', isAuthorized: true },
        config: {
          appName: APP_CONFIG.APP_NAME,
          subtitle: APP_CONFIG.SUBTITLE,
          creator: APP_CONFIG.CREATOR,
          contact: APP_CONFIG.CONTACT,
          primaryColor: APP_CONFIG.PRIMARY_COLOR
        },
        schemas: TABLE_SCHEMAS,
        viewMap: APPSHEET_VIEW_MAP,
        funcionarios: normalizeFuncionarios([]),
        navTables: {},
        driveFolders: {},
        dailyBackupEnabled: false
      };
    }

    function callLiveGas(action, payload) {
      var isRead = (action === 'bootstrap' || action === 'rows' || action === 'dashboard' || action === 'map' || action === 'listExcelBackups' || action === 'feriaAttendance' || action === 'listFeriaPdfs');
      if (!isRead) {
        invalidateCacheForWrite(action, payload);
        return fetchLiveGasRaw(action, payload);
      }

      var ck = getCacheKey(action, payload);
      var cached = readCache(ck);
      var staleMs = (action === 'bootstrap') ? 300000 : 60000;

      function refreshInBackground() {
        if (inFlight[ck]) return inFlight[ck];
        inFlight[ck] = fetchLiveGasRaw(action, payload)
          .then(function(res) {
            if (res && !res.error && res.ok !== false) {
              if (action === 'bootstrap') {
                res.schemas = TABLE_SCHEMAS;
                res.funcionarios = normalizeFuncionarios(res.funcionarios);
              }
              writeCache(ck, res);
              if (action === 'bootstrap' && window.App && App.applyLiveBootstrapUpdate) {
                App.applyLiveBootstrapUpdate(res);
              }
            }
            return res;
          })
          .finally(function() { delete inFlight[ck]; });
        return inFlight[ck];
      }

      // 1) Si hay caché: respuesta INSTANTÁNEA (0 ms) + revalidación silenciosa si está vencida
      if (cached && cached.data) {
        if (Date.now() - cached.ts > staleMs) {
          refreshInBackground().catch(function() {});
        }
        if (action === 'bootstrap') scheduleBackgroundPrefetch();
        return Promise.resolve(cached.data);
      }

      // 2) Arranque en frío de bootstrap: pinta la UI en < 10 ms y sincroniza en vivo detrás
      if (action === 'bootstrap') {
        refreshInBackground().then(function() { scheduleBackgroundPrefetch(); }).catch(function() {});
        return Promise.resolve(buildInstantBootstrapShell());
      }

      // 3) Lectura en vivo con deduplicación
      return refreshInBackground();
    }

    var DELTA_KEY = 'inspector_comercial_deltas_v3';
    var baseDbPromise = null;
    var cachedDb = null;

    function loadDeltas() {
      try {
        var raw = localStorage.getItem(DELTA_KEY);
        if (raw) return JSON.parse(raw);
      } catch(e) {}
      return { added: {}, edited: {}, deleted: {} };
    }
    function saveDeltas(d) {
      try { localStorage.setItem(DELTA_KEY, JSON.stringify(d)); } catch(e) {}
    }

    function getDbAsync() {
      if (cachedDb) return Promise.resolve(cachedDb);
      if (!baseDbPromise) {
        baseDbPromise = fetch('./data/initial-data.json?v=3')
          .then(function(r) { return r.json(); })
          .catch(function() { return {}; })
          .then(function(data) {
            cachedDb = data || {};
            return cachedDb;
          });
      }
      return baseDbPromise;
    }

    function getMergedRows(db, tableName) {
      var baseRows = (db[tableName] || []).slice();
      var deltas = loadDeltas();
      var delSet = deltas.deleted[tableName] || {};
      var editMap = deltas.edited[tableName] || {};
      var addedArr = deltas.added[tableName] || [];

      var merged = baseRows
        .filter(function(r) { return !delSet[r._RowNumber]; })
        .map(function(r) { return editMap[r._RowNumber] ? Object.assign({}, r, editMap[r._RowNumber]) : r; });

      return addedArr.concat(merged);
    }

    var localServer = {
      apiGetBootstrap: function() {
        return getDbAsync().then(function(db) {
          var funcs = getMergedRows(db, 'FUNCIONARIO');
          return {
            authorized: true,
            user: { email: 'cristopher.maciel@penaflor.cl (Datos Reales Cargados)', isAuthorized: true },
            config: { appName: APP_CONFIG.APP_NAME, subtitle: APP_CONFIG.SUBTITLE, creator: APP_CONFIG.CREATOR, contact: APP_CONFIG.CONTACT, primaryColor: APP_CONFIG.PRIMARY_COLOR },
            schemas: TABLE_SCHEMAS,
            viewMap: APPSHEET_VIEW_MAP,
            funcionarios: normalizeFuncionarios(funcs),
            navTables: {
              MENU: getMergedRows(db, 'MENU'),
              SUBMENU: getMergedRows(db, 'SUBMENU'),
              'SUBMENU 2': getMergedRows(db, 'SUBMENU 2'),
              'SUBMENU 3': getMergedRows(db, 'SUBMENU 3'),
              LISTADOS: getMergedRows(db, 'LISTADOS'),
              'FERIAS LIBRES': getMergedRows(db, 'FERIAS LIBRES'),
              DOCUMENTOS: getMergedRows(db, 'DOCUMENTOS')
            },
            driveFolders: {
              root: { url: 'https://drive.google.com/drive/folders/1ktunuY-vFv5wjz_qIlmsaGJEZL80Euw5?usp=drive_link' },
              fotos: { url: 'https://drive.google.com/drive/folders/1ktunuY-vFv5wjz_qIlmsaGJEZL80Euw5?usp=drive_link' },
              firmas: { url: 'https://drive.google.com/drive/folders/1ktunuY-vFv5wjz_qIlmsaGJEZL80Euw5?usp=drive_link' },
              pdfs: { url: 'https://drive.google.com/drive/folders/1ktunuY-vFv5wjz_qIlmsaGJEZL80Euw5?usp=drive_link' },
              excel: { url: 'https://drive.google.com/drive/folders/1ktunuY-vFv5wjz_qIlmsaGJEZL80Euw5?usp=drive_link' }
            },
            dailyBackupEnabled: true
          };
        });
      },
      apiGetTableRows: function(tableName) {
        return getDbAsync().then(function(db) {
          var rows = getMergedRows(db, tableName);
          var schema = TABLE_SCHEMAS[tableName] || { columns: [] };
          var summary = {};
          if (tableName === 'PATENTES EN MORA') {
            summary.totalPatentesMorosas = rows.reduce(function(acc, r) {
              var est = String(r['ESTADO PATENTE'] || '').trim().toUpperCase();
              var val = Number(String(r['VALOR PATENTE'] || '0').replace(/[^0-9.-]/g, '')) || 0;
              return est === 'EN MORA' ? acc + val : acc;
            }, 0);
          }
          return { tableName: tableName, headers: schema.columns.map(function(c){ return c.name; }), rows: rows, summary: summary };
        });
      },
      apiSaveRecord: function(tableName, recordData) {
        return getDbAsync().then(function() {
          var deltas = loadDeltas();
          if (recordData._RowNumber) {
            deltas.edited[tableName] = deltas.edited[tableName] || {};
            deltas.edited[tableName][recordData._RowNumber] = recordData;
          } else {
            deltas.added[tableName] = deltas.added[tableName] || [];
            recordData._RowNumber = 900000 + deltas.added[tableName].length + 1;
            if (!recordData.ID) recordData.ID = Math.random().toString(16).slice(2, 10);
            if (!recordData['FECHA HORA']) recordData['FECHA HORA'] = new Date().toLocaleString('es-CL');
            deltas.added[tableName].unshift(recordData);
          }
          saveDeltas(deltas);
          return { ok: true, isEdit: Boolean(recordData._RowNumber < 900000), record: recordData, pdf: null };
        });
      },
      apiDeleteRecord: function(tableName, rowNumber) {
        return getDbAsync().then(function() {
          var deltas = loadDeltas();
          deltas.deleted[tableName] = deltas.deleted[tableName] || {};
          deltas.deleted[tableName][rowNumber] = true;
          if (deltas.added[tableName]) {
            deltas.added[tableName] = deltas.added[tableName].filter(function(r){ return Number(r._RowNumber) !== Number(rowNumber); });
          }
          saveDeltas(deltas);
          return { ok: true };
        });
      },
      apiRegeneratePdf: function() { return Promise.resolve({ ok: true, pdf: { url: 'https://drive.google.com/drive/folders/1ktunuY-vFv5wjz_qIlmsaGJEZL80Euw5?usp=drive_link' } }); },
      apiGetDashboardStats: function() {
        return getDbAsync().then(function(db) {
          var comercio = getMergedRows(db, 'COMERCIO');
          var fueraHora = getMergedRows(db, 'FUERA DE HORA');
          var publicidad = getMergedRows(db, 'GIGANTOGRAFIA');
          var mora = getMergedRows(db, 'PATENTES EN MORA');
          var reportes = getMergedRows(db, 'REPORTE');
          var feriantes = getMergedRows(db, 'ASISTENCIA');
          var feriasNavidenas = getMergedRows(db, 'FERIAS NAVIDEÑAS');
          var notificaciones = getMergedRows(db, 'Hoja 1');

          var totalVencidosNoPago = 0, totalVencenHoy = 0, totalNotifVigentes = 0;
          notificaciones.forEach(function(r) {
            var raw = r['DIAS RESTANTES'];
            if (raw === '' || raw === null || raw === undefined) return;
            var n = Number(raw);
            if (isNaN(n)) return;
            if (n < 0) totalVencidosNoPago++;
            else if (n === 0) totalVencenHoy++;
            else totalNotifVigentes++;
          });

          var countBy = function(arr, col) {
            var m = {};
            (arr || []).forEach(function(r){
              var k = String(r[col] || 'Sin especificar').trim() || 'Sin especificar';
              m[k] = (m[k] || 0) + 1;
            });
            return Object.keys(m).map(function(k){ return { label: k, value: m[k] }; }).sort(function(a,b){ return b.value - a.value; });
          };

          var totalMora = 0, moraPorEstado = {};
          mora.forEach(function(r){
            var e = String(r['ESTADO PATENTE'] || 'Sin estado').trim().toUpperCase() || 'SIN ESTADO';
            var v = Number(String(r['VALOR PATENTE'] || '0').replace(/[^0-9.-]/g, '')) || 0;
            moraPorEstado[e] = (moraPorEstado[e] || 0) + v;
            if (e === 'EN MORA') totalMora += v;
          });

          return {
            kpis: {
              totalComercio: comercio.length,
              totalFueraHora: fueraHora.length,
              totalPublicidad: publicidad.length,
              totalMoraCount: mora.length,
              totalMoraMonto: Math.round(totalMora),
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
              valorPatentesMora: Object.keys(moraPorEstado).map(function(k){ return { label: k, value: Math.round(moraPorEstado[k]) }; }),
              totalesPatentesMora: countBy(mora.map(function(r){ return { 'ESTADO PATENTE': String(r['ESTADO PATENTE'] || 'SIN ESTADO').trim().toUpperCase() }; }), 'ESTADO PATENTE'),
              estadoNotificacionesDeudas: [
                { label: 'Vencidos por No Pago', value: totalVencidosNoPago },
                { label: 'Vencen Hoy', value: totalVencenHoy },
                { label: 'En Plazo / Vigentes', value: totalNotifVigentes }
              ]
            }
          };
        });
      },
      apiGetMapPoints: function(tableName) {
        return getDbAsync().then(function(db) {
          var rows = getMergedRows(db, tableName);
          var pts = [];
          rows.forEach(function(r){
            var m = String(r.UBICACION || r.DIRECCION || '').match(/(-?\\d+\\.\\d+)\\s*,\\s*(-?\\d+\\.\\d+)/);
            if (m) pts.push({
              lat: parseFloat(m[1]),
              lng: parseFloat(m[2]),
              title: r['NOMBRE LOCAL'] || r['NOMBRE LOCAL COMERCIAL'] || r['NOMBRE DUEÑO'] || r.FUNCIONARIO || 'Registro',
              subtitle: r.DIRECCION || '',
              inspector: r.INSPECTOR || r.FISCALIZADOR || r.FUNCIONARIO || r['CONTROLADO POR'] || '',
              fecha: r['FECHA HORA'] || r.FECHA || r['FECHA FISCALIZACION'] || ''
            });
          });
          return pts;
        });
      },
      apiRunExcelBackup: function() { return Promise.resolve({ timestamp: new Date().toLocaleString('es-CL'), files: [{ ok: true, name: 'INSPECTOR_COMERCIAL.xlsx' }] }); },
      apiListExcelBackups: function() { return Promise.resolve([]); },
      apiToggleDailyBackup: function(enable) { return Promise.resolve({ ok: true, message: enable ? 'Respaldo diario activado (23:00 hrs).' : 'Respaldo diario desactivado.' }); },
      apiGetFeriaAttendance: function(feria, fecha) {
        var fNorm = normFeriaKey(feria);
        var key = 'inspector_feria_att_v3_' + fNorm + '_' + String(fecha || '');
        var closeKey = 'inspector_feria_close_v3_' + fNorm + '_' + String(fecha || '');
        var marks = {};
        var closedRec = null;
        try { marks = JSON.parse(localStorage.getItem(key) || '{}'); } catch(e) {}
        try { closedRec = JSON.parse(localStorage.getItem(closeKey) || 'null'); } catch(e) {}
        var now = new Date();
        var today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
        var isClosed = Boolean(closedRec || (fecha && fecha < today) || (fecha === today && now.getHours() >= 17));
        return Promise.resolve({
          ok: true,
          feria: fNorm,
          fecha: fecha,
          sheetName: ('CONTROL_FERIA_' + fNorm.replace(/[^A-Z0-9]+/g, '_')).substring(0, 31),
          cutoffHour: 17,
          closure: {
            isClosed: isClosed,
            cutoffHour: 17,
            closedAt: closedRec ? closedRec.horaCierre : (isClosed ? '17:00:00' : ''),
            pdfUrl: closedRec ? closedRec.pdfUrl : '',
            pdfName: closedRec ? closedRec.pdfName : '',
            autoTriggerEnabled: false
          },
          marks: marks
        });
      },
      apiSaveFeriaAttendance: function(payload) {
        var feria = normFeriaKey(payload.feria);
        var now = new Date();
        var today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
        var fecha = String(payload.fecha || today);
        var closeKey = 'inspector_feria_close_v3_' + feria + '_' + fecha;
        var closedRec = null;
        try { closedRec = JSON.parse(localStorage.getItem(closeKey) || 'null'); } catch(e) {}
        if (closedRec || fecha < today || (fecha === today && now.getHours() >= 17)) {
          return Promise.reject(new Error('Jornada cerrada para FERIA ' + feria + ' (' + fecha + '). Tope diario: 17:00 hrs.'));
        }
        var puesto = Number(payload.puesto);
        var key = 'inspector_feria_att_v3_' + feria + '_' + fecha;
        var marks = {};
        try { marks = JSON.parse(localStorage.getItem(key) || '{}'); } catch(e) {}
        var hora = now.toTimeString().slice(0, 8);
        var mark = {
          idRegistro: (marks[puesto] && marks[puesto].idRegistro) || Math.random().toString(16).slice(2, 10).toUpperCase(),
          fecha: fecha,
          hora: hora,
          feria: feria,
          puesto: puesto,
          idFeriante: payload.idFeriante || '',
          nombre: payload.nombre || '',
          rut: payload.rut || '',
          rol: payload.rol || '',
          patente: payload.patente || '',
          giro: payload.giro || '',
          estado: payload.estado || 'ASISTE',
          medio: payload.medio || 'MANUAL',
          inspector: payload.inspector || '',
          inspectorAcompanante: payload.inspectorAcompanante || '',
          ubicacion: payload.ubicacion || '',
          observaciones: payload.observaciones || ''
        };
        marks[puesto] = mark;
        try { localStorage.setItem(key, JSON.stringify(marks)); } catch(e) {}
        return Promise.resolve({
          ok: true,
          sheetName: ('CONTROL_FERIA_' + feria.replace(/[^A-Z0-9]+/g, '_')).substring(0, 31),
          mark: mark
        });
      },
      apiCloseFeriaAttendance: function(feria, fecha, tipoCierre) {
        var fNorm = normFeriaKey(feria);
        var closeKey = 'inspector_feria_close_v3_' + fNorm + '_' + String(fecha || '');
        var now = new Date();
        var rec = { fecha: fecha, feria: fNorm, horaCierre: now.toTimeString().slice(0, 8), tipoCierre: tipoCierre || 'CIERRE_MANUAL', pdfUrl: '', pdfName: '' };
        try { localStorage.setItem(closeKey, JSON.stringify(rec)); } catch(e) {}
        return Promise.resolve({
          ok: true,
          localPrintFallback: true,
          closure: { isClosed: true, cutoffHour: 17, closedAt: rec.horaCierre, pdfUrl: '', autoTriggerEnabled: false }
        });
      },
      apiListFeriaPdfs: function() { return Promise.resolve([]); },
      apiToggleFeriaCloseTrigger: function(enable) {
        return Promise.resolve({ ok: true, enabled: Boolean(enable), message: enable ? 'Cierre automático 17:00 hrs activado.' : 'Cierre automático desactivado.' });
      }
    };

    var LIVE_MAP = {
      apiGetBootstrap: function() { return callLiveGas('bootstrap'); },
      apiGetTableRows: function(table) { return callLiveGas('rows', { table: table }); },
      apiSaveRecord: function(table, record, mediaUploads) { return callLiveGas('saveRecord', { table: table, record: record, mediaUploads: mediaUploads }); },
      apiDeleteRecord: function(table, rowNumber) { return callLiveGas('deleteRecord', { table: table, rowNumber: rowNumber }); },
      apiRegeneratePdf: function(table, rowNumber) { return callLiveGas('regeneratePdf', { table: table, rowNumber: rowNumber }); },
      apiGetDashboardStats: function() { return callLiveGas('dashboard'); },
      apiGetMapPoints: function(table) { return callLiveGas('map', { table: table }); },
      apiRunExcelBackup: function() { return callLiveGas('runExcelBackup'); },
      apiListExcelBackups: function() { return callLiveGas('listExcelBackups'); },
      apiToggleDailyBackup: function(enable) { return callLiveGas('toggleDailyBackup', { enable: enable }); },
      apiGetFeriaAttendance: function(feria, fecha) { return callLiveGas('feriaAttendance', { feria: feria, fecha: fecha }); },
      apiSaveFeriaAttendance: function(payload) { return callLiveGas('saveFeriaAttendance', { payload: payload }); },
      apiCloseFeriaAttendance: function(feria, fecha, tipoCierre) { return callLiveGas('closeFeriaAttendance', { feria: feria, fecha: fecha, tipoCierre: tipoCierre }); },
      apiListFeriaPdfs: function(feria, limit) { return callLiveGas('listFeriaPdfs', { feria: feria, limit: limit }); },
      apiToggleFeriaCloseTrigger: function(enable) { return callLiveGas('toggleFeriaCloseTrigger', { enable: enable }); }
    };

    function makeRunner(okCb, errCb) {
      var runner = {
        withSuccessHandler: function(fn) { return makeRunner(fn, errCb); },
        withFailureHandler: function(fn) { return makeRunner(okCb, fn); }
      };
      Object.keys(localServer).forEach(function(fnName) {
        runner[fnName] = function() {
          var args = Array.prototype.slice.call(arguments);
          var p = (LIVE_API_URL && LIVE_MAP[fnName])
            ? LIVE_MAP[fnName].apply(null, args)
            : localServer[fnName].apply(null, args);
          p.then(function(res) { if (okCb) okCb(res); })
           .catch(function(err) { if (errCb) errCb(err); });
        };
      });
      return runner;
    }

    window.google = { script: { run: makeRunner(null, null) } };
  })();
}
</script>
`;

const output = indexTpl
  .replace(/<\?=\s*appName\s*\?>/g, () => 'Departamento de Inspectores Comerciales')
  .replace(/<\?=\s*subtitle\s*\?>/g, () => 'Seguridad Pública Municipalidad de Peñaflor')
  .replace(/<\?!=\s*include\('Styles'\);\s*\?>/g, () => styles)
  .replace(/<\?!=\s*include\('AppScript'\);\s*\?>/g, () => previewShim + '\n' + appScript);

fs.writeFileSync(path.join(__dirname, 'index.html'), output, 'utf8');
console.log('Generated index.html successfully (' + output.length + ' bytes)');
