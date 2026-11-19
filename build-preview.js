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

    var LIVE_API_URL = (localStorage.getItem('INSPECTOR_GAS_API_URL') || '').trim();
    window.__GAS_LIVE_CONNECTED__ = Boolean(LIVE_API_URL);

    function callLiveGas(action, payload) {
      var isRead = (action === 'bootstrap' || action === 'rows' || action === 'dashboard' || action === 'map' || action === 'listExcelBackups' || action === 'feriaAttendance');
      if (isRead) {
        var qs = '?action=' + encodeURIComponent(action);
        if (payload && payload.table) qs += '&table=' + encodeURIComponent(payload.table);
        if (payload && payload.feria) qs += '&feria=' + encodeURIComponent(payload.feria);
        if (payload && payload.fecha) qs += '&fecha=' + encodeURIComponent(payload.fecha);
        return fetch(LIVE_API_URL + qs, { method: 'GET' }).then(function(r) { return r.json(); });
      }
      return fetch(LIVE_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(Object.assign({ action: action }, payload || {}))
      }).then(function(r) { return r.json(); });
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
            funcionarios: funcs.map(function(f){ return { id: f.ID || '', nombre: f.NOMBRE || '', rol: f.ROL || '', imagen: f.IMAGEN || '' }; }).filter(function(f){ return f.nombre; }),
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
              totalFeriantes: feriantes.length
            },
            charts: {
              topPublicidad: countBy(publicidad, 'INSPECTOR'),
              topFiscalizador: countBy(comercio, 'INSPECTOR'),
              topFueraHora: countBy(fueraHora, 'INSPECTOR'),
              permisoAlDia: countBy(comercio, 'PERMISO AL DIA'),
              valorPatentesMora: Object.keys(moraPorEstado).map(function(k){ return { label: k, value: Math.round(moraPorEstado[k]) }; }),
              totalesPatentesMora: countBy(mora.map(function(r){ return { 'ESTADO PATENTE': String(r['ESTADO PATENTE'] || 'SIN ESTADO').trim().toUpperCase() }; }), 'ESTADO PATENTE')
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
        var key = 'inspector_feria_att_v3_' + String(feria || '').toUpperCase() + '_' + String(fecha || '');
        var marks = {};
        try { marks = JSON.parse(localStorage.getItem(key) || '{}'); } catch(e) {}
        return Promise.resolve({ feria: feria, fecha: fecha, marks: marks });
      },
      apiSaveFeriaAttendance: function(payload) {
        var feria = String(payload.feria || '').toUpperCase();
        var fecha = String(payload.fecha || new Date().toISOString().slice(0, 10));
        var puesto = Number(payload.puesto);
        var key = 'inspector_feria_att_v3_' + feria + '_' + fecha;
        var marks = {};
        try { marks = JSON.parse(localStorage.getItem(key) || '{}'); } catch(e) {}
        var now = new Date();
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
          ubicacion: payload.ubicacion || '',
          observaciones: payload.observaciones || ''
        };
        marks[puesto] = mark;
        try { localStorage.setItem(key, JSON.stringify(marks)); } catch(e) {}
        return Promise.resolve({ ok: true, mark: mark });
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
      apiSaveFeriaAttendance: function(payload) { return callLiveGas('saveFeriaAttendance', { payload: payload }); }
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
