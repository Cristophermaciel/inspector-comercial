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
// ADAPTADOR PWA (GITHUB PAGES) -> CONECTA EN VIVO A GOOGLE APPS SCRIPT
// SI HAY URL CONFIGURADA, O USA MODO LOCAL DE VISTA PREVIA SI AÚN NO SE CONECTA
// ============================================================================
if (typeof google === 'undefined' || !google.script) {
  window.__IS_PREVIEW_SHIM__ = true;
  (function() {
    ${configGs.replace(/const /g, 'var ')}

    // Si el usuario guardó la URL de su Web App de Google Apps Script, conectar EN VIVO
    var LIVE_API_URL = (localStorage.getItem('INSPECTOR_GAS_API_URL') || '').trim();
    window.__GAS_LIVE_CONNECTED__ = Boolean(LIVE_API_URL);

    function callLiveGas(action, payload) {
      var isRead = (action === 'bootstrap' || action === 'rows' || action === 'dashboard' || action === 'map' || action === 'listExcelBackups');
      if (isRead) {
        var qs = '?action=' + encodeURIComponent(action);
        if (payload && payload.table) qs += '&table=' + encodeURIComponent(payload.table);
        return fetch(LIVE_API_URL + qs, { method: 'GET' }).then(function(r) { return r.json(); });
      }
      // Usar text/plain para evitar preflight CORS en Apps Script Web Apps
      return fetch(LIVE_API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(Object.assign({ action: action }, payload || {}))
      }).then(function(r) { return r.json(); });
    }

    var STORAGE_KEY = 'inspector_comercial_preview_v2';
    var defaultDb = {
      'FUNCIONARIO': [
        { _RowNumber: 2, ID: 'F01', NOMBRE: 'CRISTOPHER MACIEL', ROL: 'Administrador / Inspector' },
        { _RowNumber: 3, ID: 'F02', NOMBRE: 'DAVID HERNANDEZ', ROL: 'Jefatura Seguridad Pública' },
        { _RowNumber: 4, ID: 'F03', NOMBRE: 'LUIS ZAMORANO', ROL: 'Inspector Comercial' },
        { _RowNumber: 5, ID: 'F04', NOMBRE: 'NICOLAS ZAMORANO', ROL: 'Inspector Comercial' },
        { _RowNumber: 6, ID: 'F05', NOMBRE: 'TAMARA MARTINEZ', ROL: 'Inspectora Comercial' }
      ],
      'COMERCIO': [
        { _RowNumber: 2, ID: 'COM001', INSPECTOR: 'CRISTOPHER MACIEL', 'FECHA HORA': '19/11/2026 10:15', 'TIPO DE INTERVENCION': 'Fiscalización de rutina', 'TIPO FISCALIZACION': 'Inspección', DIRECCION: 'Av. Vicuña Mackenna 1850, Peñaflor', 'NOMBRE LOCAL': 'Minimarket El Parque', 'NOMBRE DUEÑO O RAZON SOCIAL': 'Juan Pérez Soto', RUT: '12.345.678-9', 'PRODUCTOS QUE VENDE': 'Minimarket , Alimentos', GIRO: 'COMERCIAL', 'DOCUMENTO FISCALIZADO': 'Patente comercial , Permiso municipal', 'CUENTA CON PATENTE': 'SI', 'EXHIBE PATENTE': 'SI', 'PERMISO AL DIA': 'SI', 'OBSERVACIONES FISCALIZACION': 'Documentación vigente y visible.', UBICACION: '-33.6063, -70.8764', TELEFONO: '+56987654321', EMAIL: 'contacto@elparque.cl' },
        { _RowNumber: 3, ID: 'COM002', INSPECTOR: 'LUIS ZAMORANO', 'FECHA HORA': '19/11/2026 11:30', 'TIPO DE INTERVENCION': 'Control de patentes', 'TIPO FISCALIZACION': 'Notificación', DIRECCION: '21 de Mayo 432, Peñaflor', 'NOMBRE LOCAL': 'Botillería Central', 'NOMBRE DUEÑO O RAZON SOCIAL': 'Comercial Central Ltda.', RUT: '76.543.210-K', 'PRODUCTOS QUE VENDE': 'Botillería', GIRO: 'ALCOHOLES', 'DOCUMENTO FISCALIZADO': 'Patente comercial', 'CUENTA CON PATENTE': 'ATRASADA', 'EXHIBE PATENTE': 'NO', 'PERMISO AL DIA': 'VENCIDO', 'OBSERVACIONES FISCALIZACION': 'Se cursa notificación por patente atrasada.', UBICACION: '-33.6085, -70.8812', TELEFONO: '+56911223344' },
        { _RowNumber: 4, ID: 'COM003', INSPECTOR: 'TAMARA MARTINEZ', 'FECHA HORA': '18/11/2026 16:40', 'TIPO DE INTERVENCION': 'Denuncia ciudadana', 'TIPO FISCALIZACION': 'Advertencia verbal', DIRECCION: 'Alcalde Luis Araya Cereceda 910, Peñaflor', 'NOMBRE LOCAL': 'Comida Rápida Express', 'NOMBRE DUEÑO O RAZON SOCIAL': 'María González', RUT: '15.987.654-3', 'PRODUCTOS QUE VENDE': 'Comida rápida', GIRO: 'PERMISO PROVISORIO', 'DOCUMENTO FISCALIZADO': 'Resolución sanitaria', 'CUENTA CON PATENTE': 'EN TRAMITE', 'EXHIBE PATENTE': 'NO CORRESPONDE', 'PERMISO AL DIA': 'EN TRAMITE', UBICACION: '-33.6031, -70.8720' }
      ],
      'FUERA DE HORA': [
        { _RowNumber: 2, ID: 'FH001', INSPECTOR: 'NICOLAS ZAMORANO', 'FECHA HORA': '18/11/2026 23:20', 'TIPO DE INTERVENCION': 'Control nocturno', 'TIPO FISCALIZACION': 'Infracción cursada', DIRECCION: 'Av. Vicuña Mackenna 2100', 'NOMBRE LOCAL': 'Pub Nocturno Peñaflor', RUT: '77.111.222-3', GIRO: 'Alcoholes', 'CUENTA CON PATENTE': 'SI', 'PERMISO AL DIA': 'si', UBICACION: '-33.6098, -70.8745' }
      ],
      'GIGANTOGRAFIA': [
        { _RowNumber: 2, ID: 'GIG001', INSPECTOR: 'CRISTOPHER MACIEL', 'FECHA HORA': '19/11/2026 09:45', 'TIPO DE INTERVENCION': 'Control de Publicidad', 'TIPO FISCALIZACION': 'Notificación', DIRECCION: 'Av. Vicuña Mackenna 1500', 'NOMBRE LOCAL': 'Ferretería Cordillera', 'METRAJE DE LA PUBLICIDAD': '4m x 2m (8 m2)', 'PUBLICIDAD LUMINOSA': 'SI', 'PAGA PUBLICIDAD': 'PENDIENTE', UBICACION: '-33.6050, -70.8790' }
      ],
      'PATENTES EN MORA': [
        { _RowNumber: 2, ID: 'PM001', NUMERO: '1045', 'NOMBRE LOCAL COMERCIAL': 'Botillería Central', PATENTE: '4-1045', ROL: '2026-1045', DIRECCION: '-33.6085, -70.8812', 'ESTADO PATENTE': 'EN MORA', 'VALOR PATENTE': 285000, 'FECHA FISCALIZACION': '19/11/2026 11:30', FISCALIZADOR: 'LUIS ZAMORANO' },
        { _RowNumber: 3, ID: 'PM002', NUMERO: '1098', 'NOMBRE LOCAL COMERCIAL': 'Taller Mecánico Sur', PATENTE: '2-1098', ROL: '2026-1098', DIRECCION: '-33.6110, -70.8830', 'ESTADO PATENTE': 'EN MORA', 'VALOR PATENTE': 195000, 'FECHA FISCALIZACION': '18/11/2026 15:00', FISCALIZADOR: 'CRISTOPHER MACIEL' },
        { _RowNumber: 4, ID: 'PM003', NUMERO: '1120', 'NOMBRE LOCAL COMERCIAL': 'Minimarket El Parque', PATENTE: '2-1120', ROL: '2026-1120', DIRECCION: '-33.6063, -70.8764', 'ESTADO PATENTE': 'AL DIA', 'VALOR PATENTE': 150000, 'FECHA FISCALIZACION': '19/11/2026 10:15', FISCALIZADOR: 'TAMARA MARTINEZ' }
      ],
      'REPORTE': [
        { _RowNumber: 2, ID: 'REP001', FUNCIONARIO: 'CRISTOPHER MACIEL', 'FECHA HORA': '19/11/2026 12:00', REPORTE: 'Operativo comercial realizado en sector centro sin incidentes.', 'A QUIEN LE INFORMA': 'DAVID HERNANDEZ , CENCOO / BRAVO', UBICACION: '-33.6063, -70.8764' }
      ],
      'ASISTENCIA': [
        { _RowNumber: 2, ID: 'FER001', FECHA: '19/11/2026 08:30', DIA: 'JUEVES', 'NOMBRE DUEÑO': 'Pedro Morales', 'RUT DUEÑO': '11.222.333-4', PATENTE: 'SI POSEE', 'N° Puesto': 12, ROL: 'F-012', GIRO: 'FERIAS LIBRE VTA AL DETALLE', FERIA: 'MODELO', UBICACION: '-33.6042, -70.8705' }
      ],
      'SEMANA PEÑAFLORINA': [],
      'Hoja 1': [
        { _RowNumber: 2, 'N° NOTIFICACION': 101, DIRECCION: '21 de Mayo 432, Peñaflor', 'LOCAL COMERCIAL': 'Botillería Central', 'FECHA DE NOTIFICACION': '15/11/2026 11:00', 'FECHA DE TERMINO': '2026-11-25', 'DIAS RESTANTES': '6' }
      ],
      'DOCUMENTOS': [
        { _RowNumber: 2, ID: 'D1', IR: 'Carpeta Drive Seguridad Pública', LINK: 'https://drive.google.com' }
      ],
      'MENU': [], 'SUBMENU': [], 'SUBMENU 2': [], 'SUBMENU 3': [], 'LISTADOS': [], 'FERIAS LIBRES': []
    };

    function loadDb() {
      try {
        var raw = localStorage.getItem(STORAGE_KEY);
        if (raw) return JSON.parse(raw);
      } catch(e) {}
      return JSON.parse(JSON.stringify(defaultDb));
    }
    function saveDb(db) {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch(e) {}
    }

    var mockServer = {
      apiGetBootstrap: function() {
        var db = loadDb();
        return {
          authorized: true,
          user: { email: 'cristopher.maciel@penaflor.cl (Vista Previa)', isAuthorized: true },
          config: { appName: APP_CONFIG.APP_NAME, subtitle: APP_CONFIG.SUBTITLE, creator: APP_CONFIG.CREATOR, contact: APP_CONFIG.CONTACT, primaryColor: APP_CONFIG.PRIMARY_COLOR },
          schemas: TABLE_SCHEMAS,
          viewMap: APPSHEET_VIEW_MAP,
          funcionarios: db.FUNCIONARIO.map(function(f){ return { id: f.ID, nombre: f.NOMBRE, rol: f.ROL, imagen: f.IMAGEN || '' }; }),
          navTables: { MENU: [], SUBMENU: [], 'SUBMENU 2': [], 'SUBMENU 3': [], LISTADOS: [], 'FERIAS LIBRES': [], DOCUMENTOS: db.DOCUMENTOS },
          driveFolders: { root: { url: 'https://drive.google.com' }, fotos: { url: 'https://drive.google.com' }, firmas: { url: 'https://drive.google.com' }, pdfs: { url: 'https://drive.google.com' }, excel: { url: 'https://drive.google.com' } },
          dailyBackupEnabled: true
        };
      },
      apiGetTableRows: function(tableName) {
        var db = loadDb();
        var rows = db[tableName] || [];
        var schema = TABLE_SCHEMAS[tableName] || { columns: [] };
        var summary = {};
        if (tableName === 'PATENTES EN MORA') {
          summary.totalPatentesMorosas = rows.reduce(function(acc, r) {
            return String(r['ESTADO PATENTE']).toUpperCase() === 'EN MORA' ? acc + (Number(r['VALOR PATENTE']) || 0) : acc;
          }, 0);
        }
        return { tableName: tableName, headers: schema.columns.map(function(c){ return c.name; }), rows: rows, summary: summary };
      },
      apiSaveRecord: function(tableName, recordData) {
        var db = loadDb();
        db[tableName] = db[tableName] || [];
        if (recordData._RowNumber) {
          var idx = db[tableName].findIndex(function(r){ return Number(r._RowNumber) === Number(recordData._RowNumber); });
          if (idx !== -1) db[tableName][idx] = Object.assign({}, db[tableName][idx], recordData);
        } else {
          recordData._RowNumber = db[tableName].length + 2;
          if (!recordData.ID) recordData.ID = 'REG-' + Math.random().toString(36).slice(2, 7).toUpperCase();
          if (!recordData['FECHA HORA']) recordData['FECHA HORA'] = new Date().toLocaleString('es-CL');
          db[tableName].unshift(recordData);
        }
        saveDb(db);
        return { ok: true, isEdit: false, record: recordData, pdf: null };
      },
      apiDeleteRecord: function(tableName, rowNumber) {
        var db = loadDb();
        db[tableName] = (db[tableName] || []).filter(function(r){ return Number(r._RowNumber) !== Number(rowNumber); });
        saveDb(db);
        return { ok: true };
      },
      apiRegeneratePdf: function() { return { ok: true, pdf: { url: 'https://drive.google.com' } }; },
      apiGetDashboardStats: function() {
        var db = loadDb();
        var countBy = function(arr, col) {
          var m = {};
          (arr || []).forEach(function(r){ var k = String(r[col] || 'Sin especificar'); m[k] = (m[k] || 0) + 1; });
          return Object.keys(m).map(function(k){ return { label: k, value: m[k] }; }).sort(function(a,b){ return b.value - a.value; });
        };
        var mora = db['PATENTES EN MORA'] || [];
        var totalMora = 0, moraPorEstado = {};
        mora.forEach(function(r){
          var e = String(r['ESTADO PATENTE'] || 'Sin estado');
          var v = Number(r['VALOR PATENTE']) || 0;
          moraPorEstado[e] = (moraPorEstado[e] || 0) + v;
          if (e.toUpperCase() === 'EN MORA') totalMora += v;
        });
        return {
          kpis: {
            totalComercio: (db['COMERCIO'] || []).length,
            totalFueraHora: (db['FUERA DE HORA'] || []).length,
            totalPublicidad: (db['GIGANTOGRAFIA'] || []).length,
            totalMoraCount: mora.length,
            totalMoraMonto: totalMora,
            totalReportes: (db['REPORTE'] || []).length,
            totalFeriantes: (db['ASISTENCIA'] || []).length
          },
          charts: {
            topPublicidad: countBy(db['GIGANTOGRAFIA'], 'INSPECTOR'),
            topFiscalizador: countBy(db['COMERCIO'], 'INSPECTOR'),
            topFueraHora: countBy(db['FUERA DE HORA'], 'INSPECTOR'),
            permisoAlDia: countBy(db['COMERCIO'], 'PERMISO AL DIA'),
            valorPatentesMora: Object.keys(moraPorEstado).map(function(k){ return { label: k, value: moraPorEstado[k] }; }),
            totalesPatentesMora: countBy(mora, 'ESTADO PATENTE')
          }
        };
      },
      apiGetMapPoints: function(tableName) {
        var db = loadDb();
        var pts = [];
        (db[tableName] || []).forEach(function(r){
          var m = String(r.UBICACION || r.DIRECCION || '').match(/(-?\\d+\\.\\d+)\\s*,\\s*(-?\\d+\\.\\d+)/);
          if (m) pts.push({ lat: parseFloat(m[1]), lng: parseFloat(m[2]), title: r['NOMBRE LOCAL'] || r['NOMBRE LOCAL COMERCIAL'] || r['NOMBRE DUEÑO'] || r.FUNCIONARIO || 'Registro', subtitle: r.DIRECCION || '', inspector: r.INSPECTOR || r.FISCALIZADOR || r.FUNCIONARIO || '', fecha: r['FECHA HORA'] || r.FECHA || '' });
        });
        return pts;
      },
      apiRunExcelBackup: function() { return { timestamp: new Date().toLocaleString('es-CL'), files: [{ ok: true, name: 'INSPECTOR_COMERCIAL_DEMO.xlsx' }, { ok: true, name: 'NOTIFICACIONES_DEUDAS_DEMO.xlsx' }] }; },
      apiListExcelBackups: function() { return []; },
      apiToggleDailyBackup: function(enable) { return { ok: true, message: enable ? 'Respaldo diario activado (23:00 hrs).' : 'Respaldo diario desactivado.' }; }
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
      apiToggleDailyBackup: function(enable) { return callLiveGas('toggleDailyBackup', { enable: enable }); }
    };

    function makeRunner(okCb, errCb) {
      var runner = {
        withSuccessHandler: function(fn) { return makeRunner(fn, errCb); },
        withFailureHandler: function(fn) { return makeRunner(okCb, fn); }
      };
      Object.keys(mockServer).forEach(function(fnName) {
        runner[fnName] = function() {
          var args = Array.prototype.slice.call(arguments);
          if (LIVE_API_URL && LIVE_MAP[fnName]) {
            LIVE_MAP[fnName].apply(null, args)
              .then(function(res) { if (okCb) okCb(res); })
              .catch(function(err) { if (errCb) errCb(err); });
          } else {
            setTimeout(function() {
              try {
                var res = mockServer[fnName].apply(null, args);
                if (okCb) okCb(res);
              } catch (e) {
                if (errCb) errCb(e);
              }
            }, 90);
          }
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
