/**
 * ============================================================================
 * INSPECTOR COMERCIAL - CONFIGURACIÓN GENERAL Y ESQUEMA DE TABLAS
 * Dirección Seguridad Pública Municipalidad de Peñaflor
 * Migrado desde AppSheet (INSPECTORCOMERCIAL-286879258) a Google Apps Script
 * ============================================================================
 */

const APP_CONFIG = {
  APP_NAME: 'Departamento de Inspectores Comerciales',
  SHORT_NAME: 'INSPECTOR COMERCIAL',
  SUBTITLE: 'Seguridad Pública Municipalidad de Peñaflor',
  CREATOR: 'CREADOR CRISTOPHER MACIEL',
  CONTACT: 'CONTACTO 1410',
  PRIMARY_COLOR: '#0c347d',
  ACCENT_COLOR: '#0b7bde',
  LOGO_URL: 'https://cristophermaciel.github.io/inspector-comercial/assets/logo-penaflor.jpg',
  TIMEZONE: 'America/Santiago',

  // Ferias Libres independientes (cada una con puestos del 1 al 100)
  FERIAS_LIST: [
    'MODELO',
    'LAS PRADERAS',
    'MALLOQUITO',
    'MANUEL GONZALES',
    'LOS GUINDOS'
  ],
  FERIA_MAX_PUESTOS: 100,
  // Horario tope diario para toma de asistencia por feria (17:00 hrs Chile)
  FERIA_ATTENDANCE_CUTOFF_HOUR: 17,
  FERIA_ATTENDANCE_CUTOFF_MINUTE: 0,
  FERIA_PDF_FOLDER_NAME: 'ASISTENCIA FERIAS LIBRES',

  // Listado oficial de funcionarios / inspectores (sin entradas de EQUIPO)
  DEFAULT_FUNCIONARIOS: [
    { nombre: 'Luis Zamorano Vásquez', rol: 'Inspector Comercial' },
    { nombre: 'JHON NAVARRETE', rol: 'INSPECTOR' },
    { nombre: 'SERGIO OYARCE', rol: 'INSPECTOR' },
    { nombre: 'ENRIQUE RODRIGUEZ', rol: 'INSPECTOR' },
    { nombre: 'NICOLAS ZAMORANO', rol: 'INSPECTOR' },
    { nombre: 'LEONARDO VALENZUELA', rol: 'INSPECTOR' },
    { nombre: 'LEONARDO HERRERA', rol: 'INSPECTOR' },
    { nombre: 'PATRICIO CANCINO', rol: 'INSPECTOR' },
    { nombre: 'TAMARA MARTINEZ', rol: 'INSPECTOR' }
  ],

  // Hojas de cálculo originales de AppSheet
  SPREADSHEETS: {
    MAIN: '1jsgsXgkJ2UfW9cZhbw02HUK795lyYjB_uxPwLizyp7I',
    DEUDAS: '1X0dTJEi6_OCznXFTEwc571DaYWFTkE_LL3AnwR0R9B4'
  },

  // Carpeta raíz independiente en Google Drive para esta aplicación
  DRIVE_ROOT_FOLDER_NAME: 'INSPECTOR_COMERCIAL_DRIVE',
  DRIVE_SUBFOLDERS: {
    FOTOS: 'FOTOS',
    FIRMAS: 'FIRMAS',
    DOCUMENTOS_PDF: 'DOCUMENTOS_PDF',
    RESPALDOS_EXCEL: 'RESPALDOS_EXCEL'
  },

  // Control de acceso (lista original de AppSheet)
  // Cambiar REQUIRE_AUTH_LIST a true si se despliega en dominio Google Workspace
  // con acceso restringido a usuarios autenticados.
  REQUIRE_AUTH_LIST: false,
  AUTHORIZED_USERS: [
    'cristopher.maciel@penaflor.cl',
    'central1410@penaflor.cl',
    'david.hernandez@penaflor.cl',
    'inspecciones@penaflor.cl',
    'luis.zamorano@penaflor.cl',
    'nicolas.zamorano@penaflor.cl',
    'tamara.martinez@penaflor.cl',
    'leonardo.herrera@penaflor.cl',
    'patricio.cancino@penaflor.cl'
  ],

  // 5 Bots de generación automática de PDF (plantillas Google Docs originales)
  PDF_BOTS: {
    'COMERCIO': {
      botName: 'comercio pdf',
      templateDocId: '1UvOA76n4T5eb6Wb-59dfs4pK1svzRWKekTZwhX6QBqg',
      folderName: 'REPORTE DE FISCALIZACION COMERCIAL DIARIA',
      prefixColumn: 'FECHA HORA',
      pdfColumn: 'PDF'
    },
    'FUERA DE HORA': {
      botName: 'fuera de hora pdf',
      templateDocId: '19W13bTRPNAm4AEDZ-hBDYQc7PD7_jfINqdc2bpsXzSk',
      folderName: 'REPORTES FUERA DE HORARIO ESTABLECIDO',
      prefixColumn: 'FECHA HORA',
      pdfColumn: 'PDF'
    },
    'GIGANTOGRAFIA': {
      botName: 'informe publicidad pdf',
      templateDocId: '1KyMkHQue4HGF0Lv_KBd58S7Yq7G6yySOPKK9aEm5J-k',
      folderName: 'INFORME DE PUBLICIDAD',
      prefixColumn: 'FECHA HORA',
      pdfColumn: 'PDF'
    },
    'PATENTES EN MORA': {
      botName: 'fiscalizaciones morosas pdf',
      templateDocId: '1glmlnVjrLGqwM8KBx_k98UgdOKhEUI8EkLeh9dvc1uA',
      folderName: 'FISCALIZACIONES MOROSAS',
      prefixColumn: 'NOMBRE LOCAL COMERCIAL',
      pdfColumn: ''
    },
    'REPORTE': {
      botName: 'reportes jefatura pdf',
      templateDocId: '1ypLF72jHMn7HeCSwndm_oF6fT7vl2ZIgpZlcuKnLYMM',
      folderName: 'REPORTES A JEFATURAS',
      prefixColumn: 'FECHA HORA',
      pdfColumn: 'PDF'
    }
  }
};

/**
 * Definición completa de las 16 tablas de usuario de INSPECTOR COMERCIAL.
 * Mantiene los nombres exactos de hojas y columnas de AppSheet para 100% compatibilidad.
 */
const TABLE_SCHEMAS = {
  'COMERCIO': {
    spreadsheetKey: 'MAIN',
    sheetName: 'COMERCIO',
    label: 'Fiscalización Comercial',
    icon: 'store',
    keyCol: 'ID',
    labelCol: 'NOMBRE LOCAL',
    hasPdfBot: true,
    formColumns: [
      'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'FECHA HORA', 'TIPO DE INTERVENCION', 'TIPO FISCALIZACION',
      'DIRECCION', 'TELEFONO', 'EMAIL', 'NOMBRE LOCAL',
      'NOMBRE DUEÑO O RAZON SOCIAL', 'RUT', 'PRODUCTOS QUE VENDE', 'GIRO',
      'DOCUMENTO FISCALIZADO', 'CUENTA CON PATENTE', 'EXHIBE PATENTE',
      'PERMISO AL DIA', 'OBSERVACIONES FISCALIZACION', 'UBICACION',
      'INTERVENCION', 'OTROS ANTECEDENTES', 'FOTO 1', 'FOTO 2', 'FOTO 3', 'FIRMA'
    ],
    tableColumns: [
      'FECHA HORA', 'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'NOMBRE LOCAL', 'RUT', 'DIRECCION',
      'CUENTA CON PATENTE', 'PERMISO AL DIA', 'PDF'
    ],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'INSPECTOR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'FECHA HORA', type: 'DateTime', initial: 'NOW' },
      { name: 'TIPO DE INTERVENCION', type: 'EnumList', values: ['Fiscalización de rutina', 'Denuncia ciudadana', 'Operativo conjunto', 'Inspección por reincidencia', 'Control nocturno', 'Control de patentes', 'Otro'] },
      { name: 'TIPO FISCALIZACION', type: 'EnumList', values: ['Inspección', 'Infracción cursada', 'Clausura', 'Notificación', 'Citación al Juzgado', 'Advertencia verbal', 'Decomiso', 'Otro'] },
      { name: 'DIRECCION', type: 'LongText' },
      { name: 'NOMBRE LOCAL', type: 'Text' },
      { name: 'NOMBRE DUEÑO O RAZON SOCIAL', type: 'Text' },
      { name: 'RUT', type: 'Text' },
      { name: 'PRODUCTOS QUE VENDE', type: 'EnumList', values: ['Alimentos', 'Botillería', 'Restaurante', 'Comida rápida', 'Ferretería', 'Ropa', 'Minimarket', 'Taller mecánico', 'Servicios', 'Otro'] },
      { name: 'GIRO', type: 'Enum', values: ['AMBULANTE', 'COMERCIAL', 'PERMISO PROVISORIO', 'INDUSTRIAL', 'ALCOHOLES', 'FERIAS LIBRES VTA AL DETALLE', 'MICROEMPRESA FAMILIAR', 'OTRO', 'SIN PATENTE'] },
      { name: 'DOCUMENTO FISCALIZADO', type: 'EnumList', values: ['Patente comercial', 'Permiso municipal', 'Resolución sanitaria', 'Inicio de actividades SII', 'Boletas / Facturas', 'Contrato de arriendo', 'Otro'] },
      { name: 'CUENTA CON PATENTE', type: 'Enum', values: ['SI', 'NO', 'EN TRAMITE', 'ATRASADA'] },
      { name: 'EXHIBE PATENTE', displayName: 'PATENTE VISIBLE', type: 'Enum', values: ['SI', 'NO', 'NO CORRESPONDE'] },
      { name: 'PERMISO AL DIA', type: 'Enum', values: ['SI', 'NO', 'EN TRAMITE', 'VENCIDO', 'NO APLICA'] },
      { name: 'OBSERVACIONES FISCALIZACION', type: 'LongText' },
      { name: 'UBICACION', displayName: 'UBICACIÓN GPS', type: 'Location' },
      { name: 'INTERVENCION', type: 'Text' },
      { name: 'OTROS ANTECEDENTES', type: 'LongText' },
      { name: 'FOTO 1', type: 'Image' },
      { name: 'FOTO 2', type: 'Image' },
      { name: 'FOTO 3', type: 'Image' },
      { name: 'FIRMA', type: 'Signature' },
      { name: 'EMAIL', type: 'Email' },
      { name: 'TELEFONO', type: 'Phone' },
      { name: 'PDF', type: 'File', readOnly: true }
    ]
  },

  'FUERA DE HORA': {
    spreadsheetKey: 'MAIN',
    sheetName: 'FUERA DE HORA',
    label: 'Fiscalización Fuera de Horario',
    icon: 'schedule',
    keyCol: 'ID',
    labelCol: 'NOMBRE LOCAL',
    hasPdfBot: true,
    formColumns: [
      'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'FECHA HORA', 'TIPO DE INTERVENCION', 'TIPO FISCALIZACION',
      'TELEFONO', 'EMAIL', 'DIRECCION', 'NOMBRE LOCAL',
      'NOMBRE DUEÑO O RAZON SOCIAL', 'RUT', 'PRODUCTOS QUE VENDE', 'GIRO',
      'DOCUMENTO FISCALIZADO', 'CUENTA CON PATENTE', 'EXHIBE PATENTE',
      'PERMISO AL DIA', 'OTROS DOCUMENTOS', 'OBSERVACIONES FISCALIZACION',
      'UBICACION', 'APOYOS O INTERVENCIONES', 'DESCRIBA INCIDENTE SI HUBO',
      'FOTO 1', 'FOTO 2', 'FOTO 3', 'FIRMA'
    ],
    tableColumns: [
      'FECHA HORA', 'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'NOMBRE LOCAL', 'RUT', 'DIRECCION',
      'CUENTA CON PATENTE', 'PERMISO AL DIA', 'PDF'
    ],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'INSPECTOR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'FECHA HORA', type: 'DateTime', initial: 'NOW' },
      { name: 'TIPO DE INTERVENCION', type: 'EnumList', values: ['Fiscalización de rutina', 'Denuncia ciudadana', 'Operativo conjunto', 'Inspección por reincidencia', 'Control nocturno', 'Control de patentes', 'Otro'] },
      { name: 'TIPO FISCALIZACION', type: 'EnumList', values: ['Inspección', 'Infracción cursada', 'Clausura', 'Notificación', 'Citación al Juzgado', 'Advertencia verbal', 'Decomiso', 'Otro'] },
      { name: 'DIRECCION', type: 'Location' },
      { name: 'NOMBRE LOCAL', type: 'Text' },
      { name: 'NOMBRE DUEÑO O RAZON SOCIAL', type: 'Text' },
      { name: 'RUT', type: 'Text' },
      { name: 'PRODUCTOS QUE VENDE', type: 'EnumList', values: ['Alimentos', 'Botillería', 'Restaurante', 'Comida rápida', 'Ferretería', 'Ropa', 'Minimarket', 'Taller mecánico', 'Servicios', 'Otro'] },
      { name: 'GIRO', type: 'Enum', values: ['Ambulante', 'Comercial', 'Permiso Provisorio', 'Industrial', 'Alcoholes', 'Ferias Libres', 'Micro Empresas', 'Otros', 'Sin Patente'] },
      { name: 'DOCUMENTO FISCALIZADO', type: 'Text' },
      { name: 'CUENTA CON PATENTE', type: 'Enum', values: ['SI', 'NO', 'EN TRAMITE'] },
      { name: 'EXHIBE PATENTE', type: 'Enum', values: ['si', 'no', 'no corresponde'] },
      { name: 'PERMISO AL DIA', type: 'Enum', values: ['si', 'no', 'vencido', 'no aplica'] },
      { name: 'OTROS DOCUMENTOS', type: 'Text' },
      { name: 'OBSERVACIONES FISCALIZACION', type: 'LongText' },
      { name: 'UBICACION', type: 'Location' },
      { name: 'APOYOS O INTERVENCIONES', type: 'Text' },
      { name: 'DESCRIBA INCIDENTE SI HUBO', type: 'LongText' },
      { name: 'FOTO 1', type: 'Image' },
      { name: 'FOTO 2', type: 'Image' },
      { name: 'FOTO 3', type: 'Image' },
      { name: 'FIRMA', type: 'Signature' },
      { name: 'PDF', type: 'File', readOnly: true },
      { name: 'EMAIL', type: 'Email' },
      { name: 'TELEFONO', type: 'Phone' }
    ]
  },

  'GIGANTOGRAFIA': {
    spreadsheetKey: 'MAIN',
    sheetName: 'GIGANTOGRAFIA',
    label: 'Fiscalización Publicidad / Gigantografía',
    icon: 'campaign',
    keyCol: 'ID',
    labelCol: 'NOMBRE LOCAL',
    hasPdfBot: true,
    formColumns: [
      'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'FECHA HORA', 'TIPO DE INTERVENCION', 'TIPO FISCALIZACION',
      'EMAIL', 'TELEFONO', 'DIRECCION', 'NOMBRE LOCAL',
      'NOMBRE DUEÑO O RAZON SOCIAL', 'RUT', 'PRODUCTOS QUE VENDE', 'GIRO',
      'METRAJE DE LA PUBLICIDAD', 'PUBLICIDAD LUMINOSA', 'PAGA PUBLICIDAD',
      'OBSERVACIONES FISCALIZACION', 'UBICACION', 'INTERVENCION',
      'OTROS ANTECEDENTES', 'FOTO 1', 'FOTO 2', 'FOTO 3', 'FIRMA'
    ],
    tableColumns: [
      'FECHA HORA', 'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'NOMBRE LOCAL', 'DIRECCION',
      'METRAJE DE LA PUBLICIDAD', 'PAGA PUBLICIDAD', 'PDF'
    ],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'INSPECTOR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'FECHA HORA', type: 'DateTime', initial: 'NOW' },
      { name: 'TIPO DE INTERVENCION', type: 'EnumList', values: ['Fiscalización de rutina', 'Denuncia ciudadana', 'Inspección por reincidencia', 'Operativo conjunto', 'Control nocturno', 'Control de patentes', 'Otro', 'Control de Publicidad'] },
      { name: 'TIPO FISCALIZACION', type: 'EnumList', values: ['Infraestructura', 'Infracción cursada', 'Clausura', 'Notificación', 'Citación al Juzgado', 'Advertencia verbal', 'Decomiso', 'Otro'] },
      { name: 'DIRECCION', type: 'Text' },
      { name: 'NOMBRE LOCAL', type: 'Text' },
      { name: 'NOMBRE DUEÑO O RAZON SOCIAL', type: 'Text' },
      { name: 'RUT', type: 'Text' },
      { name: 'PRODUCTOS QUE VENDE', type: 'EnumList', values: ['Alimentos', 'Botillería', 'Restaurante', 'Comida rápida', 'Ferretería', 'Ropa', 'Minimarket', 'Taller mecánico', 'Servicios', 'Otro'] },
      { name: 'GIRO', type: 'EnumList', values: ['Comercial', 'Industrial', 'Alcoholes', 'Profesional', 'Microempresa familiar', 'Permiso provisorio', 'Sin patente'] },
      { name: 'METRAJE DE LA PUBLICIDAD', type: 'Text' },
      { name: 'PUBLICIDAD LUMINOSA', type: 'EnumList', values: ['SI', 'NO'] },
      { name: 'PAGA PUBLICIDAD', type: 'EnumList', values: ['SI', 'NO', 'ATRASADA', 'VENCIDA', 'CONVENIO MUNICIPAL', 'PENDIENTE'] },
      { name: 'OBSERVACIONES FISCALIZACION', type: 'LongText' },
      { name: 'UBICACION', type: 'Location' },
      { name: 'INTERVENCION', type: 'Text' },
      { name: 'OTROS ANTECEDENTES', type: 'LongText' },
      { name: 'FOTO 1', type: 'Image' },
      { name: 'FOTO 2', type: 'Image' },
      { name: 'FOTO 3', type: 'Image' },
      { name: 'FIRMA', type: 'Signature' },
      { name: 'PDF', type: 'File', readOnly: true },
      { name: 'EMAIL', type: 'Email' },
      { name: 'TELEFONO', type: 'Phone' }
    ]
  },

  'PATENTES EN MORA': {
    spreadsheetKey: 'MAIN',
    sheetName: 'PATENTES EN MORA',
    label: 'Patentes en Mora',
    icon: 'payments',
    keyCol: 'ID',
    labelCol: 'NOMBRE LOCAL COMERCIAL',
    hasPdfBot: true,
    formColumns: [
      'FISCALIZADOR', 'INSPECTOR ACOMPAÑANTE', 'FECHA FISCALIZACION', 'NUMERO',
      'NOMBRE LOCAL COMERCIAL', 'PATENTE', 'ROL', 'DIRECCION',
      'ESTADO PATENTE', 'VALOR PATENTE', 'OBSERVACIONES', 'EMAIL', 'TELEFONO'
    ],
    tableColumns: [
      'NUMERO', 'NOMBRE LOCAL COMERCIAL', 'PATENTE', 'ROL',
      'ESTADO PATENTE', 'VALOR PATENTE', 'FISCALIZADOR', 'INSPECTOR ACOMPAÑANTE', 'FECHA FISCALIZACION'
    ],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'NUMERO', type: 'Text' },
      { name: 'NOMBRE LOCAL COMERCIAL', type: 'Text' },
      { name: 'PATENTE', type: 'Text' },
      { name: 'ROL', type: 'Text' },
      { name: 'DIRECCION', type: 'Location' },
      { name: 'ESTADO PATENTE', type: 'Enum', values: ['AL DIA', 'EN MORA'] },
      { name: 'VALOR PATENTE', type: 'Price' },
      { name: 'FECHA FISCALIZACION', type: 'DateTime', initial: 'NOW' },
      { name: 'FISCALIZADOR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'OBSERVACIONES', type: 'LongText' },
      { name: 'EMAIL', type: 'Email' },
      { name: 'TELEFONO', type: 'Phone' }
    ]
  },

  'REPORTE': {
    spreadsheetKey: 'MAIN',
    sheetName: 'REPORTE',
    label: 'Reportes a Jefatura',
    icon: 'assignment',
    keyCol: 'ID',
    labelCol: 'FUNCIONARIO',
    hasPdfBot: true,
    formColumns: [
      'FUNCIONARIO', 'INSPECTOR ACOMPAÑANTE', 'FECHA HORA', 'REPORTE',
      'A QUIEN LE INFORMA', 'UBICACION', 'FOTO 1', 'FOTO 2', 'FIRMA'
    ],
    tableColumns: ['FECHA HORA', 'FUNCIONARIO', 'INSPECTOR ACOMPAÑANTE', 'A QUIEN LE INFORMA', 'REPORTE', 'PDF'],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'FUNCIONARIO', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'FECHA HORA', type: 'DateTime', initial: 'NOW' },
      { name: 'REPORTE', type: 'LongText' },
      { name: 'A QUIEN LE INFORMA', type: 'EnumList', values: ['DAVID HERNANDEZ', 'ALVARO NAVARRO', 'JUAN PLAZA', 'CRISTOPHER MACIEL', 'CENCOO / BRAVO', 'JEFE DE TURNO'] },
      { name: 'UBICACION', type: 'Location' },
      { name: 'FOTO 1', type: 'Image' },
      { name: 'FOTO 2', type: 'Image' },
      { name: 'FIRMA', type: 'Signature' },
      { name: 'PDF', type: 'File', readOnly: true }
    ]
  },

  'ASISTENCIA': {
    spreadsheetKey: 'MAIN',
    sheetName: 'ASISTENCIA',
    label: 'Registro Ferias Libres / Feriantes',
    icon: 'storefront',
    keyCol: 'ID',
    labelCol: 'NOMBRE DUEÑO',
    hasPdfBot: false,
    defaultSort: { column: 'N° Puesto', asc: true },
    formColumns: [
      'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'FECHA', 'DIA', 'FERIA', 'N° Puesto',
      'NOMBRE DUEÑO', 'RUT DUEÑO', 'PATENTE', 'Foto patente', 'FOTO DUEÑO', 'ROL', 'GIRO',
      'FOTO CARNET AYUDANTE', 'FOTO CARNET AYUDANTE_15', 'FOTO CARNET AYUDANTE_16',
      'UBICACION', 'OBSERVACIONES'
    ],
    tableColumns: ['N° Puesto', 'CODIGO_QR', 'NOMBRE DUEÑO', 'DIA', 'FERIA', 'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'RUT DUEÑO', 'PATENTE', 'GIRO'],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'CODIGO_QR', displayName: 'CÓDIGO QR ÚNICO', type: 'Text', readOnly: true, hiddenInForm: true },
      { name: 'FECHA', type: 'DateTime', initial: 'NOW' },
      { name: 'DIA', type: 'Enum', values: ['MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO'] },
      { name: 'NOMBRE DUEÑO', type: 'Text' },
      { name: 'RUT DUEÑO', type: 'Text' },
      { name: 'PATENTE', type: 'Enum', values: ['SI POSEE', 'NO POSEE'] },
      { name: 'Foto patente', type: 'Image' },
      { name: 'OBSERVACIONES', type: 'LongText' },
      { name: 'N° Puesto', type: 'Number' },
      { name: 'FOTO DUEÑO', type: 'Image' },
      { name: 'ROL', type: 'Text' },
      { name: 'GIRO', type: 'Enum', values: ['AMBULANTE', 'COMERCIAL', 'PROVISORIO', 'INDUSTRIAL', 'ALCOHOLES', 'FERIAS LIBRE VTA AL DETALLE', 'MICRO EMPRESA', 'OTRO'] },
      { name: 'FERIA', type: 'Enum', values: ['MODELO', 'LAS PRADERAS', 'MALLOQUITO', 'MANUEL GONZALES', 'LOS GUINDOS'] },
      { name: 'FOTO CARNET AYUDANTE', displayName: 'Foto Ayudante 1', type: 'Image' },
      { name: 'FOTO CARNET AYUDANTE_15', displayName: 'Foto Ayudante 2', type: 'Image' },
      { name: 'FOTO CARNET AYUDANTE_16', displayName: 'Foto Ayudante 3', type: 'Image' },
      { name: 'UBICACION', type: 'Location' },
      { name: 'PDF', type: 'File', readOnly: true },
      { name: 'VISTA', type: 'Text', hiddenInForm: true },
      { name: 'INSPECTOR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' }
    ]
  },

  'SEMANA PEÑAFLORINA': {
    spreadsheetKey: 'MAIN',
    sheetName: 'SEMANA PEÑAFLORINA',
    label: 'Semana Peñaflorina',
    icon: 'festival',
    keyCol: 'ID',
    labelCol: 'CONTROLADO POR',
    hasPdfBot: false,
    formColumns: [
      'CONTROLADO POR', 'INSPECTOR ACOMPAÑANTE', 'FECHA HORA', 'ESTADO DE PAGO',
      'FOTO CARNET', 'FOTO VENTA', 'FOTO PATENTE', 'UBICACION', 'FIRMA', 'OBSERVACIONES'
    ],
    tableColumns: ['FECHA HORA', 'CONTROLADO POR', 'INSPECTOR ACOMPAÑANTE', 'ESTADO DE PAGO', 'OBSERVACIONES'],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'CONTROLADO POR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'FECHA HORA', type: 'DateTime', initial: 'NOW' },
      { name: 'ESTADO DE PAGO', type: 'Enum', values: ['NO PAGA', 'PAGA DESPUES', 'SOLO VIERNES', 'SOLO SABADO', 'SOLO DOMINGO', 'VIERNES Y SABADO', 'SABADO Y DOMINGO', 'TODOS LOS DIAS', 'VIERNES Y DOMINGO'] },
      { name: 'FOTO CARNET', type: 'Image' },
      { name: 'FOTO VENTA', displayName: 'FOTO QUE VENDE', type: 'Image' },
      { name: 'FOTO PATENTE', type: 'Image' },
      { name: 'UBICACION', type: 'Location' },
      { name: 'FIRMA', type: 'Signature' },
      { name: 'OBSERVACIONES', type: 'LongText' }
    ]
  },

  'FERIAS NAVIDEÑAS': {
    spreadsheetKey: 'MAIN',
    sheetName: 'FERIAS NAVIDEÑAS',
    label: 'Ferias Navideñas',
    icon: 'card_giftcard',
    keyCol: 'ID',
    labelCol: 'CONTROLADO POR',
    hasPdfBot: false,
    formColumns: [
      'CONTROLADO POR', 'INSPECTOR ACOMPAÑANTE', 'FECHA HORA', 'ESTADO DE PAGO',
      'FOTO CARNET', 'FOTO VENTA', 'FOTO PATENTE', 'UBICACION', 'FIRMA', 'OBSERVACIONES'
    ],
    tableColumns: ['FECHA HORA', 'CONTROLADO POR', 'INSPECTOR ACOMPAÑANTE', 'ESTADO DE PAGO', 'OBSERVACIONES'],
    columns: [
      { name: 'ID', type: 'Text', isKey: true, initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'CONTROLADO POR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'FECHA HORA', type: 'DateTime', initial: 'NOW' },
      { name: 'ESTADO DE PAGO', type: 'Enum', values: ['NO PAGA', 'PAGA DESPUES', 'SOLO VIERNES', 'SOLO SABADO', 'SOLO DOMINGO', 'VIERNES Y SABADO', 'SABADO Y DOMINGO', 'TODOS LOS DIAS', 'VIERNES Y DOMINGO'] },
      { name: 'FOTO CARNET', type: 'Image' },
      { name: 'FOTO VENTA', displayName: 'FOTO QUE VENDE', type: 'Image' },
      { name: 'FOTO PATENTE', type: 'Image' },
      { name: 'UBICACION', type: 'Location' },
      { name: 'FIRMA', type: 'Signature' },
      { name: 'OBSERVACIONES', type: 'LongText' }
    ]
  },

  'Hoja 1': {
    spreadsheetKey: 'DEUDAS',
    sheetName: 'Hoja 1',
    label: 'Notificaciones por Deudas',
    icon: 'notifications_active',
    keyCol: '_RowNumber',
    labelCol: 'LOCAL COMERCIAL',
    hasPdfBot: false,
    defaultSort: { column: 'DIAS RESTANTES', asc: true },
    formColumns: [
      'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'N° NOTIFICACION', 'LOCAL COMERCIAL',
      'DIRECCION', 'FECHA DE NOTIFICACION', 'FECHA DE TERMINO'
    ],
    tableColumns: ['N° NOTIFICACION', 'INSPECTOR', 'INSPECTOR ACOMPAÑANTE', 'LOCAL COMERCIAL', 'DIRECCION', 'FECHA DE NOTIFICACION', 'FECHA DE TERMINO', 'DIAS RESTANTES'],
    columns: [
      { name: 'N° NOTIFICACION', type: 'Number' },
      { name: 'DIRECCION', type: 'Text' },
      { name: 'LOCAL COMERCIAL', type: 'LongText' },
      { name: 'FECHA DE NOTIFICACION', type: 'DateTime', initial: 'NOW' },
      { name: 'FECHA DE TERMINO', type: 'Date', initial: 'TODAY' },
      { name: 'DIAS RESTANTES', type: 'Text', readOnly: true, formula: 'DAYS_REMAINING' },
      { name: 'INSPECTOR', displayName: 'INSPECTOR A CARGO', type: 'Ref', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' },
      { name: 'INSPECTOR ACOMPAÑANTE', displayName: 'INSPECTOR ACOMPAÑANTE DEL PROCEDIMIENTO', type: 'RefList', refTable: 'FUNCIONARIO', refCol: 'NOMBRE' }
    ]
  },

  'FUNCIONARIO': {
    spreadsheetKey: 'MAIN',
    sheetName: 'FUNCIONARIO',
    label: 'Funcionarios / Inspectores',
    icon: 'badge',
    keyCol: 'NOMBRE',
    labelCol: 'NOMBRE',
    hasPdfBot: false,
    tableColumns: ['ID', 'NOMBRE', 'ROL'],
    columns: [
      { name: 'ID', type: 'Text', initial: 'UNIQUEID' },
      { name: 'NOMBRE', type: 'Text', isKey: true },
      { name: 'IMAGEN', type: 'Image' },
      { name: 'ROL', type: 'Text' }
    ]
  },

  'DOCUMENTOS': {
    spreadsheetKey: 'MAIN',
    sheetName: 'DOCUMENTOS',
    label: 'Documentos / Drive',
    icon: 'folder_shared',
    keyCol: 'IR',
    labelCol: 'IR',
    hasPdfBot: false,
    tableColumns: ['IR', 'LINK'],
    columns: [
      { name: 'ID', type: 'Text', initial: 'UNIQUEID', hiddenInForm: true },
      { name: 'IR', type: 'Text', isKey: true },
      { name: 'LINK', type: 'Url' },
      { name: 'IMAGEN', type: 'Image' }
    ]
  },

  // Tablas de navegación de menú heredadas de AppSheet
  'MENU': { spreadsheetKey: 'MAIN', sheetName: 'MENU', keyCol: 'ID MENU', isNavTable: true, columns: [{ name: 'ID', type: 'Text' }, { name: 'ID MENU', type: 'Text' }, { name: 'IMAGEN', type: 'Image' }, { name: 'VISTA', type: 'Text' }] },
  'SUBMENU': { spreadsheetKey: 'MAIN', sheetName: 'SUBMENU', keyCol: 'ID SUBMENU', isNavTable: true, columns: [{ name: 'ID', type: 'Text' }, { name: 'ID SUBMENU', type: 'Text' }, { name: 'IMAGEN', type: 'Image' }, { name: 'VISTA', type: 'Text' }] },
  'SUBMENU 2': { spreadsheetKey: 'MAIN', sheetName: 'SUBMENU 2', keyCol: 'ID SUBMENU', isNavTable: true, columns: [{ name: 'ID', type: 'Text' }, { name: 'ID SUBMENU', type: 'Text' }, { name: 'IMAGEN', type: 'Image' }, { name: 'VISTA', type: 'Text' }] },
  'SUBMENU 3': { spreadsheetKey: 'MAIN', sheetName: 'SUBMENU 3', keyCol: 'ID SUBMENU', isNavTable: true, columns: [{ name: 'ID', type: 'Text' }, { name: 'ID SUBMENU', type: 'Text' }, { name: 'IMAGEN', type: 'Image' }, { name: 'VISTA', type: 'Text' }] },
  'LISTADOS': { spreadsheetKey: 'MAIN', sheetName: 'LISTADOS', keyCol: 'ID TABLA', isNavTable: true, columns: [{ name: 'ID', type: 'Text' }, { name: 'ID TABLA', type: 'Text' }, { name: 'IMAGEN', type: 'Image' }, { name: 'VISTA', type: 'Text' }] },
  'FERIAS LIBRES': { spreadsheetKey: 'MAIN', sheetName: 'FERIAS LIBRES', keyCol: 'ID FERIAS', isNavTable: true, columns: [{ name: 'ID', type: 'Text' }, { name: 'ID FERIAS', type: 'Text' }, { name: 'IMAGEN', type: 'Image' }, { name: 'VISTA', type: 'Text' }] }
};

/**
 * Mapeo de nombres de vistas de AppSheet -> Acciones/Vistas en la Web App
 * Permite que los menús leídos desde Google Sheets naveguen exactamente igual que en AppSheet.
 */
const APPSHEET_VIEW_MAP = {
  'MENU': { kind: 'menu', table: 'MENU', title: 'Menú Principal' },
  'SUBMENU': { kind: 'menu', table: 'SUBMENU', title: 'Submenú' },
  'SUB MENU 2': { kind: 'menu', table: 'SUBMENU 2', title: 'Submenú 2' },
  'SUB MENU 4': { kind: 'menu', table: 'SUBMENU 3', title: 'Submenú 3' },
  'LISTAS': { kind: 'menu', table: 'LISTADOS', title: 'Listados' },
  'FERIAS LIBRES': { kind: 'menu', table: 'FERIAS LIBRES', title: 'Ferias Libres' },
  'DRIVE': { kind: 'links', table: 'DOCUMENTOS', title: 'Documentos Drive' },
  'FIS. COMERCIAL': { kind: 'form', table: 'COMERCIO', title: 'Fiscalización Comercial' },
  'Fis. Fuera de Horarios': { kind: 'form', table: 'FUERA DE HORA', title: 'Fiscalización Fuera de Horario' },
  'FIS. PUBLICIDAD': { kind: 'form', table: 'GIGANTOGRAFIA', title: 'Fiscalización Publicidad' },
  'PATENTES EN DEUDA': { kind: 'form', table: 'PATENTES EN MORA', title: 'Patentes en Mora' },
  'REPORTE': { kind: 'form', table: 'REPORTE', title: 'Nuevo Reporte a Jefatura' },
  'REGISTRO FERIA': { kind: 'form', table: 'ASISTENCIA', title: 'Registro Feria Libre' },
  'LISTA EDITABLE COMERCIO': { kind: 'table', table: 'COMERCIO', title: 'Lista Editable Comercio' },
  'COMERCIO FUERA DE HORA': { kind: 'table', table: 'FUERA DE HORA', title: 'Comercio Fuera de Hora' },
  'PUBLICIDAD': { kind: 'table', table: 'GIGANTOGRAFIA', title: 'Publicidad / Gigantografía' },
  'LISTADO DE FERIANTES': { kind: 'table', table: 'ASISTENCIA', title: 'Listado de Feriantes' },
  'SEGUIMIENTOS DE MOROSOS': { kind: 'table', table: 'PATENTES EN MORA', title: 'Seguimiento de Morosos' },
  'DOCUMENTOS FISCALIZADOS': { kind: 'table', table: 'COMERCIO', title: 'Documentos Fiscalizados' },
  'SEMANA PEÑAFLORINA': { kind: 'table', table: 'SEMANA PEÑAFLORINA', title: 'Semana Peñaflorina' },
  'FERIAS NAVIDEÑAS': { kind: 'table', table: 'FERIAS NAVIDEÑAS', title: 'Ferias Navideñas' },
  'New View': { kind: 'table', table: 'Hoja 1', title: 'Notificaciones por Deudas' },
  'FISCALIZACIONES COMERCIALES': { kind: 'map', table: 'COMERCIO', title: 'Mapa de Fiscalizaciones' },
  'DASHBOARD': { kind: 'dashboard', title: 'Dashboard Estadístico' }
};
