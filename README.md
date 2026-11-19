# INSPECTOR COMERCIAL — Google Apps Script Web App

Migración completa de la aplicación **INSPECTOR COMERCIAL** (Dirección Seguridad Pública Municipalidad de Peñaflor) desde AppSheet (`INSPECTORCOMERCIAL-286879258`) hacia **Google Apps Script**, conectada directamente a tus planillas originales de Google Sheets, con **carpetas independientes en Google Drive**, **respaldo automático en Excel (`.xlsx`)** y **despliegue continuo desde GitHub**.

---

## Estructura del Proyecto

```text
inspector-comercial/
├── .clasp.json                              # Vínculo con el proyecto de Apps Script
├── .gitignore
├── .github/
│   └── workflows/
│       └── deploy-apps-script.yml           # Auto-deploy a Apps Script al hacer push a main
└── src/
    ├── appsscript.json                      # Manifiesto y permisos OAuth (Chile: America/Santiago)
    ├── Config.gs                            # IDs de planillas, 5 plantillas PDF y esquema de 16 tablas
    ├── DriveManager.gs                      # Carpetas independientes en Drive (FOTOS, FIRMAS, PDFs, EXCEL)
    ├── PdfService.gs                        # Réplica de los 5 Bots de generación automática de PDF
    ├── BackupExcel.gs                       # Exportación .xlsx manual y automática diaria (23:00 hrs)
    ├── DataService.gs                       # Motor CRUD, fórmulas virtuales, Dashboard y Mapa GPS
    ├── Code.gs                              # Punto de entrada Web App (doGet) y API
    ├── Index.html                           # Interfaz principal responsive (móvil + escritorio)
    ├── Styles.html                          # Estilos tema claro (#0b7bde, Roboto)
    └── AppScript.html                       # Lógica SPA (Menús, Formularios, Fotos, Firma, GPS, Gráficos)
```

---

## Carpetas Independientes en Google Drive

Al abrir la app por primera vez, crea automáticamente esta estructura limpia e independiente en tu Google Drive:

```text
INSPECTOR_COMERCIAL_DRIVE/
├── FOTOS/
│   ├── COMERCIO/<ID_NombreLocal>/
│   ├── FUERA DE HORA/<ID_NombreLocal>/
│   ├── GIGANTOGRAFIA/<ID_NombreLocal>/
│   ├── ASISTENCIA/<ID_NombreDueño>/
│   ├── SEMANA PEÑAFLORINA/<ID>/
│   └── REPORTE/<ID>/
├── FIRMAS/
│   └── <TABLA>/<ID_NombreLocal>/
├── DOCUMENTOS_PDF/
│   ├── REPORTE DE FISCALIZACION COMERCIAL DIARIA/
│   ├── REPORTES FUERA DE HORARIO ESTABLECIDO/
│   ├── INFORME DE PUBLICIDAD/
│   ├── FISCALIZACIONES MOROSAS/
│   └── REPORTES A JEFATURAS/
└── RESPALDOS_EXCEL/
    ├── INSPECTOR_COMERCIAL_YYYY-MM-DD_HHmm.xlsx
    └── NOTIFICACIONES_DEUDAS_YYYY-MM-DD_HHmm.xlsx
```

---

## Cómo conectar con tu GitHub y publicar en Apps Script

### 1. Crear el proyecto en Google Apps Script (cuenta `cristopher.maciel@penaflor.cl`)
1. Ingresá a [script.google.com](https://script.google.com) con `cristopher.maciel@penaflor.cl`.
2. Creá un nuevo proyecto llamado **INSPECTOR COMERCIAL**.
3. En **Configuración del proyecto** (ícono engranaje) copiá el **ID de secuencia de comandos** (`SCRIPT_ID`).
4. También activá la **API de Google Apps Script** en [script.google.com/home/usersettings](https://script.google.com/home/usersettings).

### 2. Subir este proyecto a tu repositorio de GitHub
1. Creá un repositorio en tu GitHub (por ejemplo `inspector-comercial-apps-script`).
2. Subí el contenido de esta carpeta `inspector-comercial/` a la rama `main`.

### 3. Configurar el despliegue automático desde GitHub (GitHub Actions + `clasp`)
En tu PC (con Node.js instalado):
```powershell
npm install -g @google/clasp
clasp login
```
Esto genera el archivo `C:\Users\Usuario\.clasprc.json`. Luego en tu repositorio de GitHub, andá a **Settings → Secrets and variables → Actions** y agregá:
- `SCRIPT_ID`: el ID de secuencia de comandos de tu proyecto Apps Script.
- `CLASPRC_JSON`: el contenido completo de tu archivo `.clasprc.json`.
- *(Opcional)* `DEPLOYMENT_ID`: el ID de implementación Web App una vez publicada por primera vez.

Cada vez que hagas un cambio y lo subas a `main` en GitHub, GitHub Actions actualizará tu Apps Script automáticamente.

### 4. Primera publicación como Web App
En el editor de Apps Script:
1. Hacé clic en **Implementar → Nueva implementación → Aplicación web**.
2. **Ejecutar como:** `Yo (cristopher.maciel@penaflor.cl)`.
3. **Quién tiene acceso:** `Cualquier usuario de penaflor.cl` (recomendado para usar el filtro de correos municipales) o `Cualquier usuario`.
4. Aceptá los permisos de Google Sheets, Drive y Docs. ¡Listo!
