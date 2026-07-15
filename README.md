# AjuTroquel

Aplicacion web en Google Apps Script para flujo de ajuste de troquel:

- Captura FARO previa y Post FARO.
- OCR de reportes dimensionales.
- Captura de parametros de prensa actuales y nuevos.
- Captura de calzas actuales y nuevas por numero de parte.
- Reporte comparativo antes/despues con PDF.

## Archivos principales

- `Code.gs`: logica de Apps Script, integracion con Google Sheets, OCR con Drive API y generacion de datos.
- `Index.html`: interfaz web completa.
- `appsscript.json`: manifiesto del proyecto Apps Script.

## Requisitos en Google Apps Script

1. Crear o abrir el proyecto de Apps Script.
2. Copiar `Code.gs`, `Index.html` y `appsscript.json`.
3. Activar el servicio avanzado de Google Drive:
   - Apps Script > Servicios > Drive API.
   - Tambien habilitar Google Drive API en Google Cloud si el editor lo solicita.
4. Verificar en `Code.gs` que `SPREADSHEET_ID` apunte al archivo correcto de Google Sheets.
5. Implementar como aplicacion web.

## Base de datos

El sistema usa el Spreadsheet configurado en `SPREADSHEET_ID` y espera hojas para los numeros de parte `2071154` y `2071155`, ademas de hojas auxiliares para prensa, calzas, ajustes y reportes.

## Autor

Diseno y desarrollo: Diego Crhistian Fernandez Hdez
