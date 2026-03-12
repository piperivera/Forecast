# Forecast 2026 (Provexpress)

Dashboard comercial para seguimiento de forecast 2026. Es un sitio estatico que
se autentica con Microsoft 365 (MSAL) y carga archivos Excel desde SharePoint.

## Funcionalidades
- Carga automatica desde SharePoint (Microsoft Graph).
- Vistas por rol: gerencia, director y ejecutivo.
- Panel de cambio de vista habilitado solo para `especialista.preventa`.

## Origen de datos
Ruta esperada en SharePoint (Documentos compartidos):

```
COMERCIAL/FORECAST 2026/Grupo [Director]/[Ejecutivo].xlsx
```

Reglas del archivo:
- Se toma la primera hoja que contenga "Gerencia" o "Comercial".
- La fila de encabezados se detecta por la columna "CLIENTE".
- Los nombres de columnas se normalizan internamente (por ejemplo MONTO VENTA
  CLIENTE, MONEDA 2, FECHA DIA/MES/ANO, TRM REFERENCIA, LINEA DE PRODUCTO).

## Configuracion clave
- `src/scripts/auth.js`: MSAL, roles y mapa correo -> nombre de Excel.
- `src/scripts/main.js`: carga de archivos, render y filtros.

## Ejecucion local
Abrir `index.html` en un navegador con acceso a internet. Requiere cuentas
corporativas para autenticacion en Microsoft 365.

## Notas
- No hay pruebas automatizadas.
- Si un ejecutivo no ve datos, el nombre del Excel debe coincidir con el mapa
  de correos o con el nombre del archivo en SharePoint.
