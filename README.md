# HDA Automation

Automatizacion de HDA orientada a ejecucion diaria en una VM Windows.

## Objetivo

El proyecto se esta construyendo en 2 procesos:

1. Proceso 1, madrugada:
   - iniciar sesion en HDA
   - entrar a `Payments`
   - localizar tickets `OneTime Check`
   - abrir y procesar cada ticket
   - validar datos con reglas de negocio
   - generar archivos AP15 en formato CSV
   - validar los archivos antes del envio
   - enviar correos con los archivos generados
2. Proceso 2, seguimiento:
   - leer correos de respuesta
   - interpretar la respuesta recibida
   - ejecutar acciones posteriores en HDA o en el flujo administrativo

## Estado actual

Hoy el proyecto ya hace esto para el Proceso 1:

- login a HDA
- navegacion a `Payments`
- lectura del grid sin depender del viewport
- paginacion del grid cuando hay mas de 50 tickets
- filtrado de tickets `OneTime Check`
- apertura de tickets por menu contextual
- extraccion de datos del detalle
- extraccion de `Created Date` desde la columna `Fecha` del grid
- aplicacion de reglas de negocio
- validacion local de tickets
- procesamiento de todos los tickets visibles de la corrida
- generacion de CSVs AP15 agrupados por `VendorNum`, `Currency` y grupo funcional
- clasificacion de archivos en `FMS` y `AFS`
- generacion de resumen legible para humanos por corrida
- envio SMTP de CSVs por grupo (`FMS` / `AFS`)
- envio SMTP del summary
- envio SMTP de correo de error si la corrida falla
- validacion SAP del/los CSV(s) integrada en el flujo (Test Mode)
- suspension automatica de tickets invalidos en HDA
- inspeccion automatica de resultados SAP y mapeo a tickets

Pendiente principal:

- manejo proactivo de alertas "Error inesperado" en HDA
- implementacion del Proceso 2 (Seguimiento de correos)

## Punto de entrada

El entrypoint actual del proyecto es:

```powershell
python -m src.main
```

Hoy `src.main` ejecuta el Proceso 1:

```powershell
python -m src.hda_web.ticket_processing
```

Prueba aislada del login/base SAP:

```powershell
python -m src.sap.client
```

## Flujo actual del Proceso 1

1. Login a HDA.
2. Apertura de `Payments`.
3. Lectura completa del grid, incluyendo paginacion si existe una segunda pagina.
4. Seleccion de tickets `OneTime Check`.
5. Apertura y parseo ticket por ticket.
6. Aplicacion de reglas de negocio y validacion local.
7. Separacion de tickets validos e invalidos (suspendidos localmente).
8. Generacion de CSVs AP15 candidatos.
9. Validacion SAP de cada CSV (TCode ZFIN_AP_NONPO_LUCY4).
10. Mapeo de errores SAP -> Tickets HDA.
11. Ejecucion de suspension automatica en HDA para tickets fallidos.
12. Generacion de `log_summary` final.
13. Envio de correos del summary y de los CSVs limpios por grupo.

## Reglas de negocio ya implementadas

- `Invoice Number`:
  - usa `Invoice Number`
  - si viene vacio, usa el `Id` del ticket
- `Invoice Date`:
  - usa `Invoice Date`
  - si viene vacio, usa `Created` tomado del grid
- `VendorNum`:
  - `E100 -> 900000`
  - `1000/2000 -> 900010`
  - resto -> `8000001`
- `Cost Center`:
  - quita guiones y espacios
  - si queda numerico, lo completa a 10 digitos
- `Account -> Profit Center / Cost Center`:
  - si `Account` empieza con `11`, `12`, `13`, `P1`, `P2`, `P3`, el centro se escribe en `Profit Center 10 DIGITS`
  - si `Account` empieza con `14`, `15`, `16`, `P4`, `P5`, `P6`, el centro se escribe en `Cost Center 10 DIGITS`
- `Mail Group`:
  - `FMS`: `1000`, `2000`, `E100`
  - `AFS`: companias restantes

## Estructura

```text
src/
  common/            configuracion, logger, modelos y contexto de corrida
  hda_web/           automatizacion del portal HDA
  processing/        transformaciones y reglas de negocio
  validation/        espacio reservado para motor de validacion futuro
  excel_builder/     generacion de archivos AP15 CSV
  mailer/            envio y lectura de correos
  sap/               login base y helpers para SAP GUI Scripting
  orchestrator/      espacio reservado para flujos por proceso o agenda
  pdf_processor/     legado o utilidades futuras si vuelve a usarse PDF
config/              variables de entorno y ejemplos
docs/                documentacion del proceso, reglas y referencias legacy
runtime/
  downloads/         descargas temporales
  outputs/           salidas por fecha, CSVs y summaries
  logs/              logs tecnicos por corrida
  evidence/          evidencia opcional si se reactiva
tests/               pruebas
```

## Salidas generadas

Cada corrida del Proceso 1 genera archivos dentro de:

```text
runtime/outputs/YYYYMMDD/
```

Ejemplos actuales:

- `AP15_FMS_900010_USD_process_all_tickets_20260407_145425.csv`
- `AP15_FMS_900010_CAD_process_all_tickets_20260407_145425.csv`
- `log_summary_20260407_145425.txt`

## Correos

Hoy la automatizacion ya puede enviar por SMTP:

- archivos `FMS`
- archivos `AFS`
- summary humano
- correo de error cuando una corrida falla

Los destinatarios se controlan por variables de entorno, con soporte para:

- `MAIL_TEST_RECIPIENT`
- `MAIL_FMS_RECIPIENT`
- `MAIL_AFS_RECIPIENT`
- `MAIL_SUMMARY_RECIPIENT`
- `MAIL_ERROR_RECIPIENT`
- `MAIL_BCC_RECIPIENT`

## Configuracion

Usa `config/.env.example` como base para crear `config/.env.local`.

Variables importantes:

- `HDA_URL`
- `HDA_USERNAME`
- `HDA_PASSWORD`
- `BROWSER_KEEP_OPEN`
- `OUTPUT_DIR`
- `LOG_DIR`
- `MAIL_TEST_RECIPIENT`
- `MAIL_FMS_RECIPIENT`
- `MAIL_AFS_RECIPIENT`
- `MAIL_SUMMARY_RECIPIENT`
- `MAIL_ERROR_RECIPIENT`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `SMTP_SENDER`
- `SAP_USERNAME`
- `SAP_PASSWORD`
- `SAP_CONNECTION_NAME`
- `SAP_CLIENT`
- `SAP_LANGUAGE`
- `SAP_EXECUTABLE_PATH`

SAP:

Ya existe una integracion tecnica completa en `src/sap/client.py` para:

- abrir SAP Logon y login automatico
- abrir TCode `ZFIN_AP_NONPO_LUCY4`
- carga automatica del archivo CSV generado
- llenado de formulario (Posting Date, Company Code, Separator)
- ejecucion en modo `Test` (Mandatorio)
- lectura de la tabla de resultados e inspeccion de mensajes
- mapeo de errores de SAP a tickets especificos de HDA para su suspension

## Programacion esperada

Referencia operativa actual:

- Proceso 1: una vez al dia alrededor de las 03:00 o 04:00
- Proceso 2: aproximadamente 90 minutos despues de iniciar el Proceso 1

## Referencias legacy

La app anterior basada en PDF se conserva solo como referencia historica en:

- `docs/reference/previous_app.py`
