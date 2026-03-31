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
   - enviar correos con los archivos generados
2. Proceso 2, seguimiento:
   - leer correos de respuesta
   - interpretar la respuesta recibida
   - ejecutar acciones posteriores en HDA o en el flujo administrativo

## Estado actual

Hoy el proyecto ya hace esto para el Proceso 1:

- login a HDA
- navegacion a `Payments`
- lectura de la tabla sin depender del viewport
- filtrado de tickets `OneTime Check`
- apertura de tickets
- extraccion de datos del detalle
- aplicacion de reglas de negocio
- validacion de tickets
- procesamiento de todos los tickets visibles de la corrida
- generacion de CSVs AP15 agrupados por `VendorNum` y `Currency`
- generacion de resumen legible para humanos por corrida

Pendiente principal:

- rechazo de tickets invalidos en HDA
- envio de correos del Proceso 1
- implementacion del Proceso 2
- extraccion real de `Created Date` desde HDA

## Punto de entrada

El entrypoint actual del proyecto es:

```powershell
python -m src.main
```

Hoy `src.main` ejecuta el Proceso 1:

```powershell
python -m src.hda_web.ticket_processing
```

## Estructura

```text
src/
  common/            configuracion, logger, modelos y contexto de corrida
  hda_web/           automatizacion del portal HDA
  processing/        transformaciones y reglas de negocio
  validation/        espacio reservado para motor de validacion futuro
  excel_builder/     generacion de archivos AP15 CSV
  mailer/            envio y lectura de correos
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

Ejemplos:

- `AP15_900010_USD_process_all_tickets_20260326_112115.csv`
- `log_summary_20260326_112115.txt`

## Configuracion

Usa `config/.env.example` como base para crear `config/.env.local`.

Variables importantes:

- `HDA_URL`
- `HDA_USERNAME`
- `HDA_PASSWORD`
- `BROWSER_KEEP_OPEN`
- `OUTPUT_DIR`
- `LOG_DIR`
- `MAIL_PRIMARY_RECIPIENT`
- `MAIL_SECONDARY_RECIPIENT`

## Programacion esperada

Referencia operativa actual:

- Proceso 1: una vez al dia alrededor de las 03:00 o 04:00
- Proceso 2: aproximadamente 90 minutos despues de iniciar el Proceso 1

## Referencias legacy

La app anterior basada en PDF se conserva solo como referencia historica en:

- `docs/reference/previous_app.py`
