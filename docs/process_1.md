# Process 1

## Objetivo

Procesar todos los tickets `OneTime Check` de HDA y preparar la salida AP15.

## Estado actual

Implementado:

- login
- lectura de grid
- apertura de tickets
- extraccion
- transformacion
- validacion
- agrupacion por `VendorNum` y `Currency`
- generacion de CSVs
- summary legible por corrida

Pendiente:

- rechazo de tickets invalidos en HDA
- envio de correos

## Ejecucion

```powershell
python -m src.main
```

## Salidas

- `runtime/logs/process_all_tickets_*.log`
- `runtime/outputs/YYYYMMDD/AP15_*.csv`
- `runtime/outputs/YYYYMMDD/log_summary_*.txt`
