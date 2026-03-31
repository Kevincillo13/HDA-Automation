# Business Rules

## Reglas activas hoy

Transformaciones principales:

- `CompanyCode` se obtiene del campo `Company`
- `VendorNum`:
  - `E100` -> `900000`
  - `1000` o `2000` -> `900010`
  - default -> `8000001`
- `Invoice Number`:
  - usar `Invoice Number`
  - si viene vacio, usar `ticket_id`
- `Invoice Date`:
  - usar `Invoice Date`
  - si viene vacia, usar `Created`
  - normalizar a `MM/DD/YYYY`
- `Amount`:
  - limpiar simbolos y convertir a numero
- `City`, `State`, `Zip`, `Country`:
  - derivarlos desde `City/State`
  - soportar formatos USA y Canada
- `Country`:
  - preferir lo inferido de la direccion
  - si no es claro, usar la moneda como apoyo

Validaciones activas:

- `Amount` no debe ser `0`
- `Zip` no debe venir vacio
- `Cost/Profit Center` no debe venir vacio
- `GL Account` no debe venir vacio
- si `GL Account` es numerico, debe tener 10 digitos
- si `GL Account` empieza con `P`, `VendorNum` debe ser `8000001`

## Reglas de salida AP15

- Se genera un CSV por grupo `VendorNum + Currency`
- Nombre actual:
  - `AP15_{VendorNum}_{Currency}_{run_id}.csv`
- Salida actual:
  - `runtime/outputs/YYYYMMDD/`

## Pendientes por confirmar

- De donde obtener `Created Date` real en HDA
- Regla exacta de rechazo por tipo de error
- Texto final del rechazo en HDA
- Destinatarios finales por tipo de archivo o compania
- Reglas del Proceso 2
