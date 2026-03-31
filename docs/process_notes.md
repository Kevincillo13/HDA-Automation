# Process Notes

## Vision actual

La automatizacion se dividira en 2 procesos diarios:

1. Proceso 1:
   - corre de madrugada
   - procesa tickets `OneTime Check`
   - genera AP15 CSV
   - envia correos
2. Proceso 2:
   - corre despues del Proceso 1
   - revisa correos de respuesta
   - ejecuta acciones posteriores

## Flujo real del Proceso 1

1. Abrir HDA e iniciar sesion.
2. Entrar a la vista `Payments`.
3. Leer la tabla desde el DOM para no depender del viewport.
4. Filtrar tickets cuyo `PAYMENT METHOD` sea `OneTime Check`.
5. Abrir cada ticket.
6. Extraer datos del detalle.
7. Aplicar reglas de negocio.
8. Validar cada ticket.
9. Separar tickets validos e invalidos.
10. Generar CSVs AP15 agrupados por `VendorNum` y `Currency`.
11. Generar un summary legible para humanos.
12. Pendiente: rechazar invalidos y enviar correos.

## Salidas actuales

- Logs tecnicos en `runtime/logs`
- CSVs AP15 en `runtime/outputs/YYYYMMDD`
- `log_summary` por corrida en `runtime/outputs/YYYYMMDD`

## Pendientes abiertos

- Obtener `Created Date` real desde HDA
- Implementar rechazo de tickets invalidos
- Implementar envio de correos del Proceso 1
- Diseñar e implementar el Proceso 2
