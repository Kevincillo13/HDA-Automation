# Environment Setup

## Desarrollo local

Proyecto pensado para desarrollarse primero en laptop y luego ejecutarse en VM Windows.

## Recomendacion inicial

1. Crear entorno virtual de Python.
2. Instalar dependencias de `requirements.txt`.
3. Copiar `config/.env.example` a `config/.env.local`.
4. Llenar credenciales de HDA y correo.
5. Probar `python -m src.main`.

## Consideraciones de ejecucion

- El flujo debe funcionar en VM Windows sin depender del viewport.
- La evidencia visual esta desactivada por defecto.
- La observabilidad principal hoy vive en logs y summaries.

## Programacion esperada

- Proceso 1: 03:00 o 04:00
- Proceso 2: 90 minutos despues del inicio del Proceso 1
