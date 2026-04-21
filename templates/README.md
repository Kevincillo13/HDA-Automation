Coloca aqui el archivo `template.xlsx` que usa AP15.

Ruta esperada por defecto:
`templates/template.xlsx`

El builder actual:
- carga ese template
- llena la hoja activa
- guarda un `.xlsx` por grupo
- exporta el `.csv` desde ese workbook

Si el archivo no existe, la generacion AP15 fallara con un error explicito.
