# Reporte semanal de dengue (NETLABv2)

Este repositorio genera el reporte semanal de dengue a partir de una descarga de NETLABv2. El script produce **los mismos archivos de salida** que la versión anterior, solo se reorganizó para mayor robustez y validaciones.

## Requisitos

- R (recomendado >= 4.1)
- Paquetes de R indicados en el script (se instalan automáticamente si faltan)
- Archivo Excel de entrada (por defecto `dengue_2025.xlsx`)
- Plantilla PPT opcional: `plantilla_base.pptx`

## Pasos de ejecución

1. Coloca el archivo Excel de NETLAB en la raíz del proyecto (o ajusta la ruta en el script).
2. Abre y edita **solo la sección de configuración** en `Reporte_dengue_2026.R`:
   - `archivo`, `hoja`
   - Parámetros de filtros y semanas
3. Ejecuta el script en R:

```r
source("Reporte_dengue_2026.R")
```

## Salidas

El script crea una carpeta con nombre `SE XX` (o `YYYY_SE XX` si se habilita) y genera:

- `01_IP_por_SE_HighImpact.png`
- `02_Procesamiento_por_prueba_HighImpact.png`
- `03_Tabla_positivos_por_provincia.xlsx`
- `04_Tabla_SEXX_Microred_Establecimiento.xlsx`
- `Reporte_Dengue_SEXX_Ejecutivo_V2.pptx`

## Notas

- No se modifica la lógica epidemiológica, solo la estructura y validaciones.
- Si faltan columnas clave, el script detiene la ejecución con un mensaje claro.
