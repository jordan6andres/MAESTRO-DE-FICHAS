## Descripción
Pipeline ETL desarrollado en Python que procesa automáticamente archivos Excel que contiene informacion progrmas educativos, conocidos como "Fichas" en notación SENA. Aplica reglas pertinenetes (filtros, cálculos de fechas, asignación de reglamentos) que generan de manera consolidad el reporte objetivo.

## Software y librerias
- Python 3.8+
- Pandas (manipulación y transformación de datos)
- NumPy (operaciones vectoriales)
- Pathlib (manejo de rutas)
- Openpyxl (escritura de archivos Excel)

## Flujo del proceso
1. **Extracción**: Lee todos los archivos `.xlsx`/`.xls` de un directorio, saltando filas innecesarias.
2. **Consolidación**: Unifica los datos en un solo DataFrame y elimina duplicados por identificador.
3. **Transformación**:
   - Filtra registros no deseados (niveles de formación y programas específicos).
   - Calcula columnas como `REGLAMENTO`, `FECHA_FIN_ETAPA_LECTIVA`, `FECHA_VENCIMIENTO_INICIAL` y `FECHA_VENCIMIENTO_FINAL` según reglas de negocio.
4. **Carga**: Exporta los resultados a un archivo Excel con dos hojas: `Datos_Unicos` y `Datos_Filtrados`.


## Autor
Jordan Palacios- [LinkedIn](https://linkedin.com/in/palaciosjordan)
