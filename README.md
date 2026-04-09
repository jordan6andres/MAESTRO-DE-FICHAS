## Description
ETL pipeline developed in Python for the automated processing of Excel files containing information on educational programs, referred to as "Fichas" under SENA notation. The pipeline applies relevant business rules—including filters, date calculations, and regulation assignments—to generate a consolidated target report.
## Software y librerias
- Python 3.8+
- Pandas ((data manipulation and transformation)
- NumPy (Math operations)
- Pathlib (file path handling)
- Openpyxl (Excel file writing)

## Workflow
1. **Extration**: Reads all `.xlsx`/`.xls` files from a specified directory, skipping unnecessary header rows.
2. **Consolidation**: Merges the data into a single DataFrame and removes duplicate records based on a unique identifier.
3. **Transformation**:
   - Filters out unwanted records (specific training levels and program types).
   - Calculates derived columns such as `REGLAMENTO`, `FECHA_FIN_ETAPA_LECTIVA`, `FECHA_VENCIMIENTO_INICIAL` and `FECHA_VENCIMIENTO_FINAL` in accordance with business rules.
4. **Load**: Exports the processed results to an Excel file containing two sheets: `Datos_Unicos` and `Datos_Filtrados`.


## Author
Jordan Palacios- [LinkedIn](https://linkedin.com/in/palaciosjordan)
