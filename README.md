# SENA Report Wrangling Automation

A Python-based ETL pipeline for processing and consolidating SENA (Servicio Nacional de Aprendizaje) training group reports. The script automates the extraction, deduplication, filtering, and transformation of Excel-based training records—referred to as *Fichas*—applying institutional business rules to calculate regulatory dates and compliance fields.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Workflow](#workflow)
  - [1. Extraction](#1-extraction)
  - [2. Consolidation & Deduplication](#2-consolidation--deduplication)
  - [3. Filtering & Transformation](#3-filtering--transformation)
  - [4. Export](#4-export)
- [Business Rules](#business-rules)
- [Output Schema](#output-schema)
- [Source File Requirements](#source-file-requirements)
- [Troubleshooting](#troubleshooting)
- [Author](#author)

---

## Overview

This pipeline processes multiple SENA Excel reports located in a single directory, merges them into a unified dataset, removes duplicate training groups, applies institutional filters, and computes derived date columns based on Colombian vocational training regulations (*Acuerdo 007 de 2012* and *Acuerdo 009 de 2024*).

The final deliverable is an Excel workbook containing two sheets:

- **`Datos_Unicos`** — All unique training group records after deduplication.
- **`Datos_Filtrados`** — The subset of records that remain after applying program-level and formation-level filters, enriched with calculated regulatory and date fields.

---

## Prerequisites

- **Python** 3.8 or higher
- The following Python libraries:

| Library | Purpose |
|---------|---------|
| `pandas` | Data manipulation, consolidation, and filtering |
| `numpy` | Vectorized conditional logic and calculations |
| `openpyxl` | Excel file reading and writing |
| `python-dateutil` | Accurate month-based date arithmetic |

> **Note:** `pathlib` and `datetime` are part of the Python standard library.

---

## Installation

1. Clone or download this repository.
2. Install the required dependencies:

```bash
pip install pandas numpy openpyxl python-dateutil
```

---

## Configuration

Before running the script, update the two directory variables at the top of `senareport-wrangling-automation.py`:

```python
# ===== CONFIGURATION =====
RUTA_DIRECTORIO = Path("insert source path")      # Folder containing source Excel files
RUTA_SALIDA     = Path("insert directory destination path")  # Path for the output Excel file
```

- **`RUTA_DIRECTORIO`** — Absolute or relative path to the folder that contains the source `.xlsx` or `.xls` SENA reports.
- **`RUTA_SALIDA`** — Full path (including filename and `.xlsx` extension) where the consolidated workbook will be saved. If the file already exists, the sheets will be overwritten.

---

## Workflow

### 1. Extraction

- Scans the configured input directory for all `.xlsx` and `.xls` files.
- Reads each workbook starting from **row 5** (skips the first 4 header rows).
- Loads columns **A through AZ** and treats all values as strings to preserve leading zeros and avoid implicit type casting.
- Extracts the reporting **period** from the first six characters of each filename (e.g., `202501_reporte.xlsx` → period `202501`).

### 2. Consolidation & Deduplication

- Concatenates all individual DataFrames into a single consolidated dataset.
- Validates the presence of the required column **`IDENTIFICADOR_FICHA`**.
- Sorts records by period in **descending** order and drops duplicates based on `IDENTIFICADOR_FICHA`, retaining only the most recent record for each training group.

### 3. Filtering & Transformation

**Filters applied:**

The following training levels are **excluded**:
- `PROFUNDIZACIÓN TÉCNICA`
- `EVENTO`
- `CURSO ESPECIAL`

The following special programs are **excluded**:
- `INTEGRACIÓN CON LA EDUCACIÓN MEDIA ACADÉMICA`
- `INTEGRACIÓN CON LA EDUCACIÓN MEDIA TÉCNICA`

**Calculated fields:**

| Column | Description |
|--------|-------------|
| `REGLAMENTO` | Determines which regulation applies based on the training start date (`FECHA_INICIO_FICHA`). Records starting on or after **21 November 2024** receive *Acuerdo 009 de 2024*; all earlier records receive *Acuerdo 007 de 2012*. |
| `FECHA_FIN_ETAPA_LECTIVA` | Computed by subtracting months from `FECHA_TERMINACION_FICHA`: **6 months** for `TÉCNICO` and `TECNÓLOGO` levels; **3 months** for all other levels. |
| `FECHA_VENCIMIENTO_INICIAL` | For *Acuerdo 009 de 2024*, set equal to `FECHA_TERMINACION_FICHA`; for *Acuerdo 007 de 2012*, set to `N/A`. |
| `FECHA_VENCIMIENTO_FINAL` | Computed by adding months to `FECHA_TERMINACION_FICHA`: **12 months** for *Acuerdo 009 de 2024*; **18 months** for `TÉCNICO`/`TECNÓLOGO` under *Acuerdo 007 de 2012*; **21 months** for all other levels under *Acuerdo 007 de 2012*. |

### 4. Export

- Writes the deduplicated dataset to the **`Datos_Unicos`** sheet.
- Writes the filtered and enriched dataset to the **`Datos_Filtrados`** sheet.
- Both sheets are exported to the path specified in `RUTA_SALIDA`.

---

## Business Rules

| Condition | Rule |
|-----------|------|
| **Regulation assignment** | Start date ≥ 2024-11-21 → *Acuerdo 009 de 2024*; otherwise → *Acuerdo 007 de 2012*. |
| **Lective stage end** | `TÉCNICO` / `TECNÓLOGO` → termination date − 6 months; all others → termination date − 3 months. |
| **Initial expiry** | Only defined under *Acuerdo 009 de 2024* (equals termination date). |
| **Final expiry** | *Acuerdo 009 de 2024* → termination date + 12 months; *Acuerdo 007 de 2012* + `TÉCNICO`/`TECNÓLOGO` → + 18 months; *Acuerdo 007 de 2012* + other levels → + 21 months. |

---

## Output Schema

The **`Datos_Filtrados`** sheet contains the following columns in the order listed:

1. `IDENTIFICADOR_FICHA`
2. `ESTADO_CURSO`
3. `NIVEL_FORMACION`
4. `CODIGO_PROGRAMA`
5. `VERSION_PROGRAMA`
6. `NOMBRE_PROGRAMA_FORMACION`
7. `REGLAMENTO`
8. `FECHA_INICIO_FICHA`
9. `FECHA_TERMINACION_FICHA`
10. `FECHA_FIN_ETAPA_LECTIVA`
11. `FECHA_VENCIMIENTO_INICIAL`
12. `FECHA_VENCIMIENTO_FINAL`
13. `ETAPA_FICHA`
14. `MODALIDAD_FORMACION`
15. `NOMBRE_RESPONSABLE`
16. `NOMBRE_MUNICIPIO_CURSO`
17. `NOMBRE_PROGRAMA_ESPECIAL`

> **Note:** All source files must contain these columns (plus any additional fields used during processing). The script will halt and report any missing required columns before producing output.

---

## Source File Requirements

- **Format:** `.xlsx` or `.xls`
- **Naming convention:** The first six characters of the filename must be a numeric period identifier (e.g., `202501`, `202312`).
- **Structure:** Data must begin on row 5 (the script skips the first four rows). Relevant columns must be within the A:AZ range.
- **Required column:** `IDENTIFICADOR_FICHA` must be present in every source file.

---

## Troubleshooting

| Issue | Likely Cause | Resolution |
|-------|--------------|------------|
| `⚠️ Input directory not found` | `RUTA_DIRECTORIO` points to a non-existent path. | Verify the directory path in the configuration block. |
| `⚠️ No Excel files found` | The directory exists but contains no `.xlsx` or `.xls` files. | Check the folder contents and file extensions. |
| `⚠️ Required column 'IDENTIFICADOR_FICHA' not found` | Source files are missing the expected schema or the header rows were altered. | Ensure source files follow the SENA report template. |
| `⚠️ Missing columns in data` | One or more final output columns are absent from the source data. | Verify that all source files contain the complete set of required columns. |
| `⚠️ Error exporting to Excel` | The output file is open in another application, or the destination path is invalid. | Close any open instances of the output file and check write permissions. |

---

## Author

**Jordan Palacios**  
[LinkedIn](https://linkedin.com/in/palaciosjordan)
