# Automatización de Procesamiento de Reportes SENA

Pipeline ETL desarrollado en Python para el procesamiento y consolidación de reportes de fichas de carácterización del SENA (Servicio Nacional de Aprendizaje). El script automatiza la extracción, desduplicación, filtrado y transformación de registros formativos en Excel—conocidos como *Fichas*—aplicando reglas de negocio institucionales para calcular fechas regulatorias y campos de cumplimiento normativo.

---

## Tabla de Contenidos

- [Descripción General](#descripción-general)
- [Requisitos Previos](#requisitos-previos)
- [Instalación](#instalación)
- [Configuración](#configuración)
- [Flujo de Trabajo](#flujo-de-trabajo)
  - [1. Extracción](#1-extracción)
  - [2. Consolidación y Desduplicación](#2-consolidación-y-desduplicación)
  - [3. Filtrado y Transformación](#3-filtrado-y-transformación)
  - [4. Exportación](#4-exportación)
- [Reglas de Negocio](#reglas-de-negocio)
- [Esquema de Salida](#esquema-de-salida)
- [Requisitos de los Archivos Fuente](#requisitos-de-los-archivos-fuente)
- [Solución de Problemas](#solución-de-problemas)
- [Autor](#autor)

---

## Descripción General

Este pipeline procesa múltiples reportes de Excel del SENA ubicados en un único directorio, los fusiona en un conjunto de datos unificado, elimina fichas formativas duplicadas, aplica filtros institucionales y calcula columnas derivadas de fechas con base en la normatividad colombiana de formación técnica (*Acuerdo 007 de 2012* y *Acuerdo 009 de 2024*).

El entregable final es un libro de Excel que contiene dos hojas:

- **`Datos_Unicos`** — Todos los registros únicos de fichas formativas después de la desduplicación.
- **`Datos_Filtrados`** — El subconjunto de registros que permanecen después de aplicar los filtros de nivel de formación y tipo de programa, enriquecidos con los campos regulatorios y de fechas calculados.

---

## Requisitos Previos

- **Python** 3.8 o superior
- Las siguientes bibliotecas de Python:

| Biblioteca | Propósito |
|-----------|---------|
| `pandas` | Manipulación, consolidación y filtrado de datos |
| `numpy` | Lógica condicional y cálculos vectorizados |
| `openpyxl` | Lectura y escritura de archivos Excel |
| `python-dateutil` | Aritmética de fechas basada en meses con precisión |

> **Nota:** `pathlib` y `datetime` forman parte de la biblioteca estándar de Python.

---

## Instalación

1. Clone o descargue este repositorio.
2. Instale las dependencias requeridas:

```bash
pip install pandas numpy openpyxl python-dateutil
```

---

## Configuración

Antes de ejecutar el script, actualice las dos variables de directorio ubicadas en la parte superior de `senareport-wrangling-automation.py`:

```python
# ===== CONFIGURATION =====
RUTA_DIRECTORIO = Path("inserte ruta de origen")      # Carpeta que contiene los archivos Excel fuente
RUTA_SALIDA     = Path("inserte ruta de destino")      # Ruta para el archivo Excel de salida
```

- **`RUTA_DIRECTORIO`** — Ruta absoluta o relativa a la carpeta que contiene los reportes SENA en formato `.xlsx` o `.xls`.
- **`RUTA_SALIDA`** — Ruta completa (incluyendo nombre de archivo y extensión `.xlsx`) donde se guardará el libro consolidado. Si el archivo ya existe, las hojas serán sobrescritas.

---

## Flujo de Trabajo

### 1. Extracción

- Escanea el directorio de entrada configurado en busca de todos los archivos `.xlsx` y `.xls`.
- Lee cada libro comenzando desde la **fila 5** (omite las primeras cuatro filas de encabezado).
- Carga las columnas **A hasta AZ** y trata todos los valores como cadenas de texto para preservar ceros a la izquierda y evitar conversiones de tipo implícitas.
- Extrae el **periodo** de reporte a partir de los primeros seis caracteres del nombre de cada archivo (por ejemplo, `202501_reporte.xlsx` → periodo `202501`).

### 2. Consolidación y Desduplicación

- Concatena todos los DataFrames individuales en un único conjunto de datos consolidado.
- Valida la presencia de la columna requerida **`IDENTIFICADOR_FICHA`**.
- Ordena los registros por periodo en orden **descendente** y elimina duplicados basándose en `IDENTIFICADOR_FICHA`, conservando únicamente el registro más reciente de cada ficha formativa.

### 3. Filtrado y Transformación

**Filtros aplicados:**

Se **excluyen** los siguientes niveles de formación:
- `PROFUNDIZACIÓN TÉCNICA`
- `EVENTO`
- `CURSO ESPECIAL`

Se **excluyen** los siguientes programas especiales:
- `INTEGRACIÓN CON LA EDUCACIÓN MEDIA ACADÉMICA`
- `INTEGRACIÓN CON LA EDUCACIÓN MEDIA TÉCNICA`

**Campos calculados:**

| Columna | Descripción |
|--------|-------------|
| `REGLAMENTO` | Determina qué reglamento aplica según la fecha de inicio de la ficha (`FECHA_INICIO_FICHA`). Los registros con fecha de inicio igual o posterior al **21 de noviembre de 2024** reciben *Acuerdo 009 de 2024*; todos los registros anteriores reciben *Acuerdo 007 de 2012*. |
| `FECHA_FIN_ETAPA_LECTIVA` | Calculada restando meses de `FECHA_TERMINACION_FICHA`: **6 meses** para niveles `TÉCNICO` y `TECNÓLOGO`; **3 meses** para todos los demás niveles. |
| `FECHA_VENCIMIENTO_INICIAL` | Para el *Acuerdo 009 de 2024*, se establece igual a `FECHA_TERMINACION_FICHA`; para el *Acuerdo 007 de 2012*, se establece en `N/A`. |
| `FECHA_VENCIMIENTO_FINAL` | Calculada sumando meses a `FECHA_TERMINACION_FICHA`: **12 meses** para *Acuerdo 009 de 2024*; **18 meses** para `TÉCNICO`/`TECNÓLOGO` bajo *Acuerdo 007 de 2012*; **21 meses** para los demás niveles bajo *Acuerdo 007 de 2012*. |

### 4. Exportación

- Escribe el conjunto de datos desduplicado en la hoja **`Datos_Unicos`**.
- Escribe el conjunto de datos filtrado y enriquecido en la hoja **`Datos_Filtrados`**.
- Ambas hojas se exportan a la ruta especificada en `RUTA_SALIDA`.

---

## Reglas de Negocio

| Condición | Regla |
|-----------|------|
| **Asignación de reglamento** | Fecha de inicio ≥ 2024-11-21 → *Acuerdo 009 de 2024*; de lo contrario → *Acuerdo 007 de 2012*. |
| **Fin de etapa lectiva** | `TÉCNICO` / `TECNÓLOGO` → fecha de terminación − 6 meses; demás niveles → fecha de terminación − 3 meses. |
| **Vencimiento inicial** | Solo se define bajo *Acuerdo 009 de 2024* (igual a la fecha de terminación). |
| **Vencimiento final** | *Acuerdo 009 de 2024* → fecha de terminación + 12 meses; *Acuerdo 007 de 2012* + `TÉCNICO`/`TECNÓLOGO` → + 18 meses; *Acuerdo 007 de 2012* + demás niveles → + 21 meses. |

---

## Esquema de Salida

La hoja **`Datos_Filtrados`** contiene las siguientes columnas en el orden indicado:

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

> **Nota:** Todos los archivos fuente deben contener estas columnas (además de cualquier campo adicional utilizado durante el procesamiento). El script se detendrá y reportará cualquier columna requerida faltante antes de producir la salida.

---

## Requisitos de los Archivos Fuente

- **Formato:** `.xlsx` o `.xls`
- **Convención de nomenclatura:** Los primeros seis caracteres del nombre de archivo deben ser un identificador numérico de periodo (por ejemplo, `202501`, `202312`).
- **Estructura:** Los datos deben comenzar en la fila 5 (el script omite las primeras cuatro filas). Las columnas relevantes deben estar dentro del rango A:AZ.
- **Columna requerida:** `IDENTIFICADOR_FICHA` debe estar presente en cada archivo fuente.

---

## Solución de Problemas

| Problema | Causa Probable | Solución |
|-------|--------------|------------|
| `⚠️ Input directory not found` / `⚠️ Directorio de entrada no encontrado` | `RUTA_DIRECTORIO` apunta a una ruta inexistente. | Verifique la ruta del directorio en el bloque de configuración. |
| `⚠️ No Excel files found` / `⚠️ No se encontraron archivos Excel` | El directorio existe pero no contiene archivos `.xlsx` o `.xls`. | Revise el contenido de la carpeta y las extensiones de los archivos. |
| `⚠️ Required column 'IDENTIFICADOR_FICHA' not found` / `⚠️ Columna requerida 'IDENTIFICADOR_FICHA' no encontrada` | Los archivos fuente no contienen el esquema esperado o las filas de encabezado fueron alteradas. | Asegúrese de que los archivos fuente sigan la plantilla de reporte del SENA. |
| `⚠️ Missing columns in data` / `⚠️ Columnas faltantes en los datos` | Una o más columnas de salida finales no están presentes en los datos de origen. | Verifique que todos los archivos fuente contengan el conjunto completo de columnas requeridas. |
| `⚠️ Error exporting to Excel` / `⚠️ Error al exportar a Excel` | El archivo de salida está abierto en otra aplicación, o la ruta de destino no es válida. | Cierre cualquier instancia abierta del archivo de salida y verifique los permisos de escritura. |

---

## Autor

**Jordan Palacios**  
[LinkedIn](https://linkedin.com/in/palaciosjordan)
