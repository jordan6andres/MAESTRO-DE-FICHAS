import pandas as pd
from pathlib import Path
from datetime import datetime
import numpy as np
from dateutil.relativedelta import relativedelta

# ===== CONFIGURATION =====
RUTA_DIRECTORIO = Path("insert source path")
RUTA_SALIDA = Path("insert directory destination path")

# ===== FILE PROCESSING FUNCTION =====
def procesar_archivo(archivo):
    """
    Reads an Excel file into a DataFrame and extracts the period
    from the first 6 characters of the filename.
    """
    try:
        df = pd.read_excel(
            archivo,
            skiprows=4,
            usecols="A:AZ",
            dtype=str
        )
        nombre = archivo.stem[:6]
        periodo = int(nombre) if nombre.isdigit() else None
        df['archivo_origen'] = archivo.name
        df['periodo'] = periodo
        return df
    except Exception as e:
        print(f"Error processing {archivo.name}: {e}")
        return pd.DataFrame()

# ===== DATE CALCULATION FUNCTION =====
def sumar_meses(fecha, meses):
    """
    Adds a given number of months to a date.
    Handles end-of-month overflow automatically.
    Returns a datetime object or pd.NaT for invalid inputs.
    """
    if pd.isna(fecha) or pd.isna(meses):
        return pd.NaT
    try:
        if isinstance(fecha, str):
            fecha = datetime.strptime(fecha, "%d/%m/%Y")
        return fecha + relativedelta(months=int(meses))
    except (ValueError, TypeError):
        return pd.NaT

# ===== MAIN FUNCTION =====
def main():
    print("\n" + "="*50)
    print("INITIATING DATA VERIFICATION")
    print("="*50)
    
    # Validate input directory before processing
    if not RUTA_DIRECTORIO.exists() or not RUTA_DIRECTORIO.is_dir():
        print(f"⚠️ Input directory not found: {RUTA_DIRECTORIO}")
        return
    
    # Step 1: Load data
    print("\n🔷 STEP 1: Loading and consolidating files...")
    dataframes = []
    archivos_procesados = []
    
    # Pre-filter Excel files to avoid unnecessary iterations
    excel_files = [f for f in RUTA_DIRECTORIO.iterdir() if f.suffix.lower() in ['.xlsx', '.xls']]
    if not excel_files:
        print("⚠️ No Excel files found - Verify source files")
        return
    
    for archivo in excel_files:
        print(f"Processing: {archivo.name}")
        df = procesar_archivo(archivo)
        if not df.empty:
            dataframes.append(df)
            archivos_procesados.append(archivo.name)
    
    if not dataframes:
        print("⚠️ No valid data found - Verify source files")
        return
    
    df_consolidado = pd.concat(dataframes, ignore_index=True)
    
    # Step 2: Remove duplicates
    df_consolidado['periodo'] = pd.to_numeric(df_consolidado['periodo'], errors='coerce')
    
    # Validate critical columns before deduplication
    if 'IDENTIFICADOR_FICHA' not in df_consolidado.columns:
        print("⚠️ Required column 'IDENTIFICADOR_FICHA' not found in source files")
        return
    
    df_unico = df_consolidado.sort_values('periodo', ascending=False).drop_duplicates('IDENTIFICADOR_FICHA')
    
    # ===== BLOCK 2: ADDITIONAL PROCESSING =====
    print("\n🔷 STEP 2: Filtering data and calculating new columns...")
    
    # 1. Filter unwanted records
    filtro_nivel = ~df_unico['NIVEL_FORMACION'].isin([
        'PROFUNDIZACIÓN TÉCNICA', 
        'EVENTO', 
        'CURSO ESPECIAL'
    ])
    
    filtro_programa = ~df_unico['NOMBRE_PROGRAMA_ESPECIAL'].isin([
        'INTEGRACIÓN CON LA EDUCACIÓN MEDIA ACADÉMICA',
        'INTEGRACIÓN CON LA EDUCACIÓN MEDIA TÉCNICA'
    ])
    
    df_filtrado = df_unico[filtro_nivel & filtro_programa].copy()
    
    # 2. Create new columns
    # Column: REGLAMENTO
    df_filtrado['FECHA_INICIO_FICHA'] = pd.to_datetime(
        df_filtrado['FECHA_INICIO_FICHA'], 
        dayfirst=True, 
        errors='coerce'
    )
    
    cutoff_date = pd.Timestamp('2024-11-21')
    df_filtrado['REGLAMENTO'] = np.where(
        df_filtrado['FECHA_INICIO_FICHA'] >= cutoff_date,
        'Acuerdo 009 de 2024',
        'Acuerdo 007 de 2012'
    )
    
    # Prepare FECHA_TERMINACION_FICHA as datetime for calculations
    df_filtrado['FECHA_TERMINACION_FICHA'] = pd.to_datetime(
        df_filtrado['FECHA_TERMINACION_FICHA'], 
        dayfirst=True, 
        errors='coerce'
    )
    
    # Column: FECHA_FIN_ETAPA_LECTIVA
    # Vectorized month calculation: -6 for TÉCNICO/TECNÓLOGO, otherwise -3
    meses_lectiva = np.where(
        df_filtrado['NIVEL_FORMACION'].isin(['TÉCNICO', 'TECNÓLOGO']),
        -6,
        -3
    )
    df_filtrado['FECHA_FIN_ETAPA_LECTIVA'] = [
        sumar_meses(d, int(m)) for d, m in zip(df_filtrado['FECHA_TERMINACION_FICHA'], meses_lectiva)
    ]
    
    # Column: FECHA_VENCIMIENTO_INICIAL
    df_filtrado['FECHA_VENCIMIENTO_INICIAL'] = np.where(
        df_filtrado['REGLAMENTO'] == 'Acuerdo 009 de 2024',
        df_filtrado['FECHA_TERMINACION_FICHA'],
        'N/A'
    )
    
    # Column: FECHA_VENCIMIENTO_FINAL
    # Vectorized month calculation based on REGLAMENTO and NIVEL_FORMACION
    meses_vencimiento = np.where(
        df_filtrado['REGLAMENTO'] == 'Acuerdo 009 de 2024',
        12,
        np.where(
            df_filtrado['NIVEL_FORMACION'].isin(['TÉCNICO', 'TECNÓLOGO']),
            18,
            21
        )
    )
    df_filtrado['FECHA_VENCIMIENTO_FINAL'] = [
        sumar_meses(d, int(m)) for d, m in zip(df_filtrado['FECHA_TERMINACION_FICHA'], meses_vencimiento)
    ]
    
    # 3. Select required columns
    columnas_finales = [
        'IDENTIFICADOR_FICHA', 'ESTADO_CURSO', 'NIVEL_FORMACION', 'CODIGO_PROGRAMA',
        'VERSION_PROGRAMA', 'NOMBRE_PROGRAMA_FORMACION', 'REGLAMENTO', 'FECHA_INICIO_FICHA',
        'FECHA_TERMINACION_FICHA', 'FECHA_FIN_ETAPA_LECTIVA', 'FECHA_VENCIMIENTO_INICIAL',
        'FECHA_VENCIMIENTO_FINAL', 'ETAPA_FICHA', 'MODALIDAD_FORMACION', 'NOMBRE_RESPONSABLE',
        'NOMBRE_MUNICIPIO_CURSO', 'NOMBRE_PROGRAMA_ESPECIAL'
    ]
    
    missing_cols = [c for c in columnas_finales if c not in df_filtrado.columns]
    if missing_cols:
        print(f"⚠️ Missing columns in data: {missing_cols}")
        return
    
    df_final = df_filtrado[columnas_finales].copy()
    
    # 4. Export results
    try:
        if RUTA_SALIDA.exists():
            mode = 'a'
            sheet_exists = 'replace'
        else:
            mode = 'w'
            sheet_exists = None
        
        with pd.ExcelWriter(
            RUTA_SALIDA,
            engine='openpyxl',
            mode=mode,
            if_sheet_exists=sheet_exists
        ) as writer:
            df_unico.to_excel(writer, sheet_name='Datos_Unicos', index=False)
            df_final.to_excel(writer, sheet_name='Datos_Filtrados', index=False)
    except Exception as e:
        print(f"⚠️ Error exporting to Excel: {e}")
        return
    
    # Step 3: Statistical summary
    print("\n🔷 FINAL SUMMARY:")
    print(f"Files processed: {len(archivos_procesados)}")
    print(f"Consolidated records: {len(df_consolidado)}")
    print(f"Unique records: {len(df_unico)}")
    print(f"Filtered records: {len(df_filtrado)}")
    print(f"Most recent period: {df_unico['periodo'].max()}")
    print(f"Oldest period: {df_unico['periodo'].min()}")
    
    print("\n✅ Verification completed. Open the Excel file to validate:")
    print(str(RUTA_SALIDA))

# ===== EXECUTION =====
print("Initiating verification process...")
main()
print("Process finished")
