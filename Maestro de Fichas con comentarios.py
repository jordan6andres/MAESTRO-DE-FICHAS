import pandas as pd
from pathlib import Path
import os
from datetime import datetime, timedelta #El Módulo ~datetime~ nos ayuda trabajar con fechas y horas python. 
import numpy as np

# ===== CONFIGURACIÓN =====
RUTA_DIRECTORIO = Path("C:/Users/jorda/Desktop/ETAPA PRODUCTIVA/Pecerocuatros 25")#insertar ruta de base de datos
RUTA_SALIDA = Path("C:/Users/jorda/Desktop/resultado_verificacion_2.xlsx") #insertar ruta de directorio destino

# ===== FUNCIÓN PARA PROCESAR ARCHIVOS =====
def procesar_archivo(archivo):
    try:
        df = pd.read_excel( # ~pandas.read_excel~ funcion de Pandas que lee archivo Excel en un Dataframe de pandas.
            archivo,
            skiprows=4,
            usecols="A:AZ",
            dtype=str
        )
        nombre = archivo.stem[:6] # ~PurePath.stem~ funcion de pathlib que extrae el nombre del archivo sin extensión.
        periodo = int(nombre) if nombre.isdigit() else None # ~str.isdigit~ función de string para verificar el contenido de dígitos.
        df['archivo_origen'] = archivo.name
        df['periodo'] = periodo
        return df
    except Exception as e:
        print(f"Error procesando {archivo.name}: {str(e)}")
        return pd.DataFrame()

# ===== FUNCIÓN PARA CALCULAR FECHAS =====
def sumar_meses(fecha, meses):
    try:
        fecha_dt = datetime.strptime(fecha, "%d/%m/%Y") # El metodo ~.strptime()~ convierte string → un obejeto datetime.
        # Calcula el año y mes resultante
        total_meses = fecha_dt.month + meses # El atributo ~.month~ te da directamente el número del mes para trabajar los calculos fácilmente.
        year = fecha_dt.year + total_meses // 12
        month = total_meses % 12

        '''La expresion ~[month-1]~ indexa para acceder al indice correspondiente en la lista.
        El metodo ~min()~ Garantizar que el día no exceda los días válidos del mes resultante.
        se usa para obtener los dias del mes "fecha" comparado entre los dias del mes de "month" '''

        day = min(fecha_dt.day, [31,29 if year%4==0 and not year%100==0 or year%400==0 else 28,31,30,31,30,31,31,30,31,30,31][month-1]) 
        
        return datetime(year, month, day).strftime("%d/%m/%Y") # El metodo ~.strftime()~ convierte un objeto datetime → string.
    except:
        return np.nan #El atributo ~np.nan~ se usa para manejar fechas con formatos invalidos retornando valores Null.

# ===== FUNCIÓN PRINCIPAL =====
def main():
    print("\n" + "="*50) #parametros esteticos para el titulo en la terminal
    print("INICIANDO VERIFICACIÓN DE DATOS")
    print("="*50) #parametros esteticos para el titulo en la terminal
    
    # Paso 1: Cargar datos
    print("\n🔷 PASO 1: Cargando y consolidando archivos...")
    dataframes = []
    archivos_procesados = []
    
    for archivo in RUTA_DIRECTORIO.iterdir(): # sin el metodo ~pathlib.iterdir()~ no es posible iterar sobre los elementos dentro de la ruta. Las rutas no son iterables.
        if archivo.suffix.lower() in ['.xlsx', '.xls']:
            print(f"Procesando: {archivo.name}")
            df = procesar_archivo(archivo)
            if not df.empty:
                dataframes.append(df)
                archivos_procesados.append(archivo.name)
    
    if not dataframes:
        print("⚠️ No se encontraron datos válidos - Verifica los archivos fuente")
        return
    
    df_consolidado = pd.concat(dataframes, ignore_index=True) # El parametro ~ignore_index~ cuando es 'True' ordena el indice consecutivamente evitando los repetidos.

    # Paso 2: Eliminar duplicados y exportar
    df_consolidado['periodo'] = pd.to_numeric(df_consolidado['periodo'], errors='coerce') #La función ~pandas.to_numeric~ se utiliza para convertir el argumento a valores tipo numérico, parametro 'errors' fuerza los valores no convertibles a NaN
    df_unico = df_consolidado.sort_values('periodo', ascending=False).drop_duplicates('IDENTIFICADOR_FICHA')
    
    # ===== BLOQUE 2: PROCESAMIENTO ADICIONAL =====
    print("\n🔷 PASO 2: Filtrando datos y calculando nuevas columnas...")
    
    # 1. Filtrar registros no deseados

    '''El metodo ~.isin()~  permite verificar si los elementos de una Serie o DataFrame existen en un conjunto de valores dado.
    El signo "~" es el operador de negación binaria en pandas. en este caso se usa para invertir el filtro.
    Es decir toma todos los valores excepto por lo que estan en lista'''

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
    
    # 2. Crear nuevas columnas
    # Columna: REGLAMENTO
    df_filtrado['FECHA_INICIO_FICHA'] = pd.to_datetime( #el metodo ~.to_datetime~ transforma los arg. suministrados en objetos 'datetime'.
        df_filtrado['FECHA_INICIO_FICHA'], 
        dayfirst=True, #Este argumento se usa para determinar si la primera parte de la fecha sumnistrada es el dia.
        errors='coerce'
    )
    '''Numpy.where (condición, x, y):
    En este caso, devuelve un arreglo con los elementos de x donde la condición es verdadera, y los elementos de y donde la condición es falsa.'''

    df_filtrado['REGLAMENTO'] = np.where(
        df_filtrado['FECHA_INICIO_FICHA'] >= pd.Timestamp('2024-11-21'),
        'Acuerdo 009 de 2024',
        'Acuerdo 007 de 2012'
    )
    
    # Preparar fechas para cálculos
    for col in ['FECHA_TERMINACION_FICHA']:
        df_filtrado[col] = pd.to_datetime(
            df_filtrado[col], 
            dayfirst=True, 
            errors='coerce'
        ).dt.strftime('%d/%m/%Y')
    
    # Columna: FECHA_FIN_ETAPA_LECTIVA

    '''El metodo ~.apply()~ se usa para aplicar una función a cada elemento de una Serie o cada fila o columna de un DataFrame. 
    Este ultimo dependera del parametro ~axix~ donde 0 aplica la función a cada columna. 1 aplica la función a cada fila.'''

    df_filtrado['FECHA_FIN_ETAPA_LECTIVA'] = df_filtrado.apply( 
        lambda x: sumar_meses(x['FECHA_TERMINACION_FICHA'], -6 if x['NIVEL_FORMACION'] in ['TÉCNICO', 'TECNÓLOGO'] else -3), axis=1
        )
    
    # Columna: FECHA_VENCIMIENTO_INICIAL
    df_filtrado['FECHA_VENCIMIENTO_INICIAL'] = np.where(
        df_filtrado['REGLAMENTO'] == 'Acuerdo 009 de 2024',
        df_filtrado['FECHA_TERMINACION_FICHA'],
        'N/A'
    )
    
    # Columna: FECHA_VENCIMIENTO_FINAL
    def calcular_vencimiento_final(row):
        if row['REGLAMENTO'] == 'Acuerdo 009 de 2024':
            meses = 12
        else:
            if row['NIVEL_FORMACION'] in ['TÉCNICO', 'TECNÓLOGO']:
                meses = 18
            else:
                meses = 21
        
        return sumar_meses(row['FECHA_TERMINACION_FICHA'], meses)
    
    df_filtrado['FECHA_VENCIMIENTO_FINAL'] = df_filtrado.apply(calcular_vencimiento_final, axis=1)
    
    # 3. Seleccionar columnas requeridas
    columnas_finales = [
        'IDENTIFICADOR_FICHA', 'ESTADO_CURSO', 'NIVEL_FORMACION', 'CODIGO_PROGRAMA',
        'VERSION_PROGRAMA', 'NOMBRE_PROGRAMA_FORMACION', 'REGLAMENTO', 'FECHA_INICIO_FICHA',
        'FECHA_TERMINACION_FICHA', 'FECHA_FIN_ETAPA_LECTIVA', 'FECHA_VENCIMIENTO_INICIAL',
        'FECHA_VENCIMIENTO_FINAL', 'ETAPA_FICHA', 'MODALIDAD_FORMACION', 'NOMBRE_RESPONSABLE',
        'NOMBRE_MUNICIPIO_CURSO', 'NOMBRE_PROGRAMA_ESPECIAL'
    ]
    
    df_final = df_filtrado[columnas_finales].copy()
    
    # 4. Exportar resultados
    if RUTA_SALIDA.exists(): #El metodo ~pathlib.Path.exists()~ se utiliza para verificar si una ruta de archivo o directorio especificada existe en el sistema de archivos.
        mode = 'a'
        sheet_exists = 'replace'
    else:
        mode = 'w'
        sheet_exists = None  # No aplica en modo escritura

    '''~pandas.ExcelWriter()~ se utiliza para escribir datos en archivos de Excel, permitiendo la creación de hojas nuevas y la modificación de hojas existentes. 
    Es especialmente útil cuando se necesita guardar múltiples hojas o anexar datos a un archivo Excel ya existente.'''

    '''El bloque 'with' actúa como un administrador que abre (o crea) el archivo Excel,
    escribe las dos hojas requeridas y, al terminar, guarda y cierra el archivo
    automáticamente'''
    
    with pd.ExcelWriter( 
        RUTA_SALIDA,
        engine='openpyxl',
        mode=mode,
        if_sheet_exists=sheet_exists
    ) as writer:
        df_unico.to_excel(writer, sheet_name='Datos_Unicos', index=False) #El metodo ~.to_excel~ exporta dataframes como archivos excel. ~index=False~ Evita mantener los índices del DataFrame al nuevo archivo.
        df_final.to_excel(writer, sheet_name='Datos_Filtrados', index=False)
    
    # Paso 3: Resumen estadístico
    print("\n🔷 RESUMEN FINAL:")
    print(f"Archivos procesados: {len(archivos_procesados)}")
    print(f"Registros consolidados: {len(df_consolidado)}")
    print(f"Registros únicos: {len(df_unico)}")
    print(f"Registros filtrados: {len(df_filtrado)}")
    print(f"Período más reciente: {df_unico['periodo'].max()}")
    print(f"Período más antiguo: {df_unico['periodo'].min()}")
    
    print("\n✅ Verificación completada. Abre el archivo Excel para validar:")
    print(str(RUTA_SALIDA))

# ===== EJECUCIÓN =====
print("Iniciando proceso de verificación...")
main()
print("Proceso finalizado")