# ==========================================
# Archivo: em_generator.py
# Descripción: Script de configuración y ejecución para reportes de mercado.
# Autor: Esteban Acevedo Z.
# Fecha: 06/2026
# ==========================================

import time
# Importamos el motor desde el otro archivo
from funciones_em import cargar_bases, crear_word

# 1. CONFIGURACIÓN DE ARCHIVOS
archivo_datos = 'Base_catastral_2026.csv'
db_gi = 'La Galería Inmobiliaria/Renobo_Bogotá_Bases_05-26.xlsb'
db_cu = 'Coordenada Urbana/Base General Cundinamarca May 2026.xlsx'
db_gi_comercial = 'La Galería Inmobiliaria/Bogota_Comercio_Def_05-26.xlsb'

# 2. CONFIGURACIÓN DEL PREDIO
lotcodigo_predio = '001109024007'
nombre_predio = 'Sosiego'
columna_lot = 'LotCodigo'  # No se modifica
coordenadas = (4.575164353346292, -74.079864953606)

# 3. CONFIGURACIÓN DE RADIOS Y EXCLUSIONES
radio_gi = 1.5  # Radio en km para vivienda en Galería Inmobiliaria
radio_cu = 1  # Radio en km para vivienda en Coordenada Urbana
radio_no_residencial_gi = 1.0  # Radio en km para comercio (aplica igual para GI y CU)
# Nota: El radio para usos no residenciales de CU es el mismo que el de residenciales

# Listas de proyectos (IDs) que se desean excluir:
EXCLUIR_CODIGOS_CU = [24799, 19339, 1000013393, 1000002172, 1000007996, 1000011711] # Ej: [24799, 19339, 1000013393]
EXCLUIR_CODIGOS_GI_PROYECTOS = []
EXCLUIR_CODIGOS_GI_COMERCIO = []

# 4. EJECUCIÓN DEL REPORTE
if __name__ == '__main__':
    print(f"Iniciando reporte para {nombre_predio}.")
    start_time = time.time()
    
    # Cargar bases a la memoria (Solo se hace una vez)
    df_catastral, diccionario_bases = cargar_bases(
        archivo_csv=archivo_datos, 
        db_gi=db_gi, 
        db_cu=db_cu, 
        db_gi_comercial=db_gi_comercial
    )
    
    # Generar el documento
    crear_word(
        df_catastral=df_catastral, 
        dict_dfs=diccionario_bases, 
        lotcodigo=lotcodigo_predio, 
        nombre_predio=nombre_predio, 
        columna_id=columna_lot, 
        coordenadas=coordenadas, 
        radio_gi=radio_gi, 
        radio_cu=radio_cu, 
        radio_comercio=radio_no_residencial_gi,
        excluir_codigos_gi_proyectos=EXCLUIR_CODIGOS_GI_PROYECTOS,
        excluir_codigos_cu=EXCLUIR_CODIGOS_CU,
        excluir_codigos_gi_comercio=EXCLUIR_CODIGOS_GI_COMERCIO
    )
    
    elapsed = time.time() - start_time
    print(f"Tiempo de ejecución total: {elapsed:.2f} segundos")