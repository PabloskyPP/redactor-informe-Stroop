"""
Módulo para leer datos del test Stroop desde archivos Excel

Este módulo proporciona funciones para:
- Leer datos del Excel del test Raven
- Calcular puntuaciones directas por fila y totales
- Identificar explícitamente celdas seleccionadas y no seleccionadas
- Proporcionar estructuras de datos claras y accesibles para análisis posterior
"""
import pandas as pd


def leer_datos_excel(ruta_archivo):
    """
    Lee los datos del archivo Excel del test Stroop
    
    Args:
        ruta_archivo: Ruta al archivo Excel
        
    Returns:
        dict: Diccionario con 'edad', 'sub_num', 'nombre_completo', 'nombre' y 'datos_Stroop'
    """
    try:
        # Leer hoja 'info' para obtener la edad y sub_num
        df_info = pd.read_excel(ruta_archivo, sheet_name='info')
        edad = df_info['age'].iloc[0] if 'age' in df_info.columns else None
        
        # Extraer sub_num y procesarlo
        sub_num = df_info['sub_num'].iloc[0] if 'sub_num' in df_info.columns else None
        
        # Procesar nombre completo y nombre (primer token)
        # Verificar que sub_num es válido (no None, no NaN, no cadena vacía)
        import math
        if sub_num and not (isinstance(sub_num, float) and math.isnan(sub_num)):
            nombre_completo = str(sub_num).strip()
            # Obtener solo el primer token (antes del primer espacio)
            # Verificar que hay contenido antes de dividir
            if nombre_completo:
                tokens = nombre_completo.split()
                nombre = tokens[0] if tokens else nombre_completo
            else:
                nombre = nombre_completo
        else:
            nombre_completo = None
            nombre = None
        
        # Leer hoja 'stroop' con los datos del test
        df_stroop = pd.read_excel(ruta_archivo, sheet_name='stroop')
        
        return {
            'edad': edad,
            'sub_num': sub_num,
            'nombre_completo': nombre_completo,
            'nombre': nombre,
            'datos_stroop': df_stroop
        }
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        raise


def calcular_puntuaciones_directas(datos):
    """
    Calcula las puntuaciones directas del test Stroop
    
    Estructura esperada del Excel (hoja 'stroop'):
    - Primera columna: índices (P, C, PC, E)
    - Segunda columna: valores de puntuaciones directas
    
    Args:
        datos: Dict con 'edad' y 'datos_stroop'
        
    Returns:
        dict: Diccionario con todas las puntuaciones directas
    """
    df = datos['datos_stroop']
    
    # La estructura es: primera columna = índices, segunda columna = valores
    # Obtenemos los nombres de las columnas (pueden variar)
    columnas = df.columns.tolist()
    
    if len(columnas) < 2:
        raise ValueError("El Excel debe tener al menos 2 columnas (índices y valores)")
    
    col_indice = columnas[0]  # Primera columna con los índices (P, C, PC, E)
    col_valor = columnas[1]   # Segunda columna con los valores
    
    # Diccionario para almacenar las puntuaciones por serie
    resultados = {
        'edad': datos['edad'],
        'sub_num': datos.get('sub_num'),
        'nombre_completo': datos.get('nombre_completo'),
        'nombre': datos.get('nombre'),
        'datos_stroop': df,
        
        # Puntuaciones directas por serie
        'PD_P': 0,
        'PD_C': 0,
        'PD_PC': 0,
        'PD_E': 0,  # Errores totales
        
        # Índice de interferencia
        'Indice_interferencia': 0,
        
        # Puntuaciones esperadas
        'PD_esperada_PC': 0,
    }
    
    # Leer los valores del DataFrame según el índice
    for idx, row in df.iterrows():
        indice = str(row[col_indice]).strip().upper()
        valor = row[col_valor]
        
        # Asignar el valor según el índice
        if indice == 'P':
            resultados['PD_P'] = valor
        elif indice == 'C':
            resultados['PD_C'] = valor
        elif indice == 'PC':
            resultados['PD_PC'] = valor
        elif indice == 'E':
            resultados['PD_E'] = valor
    
    # Calcular puntuaciones esperadas: PC' = (P × C) / (P + C)
    PD_P = resultados['PD_P']
    PD_C = resultados['PD_C']
    if PD_C + PD_P > 0:
        resultados['PD_esperada_PC'] = (PD_C * PD_P) / (PD_C + PD_P)
    else:
        resultados['PD_esperada_PC'] = 0
    
    # Calcular índice de interferencia: Interferencia = PC - PC'
    resultados['Indice_interferencia'] = resultados['PD_PC'] - resultados['PD_esperada_PC']
    
    return resultados


# ============================================================================
# FUNCIONES DE UTILIDAD PARA CONSULTAR RESULTADOS
# ============================================================================

def obtener_resumen_indices(resultados):
    """
    Obtiene un resumen legible de todos los índices calculados
    
    Args:
        resultados: Dict devuelto por calcular_puntuaciones_directas
        
    Returns:
        str: Texto con resumen formateado de todos los índices
    """
    lineas = []
    lineas.append("=" * 70)
    lineas.append("RESUMEN DE Stroop")
    lineas.append("=" * 70)
    lineas.append("")
    
    # Índices
    lineas.append("Puntuaciones directas")
    lineas.append(f"  PD_P: {resultados['PD_P']}")
    lineas.append(f"  PD_C: {resultados['PD_C']}")
    lineas.append(f"  PD_PC: {resultados['PD_PC']}")
    lineas.append(f"  Indice_interferencia: {resultados['Indice_interferencia']}")
    lineas.append(f"  PD_E: {resultados['PD_E']}")
    lineas.append("")

    
    return "\n".join(lineas)