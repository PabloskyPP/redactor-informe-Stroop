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
    
    Args:
        datos: Dict con 'edad' y 'datos_stroop'
        
    Returns:
        dict: Diccionario con todas las puntuaciones directas y por fila.
              Incluye estructuras explícitas para:
                - Respuestas dadas
                - Respuestas  correctas
    """
    df = datos['datos_stroop']
    
    # Extraer series del campo 'Ensayo' (P, C y PC)
    # La serie es la primera letra del ensayo
    df['serie'] = df['Ensayo'].str[0]
    
    # Diccionario para almacenar las puntuaciones por serie
    series = ['P', 'C', 'PC', 'E']
    
    resultados = {
        'edad': datos['edad'],
        'sub_num': datos.get('sub_num'),
        'nombre_completo': datos.get('nombre_completo'),
        'nombre': datos.get('nombre'),
        'datos_stroop': df,
        
        # Puntuaciones directas por serie (número de aciertos)
        'PD_P': 0,
        'PD_C': 0,
        'PD_PC': 0,
        'PD_E': 0,
        # Índices de interferencia
        'Indice_interferencia': 0,
        
        # Puntuaciones esperadas
        'PD_esperada_P': 0,
        'PD_esperada_C': 0,
        'PD_esperada_PC': 0,
    }

    # Calcular INTERFERENCIA = PC - PC' (puntuación esperada en PC) = PC - CxP /C+P
    # Calcular puntuaciones directas por serie
    for serie in series:
        resultados[f'PD_{serie}'] = df[df['serie'] == serie]['Correcto'].sum()
    
    # Calcular puntuaciones esperadas
    PD_P = resultados['PD_P']
    PD_C = resultados['PD_C']
    if PD_C + PD_P > 0:
        resultados['PD_esperada_PC'] = (PD_C * PD_P) / (PD_C + PD_P)
    else:
        resultados['PD_esperada_PC'] = 0
    
    # Calcular índice de interferencia
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