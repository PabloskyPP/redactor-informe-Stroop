"""
Módulo con las reglas psicométricas y baremos del test Stroop
"""

# Baremos base para el grupo de edad 16-44 años
# Estos baremos se ajustan para otros grupos de edad
BAREMOS_BASE = {
    80: {"P": 179, "C": 121, "PC": 83, "INT": 29.9},
    78: {"P": 175, "C": 118, "PC": 81, "INT": 28.1},
    76: {"P": 171, "C": 115, "PC": 79, "INT": 26.2},
    74: {"P": 167, "C": 113, "PC": 76, "INT": 24.4},
    72: {"P": 163, "C": 110, "PC": 74, "INT": 22.6},
    70: {"P": 159, "C": 107, "PC": 72, "INT": 20.8},
    68: {"P": 155, "C": 104, "PC": 70, "INT": 19.0},
    66: {"P": 151, "C": 101, "PC": 68, "INT": 17.2},
    64: {"P": 147, "C": 99,  "PC": 65, "INT": 15.4},
    62: {"P": 143, "C": 96, "PC": 63, "INT": 13.6},
    60: {"P": 139, "C": 93, "PC": 61, "INT": 11.8},
    58: {"P": 135, "C": 90, "PC": 59, "INT": 10.0},
    56: {"P": 131, "C": 87, "PC": 57, "INT": 8.1},
    54: {"P": 127, "C": 85, "PC": 54, "INT": 6.3},
    52: {"P": 123, "C": 82, "PC": 52, "INT": 4.5},
    50: {"P": 119, "C": 79, "PC": 50, "INT": 2.7},
    48: {"P": 115, "C": 76, "PC": 48, "INT": 0.9},
    46: {"P": 111, "C": 73, "PC": 46, "INT": -0.9},
    44: {"P": 107, "C": 71, "PC": 43, "INT": -2.7},
    42: {"P": 103, "C": 68, "PC": 41, "INT": -4.5},
    40: {"P": 99,  "C": 65, "PC": 39, "INT": -6.3},
    38: {"P": 95,  "C": 62, "PC": 37, "INT": -8.2},
    36: {"P": 91,  "C": 59, "PC": 35, "INT": -10.0},
    34: {"P": 87,  "C": 57, "PC": 32, "INT": -11.8},
    32: {"P": 83,  "C": 54, "PC": 30, "INT": -13.6},
    30: {"P": 79,  "C": 51, "PC": 28, "INT": -15.4},
    28: {"P": 75,  "C": 48, "PC": 26, "INT": -17.2},
    26: {"P": 71,  "C": 45, "PC": 24, "INT": -19.0},
    24: {"P": 67,  "C": 43, "PC": 21, "INT": -20.8},
    22: {"P": 63,  "C": 40, "PC": 19, "INT": -22.6},
    20: {"P": 59,  "C": 37, "PC": 17, "INT": -24.4}
}


def obtener_ajuste_edad(edad):
    """
    Retorna los ajustes a aplicar a las PD según la edad
    Basado en las tablas del manual de Stroop
    
    Args:
        edad: Edad del participante
        
    Returns:
        dict: Diccionario con ajustes para cada índice
    """
    if 16 <= edad <= 44:
        return {"P": 0, "C": 0, "PC": 0, "INT": 0}
    elif 45 <= edad <= 64:
        return {"P": -8, "C": -4, "PC": -5, "INT": 0}
    elif edad >= 65:
        return {"P": -14, "C": -11, "PC": -15, "INT": 0}
    else:
        # Menores de 16 años, sin baremos específicos
        return {"P": 0, "C": 0, "PC": 0, "INT": 0}


def obtener_PT_desde_PD(PD_ajustada, indice):
    """
    Obtiene la PT (Puntuación Típica) a partir de la PD ajustada
    
    Args:
        PD_ajustada: Puntuación directa ajustada por edad
        indice: 'P', 'C', 'PC' o 'INT'
        
    Returns:
        int: Puntuación típica (20-80)
    """
    # Buscar en los baremos de mayor a menor PT
    for pt in sorted(BAREMOS_BASE.keys(), reverse=True):
        if PD_ajustada >= BAREMOS_BASE[pt][indice]:
            return pt
    
    # Si es menor que el mínimo, devolver PT=20
    return 20


def clasificar_PT(PT):
    """
    Clasifica una PT en bajo, normal o alto
    
    Args:
        PT: Puntuación típica
        
    Returns:
        str: 'bajo', 'normal' o 'alto'
    """
    if PT <= 30:
        return 'bajo'
    elif PT < 70:
        return 'normal'
    else:
        return 'alto'


def clasificar_errores(num_errores):
    """
    Clasifica el número de errores según los criterios del manual
    
    E (Errores):
    - Bajo: <= 1
    - Normal: 2-4
    - Alto: > 4
    
    Args:
        num_errores: Número total de errores
        
    Returns:
        str: 'bajo', 'normal' o 'alto'
    """
    if num_errores <= 1:
        return 'bajo'
    elif num_errores <= 4:
        return 'normal'
    else:
        return 'alto'


def obtener_puntuaciones_tipicas(resultados):
    """
    Obtiene las puntuaciones típicas y clasificaciones basadas en las puntuaciones directas
    
    Args:
        resultados: Dict con puntuaciones directas y edad
        
    Returns:
        dict: Diccionario con clasificaciones en 3 niveles
    """
    edad = resultados['edad']
    
    # Obtener ajustes por edad
    ajustes = obtener_ajuste_edad(edad)
    
    # Procesar cada índice (P, C, PC, INT)
    indices = ['P', 'C', 'PC', 'INT']
    
    for indice in indices:
        # Obtener PD original
        if indice == 'INT':
            PD = resultados['Indice_interferencia']
        else:
            PD = resultados[f'PD_{indice}']
        
        # Ajustar PD por edad
        PD_ajustada = PD + ajustes[indice]
        
        # Obtener PT desde la PD ajustada
        PT = obtener_PT_desde_PD(PD_ajustada, indice)
        
        # Clasificar
        clasificacion = clasificar_PT(PT)
        
        # Guardar en resultados
        resultados[f'PT_{indice}'] = PT
        resultados[f'Clasificacion_{indice}'] = clasificacion
    
    # Clasificar errores (E)
    clasificacion_E = clasificar_errores(resultados['PD_E'])
    resultados['Clasificacion_E'] = clasificacion_E
    
    # Crear diccionario de clasificaciones para retornar
    clasificaciones = {
        'P': resultados['Clasificacion_P'],
        'C': resultados['Clasificacion_C'],
        'PC': resultados['Clasificacion_PC'],
        'INT': resultados['Clasificacion_INT'],
        'E': resultados['Clasificacion_E'],
    }
    
    return clasificaciones
