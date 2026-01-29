"""
Módulo con las reglas psicométricas y baremos del test Stroop
"""


def obtener_baremos_edad(edad):
    """
    Retorna los baremos apropiados según la edad
    Basado en las tablas A.1 a A.11 del manual
    """
    if 16 <= edad <= 44:
        return 'PT_entre 16_y_44'
    elif 45  <= edad <= 64:
        return 'PT_45_a_64'
    elif edad >= 65:
        return 'PT_65_o_mas'
    else:
        return 'Error participante menor de 16 años, grupo sin baremos. Los baremos del manual son experimentales, cabe contrastarlos con datos poblacionales propios'

        if  45  <= edad <= 64:
            baremo = baremos[grupo]
            baremo['P'] += 8
            baremo['C'] += 4
            baremo['PC'] += 5

        if edad >= 65:
            baremo = baremos[grupo]
            baremo['P'] += 14
            baremo['C'] += 11
            baremo['PC'] += 15    


def clasificar_PD(PD_total):
    """
    Clasifica la PD en 3 niveles según PT y edad
    
    Niveles:
    - 'bajo': PT ≤ 30
    - 'normal': PT 31-69
    - 'muy alto': PT ≥ 70
    """ 
baremos = {
    "pt80": {"P": 179, "C": 121, "PC": 83, "INT": 29.9},
    "pt78": {"P": 175, "C": 118, "PC": 81, "INT": 28.1},
    "pt76": {"P": 171, "C": 115, "PC": 79, "INT": 26.2},
    "pt74": {"P": 167, "C": 113, "PC": 76, "INT": 24.4},
    "pt72": {"P": 163, "C": 110, "PC": 74, "INT": 22.6},
    "pt70": {"P": 159, "C": 107, "PC": 72, "INT": 20.8},
    "pt68": {"P": 155, "C": 104, "PC": 70, "INT": 19.0},
    "pt66": {"P": 151, "C": 101, "PC": 68, "INT": 17.2},
    "pt64": {"P": 147, "C": 99,  "PC": 65, "INT": 15.4},
    "pt62": {"P": 143, "C": 96, "PC": 63, "INT": 13.6},
    "pt60": {"P": 139, "C": 93, "PC": 61, "INT": 11.8},
    "pt58": {"P": 135, "C": 90, "PC": 59, "INT": 10.0},
    "pt56": {"P": 131, "C": 87, "PC": 57, "INT": 8.1},
    "pt54": {"P": 127, "C": 85, "PC": 54, "INT": 6.3},
    "pt52": {"P": 123, "C": 82, "PC": 52, "INT": 4.5},
    "pt50": {"P": 119, "C": 79, "PC": 50, "INT": 2.7},
    "pt48": {"P": 115, "C": 76, "PC": 48, "INT": 0.9},
    "pt46": {"P": 111, "C": 73, "PC": 46, "INT": -0.9},
    "pt44": {"P": 107, "C": 71, "PC": 43, "INT": -2.7},
    "pt42": {"P": 103, "C": 68, "PC": 41, "INT": -4.5},
    "pt40": {"P": 99,  "C": 65, "PC": 39, "INT": -6.3},
    "pt38": {"P": 95,  "C": 62, "PC": 37, "INT": -8.2},
    "pt36": {"P": 91,  "C": 59, "PC": 35, "INT": -10.0},
    "pt34": {"P": 87,  "C": 57, "PC": 32, "INT": -11.8},
    "pt32": {"P": 83,  "C": 54, "PC": 30, "INT": -13.6},
    "pt30": {"P": 79,  "C": 51, "PC": 28, "INT": -15.4},
    "pt28": {"P": 75,  "C": 48, "PC": 26, "INT": -17.2},
    "pt26": {"P": 71,  "C": 45, "PC": 24, "INT": -19.0},
    "pt24": {"P": 67,  "C": 43, "PC": 21, "INT": -20.8},
    "pt22": {"P": 63,  "C": 40, "PC": 19, "INT": -22.6},
    "pt20": {"P": 59,  "C": 37, "PC": 17, "INT": -24.4}
}

    grupo = obtener_baremos_edad(edad)
    if grupo not in baremos:
        return 'normal', 50
    
    baremo = baremos[grupo]
    
    # Clasificar según PT
    if PT <= 30:
        'bajo', nivel
    elif PT < 70:
        return 'normal', nivel
    elif PT >= 90:
        return 'alto', nivel


def obtener_puntuaciones_tipicas(resultados):
    """
    Obtiene las puntuaciones típicas basadas en las puntuaciones directas
    
    Args:
        resultados: Dict con puntuaciones directas y edad
        
    Returns:
        dict: Diccionario con clasificaciones en 3 niveles
    """
    edad = resultados['edad']
    
    # Calcular índices de discrepancia
    resultados = calcular_indices_discrepancia(resultados)
    
    # Clasificar PD y obtener nivel
    clasificacion, nivel = clasificar_PD(resultados['PD'], edad)
    
    # Guardar percentil en resultados
    resultados['Nivel'] = nivel
    resultados['Clasificacion_PD'] = clasificacion
    
    clasificaciones = {
        'PD': clasificacion,
        'Nivel': nivel,
    }
    
    return clasificaciones
