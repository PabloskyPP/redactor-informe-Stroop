"""
Programa principal para generar informe del test Raven
"""
import sys
import os
from lector_datos import leer_datos_excel, calcular_puntuaciones_directas
from reglas_psicometricas import obtener_puntuaciones_tipicas
from generador_docx import crear_informe_docx, guardar_informe
from generador_pdf import generar_pdf_desde_docx


def main():
    """
    Función principal que ejecuta todo el proceso
    """
    # Configuración
    RUTA_EXCEL = r"C:\Users\Pablo\OneDrive\Escritorio\data\Pablo Prada Campello.xlsx"
    NOMBRE_CASO = "Pablo"  # Nombre de fallback si no está en el Excel (se usa sub_num si está disponible)
    
    # Determinar la ruta de la carpeta informes_generados
    script_dir = os.path.dirname(os.path.abspath(__file__))
    carpeta_informes = os.path.join(script_dir, "informes_generados")
    
    # Crear la carpeta informes_generados si no existe
    os.makedirs(carpeta_informes, exist_ok=True)
    
    # Las rutas de salida se construirán después de leer los datos (usando nombre_completo)
    
    print("=" * 70)
    print("GENERADOR DE INFORME STROOP")
    print("=" * 70)
    print()
    print(f"Carpeta de informes: {carpeta_informes}")
    print()
    
    # Paso 1: Leer datos del Excel
    print("Paso 1: Leyendo datos del archivo Excel...")
    try:
        datos = leer_datos_excel(RUTA_EXCEL)
        print(f"    Edad del evaluado: {datos['edad']} años")
        if datos.get('nombre_completo'):
            print(f"    Nombre completo: {datos['nombre_completo']}")
            print(f"    Nombre: {datos['nombre']}")
        print(f"    Datos del test Stroop cargados correctamente")
    except Exception as e:
        print(f"   X Error al leer el archivo: {e}")
        sys.exit(1)
    
    print()

    # Construir rutas de salida dentro de informes_generados usando nombre_completo del Excel
    nombre_base = "Informe_Stroop"
    nombre_caso = datos.get('nombre_completo') or NOMBRE_CASO
    RUTA_SALIDA_DOCX = os.path.join(carpeta_informes, f"{nombre_base}_{nombre_caso}.docx")
    RUTA_SALIDA_PDF = os.path.join(carpeta_informes, f"{nombre_base}_{nombre_caso}.pdf")
    
    # Paso 2: Calcular puntuaciones directas
    print("Paso 2: Calculando puntuaciones directas...")
    try:
        resultados = calcular_puntuaciones_directas(datos)
        print(f"    Puntuación directa P: {resultados['PD_P']}")
        print(f"    Puntuación directa C: {resultados['PD_C']}")
        print(f"    Puntuación directa PC: {resultados['PD_PC']}")
        print(f"    Puntuación directa E (errores): {resultados['PD_E']}")
        print(f"    Índice de interferencia (INT): {resultados['Indice_interferencia']:.2f}")
    except Exception as e:
        print(f"   X Error al calcular puntuaciones: {e}")
        sys.exit(1)
    
    print()
    
    # Paso 3: Obtener puntuaciones típicas (clasificaciones)
    print("Paso 3: Obteniendo puntuaciones típicas (baremos)...")
    try:
        clasificaciones = obtener_puntuaciones_tipicas(resultados)

        print(f"    PT_P: {resultados['PT_P']} (Clasificación: {resultados['Clasificacion_P']})")
        print(f"    PT_C: {resultados['PT_C']} (Clasificación: {resultados['Clasificacion_C']})")
        print(f"    PT_PC: {resultados['PT_PC']} (Clasificación: {resultados['Clasificacion_PC']})")
        print(f"    PT_INT: {resultados['PT_INT']} (Clasificación: {resultados['Clasificacion_INT']})")
        print(f"    Clasificación E: {resultados['Clasificacion_E']}")
    except Exception as e:
        print(f"   X Error al obtener clasificaciones: {e}")
        sys.exit(1)
    
    print()
    
    # Paso 4: Generar informe DOCX
    print("Paso 4: Generando informe en formato Word...")
    try:
        documento = crear_informe_docx(resultados, clasificaciones, nombre_caso)
        print(f"    Documento generado correctamente")
    except Exception as e:
        print(f"   X Error al generar el documento: {e}")
        sys.exit(1)
    
    print()
    
    # Paso 5: Guardar el informe Word
    print("Paso 5: Guardando informe Word...")
    try:
        guardar_informe(documento, RUTA_SALIDA_DOCX)
        print(f"    Informe Word guardado en: {RUTA_SALIDA_DOCX}")
    except Exception as e:
        print(f"   X Error al guardar el informe Word: {e}")
        sys.exit(1)
    
    print()
    
    # Paso 6: Generar PDF desde el documento Word
    print("Paso 6: Generando informe en formato PDF...")
    try:
        generar_pdf_desde_docx(RUTA_SALIDA_DOCX, RUTA_SALIDA_PDF, verbose=False)
        print(f"    Informe PDF generado en: {RUTA_SALIDA_PDF}")
    except Exception as e:
        print(f"   X Error al generar el informe PDF: {e}")
        print(f"   ! El informe Word se generó correctamente, pero la conversión a PDF falló")
        # No salir con error, el informe Word está disponible
    
    print()
    print("=" * 70)
    print("PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 70)
    print()
    print(f"Informe Word: {RUTA_SALIDA_DOCX}")
    if os.path.exists(RUTA_SALIDA_PDF):
        print(f"Informe PDF:  {RUTA_SALIDA_PDF}")


if __name__ == "__main__":
    main()
