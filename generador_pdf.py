"""
Módulo para convertir documentos Word a PDF
Este módulo proporciona funciones para exportar archivos .docx a formato PDF
manteniendo el formato del documento original.
"""
import os
import sys
import subprocess


def convertir_docx_a_pdf(ruta_docx, ruta_pdf=None):
    """
    Convierte un archivo Word (.docx) a PDF usando LibreOffice en modo headless
    
    Args:
        ruta_docx: Ruta al archivo .docx de entrada
        ruta_pdf: Ruta opcional para el archivo PDF de salida.
                  Si no se especifica, se generará en la misma ubicación con extensión .pdf
    
    Returns:
        str: Ruta al archivo PDF generado
        
    Raises:
        FileNotFoundError: Si el archivo .docx no existe
        RuntimeError: Si la conversión falla
    """
    # Verificar que el archivo .docx existe
    if not os.path.exists(ruta_docx):
        raise FileNotFoundError(f"No se encuentra el archivo: {ruta_docx}")
    
    # Si no se especifica ruta de salida, usar la misma ubicación
    if ruta_pdf is None:
        ruta_pdf = os.path.splitext(ruta_docx)[0] + '.pdf'
    
    # Obtener el directorio de salida
    directorio_salida = os.path.dirname(ruta_pdf)
    if not directorio_salida:
        directorio_salida = os.path.dirname(ruta_docx)
    
    # Asegurar que el directorio de salida existe
    os.makedirs(directorio_salida, exist_ok=True)

    try:
        if sys.platform != 'win32':
            print("    La conversión a PDF requiere Microsoft Word en Windows.")
            return False

        try:
            import win32com.client
        except ImportError:
            print("    pywin32 no está instalado. Instala con: pip install pywin32")
            return False

        # Crear instancia de Word
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False

        # Abrir el documento
        doc = word.Documents.Open(ruta_docx)

        # Guardar como PDF (formato 17 es PDF)
        doc.SaveAs(ruta_pdf, FileFormat=17)

        # Cerrar documento y Word
        doc.Close()
        word.Quit()

        return os.path.exists(ruta_pdf)

    except subprocess.TimeoutExpired:
        print("Error: La conversión de DOCX a PDF excedió el tiempo límite")
        return False
    except Exception as e:
        print(f"Error al convertir DOCX a PDF: {e}")
        return False

def generar_pdf_desde_docx(ruta_docx, ruta_pdf=None, verbose=True):
    """
    Función de alto nivel para generar un PDF desde un archivo Word
    
    Args:
        ruta_docx: Ruta al archivo .docx de entrada
        ruta_pdf: Ruta opcional para el archivo PDF de salida
        verbose: Si True, imprime mensajes de progreso
    
    Returns:
        str: Ruta al archivo PDF generado
    """
    try:
        # Resolver ruta de salida
        if ruta_pdf is None:
            ruta_pdf = os.path.splitext(ruta_docx)[0] + '.pdf'

        print("    Convirtiendo DOCX a PDF...")
        if not convertir_docx_a_pdf(ruta_docx, ruta_pdf):
            print("    Error al convertir DOCX a PDF")
            return False

        if verbose:
            print(f"    PDF generado en: {ruta_pdf}")
        return ruta_pdf

    except Exception as e:
        if verbose:
            print(f"   X Error al generar PDF: {e}", file=sys.stderr)
        raise
