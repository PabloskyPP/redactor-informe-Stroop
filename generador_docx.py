"""
Módulo para generar el informe en formato DOCX
"""
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from PIL import Image
from textos import (
    PARRAFOS_FIJOS, PARRAFOS_PT, TEXTO_ERROR_y_INT, TEXTO_FINAL_PosibleDislexia
)

def agregar_portada(doc: Document, nombre_completo: str, datos: dict) -> None:
    """
    Agrega la portada del informe
    
    Args:
        doc: Documento de Word
        nombre_completo: Nombre completo del encuestado
        datos: Diccionario con datos generales (edad, fecha_aplicacion, etc.)
    """
    # Título
    titulo = doc.add_heading('PRUEBA STROOP, TAREA DE INTERFERENCIA PALABRA-COLOR', 0)
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # Aumentar tamaño de fuente del título
    for run in titulo.runs:
        run.font.size = Pt(24)
    
    # Espacio
    doc.add_paragraph().add_run().add_break()
    
    # Información del participante
    info = doc.add_paragraph()
    info.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Nombre del encuestado
    nombre_run = info.add_run(f"Nombre del encuestado: {nombre_completo}\n")
    nombre_run.bold = True
    nombre_run.font.size = Pt(14)
    
    # Edad
    if datos.get('edad'):
        edad_run = info.add_run(f"Edad: {datos['edad']} años\n")
        edad_run.font.size = Pt(14)
    
    # Fecha de aplicación (si está disponible)
    if datos.get('fecha_aplicacion'):
        fecha_app_run = info.add_run(f"Fecha de aplicación: {datos['fecha_aplicacion']}\n")
        fecha_app_run.font.size = Pt(14)
    
    # Fecha del informe
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    fecha_informe_run = info.add_run(f"Fecha del informe: {fecha_actual}\n")
    fecha_informe_run.font.size = Pt(14)
    
    # Espacio
    doc.add_paragraph().add_run().add_break()
    
    # Nota confidencial
    nota = doc.add_paragraph()
    nota.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    nota_run = nota.add_run(f"Este es un informe de evaluación cognitiva, obtenido a partir del rendimiento de {nombre_completo} en la prueba Stroop (Tarea de Interferencia Palabra-Color).")
    doc.add_paragraph()  # Espacio
    nota.add_run("\nInforme confidencial de uso profesional y educativo")
    nota_run.italic = True
    nota_run.font.size = Pt(12)

def crear_informe_docx(resultados, clasificaciones, nombre_caso="caso", scale_factor_width=0.8, scale_factor_height=0.5):
    """
    Crea el documento Word con el informe del test Stroop
    
    Args:
        resultados: Dict con puntuaciones directas y datos del encuestado
        clasificaciones: Dict con clasificaciones (PT)
        nombre_caso: Nombre del evaluado (fallback si no está en resultados)
        
    Returns:
        Document: Documento Word generado
    """
    doc = Document()
    
    # Configurar márgenes
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Extraer nombres de resultados
    nombre_completo = resultados.get('nombre_completo') or nombre_caso
    nombre = resultados.get('nombre') or nombre_caso
    
    # Preparar datos para la portada
    datos_portada = {
        'edad': resultados.get('edad'),
        'fecha_aplicacion': resultados.get('fecha_aplicacion')
    }
    
    # 1. PORTADA (Primera página)
    agregar_portada(doc, nombre_completo, datos_portada)
    doc.add_page_break()
    
    # ========================================================================
    # TÍTULO E INTRODUCCIÓN
    # ========================================================================
    # ========================================================================
    # DESCRIPCIÓN DE LA PRUEBA
    # ========================================================================
    
    parrafo = doc.add_paragraph()
    run = parrafo.add_run(PARRAFOS_FIJOS['titulo'])
    run.bold = True
    doc.add_paragraph(PARRAFOS_FIJOS['introduccion'].format(nombre=nombre))

    parrafo = doc.add_paragraph()
    run = parrafo.add_run(PARRAFOS_FIJOS['titulo_general_prueba'])
    run.bold = True
    doc.add_paragraph(PARRAFOS_FIJOS['objetivo_prueba'])

    parrafo = doc.add_paragraph()
    run = parrafo.add_run(PARRAFOS_FIJOS['titulo_procedimiento'])
    run.bold = True
    doc.add_paragraph(PARRAFOS_FIJOS['descripcion_procedimiento1'])
    # Imagen parte P
    script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'images')
    grafico_path = os.path.join(script_dir, 'P.png')
    doc.add_paragraph(PARRAFOS_FIJOS['descripcion_procedimiento2'])
    # Imagen parte C
    script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'images')
    grafico_path = os.path.join(script_dir, 'C.png')
    doc.add_paragraph(PARRAFOS_FIJOS['descripcion_procedimiento3'])
    # Imagen parte PC
    script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'images')
    grafico_path = os.path.join(script_dir, 'PC.png')

    parrafo = doc.add_paragraph()
    run = parrafo.add_run(PARRAFOS_FIJOS['titulo_indices'])
    run.bold = True
    doc.add_paragraph(PARRAFOS_FIJOS['descripcion_indices'])
    doc.add_page_break()

    # ========================================================================
    # INSERTAR Tabla resultados
    # ========================================================================
    
    # Título de resultados
    titulo_resultados = doc.add_paragraph()
    titulo_resultados.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = titulo_resultados.add_run(PARRAFOS_FIJOS['titulo_resultados'].format(nombre_completo=nombre_completo))
    run.bold = True
    run.font.size = Pt(14)
    doc.add_paragraph()  # Espacio
    
    parrafo = doc.add_paragraph()
    run = parrafo.add_run(PARRAFOS_FIJOS['titulo_resultados'].format(nombre_completo=nombre_completo))
    run.bold = True
    doc.add_paragraph(PARRAFOS_FIJOS['texto_tabla_resultados'])

    # Tabla PDs, PTs y clasificaciones

    tabla = doc.add_table(rows=4, cols=6)
    tabla.style = 'Light Grid Accent 1'
    
    # Encabezados
    hdr_cells = tabla.rows[0].cells
    hdr_cells[0].text = ''
    hdr_cells[1].text = 'Palabras'
    hdr_cells[2].text = 'Colores'
    hdr_cells[3].text = 'Palabras-Colores'
    hdr_cells[4].text = 'Interferencia'
    hdr_cells[5].text = 'Número de errores'

    # Centrar el texto en las celdas del encabezado
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
    # Fila 1: PD
    row1 = tabla.rows[1].cells
    row1[0].text = 'Puntuación directa'
    for i, serie in enumerate(['P', 'C', 'PC', 'I', 'E']):
        row1[i+1].text = str(resultados.get(f'PD_esperada_{serie}', 0))
    
    # Fila 2: PT
    row2 = tabla.rows[2].cells
    row2[0].text = 'Puntuación típica'
    for i, serie in enumerate(['P', 'C', 'PC', 'I', 'E']):
        row2[i+1].text = str(resultados.get(f'PD_{serie}', 0))
    
    # Fila 3: clasificación
    row3 = tabla.rows[3].cells
    row3[0].text = 'Rendimiento'
    for i, serie in enumerate(['P', 'C', 'PC', 'I', 'E']):
        indice = resultados.get(f'Clasificacion_{serie}', 0)
        row3[i+1].text = str(indice)

    # Centrar el texto en las celdas de las filas 1, 2 y 3
    for row in [row1, row2, row3]:
        for cell in row:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    doc.add_paragraph()  # Espacio

    # ========================================================================
    # PÁRRAFOS CONDICIONALES
    # ========================================================================

    # ========================================================================
    # DESCRIPCIÓN DE PUNTUACIONES DIRECTAS

    # Seleccionar y añadir párrafo según clasificación
    clasificacion = resultados.get('Clasificacion_PD', 'normal')
    if clasificacion in PARRAFOS_PD:
        texto_pd = PARRAFOS_PD[clasificacion].format(Nombre=nombre, nombre=nombre)
        doc.add_paragraph(texto_pd)
    
    doc.add_paragraph()  # Espacio 


    # ========================================================================
    return doc


def guardar_informe(doc, ruta_salida):
    """
    Guarda el documento generado
    
    Args:
        doc: Documento Word
        ruta_salida: Ruta donde guardar el archivo
    """
    doc.save(ruta_salida)
    print(f"Informe generado exitosamente en: {ruta_salida}")
