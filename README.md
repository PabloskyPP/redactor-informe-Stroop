# Generador de Informes Stroop

Este programa procesa los datos del Stroop y genera un informe profesional en formato Word (.docx) y PDF.

## 📋 Requisitos

- Python 3.7 o superior
- LibreOffice (para la conversión a PDF)
- Las siguientes bibliotecas de Python:
  - pandas
  - openpyxl
  - python-docx
  - Pillow

## 🚀 Instalación

1. Instalar las dependencias de Python:
```bash
pip install -r requirements.txt
```

2. Instalar LibreOffice (para generación de PDF):
   - **Linux (Ubuntu/Debian)**:
     ```bash
     sudo apt-get install libreoffice-writer
     ```
   - **macOS**:
     ```bash
     brew install --cask libreoffice
     ```
   - **Windows**: Descargar e instalar desde [libreoffice.org](https://www.libreoffice.org/)

## 📁 Estructura del Proyecto

```
proyecto_raven/
│
├── main.py                      # Programa principal
├── lector_datos.py              # Lee datos del Excel
├── reglas_psicometricas.py      # Aplica baremos según edad
├── textos.py                    # Textos del informe
├── generador_docx.py            # Genera el documento Word
├── generador_pdf.py             # Convierte Word a PDF
├── requirements.txt             # Dependencias del proyecto
├── informes_generados/          # Carpeta para informes (ignorada en git)
└── README.md                    # Este archivo
```

## 📊 Formato del Archivo Excel

El archivo Excel debe tener la siguiente estructura:

### Hoja "info"
- `age`: Edad del evaluado (obligatorio)
- `sub_num`: Nombre completo del encuestado (opcional pero recomendado)
  - Puede ser un identificador simple o nombre completo con apellidos
  - El programa extraerá automáticamente:
    - `nombre_completo`: valor completo de sub_num
    - `nombre`: solo el primer token (antes del primer espacio)

### Hoja "Stroop"
Debe contener las siguientes columnas:
- `Ensayo`: 3 series: P, C, PC
- `PD`: Indica la puntuación directa en la serie
- `Errores`: Indica el número de errores durante la serie


## 🎯 Cómo Usar el Programa

### Paso 1: Configurar las rutas

Editar el archivo `main.py` y modificar las siguientes variables:

```python
RUTA_EXCEL = r"C:\Users\Pablo\OneDrive\Escritorio\data\María.xlsx"
RUTA_EXCEL = r"C:\Users\Pablo\OneDrive\Escritorio\data\María.xlsx"
NOMBRE_CASO = "María"  # Nombre de fallback si no está en el Excel
SCALE_FACTOR = 0.5  # Factor de escalado para la imagen (0.5 = 50% del tamaño original)
```

**Nota**: Los informes se generarán automáticamente en la carpeta `informes_generados/` dentro del directorio del proyecto.

### Paso 2: Ejecutar el programa

```bash
python main.py
```

El programa ejecutará automáticamente los siguientes pasos:
1. Crear la carpeta `informes_generados/` si no existe
2. Leer datos del archivo Excel
3. Calcular puntuaciones directas
4. Obtener puntuaciones típicas (baremos)
5. Generar informe en formato Word
6. Guardar el informe Word en `informes_generados/`
7. Convertir automáticamente a PDF
8. Guardar el informe PDF en `informes_generados/`

### Archivos Generados

Después de ejecutar el programa, encontrará los siguientes archivos en la carpeta `informes_generados/`:
- `Informe_Stroop_Resultado.docx` - Informe en formato Word
- `Informe_Stroop_Resultado.pdf` - Informe en formato PDF

## 📈 Puntuaciones Calculadas

El programa calcula las siguientes puntuaciones directas:

### Por cada fila (P, C, PC):
- **PD de cada serie**: Número de aciertos por serie
### Totales:
- **Índice de interferencia**: Diferencia, resto entre la puntuación PC-PC' = PC- CxP /C+P

## 🎓 Interpretación de Puntuaciones Típicas

El programa clasifica automáticamente las puntuaciones según baremos por edad:

### PD:
- **bajo**: Por debajo del PT30
- **normal**: Entre el PT30 y el PT70
- **alto**: Por encima del Pc70



## 📄 Contenido del Informe

El informe generado incluye:

1. **Portada** (Primera página)
   - Título del test: "Test Stroop"
   - Nombre completo del encuestado (extraído de `sub_num`)
   - Edad del evaluado
   - Fecha de aplicación (si está disponible)
   - Fecha del informe (generada automáticamente)
   - Nota de confidencialidad
2. **Introducción** predeterminada y constante
3. **Descripción de la prueba** y variables técnicas PD e índice de discrepancia
4. **Imagen de baremos** (escalada según `SCALE_FACTOR`)
5. **Tabla de resultados** con puntuaciones directas y clasificaciones
6. **Interpretación detallada** de los resultados del participante
7. **Síntesis del caso**

Los informes se generan en dos formatos:
- **Word (.docx)**: Documento editable con formato completo
- **PDF**: Versión no editable derivada del Word, manteniendo el formato original

## ⚙️ Configuración Avanzada

Ajustar según las necesidades de visualización del informe.

## 🐛 Solución de Problemas

### Error: "No se encuentra el archivo Excel"
- Verificar que la ruta en `RUTA_EXCEL` es correcta
- Usar rutas absolutas o relativas correctas
- Verificar que el archivo existe

### Error: "Columna no encontrada"
- Verificar que la hoja "info" tiene la columna "age"
- Verificar que la hoja "Ravens Matrices" tiene todas las columnas requeridas

### Error al convertir a PDF
Si la conversión a PDF falla pero el archivo Word se generó correctamente:
- Verificar que LibreOffice está instalado: `libreoffice --version`
- En Linux/Mac, asegurarse de que LibreOffice está en el PATH
- En Windows, agregar la ruta de LibreOffice al PATH del sistema
- El archivo Word estará disponible en `informes_generados/` aunque falle el PDF

### Error al instalar dependencias
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

## 📚 Baremos Utilizados

Los baremos están basados en el manual oficial de TEA 3ºEdición STROOP - Test de Colores y Palabras (Charles & Golden, Ed. Madrid 2001).

## 📞 Contacto

Para dudas o sugerencias sobre el programa, contactar al desarrollador.

## 📝 Notas Importantes

- El programa está diseñado para generar informes profesionales automáticamente
- Se recomienda revisar el informe generado antes de su uso clínico
- Los baremos son aproximaciones basadas en el manual oficial
- Siempre verificar que los datos de entrada son correctos antes de ejecutar
