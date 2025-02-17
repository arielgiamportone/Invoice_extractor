import os
import re
import json
import fitz
import pytesseract
from PIL import Image
import tkinter as tk
from tkinter import filedialog

def configure_tesseract():
    """
    Configura la ruta de Tesseract OCR:
      - Intenta usar la variable de entorno 'TESSERACT_PATH'
      - Si no está definida, usa la ruta por defecto en Windows.
      - Si la ruta por defecto no existe, solicita al usuario que seleccione el ejecutable.
    Devuelve la ruta a tesseract.exe.
    """
    # Intentar obtener la ruta desde la variable de entorno
    tesseract_path = os.environ.get('TESSERACT_PATH')
    if tesseract_path and os.path.exists(tesseract_path):
        return tesseract_path

    # Ruta por defecto en Windows
    default_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(default_path):
        tesseract_path = default_path
    else:
        # Si no se encuentra, pedir al usuario que seleccione el ejecutable
        root = tk.Tk()
        root.withdraw()  # Oculta la ventana principal
        tesseract_path = filedialog.askopenfilename(
            title="Selecciona tesseract.exe",
            filetypes=[("Ejecutable", "*.exe")]
        )
        root.destroy()
        if not tesseract_path or not os.path.exists(tesseract_path):
            raise RuntimeError("No se pudo encontrar Tesseract OCR. Por favor, instálalo o configura TESSERACT_PATH.")

    return tesseract_path

# Configurar Tesseract
tesseract_path = configure_tesseract()
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Configurar TESSDATA_PREFIX, si no está definida
if not os.environ.get('TESSDATA_PREFIX'):
    tessdata_folder = os.path.join(os.path.dirname(tesseract_path), "tessdata")
    if os.path.exists(tessdata_folder):
        os.environ['TESSDATA_PREFIX'] = tessdata_folder
    else:
        print("Advertencia: No se encontró la carpeta 'tessdata' en la ruta de Tesseract.")

class PDFExtractor:
    def __init__(self, template_path):
        """
        Inicializa el extractor cargando la plantilla de extracción (en formato JSON).
        """
        self.template = self._load_template(template_path)

    @staticmethod
    def _load_template(path):
        """
        Carga y devuelve la plantilla JSON desde el archivo especificado.
        """
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            raise RuntimeError(f"Error al cargar la plantilla: {e}")

    def extract_from_pdf(self, pdf_path):
        """
        Extrae la información definida en la plantilla desde el PDF.
        
        Extrae los campos fijos y las tablas, devolviendo un diccionario con la siguiente estructura:
            {
                "fields": { nombre_campo: texto_extraído, ... },
                "tables": { nombre_tabla: [lista_de_filas, ...] }
            }
        """
        results = {}
        # Usar context manager para asegurar que el PDF se cierra correctamente
        with fitz.open(pdf_path) as doc:
            # Extraer campos fijos
            fields_data = {}
            for field in self.template.get("fields", []):
                page_num = field.get("page", 0)
                page = doc.load_page(page_num)
                text = self._extract_text(page, field["coordinates"])
                fields_data[field["name"]] = text
            results["fields"] = fields_data

            # Extraer tablas
            tables_data = {}
            for table in self.template.get("tables", []):
                page_num = table.get("page", 0)
                page = doc.load_page(page_num)
                text = self._extract_text(page, table["coordinates"])
                table_rows = self.process_table(text, table.get("columns", []))
                tables_data[table["name"]] = table_rows
            results["tables"] = tables_data

        return results

    def _extract_text(self, page, coords):
        """
        Extrae texto de la página utilizando las coordenadas especificadas.
        
        El parámetro 'coords' puede ser:
          - Un objeto fitz.Rect
          - Una secuencia (lista/tupla) con [x1, y1, x2, y2]
          
        Primero intenta la extracción nativa. Si no obtiene texto, recurre a OCR.
        """
        # Convertir coordenadas a un objeto fitz.Rect usando float para mayor precisión
        if isinstance(coords, fitz.Rect):
            rect = coords
        else:
            try:
                rect = fitz.Rect(*map(float, coords))
            except Exception as e:
                raise ValueError(f"Coordenadas inválidas: {coords}. Error: {e}")
        
        # Intento de extracción nativa
        text = page.get_text("text", clip=rect).strip()
        if text:
            return text
        
        # Si falla la extracción nativa, se utiliza OCR
        pix = page.get_pixmap(clip=rect)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ocr_text = pytesseract.image_to_string(img).strip()
        return ocr_text

    def process_table(self, table_text, columns):
        """
        Procesa el texto de una tabla dividiendo cada fila por dos o más espacios.
        
        Se asigna cada celda a la columna definida por orden. Retorna una lista de diccionarios,
        donde cada diccionario representa una fila de la tabla.
        """
        rows = table_text.strip().splitlines()
        table_data = []
        for row in rows:
            # Dividir la fila por dos o más espacios
            cells = re.split(r'\s{2,}', row.strip())
            row_data = {}
            num_defined_columns = len(columns)
            for i in range(num_defined_columns):
                col_name = columns[i].get("name")
                # Si hay más celdas que columnas definidas, se ignoran las adicionales o se pueden agrupar
                cell_value = cells[i] if i < len(cells) else ""
                row_data[col_name] = cell_value
            # Solo agregar filas que tengan al menos una celda no vacía
            if any(cell.strip() for cell in row_data.values()):
                table_data.append(row_data)
        return table_data