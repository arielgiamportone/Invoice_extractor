import os
import re
import json
import fitz
import pytesseract
from PIL import Image
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

# Configurar la ruta de Tesseract (usando la variable de entorno o la ruta por defecto en Windows)
tesseract_path = os.environ.get('TESSERACT_PATH', r'C:\Program Files\Tesseract-OCR\tesseract.exe')
pytesseract.pytesseract.tesseract_cmd = tesseract_path

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
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def extract_from_pdf(self, pdf_path):
        """
        Extrae los campos definidos en la plantilla a partir del PDF.
        
        Para cada campo, se extrae el texto de la zona indicada en el PDF.  
        Si el campo es de extracción múltiple, se procesa por filas.
        """
        results = {}
        # Uso de context manager para asegurar el cierre del documento
        with fitz.open(pdf_path) as doc:
            for field in self.template.get("fields", []):
                page_num = field.get("page", 0)
                page = doc.load_page(page_num)
                
                if field.get("multiple", False):
                    results[field["name"]] = self._extract_multiple(page, field)
                else:
                    text = self._extract_text(page, field["coordinates"])
                    results[field["name"]] = self._clean_text(text, field)
        return results
    
    def _extract_multiple(self, page, field):
        """
        Extrae campos repetidos (por ejemplo, filas de una tabla) de una zona específica de la página.
        
        Se espera que en el diccionario del campo existan:
          - "coordinates": lista o tupla con [x1, y1, x2, y2] (definiendo la columna base)
          - "start_y": posición inicial en y para comenzar la extracción
          - "end_y": posición final en y para detener la extracción
          - "row_height": altura de cada fila a extraer
          - "row_spacing": (opcional) espaciado entre filas (por defecto se usa el valor de row_height)
        """
        items = []
        current_y = field.get("start_y", 0)
        end_y = field.get("end_y")
        row_height = field.get("row_height")
        row_spacing = field.get("row_spacing", row_height)
        
        if end_y is None or row_height is None:
            raise ValueError("Para la extracción múltiple se requieren 'end_y' y 'row_height' en la plantilla.")
        
        # Extraer cada fila dentro de la zona definida
        while current_y < end_y:
            # Se usan los valores x1 y x2 de la definición base de la columna
            x1, _, x2, _ = field["coordinates"]
            rect = fitz.Rect(x1, current_y, x2, current_y + row_height)
            text = self._extract_text(page, rect)
            if text:
                items.append(self._clean_text(text, field))
            current_y += row_spacing
        
        return items
    
    def _clean_text(self, text, field):
        """
        Realiza una limpieza básica del texto extraído y aplica validaciones según el tipo de campo.
        """
        text = text.replace('\n', ' ').strip()
        field_type = field.get("type", "").lower()
        if field_type == "monto":
            return self._clean_currency(text)
        elif field_type == "fecha":
            return self._clean_date(text)
        return text
    
    def _clean_currency(self, text):
        """
        Limpia y deja solo los dígitos y los caracteres permitidos para montos.
        """
        return ''.join(c for c in text if c.isdigit() or c in [',', '.'])
    
    def _clean_date(self, text):
        """
        Implementa la lógica de conversión de fecha si es necesario.
        Actualmente retorna el texto sin modificar.
        """
        return text
    
    def _extract_text(self, page, coords):
        """
        Extrae texto de la página usando las coordenadas especificadas.
        
        El parámetro 'coords' puede ser:
          - Una secuencia (lista/tupla) con [x1, y1, x2, y2]
          - Un objeto fitz.Rect
        """
        # Si ya es un objeto fitz.Rect, usarlo directamente; de lo contrario, convertir usando float
        if isinstance(coords, fitz.Rect):
            rect = coords
        else:
            try:
                rect = fitz.Rect(*map(float, coords))
            except Exception as e:
                raise ValueError(f"Coordenadas inválidas: {coords}. Error: {e}")
        
        # Primer intento: extracción nativa de texto
        text = page.get_text("text", clip=rect).strip()
        if text:
            return text
        
        # Si no se extrajo texto, se recurre a OCR (útil para PDFs escaneados)
        pix = page.get_pixmap(clip=rect)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ocr_text = pytesseract.image_to_string(img).strip()
        # Puedes agregar logging para indicar que se usó OCR
        return ocr_text
    
    def extract_tables(self, pdf_path, template):
        """
        Extrae tablas a partir de la definición dada en la plantilla.
        Usa pdfminer para recorrer los elementos de la página (teniendo en cuenta que usa numeración 1-based).
        """
        tables_data = {}
    
        for table_def in template.get('tables', []):
            table_elements = []
            # Recorre las páginas usando pdfminer
            for page in extract_pages(pdf_path):
                # Ajuste: pdfminer usa numeración 1-based, por ello se suma 1 al valor de la plantilla
                if page.page_number == table_def['page'] + 1:
                    for element in page:
                        if isinstance(element, LTTextContainer):
                            x0, y0, x1, y1 = element.bbox
                            # Se comprueba que el contenedor se encuentre enteramente dentro del área definida
                            if (x0 >= table_def['coordinates']['x0'] and
                                x1 <= table_def['coordinates']['x1'] and
                                y0 >= table_def['coordinates']['y0'] and
                                y1 <= table_def['coordinates']['y1']):
                                table_elements.append({
                                    'text': element.get_text().strip(),
                                    'x0': x0,
                                    'y0': y0,
                                    'x1': x1,
                                    'y1': y1
                                })
        
            # Procesar la estructura de la tabla a partir de los elementos encontrados
            structured_data = self.structure_table_data(table_elements, table_def['columns'])
            tables_data[table_def['name']] = structured_data
    
        return tables_data
    
    def structure_table_data(self, elements, columns):
        """
        Agrupa elementos de texto en filas y los asigna a las columnas definidas.
        
        Actualmente, se agrupan los elementos por su posición vertical redondeada.  
        Para mayor robustez, se podría implementar una agrupación usando una tolerancia.
        """
        rows = {}
        for elem in elements:
            # Agrupar por posición vertical; se usa round() (se puede ajustar la tolerancia)
            row_key = round(elem['y0'])
            rows.setdefault(row_key, []).append(elem)
    
        # Ordenar filas de arriba a abajo (dependiendo del origen de coordenadas en el PDF)
        sorted_rows = sorted(rows.values(), key=lambda r: -r[0]['y0'])
    
        final_data = []
        for row in sorted_rows:
            # Ordenar elementos de izquierda a derecha
            row_items = sorted(row, key=lambda x: x['x0'])
            row_data = {}
            for col in columns:
                cell_text = []
                for item in row_items:
                    # Se asigna el contenido del elemento si se encuentra dentro de los límites de la columna
                    if item['x0'] >= col['x0'] and item['x1'] <= col['x1']:
                        cell_text.append(item['text'])
                row_data[col['name']] = ' '.join(cell_text).strip()
            final_data.append(row_data)
    
        return final_data
    
    def process_table(self, table_text, columns):
        """
        Procesa el texto de una tabla utilizando las definiciones de columna.
        
        Cada columna se define mediante:
          - "name": nombre del campo.
          - "pattern": expresión regular para extraer el valor de esa columna.
        
        Retorna una lista de diccionarios, donde cada diccionario representa una fila de la tabla.
        """
        rows = table_text.strip().splitlines()
        table_data = []
        for row in rows:
            row_data = {}
            for col in columns:
                col_name = col.get("name")
                pattern = col.get("pattern")
                if col_name and pattern:
                    match = re.search(pattern, row)
                    row_data[col_name] = match.group(0).strip() if match else ""
            if row_data:
                table_data.append(row_data)
        return table_data
