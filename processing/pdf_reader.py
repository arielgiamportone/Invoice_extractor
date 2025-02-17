import fitz

def load_pdf_page(file_path, page_num=0):
    """Carga una página específica de un PDF y retorna el documento y la página."""
    doc = fitz.open(file_path)
    page = doc.load_page(page_num)
    return doc, page

# Ejemplo de renderización usando tkinter (suponiendo que 'canvas' es un tkinter.Canvas)
from PIL import Image, ImageTk
from io import BytesIO

def render_pdf_preview(page, canvas):
    """Renderiza una página del PDF en el canvas de la GUI."""
    pix = page.get_pixmap()
    # Convertir el pixmap a bytes y luego a imagen PIL
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    tk_img = ImageTk.PhotoImage(img)
    # Usar create_image para mostrar la imagen en el canvas
    canvas.create_image(0, 600, image=tk_img, anchor='nw')
    # Guardar la referencia en el canvas para evitar que se elimine
    canvas.image = tk_img