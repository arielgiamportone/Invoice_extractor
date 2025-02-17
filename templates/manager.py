import json

class TemplateManager:
    def save_template(self, template, path):
        """Guarda la plantilla en formato JSON en la ruta especificada."""
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(template, f, indent=4)
        except Exception as e:
            raise IOError(f"Error al guardar la plantilla: {e}")

    def load_template(self, path):
        """Carga y retorna la plantilla desde el archivo JSON en la ruta especificada."""
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            raise IOError(f"Error al cargar la plantilla: {e}")