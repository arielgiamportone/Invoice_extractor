import tkinter as tk
from gui.main_window import MainWindow

def main():
    root = tk.Tk()
    root.title("Extractor de Facturas")  # Puedes definir el título aquí o desde MainWindow
    # Opcional: root.iconbitmap('ruta/a/icono.ico')
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
