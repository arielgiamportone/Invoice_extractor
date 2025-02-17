# Extractor de Facturas
Extractor de Facturas es una aplicación de escritorio que permite extraer datos y tablas de facturas en formato PDF. La aplicación permite definir áreas de extracción, generar plantillas en formato JSON y, a partir de estas, procesar uno o varios archivos PDF para exportar los resultados a Excel.

## Características
Interfaz gráfica sencilla con botones e indicaciones para cargar PDFs, definir áreas de extracción y generar plantillas.
- Extracción de texto y OCR:
Utiliza PyMuPDF para extraer texto de PDFs nativos y Tesseract OCR para PDFs escaneados.
- Exportación a Excel: 
Los datos extraídos se organizan y se pueden exportar en formato Excel.
- Autónoma: 
El ejecutable es independiente, por lo que no es necesario instalar Python en la máquina del usuario.

Requisitos para el Usuario Final
Sistema Operativo: Windows (la versión empaquetada se ha probado en Windows 10).

- Tesseract OCR:
Para que la aplicación funcione correctamente, se requiere tener instalada una versión compatible de Tesseract OCR.

- Instalación recomendada: 
Descarga e instala Tesseract OCR desde el enlace anterior. La ruta por defecto es:

C:\Program Files\Tesseract-OCR\tesseract.exe

Si Tesseract no se encuentra en la ruta predeterminada, al ejecutar la aplicación se te solicitará seleccionar manualmente el ejecutable de Tesseract.

- Archivos de Entrenamiento (Language Data):
La instalación de Tesseract incluye, por defecto, los archivos de entrenamiento para el idioma inglés.

Si deseas extraer texto en otros idiomas, descarga los archivos .traineddata correspondientes (por ejemplo, para español, etc.) desde el repositorio oficial de tessdata y colócalos en la carpeta tessdata que se encuentra junto a tesseract.exe.

## Instalación y Uso


- Instalación de Tesseract OCR

Si aún no tienes instalado Tesseract OCR, descárgalo e instálalo desde este enlace.
Durante la primera ejecución, la aplicación intentará localizar Tesseract automáticamente. Si no se encuentra en la ruta predeterminada, se te solicitará que selecciones el ejecutable manualmente.
Ejecuta la Aplicación

- Descarga e Instala la Aplicación

Descarga el archivo ejecutable ExtractorDeFacturas.exe (ubicado en la carpeta dist generada con PyInstaller) junto con cualquier otro recurso necesario (por ejemplo, la carpeta assets, si no fue empaquetada dentro del ejecutable).

Haz doble clic en ExtractorDeFacturas.exe.
La aplicación se abrirá y podrás cargar un PDF, definir áreas para extraer datos, generar un template JSON y, finalmente, procesar documentos para exportar los datos a Excel.
Guía de Uso Interna


**Cargar PDF:** Selecciona un archivo PDF que contenga la factura.
**Definir Área:** Haz clic y arrastra sobre el PDF para definir un área de extracción.
**Agregar Campo/Tabla:** Ingresa el nombre del campo o tabla y, si es necesario, ajusta la configuración.
**Guardar Template:** Guarda la plantilla en formato JSON para reutilizarla.
**Procesar Documentos:** Selecciona la plantilla y uno o más PDFs para extraer los datos y exportarlos a Excel.

## Notas Adicionales
- Entorno Autónomo:
Este ejecutable se ha creado usando PyInstaller y es completamente independiente, por lo que no es necesario que tengas Python instalado en tu sistema.

- Soporte de Idiomas:
Si necesitas agregar nuevos idiomas para la extracción de texto mediante OCR, descarga los archivos .traineddata correspondientes y colócalos en la carpeta tessdata de tu instalación de Tesseract OCR.

- Actualizaciones:
Consulta la documentación de Tesseract OCR para más detalles sobre configuración, actualizaciones y soporte de idiomas.

- Preguntas y Soporte
Si encuentras algún problema o tienes alguna pregunta, por favor contacta a [giamportone1@gmail.com] o consulta la documentación incluida en este paquete.

Desarrollado por Ariel Lujan Giamportone. 
linkedin: https://www.linkedin.com/in/agiamportone/
Github: https://github.com/arielgiamportone
