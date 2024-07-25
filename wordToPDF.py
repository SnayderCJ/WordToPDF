import os
import comtypes.client
import logging

def docx_to_pdf(docx_directory, pdf_directory):
    """Convierte todos los archivos .docx en un directorio a PDF."""

    logging.info(f"Iniciando conversión de Word a PDF desde '{docx_directory}' a '{pdf_directory}'")

    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # No mostrar la interfaz de Word

        # Verificar si el directorio de entrada existe
        if not os.path.exists(docx_directory):
            logging.error(f"El directorio de entrada '{docx_directory}' no existe.")
            return  # Salir de la función si no hay directorio de entrada

        # Crear el directorio de salida si no existe
        if not os.path.exists(pdf_directory):
            os.makedirs(pdf_directory)
            logging.info(f"Se creó el directorio de salida '{pdf_directory}'.")

        # Iterar sobre los archivos .docx en el directorio de entrada
        for docx_file in os.listdir(docx_directory):
            if docx_file.endswith(".docx"):
                docx_path = os.path.join(docx_directory, docx_file)
                pdf_path = os.path.join(pdf_directory, docx_file.replace(".docx", ".pdf"))

                try:
                    doc = word.Documents.Open(docx_path)
                    doc.SaveAs(pdf_path, FileFormat=17)  # 17 es el formato PDF
                    doc.Close()
                    logging.info(f"Convertido '{docx_file}' a PDF.")
                except Exception as e:
                    logging.error(f"Error al convertir '{docx_file}': {e}")

    except Exception as e:
        logging.error(f"Error general en la conversión: {e}")

    finally:
        if word is not None:
            word.Quit()

# Configurar el registro (puedes ajustar el nivel de registro según tus necesidades)
logging.basicConfig(filename="conversion_word_a_pdf.log", level=logging.INFO)

# Directorios de entrada y salida (personaliza según tu estructura de carpetas)
input_directory = '/home/snazzyc/SNAYDERC/PyCharm/WordToPDF/WORDs'
output_directory = '/home/snazzyc/SNAYDERC/PyCharm/WordToPDF/PDFs'

# Iniciar la conversión
docx_to_pdf(input_directory, output_directory)