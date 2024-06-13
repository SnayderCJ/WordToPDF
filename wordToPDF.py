import os
import comtypes.client

def docx_to_pdf(docx_directory, pdf_directory):
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False

    if not os.path.exists(docx_directory):
        os.makedirs(docx_directory)

    for docx_file in os.listdir(docx_directory):
        if docx_file.endswith(".docx"):
            docx_path = os.path.join(docx_directory, docx_file)
            pdf_path = os.path.join(pdf_directory, docx_file.replace(".docx", ".pdf"))

            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()

    word.Quit()

input_directory = 'D:\Snayder\PyCharm\ED_A1\Converter_Word_To_Pdf\WORDs'
output_directory = 'D:\Snayder\PyCharm\ED_A1\Converter_Word_To_Pdf\PDFs'

docx_to_pdf(input_directory, output_directory)

