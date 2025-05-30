import os
from pdf2image import convert_from_path
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.draw import Frame, Image
from PIL import Image as PILImage

# Caminho do Poppler
POPPLER_PATH = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# Pastas
INPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\pdfs"
OUTPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\odss"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def pdf_to_ods(pdf_path, output_path):
    images = convert_from_path(pdf_path, dpi=150, poppler_path=POPPLER_PATH)
    spreadsheet = OpenDocumentSpreadsheet()

    temp_images = []

    for i, img in enumerate(images):
        # Salva imagem temporária
        temp_img_path = os.path.join(output_path.rsplit("\\", 1)[0], f"temp_page_{i+1}.png")
        img.save(temp_img_path, "PNG")
        temp_images.append(temp_img_path)

        href = spreadsheet.addPicture(temp_img_path)

        table = Table(name=f"Page_{i+1}")
        row = TableRow()
        cell = TableCell()

        frame = Frame(width="20cm", height="26cm", anchortype="paragraph")
        frame.addElement(Image(href=href))
        cell.addElement(frame)
        row.addElement(cell)
        table.addElement(row)
        spreadsheet.spreadsheet.addElement(table)

    spreadsheet.save(output_path)

    for img_path in temp_images:
        if os.path.exists(img_path):
            os.remove(img_path)

    print(f"Convertido: {os.path.basename(pdf_path)} → {os.path.basename(output_path)}")

def process_folder(input_folder, output_folder):
    for file in os.listdir(input_folder):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, file)
            output_file = os.path.splitext(file)[0] + ".ods"
            output_path = os.path.join(output_folder, output_file)
            pdf_to_ods(pdf_path, output_path)

# Executa a conversão
process_folder(INPUT_FOLDER, OUTPUT_FOLDER)
