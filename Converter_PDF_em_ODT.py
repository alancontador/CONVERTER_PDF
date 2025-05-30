import os
from pdf2image import convert_from_path
from odf.opendocument import OpenDocumentText
from odf.text import P
from odf.draw import Frame, Image
from odf.style import Style, GraphicProperties
from PIL import Image as PILImage

# Caminho do Poppler
POPPLER_PATH = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# Pasta de entrada e saída
INPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\pdfs"
OUTPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\odts"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def pdf_to_odt(pdf_path, output_path):
    images = convert_from_path(pdf_path, dpi=150, poppler_path=POPPLER_PATH)
    document = OpenDocumentText()

    temp_images = []

    for i, img in enumerate(images):
        # Salvar a imagem temporariamente
        temp_img_path = os.path.join(output_path.rsplit("\\", 1)[0], f"temp_page_{i+1}.png")
        img.save(temp_img_path, "PNG")
        temp_images.append(temp_img_path)

        href = document.addPicture(temp_img_path)

        # Inserir imagem em um parágrafo
        frame = Frame(width="16cm", height="22cm", anchortype="paragraph")
        frame.addElement(Image(href=href))
        p = P()
        p.addElement(frame)
        document.text.addElement(p)

    document.save(output_path)

    for img_path in temp_images:
        if os.path.exists(img_path):
            os.remove(img_path)

    print(f"Convertido: {os.path.basename(pdf_path)} → {os.path.basename(output_path)}")

def process_folder(input_folder, output_folder):
    for file in os.listdir(input_folder):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, file)
            output_file = os.path.splitext(file)[0] + ".odt"
            output_path = os.path.join(output_folder, output_file)
            pdf_to_odt(pdf_path, output_path)

# Executa a conversão em lote
process_folder(INPUT_FOLDER, OUTPUT_FOLDER)
