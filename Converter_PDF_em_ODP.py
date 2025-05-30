import os
from pdf2image import convert_from_path
from odf.opendocument import OpenDocumentPresentation
from odf.draw import Page, Frame, Image
from odf.style import PageLayout, PageLayoutProperties, MasterPage
from PIL import Image as PILImage

# Caminho para o Poppler
POPPLER_PATH = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# Pasta com os PDFs
INPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\pdfs"
# Pasta de saída
OUTPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\odps"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def pdf_to_odp(pdf_path, output_path):
    images = convert_from_path(pdf_path, dpi=150, poppler_path=POPPLER_PATH)
    presentation = OpenDocumentPresentation()

    layout = PageLayout(name="MyLayout")
    layout_props = PageLayoutProperties(margin="0cm", pagewidth="28cm", pageheight="21cm", printorientation="landscape")
    layout.addElement(layout_props)
    presentation.automaticstyles.addElement(layout)

    master_page = MasterPage(name="Default", pagelayoutname=layout)
    presentation.masterstyles.addElement(master_page)

    temp_images = []

    for i, img in enumerate(images):
        slide = Page(name=f"Slide{i+1}", masterpagename="Default")
        temp_img_path = os.path.join(output_path.rsplit("\\", 1)[0], f"temp_slide_{i+1}.png")
        img.save(temp_img_path, "PNG")
        temp_images.append(temp_img_path)

        href = presentation.addPicture(temp_img_path)
        frame = Frame(width="28cm", height="21cm", x="0cm", y="0cm")
        frame.addElement(Image(href=href))
        slide.addElement(frame)
        presentation.presentation.addElement(slide)

    presentation.save(output_path)

    for img_path in temp_images:
        if os.path.exists(img_path):
            os.remove(img_path)

    print(f"Convertido: {os.path.basename(pdf_path)} → {os.path.basename(output_path)}")

def process_folder(input_folder, output_folder):
    for file in os.listdir(input_folder):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, file)
            output_file = os.path.splitext(file)[0] + ".odp"
            output_path = os.path.join(output_folder, output_file)
            pdf_to_odp(pdf_path, output_path)

# Executa o processamento
process_folder(INPUT_FOLDER, OUTPUT_FOLDER)
