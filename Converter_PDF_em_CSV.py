import os
import pdfplumber
import pandas as pd
from pdf2image import convert_from_path
import pytesseract

# Caminho para o Poppler e Tesseract no Windows
POPPLER_PATH = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\Release-24.08.0-0\poppler-24.08.0\Library\bin"
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

def extract_text_with_ocr(pdf_path):
    try:
        images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        text = ''
        for img in images:
            text += pytesseract.image_to_string(img)
        return text
    except Exception as e:
        print(f"Erro no OCR de {os.path.basename(pdf_path)}: {e}")
        return None

def pdf_to_csv(pdf_path, output_path):
    try:
        # Primeiro tenta com pdfplumber (PDF digital)
        with pdfplumber.open(pdf_path) as pdf:
            tables = []
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    tables.append(pd.DataFrame(table))
        
        if tables:
            final_df = pd.concat(tables, ignore_index=True)
            final_df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(f"Convertido: {os.path.basename(pdf_path)} → {os.path.basename(output_path)}")
        else:
            raise ValueError("Nenhuma tabela encontrada")
    
    except:
        # Se não encontrou tabela, tenta OCR (PDF escaneado)
        print(f"Usando OCR em {os.path.basename(pdf_path)}...")
        text = extract_text_with_ocr(pdf_path)
        if text:
            with open(output_path, "w", encoding='utf-8-sig') as f:
                f.write(text)
            print(f"OCR concluído: {os.path.basename(pdf_path)} → {os.path.basename(output_path)}")
        else:
            print(f"Nada encontrado em {os.path.basename(pdf_path)}")

def process_folder(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file in os.listdir(input_folder):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, file)
            output_name = os.path.splitext(file)[0] + ".csv"
            output_path = os.path.join(output_folder, output_name)
            pdf_to_csv(pdf_path, output_path)

# Caminhos de entrada e saída
INPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\pdfs"
OUTPUT_FOLDER = r"C:\Users\alant\Documents\VSCODE\CODIGO\Converter_PDF\csvs"

process_folder(INPUT_FOLDER, OUTPUT_FOLDER)
