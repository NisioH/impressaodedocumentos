import os
from PIL import Image
import win32print, win32api
from PyPDF2 import PdfReader, PdfWriter
import time  

printer_name = win32print.GetDefaultPrinter()
folder_path = r"C:\Users\fazin\OneDrive\Documents\Nisio\ImpressaoDocumentos\Fichas_Preenchidas"


temp_folder = os.path.join(folder_path, "temp_paginas")
os.makedirs(temp_folder, exist_ok=True)

for file_name in os.listdir(folder_path):
    file_path = os.path.join(folder_path, file_name)

    # extrai a página que será impressa
    if file_path.lower().endswith(".pdf"):
        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])  

            temp_pdf_path = os.path.join(temp_folder, f"pagina1_{file_name}")
            with open(temp_pdf_path, "wb") as f:
                writer.write(f)

            win32api.ShellExecute(0, "print", temp_pdf_path, None, ".", 0)

            # tempo pra que a impressão seja iniciada
            time.sleep(10)

            os.remove(temp_pdf_path)

        except Exception as e:
            print(f"Erro ao processar {file_name}: {e}")

    
    elif file_path.lower().endswith(".docx"):
        win32api.ShellExecute(0, "print", file_path, None, ".", 0)

    
    elif file_path.lower().endswith(".png"):
        try:
            imagem = Image.open(file_path)
            imagem = imagem.rotate(90, expand=True)
            imagem.save(file_path)

            os.system(f'mspaint /p "{file_path}"')
        except Exception as e:
            print(f"Erro ao processar imagem {file_name}: {e}")


try:
    os.rmdir(temp_folder)
except OSError:
    pass  



