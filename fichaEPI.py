import pandas as pd
import io
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape 
from reportlab.lib.colors import black
import os

df = pd.read_excel("DadosFichaEPI.xlsx")
pdf_template = "Ficha de EPI - Modelo.pdf"
output_dir = "Fichas_Preenchidas"

if not os.path.exists(output_dir):
    os.makedirs(output_dir)

def create_overlay(nome, matricula, funcao, setor, admissao):
    packet = io.BytesIO()
    
    # Define a folha para Paisagem (Deitada)
    can = canvas.Canvas(packet, pagesize=landscape(A4)) 
    can.setFont("Helvetica-Bold", 12)
    can.setFillColor("black") 
    
    can.drawString(550, 475, f"{nome}")          
    can.drawString(550, 433, f"{matricula}")     
    can.drawString(550, 390, f"{funcao}")       
    can.drawString(550, 352, f"{setor}")         
    
    data_admissao = str(admissao) if pd.notna(admissao) else "N/A"
    can.drawString(550, 310, f"{data_admissao}") 
        
    can.save()
    packet.seek(0)
    return packet

for index, row in df.iterrows():
    nome = row['Nome']
    matricula = row['Matrícula']
    funcao = row['Função']
    setor = row['Setor']
    admissao = row['DataAdmissão']

    if pd.notna(admissao):
        data_admissao_formatada = pd.to_datetime(admissao).strftime("%d/%m/%Y")
    else:        data_admissao_formatada = ""
    
    print(f"Gerando ficha para: {nome}...")
    
    overlay_pdf_stream = create_overlay(nome, matricula, funcao, setor, data_admissao_formatada)
    
    template_pdf = PdfReader(open(pdf_template, "rb"))
    overlay_pdf = PdfReader(overlay_pdf_stream)
    
    output_pdf = PdfWriter()
    
    page = template_pdf.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    output_pdf.add_page(page)
    
    for page_num in range(1, len(template_pdf.pages)):
        output_pdf.add_page(template_pdf.pages[page_num])
    
    safe_name = str(nome).replace(" ", "_").replace("/", "-")
    output_filename = os.path.join(output_dir, f"Ficha_EPI_{safe_name}.pdf")
    
    with open(output_filename, "wb") as output_file:
        output_pdf.write(output_file)

print(f"\nConcluído! Abra a pasta '{output_dir}' e veja se o texto apareceu.")