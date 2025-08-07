import pandas as pd
from pdf2image import convert_from_path
from PIL import ImageDraw, ImageFont
import os

excel_file = pd.ExcelFile("dados.xlsx")

pdf_path = 'envio.pdf'
pages = convert_from_path(pdf_path)

fonte = ImageFont.truetype('arialbd.ttf', 24)

posicoes = {
    "Nome": (350, 1000),
    "Endereço": (350, 1050),
    "Cidade": (350, 1095),
    "Estado": (960, 1095),
    "CEP": (1250, 1095),
    "CNPJ": (350, 1135),
    "Inscrição Estadual": (1010, 1135),
    "Contato": (350, 1180),
    "Telefone": (1225, 1180),
    "E-mail": (580, 1225),
    "Endereço de Entrega": (580, 1260),
    "Substituto Tributário": (1300, 1300),
    "Modelo": (350, 1490),
    "Número de Série": (750, 1490),
    "Manutenção": (170, 1965),
    "Transportadora": (780, 1955),
    "Operações Comerciais": (1115, 1610),
    "Conectado": (1260, 1670),
    "Grãos": (780, 2075)
}

for i, sheet_name in enumerate(excel_file.sheet_names):
    df = excel_file.parse(sheet_name)  
    df.columns = df.columns.str.strip()  
    dados = dict(zip(df["Campo"], df["Valor"]))  

    
    imagem = pages[0].copy()  
    draw = ImageDraw.Draw(imagem)

    
    for campo, posicao in posicoes.items():
        valor = str(dados[campo]) if campo in dados else ""  
        draw.text(posicao, valor, font=fonte, fill=(0, 0, 0))


    imagem_salva = f'Instruções/Instrução_{i + 1}.png'
    if not os.path.exists('Instruções'):
        os.makedirs('Instruções')
    imagem.save(imagem_salva, 'PNG')

    
