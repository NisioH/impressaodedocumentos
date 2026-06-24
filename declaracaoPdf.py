import pandas as pd
from pdf2image import convert_from_path
from PIL import ImageDraw, ImageFont
import os

df = pd.read_excel("dadosDeclaração (1).xlsx")

df.columns = df.columns.str.strip()

pdf_path = 'DeclaraçãoIntrajornada.pdf'
pages = convert_from_path(pdf_path)

try:
    fonte = ImageFont.truetype('arialbd.Bold.ttf', 35)
except OSError:
    fonte = ImageFont.load_default(size=35)

posicoes = {
    "Nome": [(230, 580), (240, 1355)],
    "CPF": [(1050, 580)],
    "ENDERECO": [(150, 640)],
    "BAIRRO": [(840, 640)],
    "MUNICIPIO": [(1300, 640)],
    "CARGO": [(762, 755)],
}

pasta_destino = "Declarações"
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)


for i, linha in df.iterrows():
    imagem = pages[0].copy()
    draw = ImageDraw.Draw(imagem)

    # Escrever os dados do funcionário atual na imagem
    for coluna, lista_de_posicoes in posicoes.items():
        valor = str(linha[coluna]) if coluna in linha else ""

        if valor.lower() == "nan":
            valor = ""

        # Percorre todas as coordenadas cadastradas para esta coluna
        for posicao in lista_de_posicoes:
            draw.text(posicao, valor, font=fonte, fill=(0, 0, 0))

    # 4. Salvar a imagem
    imagem_salva = os.path.join(pasta_destino, f"Declaração_{linha.get('Nome', i)}.png")
    imagem.save(imagem_salva, "PNG")
    print(f"Gerada: {imagem_salva} para o funcionário {linha.get('Nome', i)}")

    
