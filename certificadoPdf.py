import pandas as pd
from pdf2image import convert_from_path
from PIL import ImageDraw, ImageFont
import os

df = pd.read_excel("dadosCertificado.xlsx")

df.columns = df.columns.str.strip()

pdf_path = 'CertificadoCorreto.pdf'
pages = convert_from_path(pdf_path)

try:
    fonte = ImageFont.truetype('arialbd.Bold.ttf', 30)
except OSError:
    fonte = ImageFont.load_default(size=30)

posicoes = {
    "Nome": [(315, 645), (300, 1340)],
    "CPF": [(1030, 645)],

}

pasta_destino = "Certificados"
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
    imagem_salva = os.path.join(pasta_destino, f"Certificado_{linha.get('Nome', i)}.png")
    imagem.save(imagem_salva, "PNG")
    print(f"Gerada: {imagem_salva} para o funcionário {linha.get('Nome', i)}")

    
