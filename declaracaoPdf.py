import os
import sys
import pandas as pd
from pdf2image import convert_from_path
from PIL import ImageDraw, ImageFont

def formatar_data_extenso(data_input):
    MESES_PT = {
        "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
        "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
        "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro",
        "1": "Janeiro", "2": "Fevereiro", "3": "Março", "4": "Abril",
        "5": "Maio", "6": "Junho", "7": "Julho", "8": "Agosto",
        "9": "Setembro"
    }
    try:
        data_str = str(data_input).strip()
        if pd.isna(data_input) or data_str.lower() == "nan" or data_str == "":
            return ""
            
        data_str = data_str.split()[0]
        
        if "-" in data_str:
            ano, mes, dia = data_str.split("-")
        elif "/" in data_str:
            dia, mes, ano = data_str.split("/")
        else:
            return data_str
        
        dia_limpo = str(int(dia))
        mes_extenso = MESES_PT.get(mes, mes)
        
        return f"{dia_limpo} de {mes_extenso} de {ano}"
    except Exception:
        return str(data_input)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if sys.platform.startswith('win'):
    CAMINHO_POPPLER = os.path.join(BASE_DIR, 'poppler-windows', 'Library', 'bin')
    FONTE_CAMINHO = os.path.join(BASE_DIR, 'arialbd.ttf') 
else:
  
    CAMINHO_POPPLER = None 
    FONTE_CAMINHO = "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf"

df = pd.read_excel("dadosDeclaração (1).xlsx")
df.columns = df.columns.str.strip()

pdf_path = 'DeclaraçãoIntrajornada.pdf'

pages = convert_from_path(pdf_path, poppler_path=CAMINHO_POPPLER)

try:
    fonte = ImageFont.truetype(FONTE_CAMINHO, 30)
except OSError:
    print(f"Aviso: Não foi possível carregar a fonte em '{FONTE_CAMINHO}'. Usando fonte padrão.")
    fonte = ImageFont.load_default(size=30)

posicoes = {
    "Nome": [(210, 666), (180, 1645)],
    "CPF": [(1225, 666)],
    "ENDERECO": [(340, 723)],
    "BAIRRO": [(1050, 723)],
    "MUNICIPIO": [(380, 783)],
    "CARGO": [(200, 895)],
    "DATA": [(1105, 2046)],
}

pasta_destino = "Declarações"
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)

for i, linha in df.iterrows():
    imagem = pages[0].copy()
    draw = ImageDraw.Draw(imagem)

    for coluna, lista_de_posicoes in posicoes.items():
        
        if coluna == "DATA":
            valor = formatar_data_extenso(linha['DATA']) if 'DATA' in linha else ""
        else:
            valor = str(linha[coluna]) if coluna in linha else ""

        if valor.lower() == "nan" or valor.strip() == "":
            valor = ""

        for posicao in lista_de_posicoes:
            draw.text(posicao, valor, font=fonte, fill=(0, 0, 0))

    nome_funcionario = str(linha.get('Nome', i)).replace('/', '-').replace('\\', '-')
    
    imagem_salva = os.path.join(pasta_destino, f"Declaração_{nome_funcionario}.pdf")
    imagem.save(imagem_salva, "PDF")
    print(f"Gerada: {imagem_salva} para o funcionário {nome_funcionario}")