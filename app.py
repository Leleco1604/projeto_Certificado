""" 
Tenho uma planilha de excel com os dados dos alunos que fizeram um curso

Quero ver a possibilidade de criar um programa usando python para automatizar enviando os dados para 
uma planilha para preencher os campos mutáveis no certificado padrão 

Nome do curso, nome do participante, tipo de participação, data do inicio, data do final, carga horaria, 
data de emissão do certificado e as assinaturas do gestor geral.


#Pegar os dados da planilha 
Nome do curso, nome do participante, tipo de participação, data do inicio, data do final, carga horaria, 
data de emissão do certificado e as assinaturas do gestor geral.

#Transferir os dados para imagem do certificado 
"""

# Pegar os dados da planilha

import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook("planilha_alunos.xlsx")
sheet_alunos = workbook_alunos["Sheet1"]


for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada célula que contém a info que precisamos
    nome_curso = linha[0].value  # nome do curso
    nome_participante = linha[1].value  # nome participante
    tipo_participacao = linha[2].value  # tipo da partpação
    data_inicio = linha[3].value  # tipo da partpação
    data_termino = linha[4].value  # data termino
    carga_horaria = linha[5].value  # carga horaria
    data_emissao = linha[6].value  # data de emissão

    # Transferir os dados da planilha para imagem do certificado
    # Definindo fonte
    font_nome = ImageFont.truetype("./tahomabd.ttf",90)
    font_geral = ImageFont.truetype("./tahoma.ttf",80)
    font_data = ImageFont.truetype("./tahoma.ttf",55)


    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020,827),nome_participante, fill = 'black', font = font_nome)
    desenhar.text((1060,950),nome_curso, fill = 'black', font = font_geral)
    desenhar.text((1435,1065),tipo_participacao, fill = 'black', font = font_geral)
    desenhar.text((1480,1182),str(carga_horaria), fill = 'black', font = font_geral)

    desenhar.text((750,1770), data_inicio, fill='black' , font= font_data)
    desenhar.text((750,1930), data_termino, fill='black' , font= font_data)

    desenhar.text((2220,1930), data_emissao, fill='black', font= font_data)


    image.save(f'./{indice} {nome_participante} certificado.png')





