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

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook("planilha_alunos.xlsx")
sheet_alunos = workbook_alunos["Sheet1"]


for linha in sheet_alunos.iter_rows(min_row=2):
    # cada célula que contém a info que precisamos
    nome_curso = linha[0].value  # nome do curso
    nome_participante = linha[1].value  # nome participante
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_termino = linha[4].value
    cara
