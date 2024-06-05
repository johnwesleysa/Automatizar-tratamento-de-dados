import pyarrow as pa
import pandas as pd
import numpy as np
import openpyxl
from validate_email_address import validate_email
  

def tratar_data_set(cidade):
  #Leitura de arquivo CSV
  template = pd.read_excel("template.xlsx")
  
  #apagando a coluna Serie-Temp
  #template.drop(columns='SERIE-TEMP', inplace=True)

  template = template.drop("SERIE-TEMP", axis="columns")
  #Trata as colunas e apaga linhas vazias
  colunas = ['Responsavel::Name', 'Responsavel::Email', 'ALUNO-NOME', 'Responsavel::ALUNO-NOME', 'ALUNO-UNIDADE', 'Responsavel::ESCOLA-UNIDADE', 'Responsavel::CIDADE-ASSINATURA', 'Responsavel::ESCOLA-CNPJ', 'Responsavel::ESCOLA-ENDERECO', 'Responsavel::ALUNO-SERIE', 'Responsavel::ALUNO-TURNO', 'Responsavel::ANUIDADE-TEXTO1X', 'Responsavel::ANUIDADE-TEXTO12X']
  template = template.dropna(subset=colunas, how='any', axis=0)

  #Validar e-mails com validate_email_address do python
  #função para validações maiores que 0
  def maior_que_zero(numero):
    if len(numero) > 0:
      return f'Indices inválidos'
    else:
      return f'Todos os indices válidados com sucesso!'

  #criando uma função para validar os e-mails
  def validar_emails(email):
    return validate_email(email, check_deliverability=True)

 #validação do endereço e CNPJ
  enderecos_elevas = {
      'Barra': [{
          'endereco': 'Av. José Silva de Azevedo Neto 309 - Barra da Tijuca - Rio de Janeiro - RJ',
          'cnpj': '20.151.362/0002-71'
      }],
      'Botafogo': [{
          'endereco': 'R. Gen. Severiano 159 - Botafogo - Rio de Janeiro - RJ',
          'cnpj': '20.151.362/0001-90'
      }],
      'Urca': [{
          'endereco': 'Avenida Joao Luiz Alves 13 - Rio de Janeiro - RJ',
          'cnpj': '20.151.362/0006-03'
      }],
      'Brasília': [{
          'endereco': 'Quadra SGAS 613/614 - LT 100 SN - Asa Sul - Brasília - DF',
          'cnpj': '20.151.362/0003-52'
      }],
      'Recife': [{
          'endereco': 'R. Alameda das Hortências 279 - Lote 1B Quadra X1 - Imbiribeira - Recife - PE',
          'cnpj': '20.151.362/0005-14'
      }],
      'São Paulo': [{
          'endereco': 'R. José Antonio Coelho, 879 – Vila Mariana – SP',
          'cnpj': '20.151.362/0007-86'
      }],
  }

  #função para validar endereco, cnpj e cidade.
  def validar_endereco(cidade):
    #verifica se todos os itens na cidade estão corretos
    #caso contrário irá armazenar os valores incorretos na variavel cidade_selecionada
    
    #valida se todas as cidades estão corretas
    if cidade == "Recife" or cidade == "São Paulo" or cidade == "Brasília":
      cidade_selecionada = template[template['Responsavel::CIDADE-ASSINATURA'] != cidade]
      if len(cidade_selecionada) > 0:
        return 'Cidade está incorreta'
      else:
        enderecos_errados = template[template['Responsavel::ESCOLA-ENDERECO'] != enderecos_elevas[cidade][0]['endereco']]
        cnpj_errado = template[template['Responsavel::ESCOLA-CNPJ'] != enderecos_elevas[cidade][0]['cnpj']]

        return maior_que_zero(enderecos_errados), maior_que_zero(cnpj_errado)
    
    else:
      cidade_selecionada = template[template['Responsavel::CIDADE-ASSINATURA'] != "Rio de Janeiro"]
      if len(cidade_selecionada) > 0:
        return 'Cidade está incorreta'
      else:
        enderecos_errados = template[template['Responsavel::ESCOLA-ENDERECO'] != enderecos_elevas[cidade][0]['endereco']]
        cnpj_errado = template[template['Responsavel::ESCOLA-CNPJ'] != enderecos_elevas[cidade][0]['cnpj']]

        return maior_que_zero(enderecos_errados), maior_que_zero(cnpj_errado)

  #Validar e-mails
  indices_emails_invalidos = template[template['Responsavel::Email'].apply(lambda x: not validar_emails(x))].index
  

  #Validar de o nome das colunas de aluno estão iguais.
  nomes_alunos_diferentes = template[template['ALUNO-NOME'] != template['Responsavel::ALUNO-NOME']].index
  

  #Validar escola unidade
  escolas_diferentes = template[template['ALUNO-UNIDADE'] != template['Responsavel::ESCOLA-UNIDADE']].index


  #retorno = [maior_que_zero(indices_emails_invalidos), maior_que_zero(nomes_alunos_diferentes),  maior_que_zero(escolas_diferentes), validar_endereco(cidade)]
  template_tratado = template
  validacoes = f"Validação e-mails: {maior_que_zero(indices_emails_invalidos)}\nValidação Nomes Alunos: {maior_que_zero(nomes_alunos_diferentes)}\nValidar Escolas Diferentes: {maior_que_zero(escolas_diferentes)}\nValidar Endereço de {cidade}: {validar_endereco(cidade)}"
  return template_tratado, validacoes
    
def gerar_arquivo(cidade, data, template):
  nome_cidade = cidade

  """# **LIV**"""

  lista_liv = ['Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5', 'Kindergarten 3', 'Kindergarten 4', 'Kindergarten 5']
  alunos_liv = template['Responsavel::ALUNO-SERIE'].isin(lista_liv)
  template_liv = template[alunos_liv]

  
  if template_liv['Responsavel::ALUNO-SERIE'].count() > 0:
    template_liv.to_csv(f'Lote-{nome_cidade}-Matriculas LIV {data}.csv', sep=';', encoding='utf-8', index=False)

  #template.dropna(subset='Responsavel::ALUNO-MATRICULA', how='all')

  """# **LIV+MD**"""

  lista_liv_md = ['Grade 6', 'Grade 7', 'Grade 8', 'Grade 9']
  alunos_liv_md = template['Responsavel::ALUNO-SERIE'].isin(lista_liv_md)
  template_liv_md = template[alunos_liv_md]

  if template_liv_md['Responsavel::ALUNO-SERIE'].count() > 0:
    template_liv_md.to_csv(f'Lote-{nome_cidade}-Matriculas LIV+MD {data}.csv', sep=';',encoding='utf-8', index=False)

  """# **SemMaterial**"""

  lista_sem_material = ['Grade 10', 'Grade 11', 'Grade 12', 'Kindergarten 1', 'Kindergarten 2']
  alunos_sem_material = template['Responsavel::ALUNO-SERIE'].isin(lista_sem_material)
  template_sem_material = template[alunos_sem_material]
  
  if template_sem_material['Responsavel::ALUNO-SERIE'].count() > 0:
    template_sem_material.to_csv(f'Lote-{nome_cidade}-Matriculas SemMaterial {data}.csv', sep=';',encoding='utf-8', index=False)

  



