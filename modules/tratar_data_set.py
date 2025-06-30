import pandas as pd
import openpyxl
import os
import shutil
from validate_email_address import validate_email
from modules.get_path_sheet import get_latest_xlsx

# Diretório de downloads do usuário
DOWLOAD_DIR = r'C:\Users\john.alencar\Downloads'
ARQUIVO = get_latest_xlsx()


#GLOBAL
CIDADE_GLOBAL = ''
CIDADE_ASSINATURA_GLOBAL = ''


def tratar_data_set():
    global CIDADE_GLOBAL
    global CIDADE_GLCIDADE_ASSINATURA_GLOBAL

    # Leitura de arquivo Excel
    template = pd.read_excel(ARQUIVO)

    #pega o nome da unidade da escola
    cidade = str(template['ALUNO-UNIDADE'].iloc[0])
    CIDADE_GLOBAL = cidade

    #pega o nome da cidade da escola
    cidade_assinatura = str(template['Responsavel::CIDADE-ASSINATURA'].iloc[0])
    CIDADE_ASSINATURA_GLOBAL = cidade_assinatura

    # Apagando a coluna Serie-Temp se existir
    if "SERIE-TEMP" in template.columns:
        template = template.drop("SERIE-TEMP", axis="columns")

    # Definindo as colunas necessárias
    colunas = [
        'Responsavel::Name', 'Responsavel::Email', 'ALUNO-NOME', 
        'Responsavel::ALUNO-NOME', 'ALUNO-UNIDADE', 
        'Responsavel::ESCOLA-UNIDADE', 'Responsavel::CIDADE-ASSINATURA', 
        'Responsavel::ESCOLA-CNPJ', 'Responsavel::ESCOLA-ENDERECO', 
        'Responsavel::ALUNO-SERIE', 'Responsavel::ALUNO-TURNO', 
        'Responsavel::ANUIDADE-TEXTO1X', 'Responsavel::ANUIDADE-TEXTO12X'
    ]

    # Apagando linhas vazias
    template = template.dropna(how='all', subset=colunas)

    # Validar e-mails
    def validar_emails(email):
        return validate_email(email, check_deliverability=True)

    # Dados públicos extraídos de fontes oficiais das unidades escolares (site, Google, redes sociais)
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

    def maior_que_zero(indices):
        return 'Existem índices inválidos' if len(indices) > 0 else 'Todos os índices validados com sucesso!'

    def validar_endereco(cidade):
        if cidade in ['Recife', 'São Paulo', 'Brasília']:
            cidade_selecionada = template[template['Responsavel::CIDADE-ASSINATURA'] == cidade]
        else:
            cidade_selecionada = template[template['Responsavel::CIDADE-ASSINATURA'] == "Rio de Janeiro"]

        if len(cidade_selecionada) == 0:
            return 'Cidade está incorreta'

        enderecos_errados = cidade_selecionada[cidade_selecionada['Responsavel::ESCOLA-ENDERECO'] != enderecos_elevas[cidade][0]['endereco']].index
        cnpj_errado = cidade_selecionada[cidade_selecionada['Responsavel::ESCOLA-CNPJ'] != enderecos_elevas[cidade][0]['cnpj']].index

        return maior_que_zero(enderecos_errados), maior_que_zero(cnpj_errado)

    # Validar e-mails
    indices_emails_invalidos = template[template['Responsavel::Email'].apply(lambda x: not validar_emails(str(x)) if pd.notna(x) else False)].index

    # Validar se o nome dos alunos estão iguais
    nomes_alunos_diferentes = template[template['ALUNO-NOME'] != template['Responsavel::ALUNO-NOME']].index

    # Validar escola unidade
    escolas_diferentes = template[template['ALUNO-UNIDADE'] != template['Responsavel::ESCOLA-UNIDADE']].index

    validacoes = (
        f"Validação e-mails: {maior_que_zero(indices_emails_invalidos)}\n"
        f"Validação Nomes Alunos: {maior_que_zero(nomes_alunos_diferentes)}\n"
        f"Validar Escolas Diferentes: {maior_que_zero(escolas_diferentes)}\n"
        f"Validar Endereço de {cidade}: {validar_endereco(cidade)}"
    )

    return template, validacoes

def gerar_arquivo(data, template):
    global CIDADE_GLOBAL
    global CIDADE_GLCIDADE_ASSINATURA_GLOBAL
    cidade = CIDADE_GLOBAL
    # Diretório base onde as pastas específicas para cada cidade estão localizadas
    diretorio_base = "C:/Users/john.alencar/Desktop/Matrículas e Rematrículas/Matriculas 2024"
    # Nome da pasta específica para a cidade
    pasta_cidade = os.path.join(diretorio_base, cidade)

    # Criar o diretório da cidade se ele ainda não existir
    if not os.path.exists(pasta_cidade):
        os.makedirs(pasta_cidade)

    # Criar o diretório da data dentro da pasta da cidade
    pasta_destino = os.path.join(pasta_cidade, data)
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    nome_cidade = cidade

    # **LIV**
    lista_liv = ['Grade 1', 'Grade 2', 'Grade 3', 'Grade 4', 'Grade 5', 'Kindergarten 3', 'Kindergarten 4', 'Kindergarten 5']
    template_liv = template[template['Responsavel::ALUNO-SERIE'].isin(lista_liv)]

    if not template_liv.empty:
        nome_arquivo_liv = f'Lote-{nome_cidade}-Matriculas LIV {data}.csv'
        caminho_arquivo_liv = os.path.join(pasta_destino, nome_arquivo_liv)
        template_liv.to_csv(caminho_arquivo_liv, sep=';', encoding='utf-8', index=False)

    # **LIV+MD**
    lista_liv_md = ['Grade 6', 'Grade 7', 'Grade 8', 'Grade 9']
    template_liv_md = template[template['Responsavel::ALUNO-SERIE'].isin(lista_liv_md)]

    if not template_liv_md.empty:
        nome_arquivo_liv_md = f'Lote-{nome_cidade}-Matriculas LIV+MD {data}.csv'
        caminho_arquivo_liv_md = os.path.join(pasta_destino, nome_arquivo_liv_md)
        template_liv_md.to_csv(caminho_arquivo_liv_md, sep=';', encoding='utf-8', index=False)

    # **SemMaterial**
    lista_sem_material = ['Grade 10', 'Grade 11', 'Grade 12', 'Kindergarten 1', 'Kindergarten 2']
    template_sem_material = template[template['Responsavel::ALUNO-SERIE'].isin(lista_sem_material)]

    if not template_sem_material.empty:
        nome_arquivo_sem_material = f'Lote-{nome_cidade}-Matriculas SemMaterial {data}.csv'
        caminho_arquivo_sem_material = os.path.join(pasta_destino, nome_arquivo_sem_material)
        template_sem_material.to_csv(caminho_arquivo_sem_material, sep=';', encoding='utf-8', index=False)

    # Mover o arquivo Excel para a pasta de destino
    destino_excel = os.path.join(pasta_destino, f"Lote-{nome_cidade}-{data}.xlsx")
    shutil.move(ARQUIVO, destino_excel)

    # Abrir a pasta de destino
    os.startfile(pasta_destino)
