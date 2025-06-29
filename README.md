# 📊 Automatização de Tratamento de Dados de Matrícula

Este repositório traz uma automação desenvolvida em **Python** para a **Inspired Education Group (Escolas Elevas Brasil)**. O objetivo é limpar, validar e organizar planilhas de matrícula em Excel (`.xlsx`), gerando arquivos `.csv` prontos para upload na plataforma **DocuSign**.

Com um único comando, todo o processo manual (e demorado) virou uma tarefa rápida, confiável e repetível. 🚀

---

## 🔎 Visão Geral

Essa automação realiza as seguintes etapas:

1. **Localiza automaticamente** o arquivo `.xlsx` mais recente na pasta de downloads  
2. **Lê e pré-processa** os dados:  
   - Remove colunas temporárias  
   - Elimina linhas vazias  
   - Valida e padroniza campos como e-mail, CNPJ, endereço, nomes e unidade escolar  
3. **Aponta inconsistências** de forma clara no terminal  
4. **Gera três arquivos `.csv` separados**, conforme o tipo de matrícula:  
   - LIV  
   - LIV+MD  
   - SemMaterial  
5. **Organiza tudo em pastas** por cidade e data  
6. **Abre a pasta final** no Explorer para que você possa conferir os arquivos

Tudo isso é feito com apenas um comando no script principal. Simples assim. 😉

---

## 📁 Estrutura do Projeto

```plaintext
automatizar-tratamento-de-dados/
│
├── modules/
│   ├── get_path_sheet.py       # Localiza o Excel mais recente
│   ├── organizar.py            # Funções auxiliares (limpar tela, separadores, etc.)
│   └── tratar_data_set.py      # Processa, valida e organiza os dados; gera os CSVs
│
├── lote_matriculas.py          # Script principal que orquestra o fluxo inteiro
├── requirements.txt            # Dependências Python
└── README.md                   # Este guia
```

---

## ⚙️ Pré-requisitos

- **Python 3.8 ou superior**
- Bibliotecas (instaláveis via `requirements.txt`):
  - `pandas`  
  - `openpyxl`  
  - `validate_email_address`  

---

## 🚀 Como Usar (inclusive para meu "eu do futuro" rs)

1. **Clone o repositório**:
   ```bash
   git clone https://github.com/johnwesleysa/Automatizar-tratamento-de-dados.git
   cd Automatizar-tratamento-de-dados
   ```

2. **Instale as dependências**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Execute o script principal**:
   ```bash
   python lote_matriculas.py
   ```

4. **Siga as instruções no terminal**:
   - A data atual será exibida.
   - O sistema mostrará os resultados da validação (e‑mail, nomes, CNPJ e endereço).
   - Você escolherá se quer ou não gerar os arquivos (`1` = Sim | `2` = Não).

5. **Arquivos gerados**:  
   Os `.csv` e o `.xlsx` original serão organizados assim:
   ```plaintext
   C:/Users/<seu-usuário>/Desktop/Matrículas e Rematrículas/Matriculas <ano>/
     └── <CIDADE>/
         └── <DD‑MM‑YYYY>/
             ├── Lote-<CIDADE>-Matriculas LIV <data>.csv
             ├── Lote-<CIDADE>-Matriculas LIV+MD <data>.csv
             ├── Lote-<CIDADE>-Matriculas SemMaterial <data>.csv
             └── Lote-<CIDADE>-<data>.xlsx
   ```

---

## 🔧 Personalizações e Configurações

- **Pasta de downloads** padrão:  
  `C:/Users/john.alencar/Downloads`  
  → Para mudar, edite `download_dir` no `get_path_sheet.py`.

- **Pasta de destino dos arquivos tratados**:  
  `C:/Users/john.alencar/Desktop/Matrículas e Rematrículas/Matriculas 2024`  
  → Para mudar, edite `diretorio_base` dentro da função `gerar_arquivo()` no `tratar_data_set.py`.

- **Validação de endereços e CNPJs**:  
  O mapeamento por cidade está no dicionário `enderecos_elevas`.  
  → Para adicionar novas unidades, inclua no mesmo padrão.

---

## 📈 Como Funciona a Validação

- **E‑mail**: verificação de entregabilidade com `validate_email_address` (`check_deliverability=True`)  
- **Nome do aluno**: compara `ALUNO-NOME` e `Responsavel::ALUNO-NOME`  
- **Unidade escolar**: compara `ALUNO-UNIDADE` e `Responsavel::ESCOLA-UNIDADE`  
- **Endereço e CNPJ**: validados conforme a cidade

O terminal mostra um resumo completo de todos os pontos validados.

---
