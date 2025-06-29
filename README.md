# 📊 Automatizar Tratamento de Dados

Automação em Python para a Inspired Educational Group (Escolas Elevas Brasil), que limpa, valida e estrutura arquivos de matrícula em Excel (`.xlsx`), gerando CSVs prontos para upload na plataforma de assinaturas digitais DocuSign.

---

## 🔎 Visão Geral

Este projeto:

1. **Identifica** o arquivo `.xlsx` mais recente na sua pasta de downloads.  
2. **Lê** e **pré‑processa** o conteúdo:  
   - Remove colunas temporárias  
   - Filtra e descarta linhas vazias  
   - Normaliza e valida campos (e‑mail, CNPJ, endereço, nomes, unidade)  
3. **Reporta** eventuais inconsistências organizadas em uma mensagem de validação.  
4. **Gera** diferentes lotes de CSV por série/nível educacional (LIV, LIV+MD, SemMaterial. Que são referentes aos tipos de matrícula de cada lote de aluno).  
5. **Move** o Excel original para uma pasta organizada por cidade e data.  
6. **Abre** a pasta de destino no Explorer para confirmação visual.

Tudo isso com um único comando no script principal.

---

## 📁 Estrutura de Diretórios

```
automatizar-tratamento-de-dados/
│
├── modules/
│   ├── get_path_sheet.py       # Busca o .xlsx mais recente
│   ├── organizar.py            # Funções utilitárias (linha separadora, clear screen)
│   └── tratar_data_set.py      # Lê, limpa e valida o DataFrame; gera CSVs e move arquivos
│
├── lote_matriculas.py          # Script “entry point” que orquestra todo o fluxo
├── requirements.txt            # Dependências do projeto
└── README.md                   # Esta documentação
```

---

## ⚙️ Pré‑requisitos

- **Python 3.8+**  
- Bibliotecas (veja `requirements.txt`):
  - pandas  
  - openpyxl  
  - validate_email_address  

---

## 🚀 Como Usar

1. **Clone** este repositório:  
   ```bash
   git clone https://github.com/johnwesleysa/Automatizar-tratamento-de-dados.git
   cd Automatizar-tratamento-de-dados
   ```

2. **Instale** as dependências:  
   ```bash
   pip install -r requirements.txt
   ```

3. **Execute** o script principal:  
   ```bash
   python lote_matriculas.py
   ```

4. **Siga as instruções** no console:
   - O sistema exibirá a data atual.
   - Apresentará o resultado das validações de e‑mail, nomes, CNPJ e endereço.
   - Perguntará se deseja gerar os arquivos (`1` = Sim, `2` = Não).

5. Se optar por gerar, os arquivos `.csv` e o `.xlsx` original serão organizados em:
   ```
   C:/Users/<seu-usuário>/Desktop/Matrículas e Rematrículas/Matriculas <ano>/
     └── <CIDADE>/
         └── <DD‑MM‑YYYY>/
             ├── Lote-<CIDADE>-Matriculas LIV <DD‑MM‑YYYY>.csv
             ├── Lote-<CIDADE>-Matriculas LIV+MD <DD‑MM‑YYYY>.csv
             ├── Lote-<CIDADE>-Matriculas SemMaterial <DD‑MM‑YYYY>.csv
             └── Lote-<CIDADE>-<DD‑MM‑YYYY>.xlsx
   ```

---

## 🔧 Configurações

- **Downloads**: por padrão usa `C:\Users\john.alencar\Downloads`.  
  — Para mudar, edite `download_dir` em `modules/get_path_sheet.py` ou parametrize via variável de ambiente.  
- **Destino**: pasta-base `"C:/Users/john.alencar/Desktop/Matrículas e Rematrículas/Matriculas 2024"`.  
  — Para ajustar, altere `diretorio_base` dentro de `gerar_arquivo()` em `modules/tratar_data_set.py`.
- **Endereços e CNPJs**: mapeados em `enderecos_elevas`.  
  — Para escolas adicionais, adicione novas chaves/valores seguindo o mesmo padrão.

---

## 📈 Lógica de Validação

- **E‑mail**: usa `validate_email_address` com `check_deliverability=True`.  
- **Nomes**: compara `ALUNO-NOME` vs. `Responsavel::ALUNO-NOME`.  
- **Unidade**: compara `ALUNO-UNIDADE` vs. `Responsavel::ESCOLA-UNIDADE`.  
- **Endereço & CNPJ**: conforme mapeamento por cidade (Recife, São Paulo, Brasília, Rio de Janeiro).

A mensagem de saída exibe, em sequência, o status de cada validação.

---

## 🤝 Contribuições

1. Faça um _fork_ do projeto  
2. Crie uma _branch_ (`feature/nome-da-feature`)  
3. Commit suas mudanças (`git commit -m 'Adiciona feature X'`)  
4. Push na branch (`git push origin feature/nome-da-feature`)  
5. Abra um _Pull Request_

---
