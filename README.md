# 🐍 projetosempyhton

Coleção de projetos pequenos e objetivos em Python, feitos para resolver problemas do dia a dia de TI.
Cada projeto fica na sua própria pasta, é independente dos outros e roda localmente.

## 📚 Projetos

| Projeto | O que faz | Pasta | Porta |
|---|---|---|---|
| [Consulta de Parcelas (Contas a Receber)](#-consulta-de-parcelas--contas-a-receber) | Consulta parcelas a receber de um cliente por CPF/CNPJ | [`consulta_parcelas`](./consulta_parcelas) | 5000 |
| [Consulta de Contas a Pagar](#-consulta-de-contas-a-pagar) | Consulta parcelas a pagar de um fornecedor por CNPJ/CPF | [`consulta_contas_pagar`](./consulta_contas_pagar) | 5001 |

Os dois usam as mesmas ideias: **importar uma planilha exportada do ERP uma única vez**, guardar só o necessário em um banco **SQLite** local e fazer **consultas instantâneas** por CPF/CNPJ em uma tela web simples.
Como as portas são diferentes, os dois podem rodar ao mesmo tempo.

---

## 💰 Consulta de Parcelas — Contas a Receber

Reproduz o painel **"Informações Parcelas Filtradas"** do ERP para consultar rapidamente a situação financeira de um cliente.

### Painel de indicadores

```
Qtde Parcelas      Total Original     Juros a Receber    Total a Receber
Primeiro Vencto    Total Saldo        Multa a Receber
Último Vencto      Total Liquidado    Total Cancelado    Total Renegociado
```

### Recursos
- 🔎 Pesquisa por **CPF** ou **CNPJ** (aceita com ou sem pontuação).
- 🎛️ Filtros: todas, em aberto, liquidadas, canceladas, renegociadas e **período de vencimento**.
- 🧮 Os indicadores são recalculados a cada filtro.
- 📋 Lista com as linhas da planilha do documento no período (Filial, Cód. Cliente, Nome, Nº Documento, Parcela, Emissão, Vencimento, Situação, valores, juros, multa, liquidação etc.).
- 👆 Clique numa linha e o painel mostra só aquela parcela. Com **Ctrl+clique** ou as caixas de seleção, você soma várias.
- 📥 Importa `.xlsx` ou `.csv`. Um arquivo com mais de 150 mil linhas foi importado em poucos segundos.

### Colunas lidas da planilha
`CNPJ / CPF`, `Data Vencimento`, `Valor Original`, `Valor Saldo`, `Valor Total Recebido`, `Juros Calculado`, `Multa Calculada`, `Situação`, mais as colunas exibidas na lista. As demais colunas do arquivo são ignoradas.

### Regras de cálculo
| Indicador | Regra |
|---|---|
| Total Original / Total Saldo | Soma de `Valor Original` / `Valor Saldo` |
| Total Liquidado | `Valor Total Recebido` (se preenchido), senão `Valor Original` das parcelas liquidadas |
| Total Cancelado / Renegociado | `Valor Original` das parcelas nessa situação |
| Juros, Multa e Total a Receber | Só parcelas em aberto. Total a Receber = Saldo + Juros + Multa |
| Período | Filtra pela `Data Vencimento` |

---

## 🧾 Consulta de Contas a Pagar

Mesma ideia, aplicada ao contas a pagar: informe o CNPJ/CPF do fornecedor e veja o resumo e as parcelas.

### Painel de indicadores

```
Total de Parcelas      Total Original      Primeiro Vencimento
                       Total Saldo         Último Vencimento
```

### Recursos
- 🔎 Pesquisa por **CNPJ** ou **CPF**.
- 🎛️ Filtro por **situação** (as opções vêm do próprio arquivo: Aberto, Parcial, Liquidado, Renegociado, Cancelado) e por **período de vencimento**.
- 📋 Lista com as linhas do arquivo do documento (Cod. Fornec., Filial, Nome, Nº Documento, Número AP, Parcela, Entrada, vencimentos, valores, situação, liquidação, histórico e tipo de despesa).
- 👆 Clique nas linhas para atualizar o painel, com as mesmas regras de seleção do projeto anterior.
- 📥 Importa `.csv` do ERP (codificação ISO-8859-1, separador `;`, textos no formato `="valor"`) ou `.xlsx`.

### Colunas lidas
`CNPJ / CPF`, `Vcto Saldo`, `Valor Original`, `Valor Saldo`, `Situação`, mais as colunas exibidas na lista.

### Observações
- O primeiro e o último vencimento e o filtro de período usam `Vcto Saldo`.
- Linhas sem CNPJ/CPF no arquivo não podem ser pesquisadas e **não são importadas**. A tela mostra quantas ficaram de fora.
- Os valores são exibidos como estão no arquivo, inclusive datas digitadas com erro (por exemplo, o ano 8202).

---

## 🚀 Como rodar

**Requisitos:** [Python 3.9+](https://www.python.org/downloads/). No Windows, marque **"Add Python to PATH"** na instalação.

```bash
# 1. clone o repositório
git clone https://github.com/<seu-usuario>/projetosempyhton.git
cd projetosempyhton/consulta_parcelas        # ou consulta_contas_pagar

# 2. instale as dependências
pip install -r requirements.txt

# 3. inicie
python app.py
```

No Windows, basta dar dois cliques em **`iniciar.bat`** dentro da pasta do projeto.
O navegador abre sozinho em `http://127.0.0.1:5000` (contas a receber) ou `http://127.0.0.1:5001` (contas a pagar).

### Primeiro uso
1. No rodapé da tela, escolha a planilha exportada do ERP e clique em **Importar / atualizar**.
2. Escolha **CPF** ou **CNPJ**, digite o número e clique em **Consultar**.
3. Ajuste os filtros e clique nas linhas para detalhar.

Os dados ficam no arquivo `.db` da pasta do projeto. Só é preciso importar de novo quando houver uma planilha nova.

## 🧰 Tecnologias
- Python 3 e [Flask](https://flask.palletsprojects.com/)
- SQLite (módulo `sqlite3` da biblioteca padrão)
- [openpyxl](https://openpyxl.readthedocs.io/) para ler `.xlsx` em fluxo
- HTML, CSS e JavaScript puros, sem frameworks de front-end

## 📁 Estrutura

```
projetosempyhton/
├── consulta_parcelas/        # contas a receber
│   ├── app.py
│   ├── requirements.txt
│   ├── iniciar.bat
│   └── LEIA-ME.txt
├── consulta_contas_pagar/    # contas a pagar
│   ├── app.py
│   ├── requirements.txt
│   ├── iniciar.bat
│   └── LEIA-ME.txt
├── .gitignore
└── README.md
```

## 🔒 Privacidade dos dados

As planilhas do ERP contêm **CPF, CNPJ e dados financeiros reais**. Por isso o `.gitignore` deste repositório bloqueia `*.db`, `*.csv`, `*.xlsx` e a pasta `_uploads/`.
**Não envie planilhas nem bancos para o GitHub.** Os aplicativos rodam apenas na máquina local (`127.0.0.1`) e não enviam nada para a internet.

## 🗺️ Ideias para o futuro
- Exportar a lista filtrada para Excel/CSV
- Escolher qual data usar no filtro de período (`Vcto Saldo` ou `Vencto Original`)
- Reunir os dois projetos em um único aplicativo com abas

---

Feito com Python 🐍 e muito café ☕
# projetosempython
