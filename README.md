# convenio-coparticipacao
Coparticipação

# 🏥 Automação de Coparticipação de Convênio Médico

Pipeline de tratamento, consolidação e enriquecimento cadastral de relatórios mensais de coparticipação médica para importação em sistemas de Recursos Humanos / Folha de Pagamento.

## 📌 Visão Geral

O projeto automatiza o processamento de faturas de convênio médico em formato `.xls` emitidas por operadoras de saúde. Ele substitui o trabalho manual de auditoria ao realizar as seguintes etapas:

1. **Varredura e Extração**: Identifica blocos de titulares, beneficiários e custos associados (colunas I e J).
2. **Consolidação Financeira**: Agrupa múltiplos procedimentos por beneficiário e calcula o valor total devido.
3. **Cruzamento de Dados Cadastrais (RH)**: Realiza o `merge` inteligente com a base interna da empresa, atribuindo a `CHAPA` do colaborador e o `NÚMERO DEPENDENTE`.
4. **Tratamento de Divergências de Nome**: Aplica normalização de texto (remoção de acentos e padronização em maiúsculas) para garantir correspondência exata mesmo com pequenas variações de grafia.
5. **Auditoria de Erros**: Gera uma aba separada de pendências para beneficiários não identificados no cadastro oficial.

---

## 🛠️ Tecnologias Utilizadas

* **Python 3.10+**
* **Pandas**: Manipulação e agregação de dados.
* **xlrd**: Leitura de arquivos legados de planilha (`.xls`).
* **openpyxl**: Exportação e formatação de arquivos Excel multi-aba (`.xlsx`).
* **Unicodedata & Re**: Expressões regulares e normalização de texto Unicode.

---

## 📁 Estrutura do Projeto

```text
convenio-coparticipacao/
│
├── .dados/                                  # Pasta de arquivos de entrada e saída
│   ├── arquivo_operadora_mes.xls            # Relatório da operadora (baixado mensalmente)
│   ├── Chapa colab e Dados dependentes.xlsx # Tabela oficial do cadastro de RH
│   └── coparticipacao_consolidada_rh.xlsx   # Resultado final gerado pelo script
│
├── venv/                                    # Ambiente virtual Python
├── .gitignore                               # Arquivos/pastas ignorados pelo Git
├── requirements.txt                         # Dependências do projeto
├── varredura.py                             # Script principal de execução
└── README.md                                # Documentação do projeto

```
---

## ⚙️ Configuração e Instalação

1. **Clone o repositório:**
```bash
git clone https://github.com/Anecy-Otta/convenio-coparticipacao.git
cd convenio-coparticipacao

```


2. **Crie e ative o ambiente virtual (`venv`):**
* **Windows:**
```bash
python -m venv venv
.\venv\Scripts\activate

```


* **Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate

```




3. **Instale as dependências:**
```bash
pip install -r requirements.txt

```



---

## 🚀 Como Executar (Rotina Mensal)

O script possui **detecção automática** do arquivo de faturamento mais recente. Não é necessário alterar o código-fonte mês a mês.

1. Baixe o relatório de faturamento do mês no portal da operadora (formato `.xls`).
2. Mova o arquivo baixado para a pasta `.dados/`.
3. Caso tenha ocorrido novas admissões ou inclusões de dependentes no mês, atualize a planilha `Chapa colab e Dados dependentes.xlsx` na pasta `.dados/`.
4. Execute o script no terminal:
```bash
python varredura.py

```



---

## 📊 Estrutura do Arquivo Gerado

O arquivo final é salvo em `.dados/coparticipacao_consolidada_rh.xlsx` e contém:

### Aba 1: `Consolidado_Folha`

Arquivo final limpo e pronto para importação no sistema de folha de pagamento:

| CHAPA | Titular | NÚMERO DEPENDENTE | Beneficiário/Dependente | Grau Parentesco | Codigo Titular | Valor Coparticipação |
| --- | --- | --- | --- | --- | --- | --- |
| `1001` | COLABORADOR EXEMPLO A | `0` | COLABORADOR EXEMPLO A | `00` | 12345 | 45.80 |
| `1001` | COLABORADOR EXEMPLO A | `1` | DEPENDENTE EXEMPLO A | `01` | 12345 | 19.78 |

### Aba 2: `CONFERIR_NO_RH` *(Gerada apenas se houver pendências)*

Lista os beneficiários presentes na fatura que não foram localizados na planilha do RH, fornecendo uma **sugestão de chapa** baseada no primeiro nome do titular para agilizar a auditoria humana.

```

```