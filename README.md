<h1 align="center">
  📊 Dashboard Contábil - Gerador V3
</h1>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white">
  <img alt="Pandas" src="https://img.shields.io/badge/Pandas-v2.0+-150458?logo=pandas&logoColor=white">
  <img alt="HTML5" src="https://img.shields.io/badge/HTML5-interativo-E34F26?logo=html5&logoColor=white">
</p>

<p align="center">
  <strong>Transformação ágil de planilhas contábeis de custos logísticos em dashboards visuais e interativos.</strong>
</p>

---

## 📖 Sobre o Projeto

O **Dashboard Contábil V3** baseia-se no script automatizador `gerar_dashboard_v3.py`. Ele foi criado para processar extensas bases de dados financeiras extraídas num formato `.xlsx` cru, estruturando-as através de processos de limpeza, tratamento de anomalias textuais e cruzamentos matemáticos. 

O resultado final é a entrega de um relátorio de alta qualidade em um **painel HTML interativo** para a avaliação e visualização ágil das métricas logísticas do evento, sem necessidade de servidores complexos ou ferramentas pagas.

## ✨ Principais Funcionalidades

- **🧮 Agrupamento Inteligente**: Tratamento de normalização de strings para destinos, áreas e grupos produtores, unificando falhas de preenchimento.
- **📈 Visualização Interativa via HTML**: Um relatório dinâmico que pode ser acessado em navegadores web locais a partir de códigos injetados pelo Python.
- **✈️ Tracking de Logística Profunda**: Dissecação dos custos entre diárias de hotéis (Hospedagens), passagens aéreas, estimativas de transporte terrestre (Uber/99), e suas médias em orçado vs. realizado.
- **🎯 Separação Nacionais/Internacional**: Algoritmo que reflete relatórios seccionais identificando fronteiras nos gastos dos colaboradores.

## 🚀 Tecnologias

A aplicação é construída majoritariamente com processamento de dados puro fornecido pelas seguintes bibliotecas de Python:

- **[Python 3.8+](https://www.python.org/)**  
- **[Pandas](https://pandas.pydata.org/)** (Tratamento analítico e computacional sobre o dataframe)
- **[Openpyxl](https://openpyxl.readthedocs.io/en/stable/)** (Leitura vital das formatações e células da Microsoft Excel)

---

## 🛠️ Como Usar

### 1. Pré-Requisitos e Clonagem
Garanta que seu sistema contenha o [Python](https://www.python.org/downloads/) corretamente alocado no `PATH`. Sendo assim, instale os componentes da aplicação:

```bash
# Baixe ou clone o repositório
git clone https://github.com/kayegomes/accounting-travel-dashboard.git

# Acesse a pasta do script principal
cd painel_contabil

# Instale os requisitos da engenharia de dados
pip install -r requirements.txt
```

### 2. Rodando o Script de Geração
A aplicação necessita de um ponto de origem. Adicione seus documentos fontes de Excel no diretório do projeto seguindo uma formatação de layout e então inicie o processamento:

```bash
python gerar_dashboard_v3.py
```

O código interpretará os dados limpos, cruzar e formatar os cálculos estatísticos e irá conceber o arquivo `dashboard_v3_8.html` (ou nomenclatura similar definida no código final) contendo os belíssimos agrupadores visuais de custos operacionais e KPIs.

---

## 📁 A Estrutura do Diretório

A organização do ecossistema que irá para o repositório baseia-se e mantém a estrutura:

```text
📦 painel_contabil
 ┣ 📜 gerar_dashboard_v3.py        # Core Engine: Script responsável pelo dashboard
 ┣ 📜 config.json                  # Parâmetros de configurações auxiliares (opcional)
 ┣ 📜 requirements.txt             # Dependências de ambiente (Pandas, Openpyxl)
 ┣ 📜 .gitignore                   # Exclusão de caches, e ocultações de bases Excel
 ┣ 📜 README.md                    # Documentação global (você está lendo isto)
 ┗ 📂 legacy/                      # Arquivos auxiliares, logs retroativos e antigas versões 
```

> **Aviso GERAL & Privacidade Logística**: Por imposição de confidencialidade de massas de dados financeiras de empresas, os arquivos com a extensão explícita em planilhas (`.xlsx`/`.csv`) ou os relatórios espelhos (`.html`) de finalizações da ferramenta **NÃO** compõe ou tracionam na lista de envios para o GitHub (veja `.gitignore`).
