# Análise da Planilha Contábil - Processos Identificados

## 1. Estrutura Geral da Planilha

A planilha **PainelContábilV2-copia.xlsx** possui **32 abas** com diferentes propósitos:

### Abas Principais Identificadas:

1. **RESUMO MENSAL CONTÁBIL** - Consolidação mensal de orçado vs realizado
2. **RESUMO VERBA GESTORES - GG** - Acompanhamento de verbas por gestor (GG)
3. **RESUMO VERBA GESTORES - MARCELA** - Acompanhamento de verbas por gestor (Marcela)
4. **Realizado Verba Gestores** - Base de dados detalhada de realizações
5. **ACOMP por PLATAFORMA** - Acompanhamento por plataforma (TV Globo, Sportv, Premiere, etc.)
6. **Elenco** - Cadastro de profissionais (narradores, comentaristas)
7. **CENTROS DE CUSTO** - Estrutura organizacional e gestores
8. **CENTRO DE RESULTADO** - Produtos/plataformas (Amazon, TV Globo, Sportv, etc.)

### Abas de Apoio/Base de Dados:
- BasePassagens_New, BaseHospedagens_New
- ACOMPANHAMENTO CONTÁBIL VIAGENS, ACOMPANHAMENTO CONTÁBIL HOTEL
- Consolidado Geral (UBER e 99)
- RESUMO LOGÍSTICA

---

## 2. Processos Identificados

### 2.1 Controle Orçamentário
**Aba: RESUMO MENSAL CONTÁBIL**

**Estrutura:**
- Orçamento total anual dividido por categorias
- Acompanhamento mensal de REALIZADO vs ORÇADO
- Cálculo de SALDO e PREVISÃO
- Percentuais de execução

**Categorias principais:**
- Verbas gestores (R$ 2.084.417)
- Custos afundados (R$ 1.808.040)
- Outras verbas (R$ 810.583)
- Produtos (R$ 13.575.967)
  - Futebol (61% do orçamento de produtos)
  - Multimodalidades (29% do orçamento de produtos)

**Problema identificado:** Descasamento entre orçado e realizado (#VALUE!)

---

### 2.2 Gestão de Verbas por Gestor
**Abas: RESUMO VERBA GESTORES - GG e RESUMO VERBA GESTORES - MARCELA**

**Estrutura:**
- Divisão por áreas/departamentos
- Orçamento mensal (Jan a Dez)
- Realizado mensal
- Comparação orçado vs realizado

**Áreas identificadas:**
- **GG:** Direção Transmissões, Futebol
- **MARCELA:** Planejamento e governança, Gestão Eventos

**Exemplo de dados:**
- Planejamento e governança: R$ 40.833/mês orçado
- Realizado variável: Jan R$ 13.022, Fev R$ 95.489, Mar R$ 39.139

---

### 2.3 Base de Dados de Realizações
**Aba: Realizado Verba Gestores**

**Estrutura:**
- Campos: visão painel, visão painel2, visão painel3, Mês, Total, Concat
- Categorias: Amazon, Aquecimento, Custos Afundados (FOOTSTATS, Copinha, Youcast, WSC)
- Valores detalhados por mês

**Exemplos:**
- Amazon (Set): R$ 8.879,49
- Aquecimento (Set): R$ 41.502,47
- Copinha (Mai): R$ 500.000
- FOOTSTATS: múltiplas entradas mensais

---

### 2.4 Acompanhamento por Plataforma/Produto
**Aba: ACOMP por PLATAFORMA**

**Estrutura:**
- Total por plataforma PNT (Passagem + Hospedagem)
- Período: 02/12 à 31/10
- Métricas: META vs REALIZADO (quantidade)

**Plataformas:**
- TV GLOBO: 2.783 viagens
- SPORTV: 1.098 viagens
- PREMIERE: 325 viagens
- AMAZON: 65 viagens
- COMBATE: 17 viagens
- **TOTAL: 4.288 viagens**

**Grupos de pessoas:**
- Elenco: 276 viagens
- Sup. eventos: 48 viagens
- Gerentes: 23 viagens
- Colaborador: 11 viagens
- Esp. eventos: 2 viagens

---

### 2.5 Cadastros de Referência

**Aba: Elenco**
- Nome completo, nome conhecido (UT)
- Função: Narrador, Comentaristas
- Faixa: N1-N4, C1-C4

**Aba: CENTROS DE CUSTO**
- Código GL
- Área Resumo e Nome
- Gestor responsável
- Estrutura: Direção Esporte, Eventos, Redação

**Aba: CENTRO DE RESULTADO**
- Código do centro de resultado
- Produtos: AMAZON, TV GLOBO, SPORTV, PREMIERE, COMBATE, GE, CARTOLA

---

## 3. Problemas e Oportunidades de Automação

### 3.1 Problemas Identificados:
1. **Descasamento de dados** - Indicação de #VALUE! na aba RESUMO MENSAL CONTÁBIL
2. **Múltiplas fontes de dados** - Dados espalhados em várias abas
3. **Falta de consolidação automática** - Processos manuais de agregação
4. **Dificuldade de visualização** - Orçado vs Realizado por grupo e produto não está consolidado

### 3.2 Oportunidades de Automação:
1. **Consolidação automática** de orçado vs realizado por:
   - Grupo de pessoas (Elenco, Gerentes, Colaboradores, etc.)
   - Produto/Plataforma (TV Globo, Sportv, Premiere, Amazon, etc.)
   - Centro de custo
   - Gestor

2. **Dashboard interativo** com:
   - Filtros por período, produto, grupo
   - Gráficos de evolução mensal
   - Alertas de desvio orçamentário
   - Percentual de execução

3. **Tratamento de dados** automatizado:
   - Limpeza de dados
   - Validação de valores
   - Cálculo automático de saldos e previsões
   - Detecção de inconsistências

---

## 4. Proposta de Solução

### Abordagem Recomendada:
Desenvolver um **script Python** que:

1. **Leia e consolide** dados de múltiplas abas
2. **Normalize e limpe** os dados
3. **Calcule métricas** de orçado vs realizado
4. **Gere relatórios** por:
   - Grupo de pessoas
   - Produto/Plataforma
   - Centro de custo
   - Período (mensal, acumulado)
5. **Exporte resultados** em formato Excel com dashboards ou em formato web interativo

### Tecnologias:
- Python + Pandas (manipulação de dados)
- OpenPyXL (leitura/escrita Excel)
- Plotly/Matplotlib (visualizações)
- Streamlit ou Flask (interface web opcional)
