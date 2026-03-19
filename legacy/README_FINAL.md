# 📊 Analisador de Painel Contábil - Versão Final

## Visão Geral

Sistema completo e inteligente para análise automatizada de dados contábeis com foco em **acompanhamento de orçado vs. realizado** por grupo de pessoas e por produto (hierarquia MACRO/MICRO).

### ✨ Recursos Principais

- ✅ **Análise Completa de Logística**: Passagens + Hospedagens + Transporte (Uber/99)
- ✅ **Hierarquia de Produtos**: MACRO (Plataformas) + MICRO (Campeonatos específicos)
- ✅ **Análise por Grupo de Pessoas**: Elenco, Tecnologia, Repórter, Produção, etc.
- ✅ **Inteligência Artificial**: Análise automática com insights e recomendações
- ✅ **Planilha Consolidada Tratada**: Arquivo Excel pronto para uso
- ✅ **Dashboard Interativo**: Visualização HTML com gráficos dinâmicos

---

## 📁 Estrutura de Arquivos

### Arquivos Principais

| Arquivo | Descrição |
|---------|-----------|
| `executar_analise_final.py` | **Script principal** - Execute este para análise completa |
| `analisador_completo_v3.py` | Motor de análise com IA |
| `gerar_dashboard_v3.py` | Gerador de dashboard HTML |
| `planilha_consolidada_tratada.xlsx` | **Planilha tratada** - Abra no Excel |
| `dashboard_v3.html` | **Dashboard interativo** - Abra no navegador |

### Arquivos de Suporte

- `README_FINAL.md` - Esta documentação
- `GUIA_RAPIDO.md` - Guia rápido de uso

---

## 🚀 Como Usar

### Opção 1: Execução Rápida (Recomendado)

```bash
python3.11 executar_analise_final.py
```

Este comando irá:
1. Analisar todos os componentes de gasto
2. Gerar planilha consolidada tratada
3. Executar análise com IA
4. Criar dashboard interativo

### Opção 2: Execução Modular

```bash
# Apenas análise e planilha tratada
python3.11 analisador_completo_v3.py

# Apenas dashboard (requer planilha tratada)
python3.11 gerar_dashboard_v3.py
```

---

## 📊 Componentes Analisados

### Total de Logística: **R$ 11.669.253,76**

| Componente | Valor | Percentual |
|------------|-------|------------|
| ✈️ Passagens | R$ 6.452.392,69 | 55,3% |
| 🏨 Hospedagens | R$ 3.545.091,59 | 30,4% |
| 🚗 Transporte (Uber/99) | R$ 1.671.769,48 | 14,3% |

---

## 🎯 Hierarquia de Produtos

### MACRO - Plataformas

Análise por plataforma de transmissão:
- TV Globo
- Sportv
- Premiere
- Combate
- Amazon
- Globoplay

### MICRO - Campeonatos

Análise por evento/campeonato específico:
- Brasileirão Série A
- Copa do Brasil
- Copa Libertadores
- Jogos Olímpicos
- WSL (World Surf League)
- Superliga de Vôlei
- E muitos outros...

---

## 👥 Grupos de Pessoas

Análise por categoria profissional:
- Elenco
- Grandes Eventos
- Tecnologia
- Repórter
- Produção de Eventos
- Repcine
- Gerentes
- Colaboradores
- E outros...

---

## 🤖 Análise com IA

O sistema utiliza **Inteligência Artificial (GPT-4.1-mini)** para:

1. **Identificar Insights**: Padrões e tendências nos dados
2. **Áreas de Atenção**: Possíveis problemas ou riscos
3. **Recomendações**: Sugestões práticas para otimização

### Exemplo de Análise IA

```
Principais Insights:
- Passagens representam 55,3% dos custos, indicando oportunidade de otimização
- Plataforma "Futebol" concentra 40% dos gastos
- Eventos de alto custo: Copa do Mundo de Clubes e Brasileirão Série A

Recomendações:
- Negociar tarifas corporativas com companhias aéreas
- Implementar políticas de viagem mais rigorosas
- Avaliar alternativas de transporte terrestre para trajetos curtos
```

---

## 📈 Arquivos Gerados

### 1. Planilha Consolidada Tratada (`planilha_consolidada_tratada.xlsx`)

Abas disponíveis:
- **Resumo Executivo**: Visão geral dos componentes
- **Por Plataforma (MACRO)**: Análise por plataforma
- **Por Campeonato (MICRO)**: Análise por campeonato
- **Por Grupo de Pessoas**: Análise por categoria profissional
- **Detalhamentos**: Passagens, Hospedagens e Transporte separados

### 2. Dashboard Interativo (`dashboard_v3.html`)

Recursos:
- Gráficos interativos (Pizza, Barras, Donut)
- Tabelas detalhadas
- Responsivo (funciona em desktop e mobile)
- Não requer internet (funciona offline)

---

## 🔄 Atualizando os Dados

Para processar uma nova planilha:

1. **Substitua o arquivo original**:
   ```bash
   cp sua_nova_planilha.xlsx /home/ubuntu/upload/PainelContábilV2-copia.xlsx
   ```

2. **Execute a análise novamente**:
   ```bash
   python3.11 executar_analise_final.py
   ```

3. **Abra os novos arquivos gerados**:
   - `planilha_consolidada_tratada.xlsx`
   - `dashboard_v3.html`

---

## ⚙️ Requisitos Técnicos

### Python 3.11+

Bibliotecas necessárias (já instaladas):
- `pandas` - Manipulação de dados
- `openpyxl` - Leitura/escrita de Excel
- `openai` - Análise com IA (opcional)

### Variáveis de Ambiente

- `OPENAI_API_KEY` - Já configurada para análise com IA

---

## 🎨 Personalização

### Modificar Análise

Edite `analisador_completo_v3.py` para:
- Adicionar novos componentes de gasto
- Alterar agrupamentos
- Modificar cálculos

### Modificar Dashboard

Edite `gerar_dashboard_v3.py` para:
- Alterar cores e estilos
- Adicionar novos gráficos
- Modificar layout

---

## 📝 Estrutura de Dados

### Entrada (Planilha Original)

Abas utilizadas:
- `BasePassagens_New` - Dados de passagens
- `BaseHospedagens_New` - Dados de hospedagens
- `Consolidado Geral (UBER e 99)` - Dados de transporte
- `Chaves - 2025` - Mapeamento de produtos
- `RESUMO MENSAL CONTÁBIL` - Orçamento e realizado

### Saída (Planilha Tratada)

Estrutura consolidada e limpa:
- Dados agregados por hierarquia
- Valores calculados e validados
- Formato pronto para análise

---

## 🔍 Troubleshooting

### Erro: "Planilha não encontrada"

**Solução**: Verifique se o arquivo está em `/home/ubuntu/upload/PainelContábilV2-copia.xlsx`

### Erro: "Cliente IA não disponível"

**Solução**: O sistema continuará funcionando sem a análise com IA. Apenas os insights automáticos não serão gerados.

### Dashboard não abre

**Solução**: Abra o arquivo `dashboard_v3.html` diretamente no navegador (Chrome, Firefox, Edge, Safari).

---

## 💡 Dicas de Uso

1. **Análise Rápida**: Use o dashboard HTML para visualização rápida
2. **Análise Detalhada**: Use a planilha Excel para análises aprofundadas
3. **Exportação**: Copie dados da planilha tratada para seus relatórios
4. **Automação**: Configure para executar periodicamente (semanal/mensal)

---

## 📞 Suporte

Para dúvidas ou problemas:
1. Consulte este README
2. Verifique o `GUIA_RAPIDO.md`
3. Revise os logs de execução no terminal

---

## 📜 Changelog

### Versão Final (V3) - 13/11/2025

- ✅ Adicionado componente de Transporte (Uber/99)
- ✅ Implementada análise com IA
- ✅ Geração de planilha consolidada tratada
- ✅ Dashboard completo com todos os componentes
- ✅ Hierarquia completa MACRO/MICRO
- ✅ Performance otimizada (~60s de execução)

### Versão 2 (V2)

- ✅ Hierarquia de produtos implementada
- ✅ Cálculos corrigidos
- ✅ Passagens e Hospedagens

### Versão 1 (V1)

- ✅ Versão inicial
- ⚠️ Componente de transporte faltando

---

## 🎉 Conclusão

Este sistema oferece uma solução completa e automatizada para análise do painel contábil, economizando tempo e fornecendo insights valiosos para tomada de decisão.

**Desenvolvido com ❤️ para otimizar a gestão financeira de eventos esportivos.**
