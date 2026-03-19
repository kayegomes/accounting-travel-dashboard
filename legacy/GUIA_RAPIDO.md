# 🚀 Guia Rápido - Analisador de Orçamento

## Executar Análise Completa

### Comando Único (Recomendado)

```bash
python3.11 /home/ubuntu/painel_contabil/executar_analise_completa.py
```

**O que este comando faz:**
1. ✅ Analisa toda a planilha contábil
2. ✅ Gera relatório Excel consolidado
3. ✅ Cria dashboard HTML interativo
4. ✅ Salva histórico de execução

**Tempo de execução:** ~3-4 segundos

---

## Visualizar Resultados

### Dashboard Interativo (Recomendado)

1. Abra o arquivo no navegador:
   ```
   /home/ubuntu/painel_contabil/dashboard.html
   ```

2. Visualize:
   - 📊 Gráficos interativos
   - 📈 Métricas principais
   - 📋 Tabelas detalhadas

### Relatório Excel

1. Abra o arquivo:
   ```
   /home/ubuntu/painel_contabil/relatorio_consolidado.xlsx
   ```

2. Navegue pelas abas:
   - **Por Grupo de Pessoas**: Viagens por grupo (Elenco, Gerentes, etc.)
   - **Por Produto**: Viagens por plataforma (TV Globo, Sportv, etc.)
   - **Orçado vs Realizado**: Comparativo financeiro
   - **Realizado por Categoria**: Detalhamento de gastos
   - **Resumo Geral**: Visão consolidada

---

## Principais Métricas Geradas

### 💰 Financeiras
- **Orçamento Total**: R$ 18.822.766,67
- **Realizado**: R$ 1.230.703,74
- **Saldo Disponível**: R$ 17.592.062,93
- **% Executado**: 6,54%

### 🚗 Operacionais
- **Total de Viagens**: 8.576
- **Elenco Cadastrado**: 72 pessoas
- **Centros de Custo**: 23

### 👥 Por Grupo de Pessoas
- Elenco: 276 viagens
- Sup. eventos: 48 viagens
- Gerentes: 23 viagens
- Colaborador: 11 viagens
- Esp. eventos: 2 viagens

### 📺 Por Produto/Plataforma
- TV GLOBO: 2.783 viagens (64,9%)
- SPORTV: 1.098 viagens (25,6%)
- PREMIERE: 325 viagens (7,6%)
- AMAZON: 65 viagens (1,5%)
- COMBATE: 17 viagens (0,4%)

---

## Atualizar Dados

1. Substitua a planilha original
2. Execute novamente:
   ```bash
   python3.11 /home/ubuntu/painel_contabil/executar_analise_completa.py
   ```
3. Os relatórios serão atualizados automaticamente

---

## Arquivos Importantes

| Arquivo | Descrição |
|---------|-----------|
| `dashboard.html` | Dashboard interativo com gráficos |
| `relatorio_consolidado.xlsx` | Relatório detalhado em Excel |
| `historico_execucoes.json` | Histórico de todas as análises |
| `README.md` | Documentação completa |

---

## Comandos Úteis

### Executar apenas análise de dados
```bash
python3.11 /home/ubuntu/painel_contabil/analisador_orcamento.py
```

### Executar apenas geração de dashboard
```bash
python3.11 /home/ubuntu/painel_contabil/gerar_dashboard.py
```

### Ver histórico de execuções
```bash
cat /home/ubuntu/painel_contabil/historico_execucoes.json
```

---

## Solução Rápida de Problemas

### Erro: Módulo não encontrado
```bash
pip3 install pandas openpyxl
```

### Erro: Permissão negada
```bash
chmod +x /home/ubuntu/painel_contabil/*.py
```

### Erro: Planilha não encontrada
Verifique se o arquivo está em:
```
/home/ubuntu/upload/PainelContábilV2-copia.xlsx
```

---

## 💡 Dica

Para melhor experiência, abra o **dashboard.html** em um navegador moderno (Chrome, Firefox, Edge) para visualizar os gráficos interativos.

---

**Precisa de mais detalhes?** Consulte o arquivo `README.md` para documentação completa.
