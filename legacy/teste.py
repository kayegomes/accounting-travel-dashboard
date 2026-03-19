import pandas as pd
from pathlib import Path

# Carregar a aba RESUMO LOGÍSTICA
caminho_planilha = Path(r"C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil\Painel Contábil V2 - copia.xlsx")

df_resumo = pd.read_excel(caminho_planilha, sheet_name='RESUMO LOGÍSTICA', header=None)

# Extrair valores orçados e realizados
# Linha 10 (índice 6): Total Logística
# Linha 14 (índice 10): Passagens
# Linha 18 (índice 14): Hospedagens
# Linha 22 (índice 18): Transporte

orcado_total = df_resumo.iloc[6, 1]  # 8.848.513
realizado_total = df_resumo.iloc[6, 2]  # 5.888.029,75

orcado_passagens = df_resumo.iloc[10, 1]  # 5.272.348
realizado_passagens = df_resumo.iloc[10, 2]  # 3.021.070,75

orcado_hospedagens = df_resumo.iloc[14, 1]  # 1.522.779
realizado_hospedagens = df_resumo.iloc[14, 2]  # 2.222.970,41

orcado_transporte = df_resumo.iloc[18, 1]  # 2.053.386
realizado_transporte = df_resumo.iloc[18, 2]  # 643.988,59

print(f"Orçado Total: R$ {orcado_total:,.2f}")
print(f"Realizado Total: R$ {realizado_total:,.2f}")
print(f"\nOrçado Passagens: R$ {orcado_passagens:,.2f}")
print(f"Realizado Passagens: R$ {realizado_passagens:,.2f}")
print(f"\nOrçado Hospedagens: R$ {orcado_hospedagens:,.2f}")
print(f"Realizado Hospedagens: R$ {realizado_hospedagens:,.2f}")
print(f"\nOrçado Transporte: R$ {orcado_transporte:,.2f}")
print(f"Realizado Transporte: R$ {realizado_transporte:,.2f}")
