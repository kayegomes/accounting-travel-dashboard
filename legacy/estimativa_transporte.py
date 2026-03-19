import pandas as pd
import numpy as np
from datetime import datetime

caminho = r"C:\Users\ligomes\Downloads\Painel Contábil V2 - copia.xlsx"

print("=" * 80)
print("🚗 ESTIMATIVA DETALHADA DE TRANSPORTE POR CAMPEONATO")
print("=" * 80)
print("\nMétodo: Distribuição proporcional baseada em Passagens + Hospedagens")
print("=" * 80)

# 1. Carregar dados de Transporte (consolidados mensais)
print("\n📊 Carregando dados de Transporte...")
df_trans = pd.read_excel(caminho, sheet_name='Consolidado Geral (UBER e 99)')

# Aplicar correções
palavras_filtro = ['total', 'subtotal', 'geral', 'soma']
mask = df_trans['ÁREA'].astype(str).str.lower().str.contains('|'.join(palavras_filtro), na=False)
df_trans = df_trans[~mask]

df_trans['Dia'] = df_trans['Mês pagamento'].dt.day
df_trans = df_trans[df_trans['Dia'] == 1]  # Apenas consolidados mensais
df_trans['Total'] = pd.to_numeric(df_trans['Total'], errors='coerce')

print(f"✅ Transporte carregado: R$ {df_trans['Total'].sum():,.2f} em {len(df_trans)} consolidados mensais")

# 2. Carregar Passagens
print("\n📊 Carregando dados de Passagens...")
df_pass = pd.read_excel(caminho, sheet_name='BasePassagens_New')
df_pass['VALOR AJUSTADO'] = pd.to_numeric(df_pass['VALOR AJUSTADO'], errors='coerce')
df_pass = df_pass[df_pass['VALOR AJUSTADO'].notna()]

# Extrair mês/ano da data
df_pass['Data'] = pd.to_datetime(df_pass['Data'], errors='coerce')
df_pass['Mês'] = df_pass['Data'].dt.month
df_pass['Ano'] = df_pass['Data'].dt.year

print(f"✅ Passagens carregadas: R$ {df_pass['VALOR AJUSTADO'].sum():,.2f}")

# 3. Carregar Hospedagens
print("\n📊 Carregando dados de Hospedagens...")
df_hosp = pd.read_excel(caminho, sheet_name='BaseHospedagens_New')
df_hosp['TOTAL AJUSTADO'] = pd.to_numeric(df_hosp['TOTAL AJUSTADO'], errors='coerce')
df_hosp = df_hosp[df_hosp['TOTAL AJUSTADO'].notna()]

df_hosp['Data'] = pd.to_datetime(df_hosp['Data'], errors='coerce')
df_hosp['Mês'] = df_hosp['Data'].dt.month
df_hosp['Ano'] = df_hosp['Data'].dt.year

print(f"✅ Hospedagens carregadas: R$ {df_hosp['TOTAL AJUSTADO'].sum():,.2f}")

# 4. Criar base para distribuição
print("\n" + "=" * 80)
print("🔄 CALCULANDO DISTRIBUIÇÃO PROPORCIONAL")
print("=" * 80)

# Consolidar Passagens + Hospedagens por Mês/Ano/Área/Campeonato
df_pass_group = df_pass.groupby(['Mês', 'Ano', 'Área', 'Nome Projeto'])['VALOR AJUSTADO'].sum().reset_index()
df_pass_group.columns = ['Mês', 'Ano', 'Área', 'Campeonato', 'Valor_Passagens']

df_hosp_group = df_hosp.groupby(['Mês', 'Ano', 'Área', 'Nome Projeto'])['TOTAL AJUSTADO'].sum().reset_index()
df_hosp_group.columns = ['Mês', 'Ano', 'Área', 'Campeonato', 'Valor_Hospedagens']

# Merge
df_base = pd.merge(df_pass_group, df_hosp_group, 
                   on=['Mês', 'Ano', 'Área', 'Campeonato'], 
                   how='outer').fillna(0)

df_base['Valor_Total_Base'] = df_base['Valor_Passagens'] + df_base['Valor_Hospedagens']

print(f"\n✅ Base criada com {len(df_base)} registros únicos (Mês/Ano/Área/Campeonato)")

# 5. Distribuir transporte proporcionalmente
estimativas = []

for _, trans_row in df_trans.iterrows():
    mes = trans_row['Mês']
    ano = trans_row['Ano']
    area = trans_row['ÁREA']
    valor_trans = trans_row['Total']
    
    # Filtrar registros do mesmo mês/ano/área
    df_filtrado = df_base[
        (df_base['Mês'] == mes) & 
        (df_base['Ano'] == ano) & 
        (df_base['Área'] == area)
    ].copy()
    
    if len(df_filtrado) > 0:
        # Calcular proporção baseada em Passagens + Hospedagens
        total_base = df_filtrado['Valor_Total_Base'].sum()
        
        if total_base > 0:
            df_filtrado['Proporcao'] = df_filtrado['Valor_Total_Base'] / total_base
            df_filtrado['Transporte_Estimado'] = df_filtrado['Proporcao'] * valor_trans
            
            for _, row in df_filtrado.iterrows():
                estimativas.append({
                    'Mês': int(mes),
                    'Ano': int(ano),
                    'Área': area,
                    'Campeonato': row['Campeonato'],
                    'Passagens': row['Valor_Passagens'],
                    'Hospedagens': row['Valor_Hospedagens'],
                    'Transporte_Estimado': row['Transporte_Estimado'],
                    'Total_Estimado': row['Valor_Passagens'] + row['Valor_Hospedagens'] + row['Transporte_Estimado'],
                    'Proporcao': row['Proporcao']
                })

df_estimativa = pd.DataFrame(estimativas)

print(f"\n✅ Estimativa gerada para {len(df_estimativa)} registros")

# 6. Validação
total_estimado = df_estimativa['Transporte_Estimado'].sum()
total_real = df_trans['Total'].sum()
diferenca = abs(total_estimado - total_real)

print("\n" + "=" * 80)
print("✅ VALIDAÇÃO DA ESTIMATIVA")
print("=" * 80)
print(f"Transporte Real (consolidados):    R$ {total_real:,.2f}")
print(f"Transporte Estimado (distribuído): R$ {total_estimado:,.2f}")
print(f"Diferença:                         R$ {diferenca:,.2f} ({(diferenca/total_real)*100:.2f}%)")

# 7. Consolidar por Campeonato
print("\n" + "=" * 80)
print("📊 TOP 20 CAMPEONATOS - ESTIMATIVA DE TRANSPORTE")
print("=" * 80)

por_campeonato = df_estimativa.groupby('Campeonato').agg({
    'Passagens': 'sum',
    'Hospedagens': 'sum',
    'Transporte_Estimado': 'sum',
    'Total_Estimado': 'sum'
}).reset_index()

por_campeonato = por_campeonato.sort_values('Total_Estimado', ascending=False)

print("\n{:<50} {:>15} {:>15} {:>15} {:>15}".format(
    "Campeonato", "Passagens", "Hospedagens", "Transporte", "TOTAL"
))
print("-" * 110)

for _, row in por_campeonato.head(20).iterrows():
    camp = row['Campeonato'][:47] + "..." if len(row['Campeonato']) > 50 else row['Campeonato']
    print("{:<50} R$ {:>12,.2f} R$ {:>12,.2f} R$ {:>12,.2f} R$ {:>12,.2f}".format(
        camp,
        row['Passagens'],
        row['Hospedagens'],
        row['Transporte_Estimado'],
        row['Total_Estimado']
    ))

# 8. Salvar resultado
caminho_saida = r"C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil\transporte_estimado_por_campeonato.xlsx"

with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
    # Aba 1: Por Campeonato
    por_campeonato.to_excel(writer, sheet_name='Por Campeonato', index=False)
    
    # Aba 2: Por Área
    por_area = df_estimativa.groupby('Área').agg({
        'Passagens': 'sum',
        'Hospedagens': 'sum',
        'Transporte_Estimado': 'sum',
        'Total_Estimado': 'sum'
    }).reset_index().sort_values('Total_Estimado', ascending=False)
    por_area.to_excel(writer, sheet_name='Por Área', index=False)
    
    # Aba 3: Detalhado (Mês/Ano/Área/Campeonato)
    df_estimativa_exportar = df_estimativa.copy()
    df_estimativa_exportar = df_estimativa_exportar.sort_values('Total_Estimado', ascending=False)
    df_estimativa_exportar.to_excel(writer, sheet_name='Detalhado', index=False)
    
    # Aba 4: Resumo
    resumo_data = {
        'Componente': ['Passagens', 'Hospedagens', 'Transporte (Estimado)', 'TOTAL'],
        'Valor': [
            df_estimativa['Passagens'].sum(),
            df_estimativa['Hospedagens'].sum(),
            df_estimativa['Transporte_Estimado'].sum(),
            df_estimativa['Total_Estimado'].sum()
        ]
    }
    df_resumo = pd.DataFrame(resumo_data)
    df_resumo.to_excel(writer, sheet_name='Resumo', index=False)

print("\n" + "=" * 80)
print("💾 ARQUIVO GERADO")
print("=" * 80)
print(f"📁 {caminho_saida}")
print("\nAbas incluídas:")
print("  1. Por Campeonato - Totais consolidados")
print("  2. Por Área - Totais por grupo")
print("  3. Detalhado - Mês a mês com todas as informações")
print("  4. Resumo - Visão geral")

print("\n" + "=" * 80)
print("✅ ESTIMATIVA CONCLUÍDA COM SUCESSO!")
print("=" * 80)

print("\n💡 METODOLOGIA:")
print("   O transporte foi distribuído proporcionalmente aos gastos de")
print("   Passagens + Hospedagens de cada campeonato no mesmo mês/área.")
print("   Exemplo: Se um campeonato representa 30% dos gastos de um mês,")
print("   ele recebe 30% do transporte daquele mês.")
print("=" * 80)