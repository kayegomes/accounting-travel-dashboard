import pandas as pd

caminho = r"C:\Users\ligomes\Downloads\Painel Contábil V2 - copia.xlsx"

# Passagens
print("="*80)
print("PASSAGENS - Análise de Colunas")
print("="*80)
df = pd.read_excel(caminho, sheet_name='BasePassagens_New')

print(f"\nTotal linhas: {len(df)}")
print(f"\nColunas disponíveis: {df.columns.tolist()}")

# Mostrar valores únicos de Natureza
if 'Natureza' in df.columns:
    print(f"\nValores únicos de Natureza:")
    print(df['Natureza'].value_counts())

# Mostrar distribuição por mês
for col in ['Z', 'MÊS', 'Mês']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        print(f"\nDistribuição de {col}:")
        print(df[col].value_counts().sort_index())
        print(f"\nRegistros com {col} <= 8: {(df[col] <= 8).sum()}")
        print(f"Valor total com {col} <= 8: R$ {df[df[col] <= 8]['VALOR AJUSTADO'].sum():,.2f}")

# Hospedagens
print("\n" + "="*80)
print("HOSPEDAGENS - Análise de Colunas")
print("="*80)
df = pd.read_excel(caminho, sheet_name='BaseHospedagens_New')

print(f"\nTotal linhas: {len(df)}")
print(f"\nColunas disponíveis: {df.columns.tolist()}")

if 'Natureza' in df.columns:
    print(f"\nValores únicos de Natureza:")
    print(df['Natureza'].value_counts())

for col in ['MÊS', 'Mês', 'AD']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        print(f"\nDistribuição de {col}:")
        print(df[col].value_counts().sort_index())
        print(f"\nRegistros com {col} <= 8: {(df[col] <= 8).sum()}")
        print(f"Valor total com {col} <= 8: R$ {df[df[col] <= 8]['TOTAL AJUSTADO'].sum():,.2f}")