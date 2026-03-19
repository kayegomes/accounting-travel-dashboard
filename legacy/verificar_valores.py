"""Diagnóstico: Estrutura da aba RESUMO LOGÍSTICA"""

import pandas as pd
from pathlib import Path
import os

def diagnosticar_resumo_logistica():
    """Verifica a estrutura exata da aba RESUMO LOGÍSTICA"""
    
    if os.name == 'nt':
        onedrive_base = Path(r"C:\Users\ligomes\OneDrive - Globo Comunicação e Participações sa")
        caminho_painel = onedrive_base / "Gestão de Eventos - Documentos" / "Gestão de Eventos_planejamento" / "Painel Contábil" / "2025" / "Painel Contábil V2.xlsx"
    else:
        caminho_painel = Path('/home/ubuntu/upload/PainelContábilV2-copia.xlsx')
    
    print("="*80)
    print("🔍 DIAGNÓSTICO: Estrutura da aba RESUMO LOGÍSTICA")
    print("="*80)
    
    try:
        # Ler sem header para ver a estrutura crua
        df = pd.read_excel(caminho_painel, sheet_name='RESUMO LOGÍSTICA', header=None)
        
        print(f"\n✅ Planilha carregada: {len(df)} linhas x {len(df.columns)} colunas")
        
        # Mostrar as primeiras 20 linhas
        print(f"\n📋 Primeiras 20 linhas (índice | coluna A | coluna B | coluna C):")
        print("-"*80)
        
        for idx in range(min(20, len(df))):
            col_a = str(df.iloc[idx, 0])[:40] if pd.notna(df.iloc[idx, 0]) else 'NaN'
            col_b = str(df.iloc[idx, 1])[:40] if pd.notna(df.iloc[idx, 1]) else 'NaN'
            col_c = str(df.iloc[idx, 2])[:40] if len(df.columns) > 2 and pd.notna(df.iloc[idx, 2]) else 'NaN'
            
            print(f"Linha {idx:2d} | {col_a:40s} | {col_b:40s} | {col_c:40s}")
        
        print("\n" + "="*80)
        print("🔍 PROCURANDO VALORES ESPECÍFICOS:")
        print("="*80)
        
        # Procurar linhas com "TOTAL LOGÍSTICA", "PASSAGENS", etc.
        keywords = ['TOTAL LOGÍSTICA', 'PASSAGENS', 'HOSPEDAGENS', 'TRANSPORTE', 'Realizado', 'Orçamento']
        
        for keyword in keywords:
            print(f"\n🔎 Procurando por '{keyword}':")
            for idx in range(len(df)):
                for col_idx in range(min(5, len(df.columns))):
                    cell_value = str(df.iloc[idx, col_idx]).upper()
                    if keyword.upper() in cell_value:
                        print(f"   Encontrado na linha {idx}, coluna {col_idx}")
                        print(f"   → Linha completa: {list(df.iloc[idx, :5])}")
                        break
        
        print("\n" + "="*80)
        print("💡 SUGESTÃO DE ÍNDICES:")
        print("="*80)
        
        # Tentar encontrar automaticamente
        for idx in range(len(df)):
            linha_str = str(df.iloc[idx, 0]).upper()
            
            if 'TOTAL' in linha_str and 'LOGÍSTICA' in linha_str:
                print(f"\n📍 TOTAL LOGÍSTICA encontrado na linha {idx}")
                # Procurar "Realizado" nas linhas seguintes
                for offset in range(1, 5):
                    if idx + offset < len(df):
                        valor_realizado = df.iloc[idx + offset, 1]
                        print(f"   Linha {idx + offset}, Coluna B: {valor_realizado}")
                        if pd.notna(valor_realizado) and isinstance(valor_realizado, (int, float)):
                            print(f"   ✅ Possível valor realizado: {valor_realizado:,.2f}")
            
            elif 'PASSAGENS' in linha_str and 'HOSPEDAGENS' not in linha_str:
                print(f"\n📍 PASSAGENS encontrado na linha {idx}")
                for offset in range(1, 5):
                    if idx + offset < len(df):
                        valor_realizado = df.iloc[idx + offset, 1]
                        print(f"   Linha {idx + offset}, Coluna B: {valor_realizado}")
                        if pd.notna(valor_realizado) and isinstance(valor_realizado, (int, float)):
                            print(f"   ✅ Possível valor realizado: {valor_realizado:,.2f}")
            
            elif 'HOSPEDAGENS' in linha_str:
                print(f"\n📍 HOSPEDAGENS encontrado na linha {idx}")
                for offset in range(1, 5):
                    if idx + offset < len(df):
                        valor_realizado = df.iloc[idx + offset, 1]
                        print(f"   Linha {idx + offset}, Coluna B: {valor_realizado}")
                        if pd.notna(valor_realizado) and isinstance(valor_realizado, (int, float)):
                            print(f"   ✅ Possível valor realizado: {valor_realizado:,.2f}")
            
            elif 'TRANSPORTE' in linha_str:
                print(f"\n📍 TRANSPORTE encontrado na linha {idx}")
                for offset in range(1, 5):
                    if idx + offset < len(df):
                        valor_realizado = df.iloc[idx + offset, 1]
                        print(f"   Linha {idx + offset}, Coluna B: {valor_realizado}")
                        if pd.notna(valor_realizado) and isinstance(valor_realizado, (int, float)):
                            print(f"   ✅ Possível valor realizado: {valor_realizado:,.2f}")
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    diagnosticar_resumo_logistica()