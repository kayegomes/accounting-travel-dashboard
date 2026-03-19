"""Diagnóstico: Valores da coluna Tipo_Solicitação"""

import pandas as pd
from pathlib import Path
import os

def diagnosticar_tipo_solicitacao():
    """Verifica os valores únicos da coluna Tipo_Solicitação"""
    
    if os.name == 'nt':
        onedrive_base = Path(r"C:\Users\ligomes\OneDrive - Globo Comunicação e Participações sa")
        caminho_painel = onedrive_base / "Gestão de Eventos - Documentos" / "Gestão de Eventos_planejamento" / "Painel Contábil" / "2025" / "Painel Contábil V2.xlsx"
    else:
        caminho_painel = Path('/home/ubuntu/upload/PainelContábilV2-copia.xlsx')
    
    print("="*80)
    print("🔍 DIAGNÓSTICO: Coluna Tipo_Solicitação")
    print("="*80)
    
    # PASSAGENS
    print("\n📊 PASSAGENS:")
    try:
        df_pass = pd.read_excel(caminho_painel, sheet_name='BasePassagens_New')
        print(f"   Total de registros: {len(df_pass)}")
        
        if 'Tipo_Solicitação' in df_pass.columns:
            valores_unicos = df_pass['Tipo_Solicitação'].value_counts()
            print(f"\n   Valores únicos na coluna 'Tipo_Solicitação':")
            for valor, qtd in valores_unicos.items():
                print(f"      '{valor}': {qtd} registros")
        else:
            print("   ❌ Coluna 'Tipo_Solicitação' NÃO encontrada!")
            print(f"   Colunas disponíveis: {list(df_pass.columns)[:20]}")
            
            # Procurar colunas similares
            colunas_tipo = [col for col in df_pass.columns if 'TIPO' in str(col).upper() or 'SOLICIT' in str(col).upper()]
            if colunas_tipo:
                print(f"\n   Colunas com 'TIPO' ou 'SOLICIT': {colunas_tipo}")
                for col in colunas_tipo:
                    print(f"\n   Valores de '{col}':")
                    valores = df_pass[col].value_counts()
                    for valor, qtd in valores.head(10).items():
                        print(f"      '{valor}': {qtd} registros")
    except Exception as e:
        print(f"   ❌ Erro: {e}")
    
    # HOSPEDAGENS
    print("\n" + "="*80)
    print("📊 HOSPEDAGENS:")
    try:
        df_hosp = pd.read_excel(caminho_painel, sheet_name='BaseHospedagens_New')
        print(f"   Total de registros: {len(df_hosp)}")
        
        if 'Tipo_Solicitação' in df_hosp.columns:
            valores_unicos = df_hosp['Tipo_Solicitação'].value_counts()
            print(f"\n   Valores únicos na coluna 'Tipo_Solicitação':")
            for valor, qtd in valores_unicos.items():
                print(f"      '{valor}': {qtd} registros")
        else:
            print("   ❌ Coluna 'Tipo_Solicitação' NÃO encontrada!")
            print(f"   Colunas disponíveis: {list(df_hosp.columns)[:20]}")
            
            # Procurar colunas similares
            colunas_tipo = [col for col in df_hosp.columns if 'TIPO' in str(col).upper() or 'SOLICIT' in str(col).upper()]
            if colunas_tipo:
                print(f"\n   Colunas com 'TIPO' ou 'SOLICIT': {colunas_tipo}")
                for col in colunas_tipo:
                    print(f"\n   Valores de '{col}':")
                    valores = df_hosp[col].value_counts()
                    for valor, qtd in valores.head(10).items():
                        print(f"      '{valor}': {qtd} registros")
    except Exception as e:
        print(f"   ❌ Erro: {e}")
    
    print("\n" + "="*80)


if __name__ == '__main__':
    diagnosticar_tipo_solicitacao()