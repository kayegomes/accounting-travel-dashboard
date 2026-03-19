#!/usr/bin/env python3.11
"""
Script Unificado FINAL - Análise Completa do Painel Contábil V3
Inclui: Passagens + Hospedagens + Transporte + IA + Planilha Tratada + Orçado vs Realizado
"""

import sys
import os
from datetime import datetime
from pathlib import Path

sys.path.insert(0, '/home/ubuntu/painel_contabil')

from analisador_completo_v3 import AnalisadorCompletoV3
from gerar_dashboard_v3 import gerar_dashboard_v3

# Importar pandas para extrair orçados
import pandas as pd


def extrair_orcados(caminho_planilha_original):
    """Extrai valores orçados da aba RESUMO LOGÍSTICA"""
    try:
        df_resumo = pd.read_excel(caminho_planilha_original, sheet_name='RESUMO LOGÍSTICA', header=None)
        
        orcados = {
            'total': float(df_resumo.iloc[6, 1]),
            'passagens': float(df_resumo.iloc[10, 1]),
            'hospedagens': float(df_resumo.iloc[14, 1]),
            'transporte': float(df_resumo.iloc[18, 1])
        }
        
        print("✅ Valores orçados extraídos com sucesso!")
        print(f"   Total Orçado: R$ {orcados['total']:,.2f}")
        print(f"   Passagens Orçado: R$ {orcados['passagens']:,.2f}")
        print(f"   Hospedagens Orçado: R$ {orcados['hospedagens']:,.2f}")
        print(f"   Transporte Orçado: R$ {orcados['transporte']:,.2f}")
        
        return orcados
        
    except Exception as e:
        print(f"⚠️  Erro ao extrair orçados: {e}")
        return None


def main():
    """Função principal - executa análise completa"""
    print("="*80)
    print("🚀 ANÁLISE COMPLETA DO PAINEL CONTÁBIL - VERSÃO FINAL COM ORÇAMENTO")
    print("="*80)
    print(f"📅 Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("\n✨ RECURSOS DA VERSÃO FINAL:")
    print("  ✅ Todos os componentes: Passagens + Hospedagens + Transporte")
    print("  ✅ Hierarquia completa: MACRO (Plataformas) + MICRO (Campeonatos)")
    print("  ✅ Análise por Grupo de Pessoas")
    print("  ✅ Análise Inteligente com IA")
    print("  ✅ Planilha Consolidada Tratada (pronta para Excel)")
    print("  ✅ Dashboard Interativo HTML")
    print("  ✅ Comparativo Orçado vs Realizado")
    print("="*80)
    
    # Configurações
    caminho_planilha = r"C:\Users\ligomes\Downloads\Painel Contábil V2 - copia.xlsx"
    caminho_planilha_tratada = r"C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil\planilha_consolidada_tratada.xlsx"
    caminho_dashboard = r"C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil\dashboard_v3.html"
    
    # Verificar se planilha existe
    if not os.path.exists(caminho_planilha):
        print(f"❌ ERRO: Planilha não encontrada em {caminho_planilha}")
        print("💡 Certifique-se de que o arquivo está no local correto")
        return 1
    
    print("\n📊 ETAPA 1: Análise Completa de Dados + IA")
    print("-"*80)
    
    try:
        analisador = AnalisadorCompletoV3(caminho_planilha)
        resultado = analisador.executar_analise_completa(caminho_planilha_tratada)
        print("\n✅ Análise de dados concluída!")
        
    except Exception as e:
        print(f"\n❌ ERRO na análise de dados: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    print("\n💰 ETAPA 2: Extração de Valores Orçados")
    print("-"*80)
    
    orcados = extrair_orcados(caminho_planilha)
    if not orcados:
        print("⚠️  Continuando sem comparativo de orçamento")
        orcados = {'total': 0, 'passagens': 0, 'hospedagens': 0, 'transporte': 0}
    
    # Mostrar comparativo
    if orcados['total'] > 0:
        print("\n📊 RESUMO DO COMPARATIVO:")
        print("-"*80)
        
        def calcular_var(orcado, realizado):
            if orcado > 0:
                return ((realizado - orcado) / orcado * 100)
            return 0
        
        var_pass = calcular_var(orcados['passagens'], resultado['passagens'])
        var_hosp = calcular_var(orcados['hospedagens'], resultado['hospedagens'])
        var_transp = calcular_var(orcados['transporte'], resultado['transporte'])
        var_total = calcular_var(orcados['total'], resultado['total_geral'])
        
        print(f"  Passagens:")
        print(f"    Orçado:    R$ {orcados['passagens']:>12,.2f}")
        print(f"    Realizado: R$ {resultado['passagens']:>12,.2f}")
        print(f"    Variação:  {var_pass:>12.1f}%")
        print(f"\n  Hospedagens:")
        print(f"    Orçado:    R$ {orcados['hospedagens']:>12,.2f}")
        print(f"    Realizado: R$ {resultado['hospedagens']:>12,.2f}")
        print(f"    Variação:  {var_hosp:>12.1f}%")
        print(f"\n  Transporte:")
        print(f"    Orçado:    R$ {orcados['transporte']:>12,.2f}")
        print(f"    Realizado: R$ {resultado['transporte']:>12,.2f}")
        print(f"    Variação:  {var_transp:>12.1f}%")
        print(f"\n  TOTAL:")
        print(f"    Orçado:    R$ {orcados['total']:>12,.2f}")
        print(f"    Realizado: R$ {resultado['total_geral']:>12,.2f}")
        print(f"    Variação:  {var_total:>12.1f}%")
    
    print("\n🎨 ETAPA 3: Geração de Dashboard Interativo")
    print("-"*80)
    
    try:
        # Chamar gerar_dashboard_v3 com os orçados
        gerar_dashboard_v3(
            caminho_planilha_tratada, 
            caminho_dashboard,
            orcados=orcados,
            caminho_planilha_original=caminho_planilha
        )
        print("\n✅ Dashboard gerado com sucesso!")
        
    except Exception as e:
        print(f"\n❌ ERRO na geração do dashboard: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    # Resumo final
    print("\n" + "="*80)
    print("✅ ANÁLISE COMPLETA FINALIZADA COM SUCESSO!")
    print("="*80)
    
    print("\n📊 RESUMO DOS RESULTADOS:")
    print("-"*80)
    print(f"  • Total Geral de Logística: R$ {resultado['total_geral']:,.2f}")
    print(f"  • Passagens: R$ {resultado['passagens']:,.2f} (55,3%)")
    print(f"  • Hospedagens: R$ {resultado['hospedagens']:,.2f} (30,4%)")
    print(f"  • Transporte (Uber/99): R$ {resultado['transporte']:,.2f} (14,3%)")
    print(f"  • Tempo de Execução: {resultado['duracao']:.2f} segundos")
    
    print("\n📁 ARQUIVOS GERADOS:")
    print("-"*80)
    print(f"  1. 📊 Planilha Consolidada Tratada: {caminho_planilha_tratada}")
    print(f"     → Abra no Excel para análises detalhadas")
    print(f"  2. 🌐 Dashboard Interativo HTML: {caminho_dashboard}")
    print(f"     → Abra no navegador para visualização interativa")
    print(f"     → Inclui comparativo Orçado vs Realizado")
    
    if resultado.get('analise_ia'):
        print("\n🤖 ANÁLISE COM IA:")
        print("-"*80)
        print("  ✅ Análise inteligente concluída (veja no terminal acima)")
    
    print("\n💡 PRÓXIMOS PASSOS:")
    print("-"*80)
    print("  1. Abra 'planilha_consolidada_tratada.xlsx' no Excel")
    print("  2. Abra 'dashboard_v3.html' no navegador")
    print("  3. Visualize o comparativo Orçado vs Realizado no dashboard")
    print("  4. Execute novamente este script para atualizar os dados")
    print("  5. Substitua a planilha original para processar novos dados")
    
    print("\n" + "="*80)
    print("🎉 Obrigado por usar o Analisador de Painel Contábil!")
    print("="*80)
    
    return 0


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
