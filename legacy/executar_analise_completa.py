#!/usr/bin/env python3.11
"""
Script Unificado - Análise Completa do Painel Contábil
Executa análise de dados e gera relatórios + dashboard em um único comando
"""

import sys
import os
from datetime import datetime

# Adicionar diretório ao path
sys.path.insert(0, '/home/ubuntu/painel_contabil')

from analisador_orcamento import AnalisadorOrcamento
from gerar_dashboard import GeradorDashboard


def main():
    """Função principal - executa análise completa"""
    print("="*80)
    print("🚀 ANÁLISE COMPLETA DO PAINEL CONTÁBIL")
    print("="*80)
    print(f"📅 Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("="*80)
    
    # Configurações
    caminho_planilha = '/home/ubuntu/upload/PainelContábilV2-copia.xlsx'
    caminho_relatorio = '/home/ubuntu/painel_contabil/relatorio_consolidado.xlsx'
    caminho_historico = '/home/ubuntu/painel_contabil/historico_execucoes.json'
    caminho_dashboard = '/home/ubuntu/painel_contabil/dashboard.html'
    
    # Verificar se planilha existe
    if not os.path.exists(caminho_planilha):
        print(f"❌ ERRO: Planilha não encontrada em {caminho_planilha}")
        print("💡 Certifique-se de que o arquivo está no local correto")
        return 1
    
    print("\n📊 ETAPA 1: Análise de Dados")
    print("-"*80)
    
    try:
        # Criar analisador e executar análise
        analisador = AnalisadorOrcamento(caminho_planilha)
        resumo = analisador.executar_analise_completa(caminho_relatorio, caminho_historico)
        
        print("\n✅ Análise de dados concluída!")
        
    except Exception as e:
        print(f"\n❌ ERRO na análise de dados: {e}")
        return 1
    
    print("\n🎨 ETAPA 2: Geração de Dashboard")
    print("-"*80)
    
    try:
        # Gerar dashboard
        gerador = GeradorDashboard(caminho_relatorio, caminho_historico)
        gerador.gerar(caminho_dashboard)
        
        print("\n✅ Dashboard gerado com sucesso!")
        
    except Exception as e:
        print(f"\n❌ ERRO na geração do dashboard: {e}")
        return 1
    
    # Resumo final
    print("\n" + "="*80)
    print("✅ ANÁLISE COMPLETA FINALIZADA COM SUCESSO!")
    print("="*80)
    
    print("\n📊 RESUMO DOS RESULTADOS:")
    print("-"*80)
    for chave, valor in resumo.items():
        if valor is not None:
            if isinstance(valor, float):
                if 'Percentual' in chave:
                    print(f"  • {chave}: {valor:.2f}%")
                elif 'Total' in chave and chave != 'Total_Viagens' and chave != 'Total_Elenco' and chave != 'Total_Centros_Custo':
                    print(f"  • {chave}: R$ {valor:,.2f}")
                else:
                    print(f"  • {chave}: {valor:,}")
            else:
                print(f"  • {chave}: {valor}")
    
    print("\n📁 ARQUIVOS GERADOS:")
    print("-"*80)
    print(f"  1. Relatório Excel: {caminho_relatorio}")
    print(f"  2. Dashboard HTML: {caminho_dashboard}")
    print(f"  3. Histórico JSON: {caminho_historico}")
    
    print("\n💡 PRÓXIMOS PASSOS:")
    print("-"*80)
    print("  • Abra o arquivo 'dashboard.html' no navegador para visualização interativa")
    print("  • Consulte o arquivo 'relatorio_consolidado.xlsx' para análises detalhadas")
    print("  • Execute novamente este script para atualizar os dados")
    
    print("\n" + "="*80)
    
    return 0


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
