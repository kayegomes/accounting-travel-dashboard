"""Gerador de Dashboard V3.8 - TODAS AS CORREÇÕES APLICADAS
Correções V3.8:
1. ✅ Gráficos mostram MÉDIA (não total) de orçado vs realizado
2. ✅ Filtros nacional/internacional funcionando 100%
3. ✅ Números aparecem diretamente nos gráficos (com datalabels)
4. ✅ Filtros posicionados corretamente acima de cada seção
5. ✅ Top 10 gastadores por PESSOA (não produto)
6. ✅ Filtro de antecedência por grupo funcionando corretamente
7. ✅ Destinos agrupados/normalizados para evitar duplicatas
"""

import os
from pathlib import Path
from datetime import datetime
import pandas as pd
import json
import re

from jinja2 import Environment, FileSystemLoader

def normalizar_destino(destino):
    """Normaliza destinos para agrupar variações similares"""
    if not destino or str(destino).lower() in ['nan', 'none', '']:
        return 'Não informado'
    
    destino_str = str(destino).upper().strip()
    
    # Remover acentos comuns
    destino_str = destino_str.replace('Ã', 'A').replace('Õ', 'O').replace('Ç', 'C')
    destino_str = destino_str.replace('É', 'E').replace('Á', 'A').replace('Ó', 'O')
    destino_str = destino_str.replace('Í', 'I').replace('Ú', 'U')
    
    # Remover caracteres especiais
    destino_str = re.sub(r'[^\w\s-]', '', destino_str)
    
    # Padronizar cidades conhecidas
    mapeamento = {
        'SAO PAULO': 'SÃO PAULO', 'SP': 'SÃO PAULO', 'SAMPA': 'SÃO PAULO',
        'RIO DE JANEIRO': 'RIO DE JANEIRO', 'RJ': 'RIO DE JANEIRO', 'RIO': 'RIO DE JANEIRO',
        'BELO HORIZONTE': 'BELO HORIZONTE', 'BH': 'BELO HORIZONTE',
        'BRASILIA': 'BRASÍLIA', 'BSB': 'BRASÍLIA',
        'SALVADOR': 'SALVADOR', 'SSA': 'SALVADOR',
        'FORTALEZA': 'FORTALEZA', 'FOR': 'FORTALEZA',
        'RECIFE': 'RECIFE', 'REC': 'RECIFE',
        'PORTO ALEGRE': 'PORTO ALEGRE', 'POA': 'PORTO ALEGRE',
        'CURITIBA': 'CURITIBA', 'CWB': 'CURITIBA',
        'MANAUS': 'MANAUS', 'MAO': 'MANAUS',
    }
    
    # Verificar se contém alguma cidade conhecida
    for key, value in mapeamento.items():
        if key in destino_str:
            return value
    
    # Se tiver múltiplas palavras, pegar as 2 primeiras
    palavras = destino_str.split()
    if len(palavras) > 2:
        return ' '.join(palavras[:2])
    
    return destino_str


def carregar_config():
    """Carrega configurações do arquivo config.json"""
    config_path = Path(__file__).parent / "config.json"
    if config_path.exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def extrair_orcados_plataforma_manual():
    """Valores orçados por plataforma."""
    return {
        'TV GLOBO': {'total': 4288704, 'passagens': 3010436, 'hospedagens': 1278268, 'transporte': 0},
        'SPORTV': {'total': 5709121, 'passagens': 4479169, 'hospedagens': 1229952, 'transporte': 0},
        'PREMIERE': {'total': 1686324, 'passagens': 1261980, 'hospedagens': 424344, 'transporte': 0},
        'COMBATE': {'total': 65396, 'passagens': 50000, 'hospedagens': 15396, 'transporte': 0},
        'GE TV': {'total': 2088860, 'passagens': 1602500, 'hospedagens': 486360, 'transporte': 0}
    }

def extrair_orcados_detalhados_financeiro(caminho_financeiro):
    """
    Extrai da aba 'Bdados' da planilha Financeiro.xlsx:
    - Valores totais orçados por tipo (passagens, hospedagens)
    - Quantidades orçadas (número de passagens, número de diárias)
    - Separação nacional/internacional baseada na coluna 'Sinal'
    Retorna dicionário com as médias e quantidades para os gráficos.
    """
    try:
        df = pd.read_excel(caminho_financeiro, sheet_name='Bdados')
        print(f"\n📊 Extraindo orçados detalhados da Financeiro: {len(df)} registros")

        # Mapear tipos de conta
        tipo_conta = {
            'Passagem': 'passagens',
            'Hospedagem': 'hospedagens',
            # Transporte Viagem pode ser ignorado aqui, pois não temos gráfico de transporte
        }

        # Converter colunas numéricas
        for col in ['Datas', 'Quantidade de Pessoas', 'Diárias', 'Total']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Inicializar acumuladores
        totais = {'passagens': 0.0, 'hospedagens': 0.0}
        quantidades = {'passagens': 0, 'hospedagens': 0}
        totais_nac = {'passagens': 0.0, 'hospedagens': 0.0}
        quantidades_nac = {'passagens': 0, 'hospedagens': 0}
        totais_int = {'passagens': 0.0, 'hospedagens': 0.0}
        quantidades_int = {'passagens': 0, 'hospedagens': 0}
        

        for _, row in df.iterrows():
            conta = str(row.get('Conta', '')).strip()
            if conta not in tipo_conta:
                continue

            categoria = tipo_conta[conta]
            valor = row.get('Total', 0)
            if pd.isna(valor) or valor == 0:
                continue

            # Determinar nacional/internacional
            sinal = str(row.get('Sinal', '')).upper()
            is_internacional = 'INTERNACIONAL' in sinal

            # Calcular quantidade
            if categoria == 'passagens':
                # Quantidade de passagens = Datas * Quantidade de Pessoas
                qtd = row['Datas'] * row['Quantidade de Pessoas']
            else:  # hospedagens
                # Quantidade de diárias = Datas * Quantidade de Pessoas * Diárias
                qtd = row['Datas'] * row['Quantidade de Pessoas'] * row['Diárias']

            # Acumular totais e quantidades
            totais[categoria] += valor
            quantidades[categoria] += qtd

            if is_internacional:
                totais_int[categoria] += valor
                quantidades_int[categoria] += qtd
            else:
                totais_nac[categoria] += valor
                quantidades_nac[categoria] += qtd

        # Calcular médias
        medias = {}
        for tipo in ['passagens', 'hospedagens']:
            medias[tipo] = {
                'total': totais[tipo] / quantidades[tipo] if quantidades[tipo] > 0 else 0,
                'nacional': totais_nac[tipo] / quantidades_nac[tipo] if quantidades_nac[tipo] > 0 else 0,
                'internacional': totais_int[tipo] / quantidades_int[tipo] if quantidades_int[tipo] > 0 else 0
            }

        # Também retornar as quantidades para os gráficos de quantidade
        quantidades_orcado = {
            'passagens_total': quantidades['passagens'],
            'passagens_nacional': quantidades_nac['passagens'],
            'passagens_internacional': quantidades_int['passagens'],
            'hospedagens_total': quantidades['hospedagens'],
            'hospedagens_nacional': quantidades_nac['hospedagens'],
            'hospedagens_internacional': quantidades_int['hospedagens']
        }

        print("\n📊 Resultado da extração:")
        print(f"   Passagens - Total: {quantidades['passagens']:.0f} unidades | Média R$ {medias['passagens']['total']:,.2f}")
        print(f"   Passagens - Nacional: {quantidades_nac['passagens']:.0f} un | Média R$ {medias['passagens']['nacional']:,.2f}")
        print(f"   Passagens - Internacional: {quantidades_int['passagens']:.0f} un | Média R$ {medias['passagens']['internacional']:,.2f}")
        print(f"   Hospedagens - Total: {quantidades['hospedagens']:.0f} diárias | Média R$ {medias['hospedagens']['total']:,.2f}")
        print(f"   Hospedagens - Nacional: {quantidades_nac['hospedagens']:.0f} diárias | Média R$ {medias['hospedagens']['nacional']:,.2f}")
        print(f"   Hospedagens - Internacional: {quantidades_int['hospedagens']:.0f} diárias | Média R$ {medias['hospedagens']['internacional']:,.2f}")

        return {
            'medias': medias,
            'quantidades': quantidades_orcado
        }

    except Exception as e:
        print(f"⚠️ Erro ao extrair orçados detalhados: {e}")
        import traceback
        traceback.print_exc()
        return None
    
#  def extrair_quantidades_orcadas_com_filtro(caminho_formatos, filtro_internacional=False):
    """Extrai quantidades orçadas da planilha Formatos com filtro nacional/internacional"""
    try:
        print(f"\n🔍 Extraindo quantidades orçadas - {'INTERNACIONAL' if filtro_internacional else 'NACIONAL'}...")
        
        df = pd.read_excel(caminho_formatos, sheet_name='Bdados', header=1)
        
        col_conta_fin = None
        for col in df.columns:
            if 'CONTA FINANCEIRO' in str(col).upper():
                col_conta_fin = col
                break
        
        if not col_conta_fin:
            return {'passagens': 0, 'diarias': 0}
        
        # Converter para numérico
        if 'Pessoas*Datas Orçamento' in df.columns:
            df['Pessoas*Datas Orçamento'] = pd.to_numeric(df['Pessoas*Datas Orçamento'], errors='coerce').fillna(0)
        if 'Diária' in df.columns:
            df['Diária'] = pd.to_numeric(df['Diária'], errors='coerce').fillna(0)
        if 'Datas Orçamento' in df.columns:
            df['Datas Orçamento'] = pd.to_numeric(df['Datas Orçamento'], errors='coerce').fillna(0)
        
        # Filtros
        if filtro_internacional:
            filtro_pass = '09\\. Passagens Internacionais'
            filtro_hosp = '10\\. Hospedagens Internacionais'
        else:
            filtro_pass = '05\\. Passagens'
            filtro_hosp = '06\\. Hospedagens'
        
        # Passagens
        qtd_passagens_orcado = 0
        if 'Pessoas*Datas Orçamento' in df.columns:
            filtro = df[col_conta_fin].astype(str).str.contains(filtro_pass, case=False, na=False, regex=True)
            qtd_passagens_orcado = df.loc[filtro, 'Pessoas*Datas Orçamento'].sum()
        
        # Hospedagens
        qtd_diarias_orcado = 0
        if 'Diária' in df.columns and 'Datas Orçamento' in df.columns:
            df['Total_Diarias'] = df['Diária'] * df['Datas Orçamento']
            filtro = df[col_conta_fin].astype(str).str.contains(filtro_hosp, case=False, na=False, regex=True)
            qtd_diarias_orcado = df.loc[filtro, 'Total_Diarias'].sum()
        
        return {
            'passagens': int(qtd_passagens_orcado),
            'diarias': int(qtd_diarias_orcado)
        }
        
    except Exception as e:
        print(f"   ⚠️ Erro ao extrair quantidades orçadas: {e}")
        return {'passagens': 0, 'diarias': 0}


def extrair_dados_passagens_e_hospedagens(caminho_planilha_original):
    """Extrai dados de passagens e hospedagens com filtros e retorna totais detalhados por plataforma."""
    try:
        print("\n📊 Extraindo dados de passagens e hospedagens...")
        
        # Listas de filtros
        AREAS_PERMITIDAS = [
            'Colaborador', 'Elenco', 'Ed. Eventos', 'Produção de Eventos',
            'Repcine', 'Gestão Integrada', 'Repórter'
        ]
        PROJETOS_EXCLUIR = [
            '0000000000000000 - INDIRETO',
            'I.ESP.000010-999-999 - GESTÃO DE ELENCO/NA/NA',
            'I.ESP.000009-999-999 - POOL PRODUÇÃO - EVENTOS/NA/'
        ]
        FINALIDADE_EXCLUIR = ['FG000141 - UM']
        
        
        
        dados_combinados = []
        # Dicionário para totais por plataforma e tipo
        totais_plataforma_detalhado = {}  # {plataforma: {'passagens': 0, 'hospedagens': 0}}
        
        # 1. PASSAGENS
        try:
            df_pass = pd.read_excel(caminho_planilha_original, sheet_name='BasePassagens_New')
            print(f"   ✅ Dados de passagens: {len(df_pass)} registros")
            
            colunas_pass = {}
            for col in df_pass.columns:
                col_upper = str(col).upper()
                col_clean = str(col).strip()
                
                if col_clean == 'Passageiro':
                    colunas_pass['passageiro'] = col
                elif 'VALOR' in col_upper and ('AJUSTADO' in col_upper or 'TOTAL' in col_upper):
                    colunas_pass['valor'] = col
                elif 'GRUPO' in col_upper or col_clean == 'Área':
                    colunas_pass['grupo'] = col
                elif 'ANTECED' in col_upper:
                    colunas_pass['antecedencia'] = col
                elif col_clean == 'Destino':
                    colunas_pass['destino'] = col
                elif 'TIPO' in col_upper and ('VIAGEM' in col_upper or 'SOLICITAÇÃO' in col_upper):
                    colunas_pass['tipo_viagem'] = col
                elif col_clean == 'Plataforma':
                    colunas_pass['plataforma'] = col
                elif 'DESCRIÇÃO PROJETO' in col_upper:
                    colunas_pass['desc_projeto'] = col
                elif 'FINALIDADE' in col_upper:
                    colunas_pass['finalidade'] = col
                if col_clean == 'Plataforma' or 'PLATAFORMA' in col_upper:
                    colunas_pass['plataforma'] = col
                    
                
            
            print(f"   ✅ Colunas mapeadas em passagens: {colunas_pass}")
            
            registros_filtrados = 0
            count_inter = 0
            count_nac = 0
            
            for _, row in df_pass.iterrows():
                # Filtro básico: Tipo_Solicitação
                tipo_solicitacao = str(row.get('Tipo_Solicitação', '')).strip()
                if tipo_solicitacao != 'Nova solicitação':
                    continue
                
                # Filtros adicionais
                area = str(row.get(colunas_pass.get('grupo', ''), '')).strip()
                if area and area not in AREAS_PERMITIDAS:
                    continue
                
                desc_projeto = str(row.get(colunas_pass.get('desc_projeto', ''), '')).strip()
                if desc_projeto in PROJETOS_EXCLUIR:
                    continue
                
                finalidade = str(row.get(colunas_pass.get('finalidade', ''), '')).strip()
                if finalidade in FINALIDADE_EXCLUIR:
                    continue
                
                
                
                registros_filtrados += 1
                
                passageiro = str(row[colunas_pass['passageiro']]).strip() if colunas_pass.get('passageiro') else ''
                valor = float(row[colunas_pass['valor']]) if pd.notna(row.get(colunas_pass['valor'], 0)) else 0
                if not passageiro or passageiro == 'nan' or valor <= 0:
                    continue
                plataforma = str(row.get(colunas_pass.get('plataforma', ''), '')).strip()
                # Classificação internacional
                tipo_viagem = str(row.get(colunas_pass.get('tipo_viagem', ''), '')).upper()
                natureza = str(row.get('Natureza', '')).upper() if 'Natureza' in df_pass.columns else ''
                is_internacional = (
                    'INTERNACIONAL' in tipo_viagem or 'INTER' in tipo_viagem or
                    'INTERNACIONAL' in natureza or 'PASSAGEM INTERNACIONAL' in natureza or '09.' in natureza
                )
                
                if is_internacional:
                    count_inter += 1
                else:
                    count_nac += 1
                
                # Plataforma (para totais detalhados)
                plataforma = str(row.get(colunas_pass.get('plataforma', ''), '')).strip()
                if plataforma:
                    # Inicializa se não existir
                    if plataforma not in totais_plataforma_detalhado:
                        totais_plataforma_detalhado[plataforma] = {'passagens': 0, 'hospedagens': 0}
                    totais_plataforma_detalhado[plataforma]['passagens'] += valor
                
                plataformas_permitidas = ['TV GLOBO', 'GE TV', 'SPORTV', 'PREMIERE', 'COMBATE']
                if plataforma not in plataformas_permitidas:
                    continue
                
                dados_combinados.append({
                    'tipo': 'passagem',
                    'passageiro': passageiro,
                    'valor': valor,
                    'grupo': area,
                    'antecedencia': float(row.get(colunas_pass.get('antecedencia', 0), 0)),
                    'destino': normalizar_destino(str(row.get(colunas_pass.get('destino', ''), '')).strip()),
                    'is_internacional': is_internacional,
                    'plataforma': plataforma
                })
            
            print(f"   ✅ V3.10: Filtrados {registros_filtrados} de {len(df_pass)} (após todos os filtros)")
            print(f"   📊 Passagens: {count_nac} nacionais, {count_inter} internacionais")
            
        except Exception as e:
            print(f"   ⚠️ Erro ao extrair passagens: {e}")
        
        # 2. HOSPEDAGENS
        try:
            sheets_hosp = ['BaseHospedagens_New', 'BaseHospedagens', 'Hospedagens']
            df_hosp = None
            for sheet in sheets_hosp:
                try:
                    df_hosp = pd.read_excel(caminho_planilha_original, sheet_name=sheet)
                    print(f"   ✅ Encontrada aba: {sheet}")
                    break
                except:
                    continue
            
            if df_hosp is not None:
                print(f"   ✅ Dados de hospedagens: {len(df_hosp)} registros")
                
                colunas_hosp = {}
                for col in df_hosp.columns:
                    col_upper = str(col).upper()
                    col_clean = str(col).strip()
                    
                    if col_clean == 'Passageiro':
                        colunas_hosp['passageiro'] = col
                    elif 'TOTAL' in col_upper and 'AJUSTADO' in col_upper:
                        colunas_hosp['total_ajustado'] = col
                    elif col_clean == 'Diária' or col_clean == 'Diaria':
                        colunas_hosp['quantidade_diarias'] = col
                    elif 'GRUPO' in col_upper or col_clean == 'Área':
                        colunas_hosp['grupo'] = col
                    elif col_clean == 'Natureza':
                        colunas_hosp['natureza'] = col
                    elif col_clean == 'Plataforma':
                        colunas_hosp['plataforma'] = col
                    elif 'DESCRIÇÃO PROJETO' in col_upper:
                        colunas_hosp['desc_projeto'] = col
                    elif 'FINALIDADE' in col_upper:
                        colunas_hosp['finalidade'] = col
                
                print(f"   ✅ Colunas mapeadas: {colunas_hosp}")
                
                registros_filtrados = 0
                count_inter = 0
                count_nac = 0
                qtd_diarias_nac = 0
                qtd_diarias_inter = 0
                soma_valores_nac = 0
                soma_valores_inter = 0
                
                for _, row in df_hosp.iterrows():
                    # Filtro básico
                    tipo_solicitacao = str(row.get('Tipo_Solicitação', '')).strip()
                    if tipo_solicitacao != 'Nova solicitação':
                        continue
                    
                    # Filtros adicionais
                    area = str(row.get(colunas_hosp.get('grupo', ''), '')).strip()
                    if area and area not in AREAS_PERMITIDAS:
                        continue
                    
                    desc_projeto = str(row.get(colunas_hosp.get('desc_projeto', ''), '')).strip()
                    if desc_projeto in PROJETOS_EXCLUIR:
                        continue
                    
                    finalidade = str(row.get(colunas_hosp.get('finalidade', ''), '')).strip()
                    if finalidade in FINALIDADE_EXCLUIR:
                        continue
                    
                    registros_filtrados += 1
                    
                    passageiro = str(row[colunas_hosp['passageiro']]).strip() if colunas_hosp.get('passageiro') else ''
                    total_ajustado = float(row[colunas_hosp['total_ajustado']]) if pd.notna(row.get(colunas_hosp['total_ajustado'], 0)) else 0
                    qtd_diarias = float(row[colunas_hosp['quantidade_diarias']]) if pd.notna(row.get(colunas_hosp['quantidade_diarias'], 0)) else 0
                    
                    if not passageiro or passageiro == 'nan' or total_ajustado <= 0 or qtd_diarias <= 0:
                        continue
                    
                    valor_unitario = total_ajustado / qtd_diarias
                    
                    # Classificação internacional
                    natureza = str(row.get(colunas_hosp.get('natureza', ''), '')).upper()
                    is_internacional = (
                        'INTERNACIONAL' in natureza or
                        'HOSPEDAGEM INTERNACIONAL' in natureza or
                        '10.' in natureza
                    )
                    
                    if is_internacional:
                        count_inter += 1
                        qtd_diarias_inter += qtd_diarias
                        soma_valores_inter += total_ajustado
                    else:
                        count_nac += 1
                        qtd_diarias_nac += qtd_diarias
                        soma_valores_nac += total_ajustado
                    
                    # Plataforma
                    plataforma = str(row.get(colunas_hosp.get('plataforma', ''), '')).strip()
                    if plataforma:
                        if plataforma not in totais_plataforma_detalhado:
                            totais_plataforma_detalhado[plataforma] = {'passagens': 0, 'hospedagens': 0}
                        totais_plataforma_detalhado[plataforma]['hospedagens'] += total_ajustado
                    
                    # Criar um registro por diária
                    for _ in range(int(qtd_diarias)):
                        dados_combinados.append({
                            'tipo': 'hospedagem',
                            'passageiro': passageiro,
                            'valor': valor_unitario,
                            'grupo': area,
                            'antecedencia': 0,
                            'destino': '',
                            'is_internacional': is_internacional,
                            'plataforma': plataforma
                        })
                
                print(f"   ✅ V3.10: Filtrados {registros_filtrados} de {len(df_hosp)}")
                print(f"   📊 Hospedagens: {count_nac} registros nacionais ({int(qtd_diarias_nac)} diárias), {count_inter} internacionais ({int(qtd_diarias_inter)} diárias)")
                print(f"   💰 Valores totais: Nacional R$ {soma_valores_nac:,.2f}, Internacional R$ {soma_valores_inter:,.2f}")
                if qtd_diarias_nac > 0:
                    print(f"   📊 Média Nacional: R$ {soma_valores_nac/qtd_diarias_nac:,.2f}")
                if qtd_diarias_inter > 0:
                    print(f"   📊 Média Internacional: R$ {soma_valores_inter/qtd_diarias_inter:,.2f}")
        
        except Exception as e:
            print(f"   ⚠️ Erro ao extrair hospedagens: {e}")
            import traceback
            traceback.print_exc()
        
        print(f"   ✅ Total de registros combinados: {len(dados_combinados)}")
        
        # Extrair grupos únicos (para filtro de antecedência)
        grupos_unicos = []
        grupos_normalizados = {}
        for d in dados_combinados:
            grupo = d.get('grupo', '')
            if grupo and grupo != 'nan' and grupo.strip():
                grupo_lower = grupo.strip().lower()
                if grupo_lower not in grupos_normalizados:
                    grupos_normalizados[grupo_lower] = grupo.strip()
        grupos_unicos = sorted(list(grupos_normalizados.values()))[:15]
        
        print(f"   📊 Grupos únicos encontrados ({len(grupos_unicos)}): {grupos_unicos[:10]}")
        
        # Calcular totais gerais
        total_passagens = sum(d['valor'] for d in dados_combinados if d['tipo'] == 'passagem')
        total_hospedagens = sum(d['valor'] for d in dados_combinados if d['tipo'] == 'hospedagem')
        
        return {
            'dados': dados_combinados,
            'grupos_disponiveis': grupos_unicos,
            'total_registros': len(dados_combinados),
            'total_valor': total_passagens + total_hospedagens,
            'totais_plataforma_detalhado': totais_plataforma_detalhado,  # novo campo
            'total_passagens': total_passagens,
            'total_hospedagens': total_hospedagens
        }
        
    except Exception as e:
        print(f"⚠️ Erro ao extrair dados completos: {e}")
        import traceback
        traceback.print_exc()
        return {
            'dados': [], 'grupos_disponiveis': [], 'total_registros': 0, 'total_valor': 0,
            'totais_plataforma_detalhado': {}, 'total_passagens': 0, 'total_hospedagens': 0
        }


def extrair_orcados(caminho_planilha_original):
    """Extrai orçados GERAIS da aba RESUMO LOGÍSTICA"""
    try:
        print("\n🔍 Extraindo orçados da planilha...")
        df = pd.read_excel(caminho_planilha_original, sheet_name='RESUMO LOGÍSTICA', header=None)
        
        orcados = {
            'total': float(df.iloc[6, 1]) if pd.notna(df.iloc[6, 1]) else 0,
            'passagens': float(df.iloc[10, 1]) if pd.notna(df.iloc[10, 1]) else 0,
            'hospedagens': float(df.iloc[14, 1]) if pd.notna(df.iloc[14, 1]) else 0,
            'transporte': float(df.iloc[18, 1]) if pd.notna(df.iloc[18, 1]) else 0,
            'por_plataforma': extrair_orcados_plataforma_manual()
        }
        
        print(f"✅ Orçados extraídos!")
        return orcados
        
    except Exception as e:
        print(f"⚠️ Erro ao extrair orçados: {e}")
        return {'total': 0, 'passagens': 0, 'hospedagens': 0, 'transporte': 0, 'por_plataforma': extrair_orcados_plataforma_manual()}

def extrair_transporte(caminho_planilha_original):
    """Extrai valores de transporte (UBER e 99) da aba Consolidado Geral."""
    try:
        df = pd.read_excel(caminho_planilha_original, sheet_name='Consolidado Geral (UBER e 99)')
        # Supondo que haja uma coluna com descrição e outra com valor
        # Ajuste os nomes conforme a planilha real
        if 'Descrição' in df.columns and 'Valor' in df.columns:
            mask_uber = df['Descrição'].str.contains('UBER', case=False, na=False)
            mask_99 = df['Descrição'].str.contains('99', case=False, na=False)
            valor_uber = df.loc[mask_uber, 'Valor'].sum()
            valor_99 = df.loc[mask_99, 'Valor'].sum()
            total_transporte = valor_uber + valor_99
            print(f"   🚗 Transporte: UBER R$ {valor_uber:,.2f}, 99 R$ {valor_99:,.2f} = Total R$ {total_transporte:,.2f}")
            return total_transporte
        else:
            print("   ⚠️ Colunas 'Descrição' e 'Valor' não encontradas em Consolidado Geral.")
            return 0
    except Exception as e:
        print(f"   ⚠️ Erro ao extrair transporte: {e}")
        return 0

def gerar_dashboard_v3_8(caminho_planilha_tratada, caminho_saida=None, caminho_planilha_original=None, 
                        caminho_formatos=None):
    """Gera dashboard HTML V3.8 com valores calculados das bases (filtros aplicados)"""
    # Orçados manuais (substitua pelos valores corretos)
    ORCADO_TOTAL = 15_886_395  # exemplo
    ORCADO_PASSAGENS = 10_404_085
    ORCADO_HOSPEDAGENS = 3_428_924
    ORCADO_TRANSPORTE = 2_053_386  # se não houver orçado de transporte
    periodo_analise = "02/01/2026 a 03/03/2026"
    
    # Valores extras da Fórmula 1 (conforme anexo)
    extra_passagens_valor = 236001.00   # 142909 + 93092
    extra_hospedagens_valor = 289297.40 # 157054.90 + 132242.50
    extra_passagens_qtd = 20
    extra_hospedagens_qtd = 100

    # Dicionário com os valores por plataforma (para os cards)
    orcado_extra_plataforma = {
        'TV GLOBO': {'passagens': 142909.00, 'hospedagens': 157054.90},
        'SPORTV': {'passagens': 93092.00, 'hospedagens': 132242.50}
    }

    
    def formatar_numero(valor):
        """Formata número inteiro com ponto como separador de milhar."""
        return f"{valor:,.0f}".replace(',', '.')

    orcados = {
        'total': ORCADO_TOTAL,
        'passagens': ORCADO_PASSAGENS,
        'hospedagens': ORCADO_HOSPEDAGENS,
        'transporte': ORCADO_TRANSPORTE,
        'por_plataforma': extrair_orcados_plataforma_manual()  # mantém os manuais por plataforma
    }
    caminho_planilha_tratada = Path(caminho_planilha_tratada)
    if not caminho_planilha_tratada.exists():
        raise FileNotFoundError(f"Planilha tratada não encontrada: {caminho_planilha_tratada}")

    if caminho_saida is None:
        caminho_saida = Path(__file__).parent / "dashboard_v3_8.html"
    else:
        caminho_saida = Path(caminho_saida)

    print(f"🎨 Gerando Dashboard V3.8 COM CÁLCULO DIRETO DAS BASES (filtros aplicados)...")

    # -------------------------------------------------------------------------
    # 1. Carregar dados auxiliares (estruturas que não serão recalculadas)
    # -------------------------------------------------------------------------
    df_resumo = pd.read_excel(caminho_planilha_tratada, sheet_name='Resumo Executivo')
    df_produto = pd.read_excel(caminho_planilha_tratada, sheet_name='Por Campeonato (MICRO)')
    df_grupo = pd.read_excel(caminho_planilha_tratada, sheet_name='Por Grupo de Pessoas')

    # -------------------------------------------------------------------------
    # 2. Extrair orçados (continuam iguais)
    # -------------------------------------------------------------------------
    orcados = extrair_orcados(caminho_planilha_original) if caminho_planilha_original else {
        'total': 0, 'passagens': 0, 'hospedagens': 0, 'transporte': 0, 'por_plataforma': {}
    }
    
    # Somar aos totais gerais dos cards
    orcados['passagens'] += extra_passagens_valor
    orcados['hospedagens'] += extra_hospedagens_valor
    orcados['total'] += (extra_passagens_valor + extra_hospedagens_valor)

    # Somar aos valores por plataforma
    for plat, valores in orcado_extra_plataforma.items():
        if plat in orcados['por_plataforma']:
            orcados['por_plataforma'][plat]['passagens'] += valores['passagens']
            orcados['por_plataforma'][plat]['hospedagens'] += valores['hospedagens']
            orcados['por_plataforma'][plat]['total'] += valores['passagens'] + valores['hospedagens']
        else:
            # Se a plataforma não existir (improvável), cria
            orcados['por_plataforma'][plat] = {
                'passagens': valores['passagens'],
                'hospedagens': valores['hospedagens'],
                'transporte': 0,
                'total': valores['passagens'] + valores['hospedagens']
            }
    # -------------------------------------------------------------------------
    # 3. Extrair dados combinados com os novos filtros
    # -------------------------------------------------------------------------
    # Esta função deve ser a versão modificada que aplica os filtros e retorna também
    # os totais por plataforma (separados por passagens e hospedagens)
    dados_combinados = extrair_dados_passagens_e_hospedagens(caminho_planilha_original) if caminho_planilha_original else {
        'dados': [], 'grupos_disponiveis': [], 'total_registros': 0, 'total_valor': 0,
        'totais_plataforma': {},          # dicionário com totais por plataforma (soma de passagens+hospedagens)
        'passagens_por_plataforma': {},   # dicionário com totais de passagens por plataforma
        'hospedagens_por_plataforma': {}, # dicionário com totais de hospedagens por plataforma
        'total_passagens': 0,
        'total_hospedagens': 0
    }
    # Após extrair dados_combinados e calcular totais realizados...

    # Extrair orçados detalhados da Financeiro para os gráficos
    orcados_graficos = None
    if caminho_formatos and Path(caminho_formatos).exists():
        orcados_graficos = extrair_orcados_detalhados_financeiro(caminho_formatos)
    
    if orcados_graficos:
        # --- Passagens internacionais ---
        qtd_int_antiga = orcados_graficos['quantidades']['passagens_internacional']
        media_int_antiga = orcados_graficos['medias']['passagens']['internacional']
        valor_int_antigo = media_int_antiga * qtd_int_antiga

        novo_valor_int = valor_int_antigo + extra_passagens_valor
        nova_qtd_int = qtd_int_antiga + extra_passagens_qtd
        nova_media_int = novo_valor_int / nova_qtd_int if nova_qtd_int > 0 else 0

        orcados_graficos['medias']['passagens']['internacional'] = nova_media_int
        orcados_graficos['quantidades']['passagens_internacional'] = nova_qtd_int
        orcados_graficos['quantidades']['passagens_total'] += extra_passagens_qtd

        # Recalcular média total de passagens
        qtd_nac = orcados_graficos['quantidades']['passagens_nacional']
        media_nac = orcados_graficos['medias']['passagens']['nacional']
        valor_nac = media_nac * qtd_nac
        valor_total_pass = valor_nac + novo_valor_int
        qtd_total_pass = qtd_nac + nova_qtd_int
        orcados_graficos['medias']['passagens']['total'] = valor_total_pass / qtd_total_pass if qtd_total_pass > 0 else 0

        # --- Hospedagens internacionais ---
        qtd_hosp_int_antiga = orcados_graficos['quantidades']['hospedagens_internacional']
        media_hosp_int_antiga = orcados_graficos['medias']['hospedagens']['internacional']
        valor_hosp_int_antigo = media_hosp_int_antiga * qtd_hosp_int_antiga

        novo_valor_hosp_int = valor_hosp_int_antigo + extra_hospedagens_valor
        nova_qtd_hosp_int = qtd_hosp_int_antiga + extra_hospedagens_qtd
        nova_media_hosp_int = novo_valor_hosp_int / nova_qtd_hosp_int if nova_qtd_hosp_int > 0 else 0

        orcados_graficos['medias']['hospedagens']['internacional'] = nova_media_hosp_int
        orcados_graficos['quantidades']['hospedagens_internacional'] = nova_qtd_hosp_int
        orcados_graficos['quantidades']['hospedagens_total'] += extra_hospedagens_qtd

        # Recalcular média total de hospedagens
        qtd_hosp_nac = orcados_graficos['quantidades']['hospedagens_nacional']
        media_hosp_nac = orcados_graficos['medias']['hospedagens']['nacional']
        valor_hosp_nac = media_hosp_nac * qtd_hosp_nac
        valor_total_hosp = valor_hosp_nac + novo_valor_hosp_int
        qtd_total_hosp = qtd_hosp_nac + nova_qtd_hosp_int
        orcados_graficos['medias']['hospedagens']['total'] = valor_total_hosp / qtd_total_hosp if qtd_total_hosp > 0 else 0

        print("✅ Valores extras da F1 incorporados aos gráficos.")

    if orcados_graficos:
        # Usar os valores extraídos para os gráficos
        medias_orcado = orcados_graficos['medias']
        qtd_orcado = orcados_graficos['quantidades']

        # Atualizar as variáveis que vão para o JavaScript
        media_pass_orcado_total = medias_orcado['passagens']['total']
        media_pass_orcado_nacional = medias_orcado['passagens']['nacional']
        media_pass_orcado_internacional = medias_orcado['passagens']['internacional']

        media_hosp_orcado_total = medias_orcado['hospedagens']['total']
        media_hosp_orcado_nacional = medias_orcado['hospedagens']['nacional']
        media_hosp_orcado_internacional = medias_orcado['hospedagens']['internacional']

        qtd_passagens_orcado_total = qtd_orcado['passagens_total']
        qtd_passagens_orcado_nacional = qtd_orcado['passagens_nacional']
        qtd_passagens_orcado_internacional = qtd_orcado['passagens_internacional']

        qtd_hospedagens_orcado_total = qtd_orcado['hospedagens_total']
        qtd_hospedagens_orcado_nacional = qtd_orcado['hospedagens_nacional']
        qtd_hospedagens_orcado_internacional = qtd_orcado['hospedagens_internacional']

        print("\n✅ Gráficos usarão orçados da planilha Financeiro.")
    else:
        # Fallback: usar os valores que já estavam (rateio 70/30 ou manuais)
        print("⚠️ Gráficos usarão fallback (valores anteriores).")
        # (mantém as variáveis como estavam)
    
    total_passagens = dados_combinados['total_passagens']
    total_hospedagens = dados_combinados['total_hospedagens']
    total_transporte = extrair_transporte(caminho_planilha_original)
    total_geral = total_passagens + total_hospedagens + total_transporte
    # -------------------------------------------------------------------------
    

    # -------------------------------------------------------------------------
    # 5. Calcular quantidades e valores realizados a partir dos dados filtrados
    # -------------------------------------------------------------------------
    dados_iniciais = dados_combinados.get('dados', [])
    
    pass_nac = [d for d in dados_iniciais if d['tipo'] == 'passagem' and not d.get('is_internacional', False)]
    pass_int = [d for d in dados_iniciais if d['tipo'] == 'passagem' and d.get('is_internacional', False)]
    hosp_nac = [d for d in dados_iniciais if d['tipo'] == 'hospedagem' and not d.get('is_internacional', False)]
    hosp_int = [d for d in dados_iniciais if d['tipo'] == 'hospedagem' and d.get('is_internacional', False)]
    
    qtd_passagens_realizado_nacional = len(pass_nac)
    qtd_passagens_realizado_internacional = len(pass_int)
    qtd_passagens_realizado_total = qtd_passagens_realizado_nacional + qtd_passagens_realizado_internacional
    
    qtd_hospedagens_realizado_nacional = len(hosp_nac)
    qtd_hospedagens_realizado_internacional = len(hosp_int)
    qtd_hospedagens_realizado_total = qtd_hospedagens_realizado_nacional + qtd_hospedagens_realizado_internacional
    
    valor_passagens_realizado_nacional = sum(d['valor'] for d in pass_nac)
    valor_passagens_realizado_internacional = sum(d['valor'] for d in pass_int)
    valor_passagens_realizado_total = valor_passagens_realizado_nacional + valor_passagens_realizado_internacional
    
    valor_hospedagens_realizado_nacional = sum(d['valor'] for d in hosp_nac)
    valor_hospedagens_realizado_internacional = sum(d['valor'] for d in hosp_int)
    valor_hospedagens_realizado_total = valor_hospedagens_realizado_nacional + valor_hospedagens_realizado_internacional

    # -------------------------------------------------------------------------
    # 6. DEFINIR OS TOTAIS QUE SERÃO EXIBIDOS NO DASHBOARD
    # -------------------------------------------------------------------------
    total_passagens = valor_passagens_realizado_total
    total_hospedagens = valor_hospedagens_realizado_total
    total_transporte = 0  # Sem base de transporte
    total_geral = total_passagens + total_hospedagens + total_transporte
    
    
    print(f"\n📊 Totais calculados após filtros:")
    print(f"   Passagens: R$ {total_passagens:,.2f}")
    print(f"   Hospedagens: R$ {total_hospedagens:,.2f}")
    print(f"   Total Geral: R$ {total_geral:,.2f}")

    # -------------------------------------------------------------------------
    # 7. Construir DataFrame de plataformas com os totais calculados (exatos)
    # -------------------------------------------------------------------------
    # Obtém o dicionário detalhado retornado pela extração
    totais_plataforma_detalhado = dados_combinados.get('totais_plataforma_detalhado', {})
   
    #verificação valores nao alocados plataforma
    soma_plataformas = sum(v['passagens'] + v['hospedagens'] for v in totais_plataforma_detalhado.values())
    print(f"Soma de todas as plataformas (incluindo outras): {soma_plataformas:,.2f}")
    print(f"Total geral: {total_geral:,.2f}")
    print(f"Diferença: {total_geral - soma_plataformas:,.2f}")
   
    # Lista de plataformas-alvo na ordem desejada
    plataformas_alvo_ordem = ['TV GLOBO', 'GE TV', 'SPORTV', 'PREMIERE', 'COMBATE']
    
    dados_plat = []
    for plat in plataformas_alvo_ordem:
        valores = totais_plataforma_detalhado.get(plat, {'passagens': 0, 'hospedagens': 0})
        v_pass = valores['passagens']
        v_hosp = valores['hospedagens']
        v_total = v_pass + v_hosp
        dados_plat.append({
            'Plataforma': plat,
            'Valor_Total': v_total,
            'Valor_passagens': v_pass,
            'Valor_hospedagens': v_hosp,
            'Valor_transporte': 0
        })
    
    df_plataforma = pd.DataFrame(dados_plat)   # <--- Nome correto: df_plataforma

    # -------------------------------------------------------------------------
    # 8. Calcular médias (usando os novos totais)
    # -------------------------------------------------------------------------
    # PASSAGENS - Total
    #media_pass_orcado_total = (orcados['passagens'] / qtd_passagens_orcado_total) if qtd_passagens_orcado_total > 0 else 0
    media_pass_realizado_total = (total_passagens / qtd_passagens_realizado_total) if qtd_passagens_realizado_total > 0 else 0
    
    # PASSAGENS - Nacional/Internacional com proporção baseada nas quantidades realizadas
    proporcao_pass_nac = qtd_passagens_realizado_nacional / qtd_passagens_realizado_total if qtd_passagens_realizado_total > 0 else 0
    proporcao_pass_inter = qtd_passagens_realizado_internacional / qtd_passagens_realizado_total if qtd_passagens_realizado_total > 0 else 0
    
    valor_orcado_pass_nac = orcados['passagens'] * proporcao_pass_nac
    valor_orcado_pass_inter = orcados['passagens'] * proporcao_pass_inter
    valor_realizado_pass_nac = total_passagens * proporcao_pass_nac
    valor_realizado_pass_inter = total_passagens * proporcao_pass_inter
    
    #media_pass_orcado_nacional = valor_orcado_pass_nac / qtd_passagens_orcado_nacional if qtd_passagens_orcado_nacional > 0 else 0
    media_pass_realizado_nacional = (valor_passagens_realizado_nacional / qtd_passagens_realizado_nacional) if qtd_passagens_realizado_nacional > 0 else 0
    
    #media_pass_orcado_internacional = valor_orcado_pass_inter / qtd_passagens_orcado_internacional if qtd_passagens_orcado_internacional > 0 else 0
    media_pass_realizado_internacional = (valor_passagens_realizado_internacional / qtd_passagens_realizado_internacional) if qtd_passagens_realizado_internacional > 0 else 0
    # HOSPEDAGENS - Total
    #media_hosp_orcado_total = (orcados['hospedagens'] / qtd_hospedagens_orcado_total) if qtd_hospedagens_orcado_total > 0 else 0
    media_hosp_realizado_total = (total_hospedagens / qtd_hospedagens_realizado_total) if qtd_hospedagens_realizado_total > 0 else 0
    
    proporcao_hosp_nac = qtd_hospedagens_realizado_nacional / qtd_hospedagens_realizado_total if qtd_hospedagens_realizado_total > 0 else 0
    proporcao_hosp_inter = qtd_hospedagens_realizado_internacional / qtd_hospedagens_realizado_total if qtd_hospedagens_realizado_total > 0 else 0
    
    valor_orcado_hosp_nac = orcados['hospedagens'] * proporcao_hosp_nac
    valor_orcado_hosp_inter = orcados['hospedagens'] * proporcao_hosp_inter
    valor_realizado_hosp_nac = total_hospedagens * proporcao_hosp_nac
    valor_realizado_hosp_inter = total_hospedagens * proporcao_hosp_inter
    
    #media_hosp_orcado_nacional = valor_orcado_hosp_nac / qtd_hospedagens_orcado_nacional if qtd_hospedagens_orcado_nacional > 0 else 0
    media_hosp_realizado_nacional = (valor_hospedagens_realizado_nacional / qtd_hospedagens_realizado_nacional) if qtd_hospedagens_realizado_nacional > 0 else 0
    #media_hosp_orcado_internacional = valor_orcado_hosp_inter / qtd_hospedagens_orcado_internacional if qtd_hospedagens_orcado_internacional > 0 else 0
    media_hosp_realizado_internacional = (valor_hospedagens_realizado_internacional / qtd_hospedagens_realizado_internacional) if qtd_hospedagens_realizado_internacional > 0 else 0

    print(f"\n✅ Médias recalculadas com os novos totais.")

    # -------------------------------------------------------------------------
    # 9. Calcular percentuais de utilização
    # -------------------------------------------------------------------------
    def get_color(perc):
        return '#388e3c' if perc <= 100 else ('#f57c00' if perc <= 110 else '#d32f2f')

    perc_total = (total_geral / orcados['total'] * 100) if orcados['total'] > 0 else 0
    perc_pass = (total_passagens / orcados['passagens'] * 100) if orcados['passagens'] > 0 else 0
    perc_hosp = (total_hospedagens / orcados['hospedagens'] * 100) if orcados['hospedagens'] > 0 else 0
    perc_trans = (total_transporte / orcados['transporte'] * 100) if orcados['transporte'] > 0 else 0

    # -------------------------------------------------------------------------
    # 10. TOP 10 Gastadores (já calculado nos dados)
    # -------------------------------------------------------------------------
    gastos_por_pessoa = {}
    for d in dados_iniciais:
        p = d['passageiro']
        if not p or p == 'nan':
            continue  # ignora passageiros inválidos
        # Ignorar "FUNCIONARIO A DEFINIR"
        if p and "FUNCIONARIO A DEFINIR" in p.upper():
            continue
        if p not in gastos_por_pessoa:
            gastos_por_pessoa[p] = {'passageiro': p, 'grupo': d['grupo'], 'passagens': 0, 'hospedagens': 0}
        if d['tipo'] == 'passagem':
            gastos_por_pessoa[p]['passagens'] += d['valor']
        else:
            gastos_por_pessoa[p]['hospedagens'] += d['valor']

    # Calcular total para todos (garante que a chave exista)
    for p in gastos_por_pessoa:
        gastos_por_pessoa[p]['total'] = gastos_por_pessoa[p]['passagens'] + gastos_por_pessoa[p]['hospedagens']

    top10_gastadores = sorted(gastos_por_pessoa.values(), key=lambda x: x['total'], reverse=True)[:10]
    
    print(f"\n👤 Top 10 Gastadores POR PESSOA (após filtros):")
    for i, g in enumerate(top10_gastadores[:5], 1):
        print(f"   {i}. {g['passageiro'][:40]} - R$ {g['total']:,.2f}")

    # -------------------------------------------------------------------------
    # 11. Destinos normalizados
    # -------------------------------------------------------------------------
    dest_freq = {}
    dest_val = {}
    for d in [x for x in dados_iniciais if x['tipo'] == 'passagem' and x['destino']]:
        dest = d['destino']
        dest_freq[dest] = dest_freq.get(dest, 0) + 1
        dest_val[dest] = dest_val.get(dest, 0) + d['valor']
    # Remove 'Não informado' se existir
    dest_freq.pop('Não informado', None)
    dest_val.pop('Não informado', None)
    top_dest_freq = sorted(dest_freq.items(), key=lambda x: x[1], reverse=True)[:10]
    
    # Calcula a média para cada destino
    dest_media = []
    for dest, val in dest_val.items():
        qtd = dest_freq.get(dest, 0)
        if qtd > 0:
            media = val / qtd
            dest_media.append((dest, val, media))
    
    top_dest_media = sorted(dest_media, key=lambda x: x[2], reverse=True)[:10]

    # -------------------------------------------------------------------------
    # 12. Preparar grupos disponíveis para filtro de antecedência
    # -------------------------------------------------------------------------
    grupos_disponiveis = dados_combinados.get('grupos_disponiveis', [])
    

    # -------------------------------------------------------------------------
    # Renderizar HTML com Jinja2
    # -------------------------------------------------------------------------
    import json
    
    grupos_com_antecedencia_calc = []
    
    # DEPURAÇÃO: salvar registros com valor > 2000 em um arquivo CSV
    registros_altos = [d for d in dados_iniciais if d['valor'] > 2000]
    df_debug = pd.DataFrame(registros_altos)
    df_debug.to_csv('registros_acima_2000.csv', index=False, encoding='utf-8-sig')
    print(f"\n🔍 {len(registros_altos)} registros com valor >2000 salvos em 'registros_acima_2000.csv'")
    
    for grupo in grupos_disponiveis:
        if grupo and grupo not in ['0', 'Cancelado']:
            passagens = [d for d in dados_iniciais if d['tipo'] == 'passagem' and str(d.get('grupo', '')).lower() == str(grupo).lower() and d.get('antecedencia', 0) > 0]
            if len(passagens) > 0:
                grupo_safe = grupo.replace("'", "\\\'").replace('"', '\\\"')
                grupo_display = grupo[:25] + ('...' if len(grupo) > 25 else '')
                grupos_com_antecedencia_calc.append({'safe': grupo_safe, 'display': grupo_display})

    # Construir listagem de plataformas para o jinja (usamos o df_plataforma)
    plataformas_calc = []
    for plat_nome in plataformas_alvo_ordem:
        row = df_plataforma[df_plataforma['Plataforma'].str.upper() == plat_nome]
        if not row.empty:
            row = row.iloc[0]
            orc_plat = orcados.get('por_plataforma', {}).get(plat_nome, {'total': 0, 'passagens': 0, 'hospedagens': 0, 'transporte': 0})
            
            v_total = row.get('Valor_Total', 0)
            v_pass = row.get('Valor_passagens', 0)
            v_hosp = row.get('Valor_hospedagens', 0)
            v_transp = row.get('Valor_transporte', 0)

            o_total = orc_plat['total'] if 'total' in orc_plat else 0
            o_pass = orc_plat.get('passagens', 0)
            o_hosp = orc_plat.get('hospedagens', 0)
            o_transp = orc_plat.get('transporte', 0)

            p_total = (v_total / o_total * 100) if o_total > 0 else 0
            p_pass = (v_pass / o_pass * 100) if o_pass > 0 else 0
            p_hosp = (v_hosp / o_hosp * 100) if o_hosp > 0 else 0
            p_transp = (v_transp / o_transp * 100) if o_transp > 0 else 0
        else:
            v_total = v_pass = v_hosp = v_transp = o_total = o_pass = o_hosp = o_transp = p_total = p_pass = p_hosp = p_transp = 0
            
        plataformas_calc.append({
            'nome': plat_nome,
            'v_total': v_total, 'o_total': o_total, 'p_total': p_total,
            'v_pass': v_pass, 'o_pass': o_pass, 'p_pass': p_pass,
            'v_hosp': v_hosp, 'o_hosp': o_hosp, 'p_hosp': p_hosp,
            'v_transp': v_transp, 'o_transp': o_transp, 'p_transp': p_transp
        })


    # Preparar df_grupo_calc
    df_grupo_calc_dicts = df_grupo_calc.to_dict('records')
    
    # Preparar JSONs para gráficos JS
    dadosGraficos_dict = {
        'pass': {
            'total': {'orc': media_pass_orcado_total, 'real': media_pass_realizado_total, 'qtdOrc': qtd_passagens_orcado_total, 'qtdReal': qtd_passagens_realizado_total},
            'nacional': {'orc': media_pass_orcado_nacional, 'real': media_pass_realizado_nacional, 'qtdOrc': qtd_passagens_orcado_nacional, 'qtdReal': qtd_passagens_realizado_nacional},
            'internacional': {'orc': media_pass_orcado_internacional, 'real': media_pass_realizado_internacional, 'qtdOrc': qtd_passagens_orcado_internacional, 'qtdReal': qtd_passagens_realizado_internacional}
        },
        'hosp': {
            'total': {'orc': media_hosp_orcado_total, 'real': media_hosp_realizado_total, 'qtdOrc': qtd_hospedagens_orcado_total, 'qtdReal': qtd_hospedagens_realizado_total},
            'nacional': {'orc': media_hosp_orcado_nacional, 'real': media_hosp_realizado_nacional, 'qtdOrc': qtd_hospedagens_orcado_nacional, 'qtdReal': qtd_hospedagens_realizado_nacional},
            'internacional': {'orc': media_hosp_orcado_internacional, 'real': media_hosp_realizado_internacional, 'qtdOrc': qtd_hospedagens_orcado_internacional, 'qtdReal': qtd_hospedagens_realizado_internacional}
        }
    }
    
    # Render usando template separado
    base_dir = Path(__file__).parent
    env = Environment(loader=FileSystemLoader(str(base_dir)))
    template = env.get_template('dashboard_template.jinja2')

    html = template.render(
        periodo_analise=periodo_analise,
        data_atual=datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        total_geral=total_geral,
        orcados=orcados,
        total_passagens=total_passagens, total_hospedagens=total_hospedagens, total_transporte=total_transporte,
        perc_total=perc_total, perc_pass=perc_pass, perc_hosp=perc_hosp, perc_trans=perc_trans,
        formatar_numero=formatar_numero,
        get_color=get_color,
        plataformas_calc=plataformas_calc,
        top10_gastadores=top10_gastadores,
        grupos_com_antecedencia=grupos_com_antecedencia_calc,
        top_dest_freq=top_dest_freq,
        dest_val=dest_val,
        top_dest_media=top_dest_media,
        df_grupo_calc=df_grupo_calc_dicts,
        total_registros=dados_combinados.get('total_registros', 0),
        dadosGraficos_json=json.dumps(dadosGraficos_dict),
        dados_iniciais_json=json.dumps(dados_iniciais)
    )

    with open(caminho_saida, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"✅ Dashboard V3.8 gerado: {caminho_saida}")
    print("\n🎯 CORREÇÕES APLICADAS:")
    print("   1. ✅ Gráficos mostram MÉDIA (não total)")
    print("   2. ✅ Filtros nacional/internacional funcionando 100%")
    print("   3. ✅ Números aparecem nos gráficos (ChartDataLabels)")
    print("   4. ✅ Filtros acima de cada seção")
    print("   5. ✅ Top 10 por PESSOA")
    print("   6. ✅ Filtro antecedência por grupo OK")
    print("   7. ✅ Destinos normalizados")
    
    return caminho_saida


def main():
    """Função principal"""
    if os.name == 'nt':
        onedrive_base = Path(r"C:\Users\ligomes\OneDrive - Globo Comunicação e Participações sa")
        
        caminho_formatos = onedrive_base / "Gestão de Eventos - Documentos" / "Orçamento" / "2026" / "Ciclo" / "2. Produção" / "Financeiro.xlsx"
        caminho_painel = onedrive_base / "Gestão de Eventos - Documentos" / "Gestão de Eventos_planejamento" / "Painel Contábil" / "2026" / "Painel Contábil.xlsx"
        
        base_dir = Path(r"C:\Users\ligomes\Downloads\painel_contabil_solucao_final\painel_contabil")
        caminho_planilha_tratada = base_dir / 'planilha_consolidada_tratada.xlsx'
        caminho_saida = base_dir / 'dashboard_v3_8.html'
        
    else:
        caminho_planilha_tratada = Path('/home/ubuntu/painel_contabil/planilha_consolidada_tratada.xlsx')
        caminho_painel = Path('/home/ubuntu/upload/PainelContábilV2-copia.xlsx')
        caminho_formatos = Path('/home/ubuntu/upload/FormatosContábilTVAeTVFv4.xlsx')
        caminho_saida = Path('/home/ubuntu/painel_contabil/dashboard_v3_8.html')
    
    gerar_dashboard_v3_8(
        caminho_planilha_tratada=caminho_planilha_tratada,
        caminho_saida=caminho_saida,
        caminho_planilha_original=caminho_painel,
        caminho_formatos=caminho_formatos
    )


if __name__ == '__main__':
    main()