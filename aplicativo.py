import sys
from types import ModuleType

# Correção técnica para compatibilidade com Python 3.13 (imghdr)
if 'imghdr' not in sys.modules:
    sys.modules['imghdr'] = ModuleType('imghdr')

import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font
from io import BytesIO

st.set_page_config(page_title="Netmania Optimizer", layout="wide")
st.title("📊 Estruturador de Planilhas Personalizado")

# --- SEÇÃO DE UPLOAD ---
col1, col2 = st.columns(2)

with col1:
    arquivo_ativacao = st.file_uploader("1. Planilha de Ativação de Contrato", type=['xlsx', 'csv', 'xlsm'])

with col2:
    arquivo_protocolos = st.file_uploader("2. Planilha de Protocolos", type=['xlsx', 'csv', 'xlsm'])

# Balão informativo obrigatório
st.info("""
**💡 Instruções para Planilha de Protocolos:**
A planilha de protocolos deve conter obrigatoriamente uma coluna com o nome do **Responsável** pelo ganho de venda. 
Certifique-se de que o arquivo já contenha os filtros aplicados: *Protocolo Encerrado* e *Equipe Comercial Interno/Externo*.
""")

if arquivo_ativacao and arquivo_protocolos:
    try:
        # Funcao auxiliar para ler arquivos
        def carregar_dados(arq):
            if arq.name.lower().endswith('.csv'):
                return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1')
            return pd.read_excel(arq)

        df_ativacao = carregar_dados(arquivo_ativacao)
        df_protocolos = carregar_dados(arquivo_protocolos)

        # Limpeza básica de nomes de colunas
        df_ativacao.columns = [str(c).strip() for c in df_ativacao.columns]
        df_protocolos.columns = [str(c).strip() for c in df_protocolos.columns]

        # 1. Filtro de Status na Ativação
        if 'Status Contrato' in df_ativacao.columns:
            df_ativacao = df_ativacao[df_ativacao['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # 2. Cruzamento de Dados (Merge)
        # Buscamos 'Responsavel' na planilha de protocolos usando 'Nome Cliente' como chave
        if 'Nome Cliente' in df_ativacao.columns and 'Nome Cliente' in df_protocolos.columns:
            if 'Responsavel' in df_protocolos.columns:
                # Removemos duplicatas de protocolos para não gerar linhas extras no merge
                df_prot_clean = df_protocolos.drop_duplicates(subset=['Nome Cliente'])[['Nome Cliente', 'Responsavel']]
                
                # Unifica as planilhas
                df = pd.merge(df_ativacao, df_prot_clean, on='Nome Cliente', how='left', suffixes=('', '_prot'))
                
                # Se já existia uma coluna Responsavel vazia, ela é atualizada
                if 'Responsavel' in df.columns:
                    df['Responsavel'] = df['Responsavel'].fillna(df.get('Responsavel_prot', ''))
            else:
                st.error("⚠️ A coluna 'Responsavel' não foi encontrada na planilha de Protocolos.")
                df = df_ativacao
        else:
            st.warning("⚠️ Coluna 'Nome Cliente' não encontrada em ambas as planilhas para vincular os dados.")
            df = df_ativacao

        # --- SEÇÃO DE PERSONALIZAÇÃO E ORDEM ---
        st.subheader("⚙️ Personalize sua exportação")
        
        # Coluna Responsavel forçada na posição H (índice 7) conforme solicitado
        ordem_padrao = [
            'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
            'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
            'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
            'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
            'Origem', 'Valor Primeira Mensalidade'
        ]
        
        colunas_disponiveis = list(df.columns)
        selecao_inicial = [c for c in ordem_padrao if c in colunas_disponiveis]

        colunas_selecionadas = st.multiselect(
            "Selecione e ordene as colunas:",
            options=colunas_disponiveis,
            default=selecao_inicial
        )

        if not colunas_selecionadas:
            st.warning("⚠️ Selecione pelo menos uma coluna.")
        else:
            df_final = df[colunas_selecionadas]
            st.dataframe(df_final, use_container_width=True)

            # 4. Processamento com Estilos
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Planilha')
                ws = writer.sheets['Planilha']
                
                amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
                fonte = Font(name='Calibri', size=11, bold=False)

                for col_idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=col_idx)
                    nome = str(header.value).strip()
                    
                    # REGRAS DE CORES ATUALIZADAS:
                    # Coluna E (5) e Coluna O (15) em VERDE
                    if col_idx == 5 or col_idx == 15:
                        header.fill = verde
                    # Outras colunas finais em VERDE
                    elif col_idx > len(colunas_selecionadas) - 4:
                        header.fill = verde
                    # Iniciais e Status em AMARELO
                    elif col_idx <= 9 or nome == "Status Contrato":
                        header.fill = amarelo
                    
                    for cell in col_cells:
                        cell.font = fonte
                    ws.column_dimensions[header.column_letter].width = 22

            st.success(f"✅ Processamento concluído! {len(df_final)} linhas geradas.")
            st.download_button(
                label="📥 Baixar Planilha Final",
                data=output.getvalue(),
                file_name="PLANILHA_CONSOLIDADA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar arquivos: {e}")
