import sys
from types import ModuleType

# Correção técnica para compatibilidade com Python 3.13
if 'imghdr' not in sys.modules:
    sys.modules['imghdr'] = ModuleType('imghdr')

import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from io import BytesIO

st.set_page_config(page_title="Netmania Optimizer", layout="wide")

st.title("📊 Estruturador de Planilhas Personalizado")

# --- SEÇÃO DE UPLOAD (TRÊS PLANILHAS) ---
col1, col2, col3 = st.columns(3)

with col1:
    arquivo_ativacao = st.file_uploader("1. Planilha de Ativação", type=['xlsx', 'csv', 'xlsm'])

with col2:
    arquivo_protocolos = st.file_uploader("2. Planilha de Protocolos (Abertura)", type=['xlsx', 'csv', 'xlsm'])

with col3:
    arquivo_reativacao = st.file_uploader("3. Relatório de Reativações", type=['xlsx', 'csv', 'xlsm'])

if arquivo_ativacao and arquivo_protocolos:
    try:
        def carregar_dados(arq):
            if arq.name.lower().endswith('.csv'):
                return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1')
            return pd.read_excel(arq)

        df_ativacao = carregar_dados(arquivo_ativacao)
        df_protocolos = carregar_dados(arquivo_protocolos)

        df_ativacao.columns = [str(c).strip() for c in df_ativacao.columns]
        df_protocolos.columns = [str(c).strip() for c in df_protocolos.columns]

        # 1. Filtro Inicial
        if 'Status Contrato' in df_ativacao.columns:
            df_ativacao = df_ativacao[df_ativacao['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # 2. Cruzamento Base (Ativação + Protocolos)
        if 'Nome Cliente' in df_ativacao.columns and 'Cliente' in df_protocolos.columns:
            df_ativacao['_JOIN_KEY'] = df_ativacao['Nome Cliente'].astype(str).str.strip().str.upper()
            df_protocolos['_JOIN_KEY'] = df_protocolos['Cliente'].astype(str).str.strip().str.upper()

            df_prot_clean = df_protocolos.drop_duplicates(subset=['_JOIN_KEY'])[['_JOIN_KEY', 'Responsavel']]
            df = pd.merge(df_ativacao, df_prot_clean, on='_JOIN_KEY', how='left')
            
            # Segurança Vendedor 1
            if 'Vendedor 1' in df.columns:
                df['Responsavel'] = df['Responsavel'].fillna(df['Vendedor 1'])
                df.loc[df['Responsavel'].astype(str).str.strip() == "", 'Responsavel'] = df['Vendedor 1']

            # --- INTEGRAÇÃO DA TERCEIRA PLANILHA (REATIVAÇÕES) ---
            if arquivo_reativacao:
                df_reat = carregar_dados(arquivo_reativacao)
                df_reat.columns = [str(c).strip() for c in df_reat.columns]
                
                if 'Cliente' in df_reat.columns:
                    df_reat['_JOIN_KEY'] = df_reat['Cliente'].astype(str).str.strip().str.upper()
                    
                    # Selecionamos colunas úteis da sua nova estrutura para identificar a reativação
                    # Você mencionou: Tipo Solicitacao, Situacao, Conclusao, etc.
                    colunas_reat = ['_JOIN_KEY', 'Tipo Solicitacao', 'Situacao', 'Conclusao']
                    colunas_existentes = [c for c in colunas_reat if c in df_reat.columns]
                    
                    df_reat_clean = df_reat.drop_duplicates(subset=['_JOIN_KEY'])[colunas_existentes]
                    
                    # Merge com a base principal
                    df = pd.merge(df, df_reat_clean, on='_JOIN_KEY', how='left')
                    st.toast("✅ Relatório de Reativações integrado com sucesso!", icon="🔄")

            df = df.drop(columns=['_JOIN_KEY'])
        else:
            df = df_ativacao

        # --- PERSONALIZAÇÃO E EXPORTAÇÃO ---
        st.subheader("⚙️ Configurações da Planilha Final")
        
        colunas_disponiveis = list(df.columns)
        col_selecionadas = st.multiselect("Selecione as colunas para o arquivo final:", options=colunas_disponiveis, default=colunas_disponiveis[:19])

        if col_selecionadas:
            df_final = df[col_selecionadas]
            st.dataframe(df_final, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Planilha')
                ws = writer.sheets['Planilha']
                
                amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
                fonte = Font(name='Calibri', size=11)
                centralizado = Alignment(horizontal='center', vertical='center')
                sem_bordas = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

                for col_idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=col_idx)
                    nome_col = str(header.value).strip()
                    
                    # Lógica de cores
                    if nome_col == "Status Contrato":
                        header.fill = amarelo
                    elif col_idx == 5 or col_idx == 15:
                        header.fill = verde
                    elif col_idx > len(col_selecionadas) - 4:
                        header.fill = verde
                    elif col_idx <= 9:
                        header.fill = amarelo
                    
                    for cell in col_cells:
                        cell.font = fonte
                        cell.alignment = centralizado
                        cell.border = sem_bordas
                    ws.column_dimensions[header.column_letter].width = 25

            st.download_button(label="📥 Baixar Planilha Consolidada", data=output.getvalue(), file_name="PLANILHA_NETMANIA_ETAPA3.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")

# --- TUTORIAL ATUALIZADO NO RODAPÉ ---
st.divider()
st.subheader("📖 Tutorial de Uso - Etapa 3")
t1, t2, t3 = st.columns(3)

with t1:
    st.markdown("### 1. Upload Triplo")
    st.write("Agora você pode subir a planilha de **Ativação**, **Protocolos** e o **Relatório de Reativações** simultaneamente.")

with t2:
    st.markdown("### 2. Cruzamento Inteligente")
    st.write("O sistema identifica o cliente nas três bases. Dados de reativação (como tipo de solicitação e situação) são anexados automaticamente.")

with t3:
    st.markdown("### 3. Ajustes Dinâmicos")
    st.write("Como a terceira planilha tem muitas colunas, use o seletor acima para escolher quais informações de reativação deseja manter no relatório final.")
