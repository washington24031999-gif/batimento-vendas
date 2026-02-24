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

# --- SEÇÃO DE UPLOAD ---
col1, col2 = st.columns(2)

with col1:
    arquivo_ativacao = st.file_uploader("1. Planilha de Ativação de Contrato", type=['xlsx', 'csv', 'xlsm'])

with col2:
    arquivo_protocolos = st.file_uploader("2. Planilha de Protocolos", type=['xlsx', 'csv', 'xlsm'])

if arquivo_ativacao and arquivo_protocolos:
    try:
        def carregar_dados(arq):
            if arq.name.lower().endswith('.csv'):
                return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1')
            return pd.read_excel(arq)

        df_ativacao = carregar_dados(arquivo_ativacao)
        df_protocolos = carregar_dados(arquivo_protocolos)

        # Limpeza agressiva de nomes de colunas
        df_ativacao.columns = [str(c).strip() for c in df_ativacao.columns]
        df_protocolos.columns = [str(c).strip() for c in df_protocolos.columns]

        # 1. Filtro de Status
        if 'Status Contrato' in df_ativacao.columns:
            df_ativacao = df_ativacao[df_ativacao['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # 2. Cruzamento de Dados
        col_cliente_ativ = 'Nome Cliente'
        col_cliente_prot = 'Cliente'
        
        if col_cliente_ativ in df_ativacao.columns and col_cliente_prot in df_protocolos.columns:
            if 'Responsavel' in df_protocolos.columns:
                # Normalização para o Join
                df_ativacao['_JOIN_KEY'] = df_ativacao[col_cliente_ativ].astype(str).str.strip().str.upper()
                df_protocolos['_JOIN_KEY'] = df_protocolos[col_cliente_prot].astype(str).str.strip().str.upper()

                df_prot_clean = df_protocolos.drop_duplicates(subset=['_JOIN_KEY'])[['_JOIN_KEY', 'Responsavel']]
                df = pd.merge(df_ativacao, df_prot_clean, on='_JOIN_KEY', how='left')
                
                # Regra de Segurança: Vendedor 1 assume se Responsavel for nulo
                if 'Vendedor 1' in df.columns:
                    df['Responsavel'] = df['Responsavel'].fillna(df['Vendedor 1'])
                    df.loc[df['Responsavel'].astype(str).str.strip() == "", 'Responsavel'] = df['Vendedor 1']
                
                df = df.drop(columns=['_JOIN_KEY'])
            else:
                st.error("⚠️ Coluna 'Responsavel' não encontrada na planilha de Protocolos.")
                df = df_ativacao
        else:
            st.warning(f"⚠️ Colunas de vínculo ({col_cliente_ativ} / {col_cliente_prot}) não encontradas.")
            df = df_ativacao

        # --- PERSONALIZAÇÃO ---
        st.subheader("⚙️ Personalize sua exportação")
        
        ordem_padrao = [
            'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
            'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
            'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
            'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
            'Origem', 'Valor Primeira Mensalidade'
        ]
        
        colunas_disponiveis = list(df.columns)
        selecao_inicial = [c for c in ordem_padrao if c in colunas_disponiveis]

        col_selecionadas = st.multiselect("Selecione e ordene as colunas:", options=colunas_disponiveis, default=selecao_inicial)

        if col_selecionadas:
            df_final = df[col_selecionadas]
            st.dataframe(df_final, use_container_width=True)

            # Estilização Excel
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

            st.success("✅ Estilização concluída: Itens centralizados e títulos sem bordas.")
            st.download_button(
                label="📥 Baixar Planilha Final",
                data=output.getvalue(),
                file_name="PLANILHA_CONSOLIDADA_NETMANIA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")

# --- TUTORIAL NO RODAPÉ ---
st.divider()
st.subheader("📖 Tutorial de Uso")
t1, t2, t3 = st.columns(3)

with t1:
    st.markdown("### 1. Preparação")
    st.write("Filtre a planilha de Protocolos por **Abertura** e **Equipe Comercial**. Ela deve conter a coluna 'Responsavel'.")

with t2:
    st.markdown("### 2. Cruzamento")
    st.write("O sistema une os dados pelo nome do cliente. Se não houver protocolo, o **Vendedor 1** será o responsável.")

with t3:
    st.markdown("### 3. Download")
    st.write("Clique no botão azul acima para baixar. O arquivo virá centralizado, sem bordas e colorido conforme as regras.")
