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

# Balão informativo com as instruções exatas solicitadas
st.info("""
**💡 Instruções para Planilha de Protocolos:**
* A planilha deve conter a coluna **'Responsavel'** logo após o nome do cliente.
* **Filtros obrigatórios pré-upload:** Protocolo Abertura e Equipe Comercial (Interno/Externo).
* **Regra de Responsabilidade:** O sistema busca o responsável pelo ganho na coluna 'Responsavel'.
* **Vínculo:** O sistema cruza **'Nome Cliente'** (Ativação) com **'Cliente'** (Protocolos).
* **Segurança:** Caso o responsável não seja encontrado nos protocolos, o sistema usará o **Vendedor 1** automaticamente.
""")

if arquivo_ativacao and arquivo_protocolos:
    try:
        def carregar_dados(arq):
            if arq.name.lower().endswith('.csv'):
                return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1')
            return pd.read_excel(arq)

        df_ativacao = carregar_dados(arquivo_ativacao)
        df_protocolos = carregar_dados(arquivo_protocolos)

        # Limpeza de nomes de colunas
        df_ativacao.columns = [str(c).strip() for c in df_ativacao.columns]
        df_protocolos.columns = [str(c).strip() for c in df_protocolos.columns]

        # 1. Filtro de Status na Ativação
        if 'Status Contrato' in df_ativacao.columns:
            df_ativacao = df_ativacao[df_ativacao['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # 2. Cruzamento de Dados (Merge)
        if 'Nome Cliente' in df_ativacao.columns and 'Cliente' in df_protocolos.columns:
            if 'Responsavel' in df_protocolos.columns:
                
                # Normalização das chaves para vínculo
                df_ativacao['_JOIN_KEY'] = df_ativacao['Nome Cliente'].astype(str).str.strip().str.upper()
                df_protocolos['_JOIN_KEY'] = df_protocolos['Cliente'].astype(str).str.strip().str.upper()

                # Prepara protocolos
                df_prot_clean = df_protocolos.drop_duplicates(subset=['_JOIN_KEY'])[['_JOIN_KEY', 'Responsavel']]
                
                # Merge
                df = pd.merge(df_ativacao, df_prot_clean, on='_JOIN_KEY', how='left', suffixes=('_orig', ''))
                
                # Segurança: Vendedor 1 no lugar de Responsavel vazio
                if 'Responsavel' in df.columns and 'Vendedor 1' in df.columns:
                    df['Responsavel'] = df['Responsavel'].fillna(df['Vendedor 1'])
                    df.loc[df['Responsavel'].astype(str).str.strip() == "", 'Responsavel'] = df['Vendedor 1']
                
                df = df.drop(columns=['_JOIN_KEY'])
            else:
                st.error("⚠️ Coluna 'Responsavel' não encontrada na planilha de Protocolos.")
                df = df_ativacao
        else:
            st.error("⚠️ Verifique as colunas de vínculo: 'Nome Cliente' e 'Cliente'.")
            df = df_ativacao

        # --- SEÇÃO DE PERSONALIZAÇÃO ---
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

            # 4. Processamento com Estilos Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Planilha')
                ws = writer.sheets['Planilha']
                
                amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
                fonte = Font(name='Calibri', size=11, bold=False)

                for col_idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=col_idx)
                    nome_col = str(header.value).strip()
                    
                    # REGRAS DE CORES:
                    # 1. Se for 'Status Contrato', SEMPRE AMARELO
                    if nome_col == "Status Contrato":
                        header.fill = amarelo
                    # 2. Se for Coluna E (5) ou Coluna O (15) SEM ser Status Contrato, fica VERDE
                    elif col_idx == 5 or col_idx == 15:
                        header.fill = verde
                    # 3. Se forem as últimas 4 colunas, fica VERDE
                    elif col_idx > len(colunas_selecionadas) - 4:
                        header.fill = verde
                    # 4. Se estiver entre as 9 primeiras, fica AMARELO
                    elif col_idx <= 9:
                        header.fill = amarelo
                    
                    for cell in col_cells:
                        cell.font = fonte
                    ws.column_dimensions[header.column_letter].width = 22

            st.success("✅ Estilização aplicada: 'Status Contrato' em destaque amarelo.")
            st.download_button(
                label="📥 Baixar Planilha Final",
                data=output.getvalue(),
                file_name="PLANILHA_CONSOLIDADA_NETMANIA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
