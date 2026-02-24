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

st.title("📊 Estruturador de Planilhas Personalizado - Etapa 3")

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

        # Limpeza agressiva de nomes de colunas
        df_ativacao.columns = [str(c).strip() for c in df_ativacao.columns]
        df_protocolos.columns = [str(c).strip() for c in df_protocolos.columns]

        # 1. Filtro Inicial (Remover Cancelados da Ativação)
        if 'Status Contrato' in df_ativacao.columns:
            df_ativacao = df_ativacao[df_ativacao['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # 2. Cruzamento Base (Ativação + Protocolos)
        if 'Nome Cliente' in df_ativacao.columns and 'Cliente' in df_protocolos.columns:
            df_ativacao['_JOIN_KEY'] = df_ativacao['Nome Cliente'].astype(str).str.strip().str.upper()
            df_protocolos['_JOIN_KEY'] = df_protocolos['Cliente'].astype(str).str.strip().str.upper()

            # Captura o Responsável do protocolo de abertura
            df_prot_clean = df_protocolos.drop_duplicates(subset=['_JOIN_KEY'])[['_JOIN_KEY', 'Responsavel']]
            df = pd.merge(df_ativacao, df_prot_clean, on='_JOIN_KEY', how='left')
            
            # REGRA DE SEGURANÇA: Vendedor 1 assume se Responsavel estiver vazio
            if 'Vendedor 1' in df.columns:
                df['Responsavel'] = df['Responsavel'].fillna(df['Vendedor 1'])
                df.loc[df['Responsavel'].astype(str).str.strip() == "", 'Responsavel'] = df['Vendedor 1']

            # 3. Integração do Relatório de Reativações (Se houver upload)
            if arquivo_reativacao:
                df_reat = carregar_dados(arquivo_reativacao)
                df_reat.columns = [str(c).strip() for c in df_reat.columns]
                
                if 'Cliente' in df_reat.columns:
                    df_reat['_JOIN_KEY'] = df_reat['Cliente'].astype(str).str.strip().str.upper()
                    
                    # Selecionamos colunas chave da estrutura de reativação para o merge inicial
                    colunas_reat_desejadas = ['_JOIN_KEY', 'Tipo Solicitacao', 'Situacao', 'Protocolo']
                    colunas_existentes = [c for c in colunas_reat_desejadas if c in df_reat.columns]
                    
                    df_reat_clean = df_reat.drop_duplicates(subset=['_JOIN_KEY'])[colunas_existentes]
                    # Merge com sufixo para evitar conflito com 'Protocolo' da planilha 2
                    df = pd.merge(df, df_reat_clean, on='_JOIN_KEY', how='left', suffixes=('', '_Reat'))
                    st.toast("✅ Reativações vinculadas!", icon="🔄")

            df = df.drop(columns=['_JOIN_KEY'])
        else:
            st.warning("⚠️ Coluna 'Nome Cliente' ou 'Cliente' não encontrada para o vínculo.")
            df = df_ativacao

        # --- SEÇÃO DE PERSONALIZAÇÃO ---
        st.subheader("⚙️ Configurações da Exportação")
        
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

            # Estilização Excel (Centralizado, Sem Bordas, Cores Específicas)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Consolidado')
                ws = writer.sheets['Consolidado']
                
                amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
                fonte = Font(name='Calibri', size=11)
                centralizado = Alignment(horizontal='center', vertical='center')
                sem_bordas = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

                for col_idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=col_idx)
                    nome_col = str(header.value).strip()
                    
                    # Regras de Cores do Cabeçalho
                    if nome_col == "Status Contrato":
                        header.fill = amarelo
                    elif col_idx == 5 or col_idx == 15:
                        header.fill = verde
                    elif col_idx > len(col_selecionadas) - 4:
                        header.fill = verde
                    elif col_idx <= 9:
                        header.fill = amarelo
                    
                    # Aplicação Geral de Estilo
                    for cell in col_cells:
                        cell.font = fonte
                        cell.alignment = centralizado
                        cell.border = sem_bordas
                    ws.column_dimensions[header.column_letter].width = 25

            st.download_button(label="📥 Baixar Planilha Final (Etapa 3)", data=output.getvalue(), file_name="FINAL_NETMANIA_OTIMIZADA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")

# --- TUTORIAL FIXO NO RODAPÉ ---
st.divider()
st.subheader("📖 Guia de Procedimentos e Tutorial")

col_t1, col_t2 = st.columns(2)

with col_t1:
    st.markdown("""
    #### 💡 Instruções para Planilha de Protocolos:
    * **Estrutura:** Deve conter a coluna **'Responsavel'** logo após o nome do cliente.
    * **Filtros Obrigatórios:** O arquivo deve ser extraído com os filtros *Protocolo Abertura* e *Equipe Comercial (Interno/Externo)*.
    * **Vínculo:** O sistema cruza **'Nome Cliente'** (Ativação) com **'Cliente'** (Protocolos).
    * **Segurança:** Caso o responsável não seja encontrado nos protocolos, o sistema usará o **Vendedor 1** automaticamente para evitar células vazias.
    """)

with col_t2:
    st.markdown("""
    #### 🔄 Relatório de Reativações (Planilha 3):
    * **Objetivo:** Identificar clientes que reativaram serviços.
    * **Processo:** O sistema anexa dados como *Tipo de Solicitação* e *Situação* ao relatório principal.
    * **Customização:** Utilize o seletor de colunas acima para incluir campos adicionais (SLA, Cidade, Contrato, etc.) da estrutura de reativação.
    """)

st.info("⚠️ Verifique sempre se os nomes das colunas nas planilhas originais não possuem caracteres especiais extras se o sistema indicar que a coluna não foi encontrada.")
