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

# Balão informativo com as instruções detalhadas e atualizadas para "Responsavel"
st.info("""
**💡 Instruções para Planilha de Protocolos:**
* A planilha deve conter a coluna **'Responsavel'** logo após o nome do cliente.
* **Filtros obrigatórios pré-upload:** Protocolo Abertura e Equipe Comercial (Interno/Externo).
* **Regra de Responsabilidade:** O sistema considera o status de abertura para identificar o responsável pelo ganho da venda através da coluna 'Responsavel'.
* **Vínculo:** O sistema cruza **'Nome Cliente'** (Ativação) com **'Cliente'** (Protocolos).
""")

if arquivo_ativacao and arquivo_protocolos:
    try:
        def carregar_dados(arq):
            if arq.name.lower().endswith('.csv'):
                return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1')
            return pd.read_excel(arq)

        df_ativacao = carregar_dados(arquivo_ativacao)
        df_protocolos = carregar_dados(arquivo_protocolos)

        # Limpeza de nomes de colunas (remover espaços em branco nas pontas)
        df_ativacao.columns = [str(c).strip() for c in df_ativacao.columns]
        df_protocolos.columns = [str(c).strip() for c in df_protocolos.columns]

        # 1. Filtro de Status na Ativação
        if 'Status Contrato' in df_ativacao.columns:
            df_ativacao = df_ativacao[df_ativacao['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # 2. Cruzamento de Dados (Merge)
        if 'Nome Cliente' in df_ativacao.columns and 'Cliente' in df_protocolos.columns:
            # AJUSTE: Agora procurando por 'Responsavel' na planilha de protocolos
            if 'Responsavel' in df_protocolos.columns:
                
                # Normalização das chaves para garantir o vínculo (Maiúsculo e sem espaços)
                df_ativacao['_JOIN_KEY'] = df_ativacao['Nome Cliente'].astype(str).str.strip().str.upper()
                df_protocolos['_JOIN_KEY'] = df_protocolos['Cliente'].astype(str).str.strip().str.upper()

                # Seleciona apenas as colunas necessárias e remove duplicatas de clientes nos protocolos
                df_prot_clean = df_protocolos.drop_duplicates(subset=['_JOIN_KEY'])[['_JOIN_KEY', 'Responsavel']]
                
                # Realiza o cruzamento (Merge/PROCV)
                # O sufixo é tratado caso já exista uma coluna 'Responsavel' na ativação
                df = pd.merge(df_ativacao, df_prot_clean, on='_JOIN_KEY', how='left', suffixes=('_orig', ''))
                
                # Limpeza da chave temporária
                df = df.drop(columns=['_JOIN_KEY'])
            else:
                st.error("⚠️ Coluna 'Responsavel' não encontrada na planilha de Protocolos.")
                df = df_ativacao
        else:
            st.error("⚠️ Verifique os nomes das colunas: 'Nome Cliente' (Ativação) ou 'Cliente' (Protocolos) não encontrados.")
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
                    # Coluna E (5) e O (15) em VERDE
                    if col_idx == 5 or col_idx == 15: 
                        header.fill = verde
                    # 4 últimas em VERDE
                    elif col_idx > len(colunas_selecionadas) - 4: 
                        header.fill = verde
                    # Iniciais (até a 9ª) e Status em AMARELO
                    elif col_idx <= 9 or nome_col == "Status Contrato": 
                        header.fill = amarelo
                    
                    for cell in col_cells:
                        cell.font = fonte
                    ws.column_dimensions[header.column_letter].width = 22

            st.success("✅ Cruzamento concluído! Coluna 'Responsavel' vinculada com sucesso.")
            st.download_button(
                label="📥 Baixar Planilha Consolidada",
                data=output.getvalue(),
                file_name="NETMANIA_OPTIMIZER_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
