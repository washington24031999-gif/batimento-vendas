import sys
from types import ModuleType

# Correção técnica para compatibilidade com Python 3.13
if 'imghdr' not in sys.modules:
    sys.modules['imghdr'] = ModuleType('imghdr')

import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import column_index_from_string
from io import BytesIO

st.set_page_config(page_title="Netmania Optimizer Pro", layout="wide")
st.title("📊 Estruturador de Planilhas + Relatório de Protocolos")

# --- SEÇÃO DE UPLOAD ---
col1, col2 = st.columns(2)
with col1:
    arquivo_base = st.file_uploader("Selecione a Base Principal", type=['xlsx', 'csv', 'xlsm'])
with col2:
    arquivo_proto = st.file_uploader("Selecione o Relatório de Protocolos (Opcional)", type=['xlsx', 'xlsm'])

if arquivo_base:
    try:
        # 1. Leitura do Arquivo Base
        if arquivo_base.name.lower().endswith('.csv'):
            df = pd.read_csv(arquivo_base, sep=None, engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(arquivo_base)
        
        df.columns = [str(c).strip() for c in df.columns]

        # 2. Filtro de Status (Regra Original)
        if 'Status Contrato' in df.columns:
            df = df[df['Status Contrato'].str.lower() != 'cancelado']

        # 3. Lógica se houver Relatório de Protocolos (DE/PARA Automático)
        if arquivo_proto:
            st.info("🔄 Relatório de Protocolos detectado. Aplicando mapeamento fixo...")
            df_proto = pd.read_excel(arquivo_proto)
            
            # Criamos o DataFrame de saída com colunas A até N (14 colunas)
            colunas_resultado = [chr(65 + i) for i in range(14)] # A, B, C... N
            df_final = pd.DataFrame(columns=colunas_resultado, index=range(len(df_proto)))

            # Dicionário DE (Letra no Protocolo) -> PARA (Letra no Resultado)
            mapeamento = {
                'AK': 'B', 'H': 'C', 'I': 'D', 'J': 'E',
                'K': 'F', 'P': 'G', 'AO': 'H', 'AU': 'I',
                'AV': 'J', 'AS': 'K', 'AQ': 'L', 'AL': 'N'
            }

            for de, para in mapeamento.items():
                try:
                    idx_origem = column_index_from_string(de) - 1
                    idx_destino = column_index_from_string(para) - 1
                    
                    if idx_origem < len(df_proto.columns):
                        df_final.iloc[:, idx_destino] = df_proto.iloc[:, idx_origem]
                except:
                    continue
            
            st.success("✅ Mapeamento do Relatório de Protocolos aplicado!")
        
        else:
            # Fluxo Original: Seleção Manual de Colunas
            st.subheader("⚙️ Personalize sua exportação manual")
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
                "Selecione as colunas da Base Principal:",
                options=colunas_disponiveis,
                default=selecao_inicial
            )
            df_final = df[colunas_selecionadas] if colunas_selecionadas else df

        # --- PREPARAÇÃO DO DOWNLOAD ---
        if not df_final.empty:
            st.write("Prévia dos dados:")
            st.dataframe(df_final.head(10))

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Planilha')
                ws = writer.sheets['Planilha']
                
                # Estilização básica
                amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                fonte = Font(name='Calibri', size=11)

                for col_idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=col_idx)
                    header.fill = amarelo
                    for cell in col_cells:
                        cell.font = fonte
                    ws.column_dimensions[header.column_letter].width = 20

            st.download_button(
                label="📥 Baixar Planilha Finalizada",
                data=output.getvalue(),
                file_name="PLANILHA_PROCESSADA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro: {e}")
