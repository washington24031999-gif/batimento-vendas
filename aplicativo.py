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

st.set_page_config(page_title="Netmania Multi-Abas Full", layout="wide")
st.title("📊 Gerador de Planilha (Ativações & Reativações)")

col1, col2 = st.columns(2)
with col1:
    arquivo_base = st.file_uploader("1. Selecione a BASE PRINCIPAL", type=['xlsx', 'csv', 'xlsm'])
with col2:
    arquivo_proto = st.file_uploader("2. Selecione o RELATÓRIO DE PROTOCOLOS", type=['xlsx', 'xlsm'])

if arquivo_base and arquivo_proto:
    try:
        # --- 1. PROCESSAMENTO DA ABA 'ATIVAÇÕES' ---
        if arquivo_base.name.lower().endswith('.csv'):
            df_ativacoes = pd.read_csv(arquivo_base, sep=None, engine='python', encoding='latin-1')
        else:
            df_ativacoes = pd.read_excel(arquivo_base)
        
        df_ativacoes.columns = [str(c).strip() for c in df_ativacoes.columns]

        # Garantir que o filtro não seja excessivo
        if 'Status Contrato' in df_ativacoes.columns:
            df_ativacoes = df_ativacoes[df_ativacoes['Status Contrato'].astype(str).str.lower() != 'cancelado']

        # --- 2. PROCESSAMENTO DA ABA 'REATIVAÇÕES' ---
        # Removido o 'nrows' para garantir leitura total
        df_proto_raw = pd.read_excel(arquivo_proto)

        # Criamos o DataFrame estruturado com o exato número de linhas do arquivo de origem
        colunas_letras = [chr(65 + i) for i in range(14)] # A até N
        df_reativacoes = pd.DataFrame(columns=colunas_letras, index=range(len(df_proto_raw)))

        mapeamento = {
            'AK': 'B', 'H': 'C', 'I': 'D', 'J': 'E',
            'K': 'F', 'P': 'G', 'AO': 'H', 'AU': 'I',
            'AV': 'J', 'AS': 'K', 'AQ': 'L', 'AL': 'N'
        }

        # Preenchimento total dos dados
        for de, para in mapeamento.items():
            try:
                idx_origem = column_index_from_string(de) - 1
                idx_destino = column_index_from_string(para) - 1
                
                if idx_origem < len(df_proto_raw.columns):
                    # Usamos .values para garantir a cópia de toda a coluna
                    df_reativacoes.iloc[:, idx_destino] = df_proto_raw.iloc[:, idx_origem].values
            except Exception as e:
                continue

        # --- 3. GERAÇÃO DO ARQUIVO FINAL ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_ativacoes.to_excel(writer, index=False, sheet_name='ATIVAÇÕES')
            df_reativacoes.to_excel(writer, index=False, sheet_name='REATIVAÇÕES')

            amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            fonte_bold = Font(bold=True, name='Calibri')

            for aba in ['ATIVAÇÕES', 'REATIVAÇÕES']:
                ws = writer.sheets[aba]
                for col_idx in range(1, ws.max_column + 1):
                    header = ws.cell(row=1, column=col_idx)
                    header.fill = amarelo
                    header.font = fonte_bold
                    header.alignment = Alignment(horizontal='center')
                    ws.column_dimensions[header.column_letter].width = 22

        st.success(f"✅ Processamento concluído! Ativações: {len(df_ativacoes)} linhas | Reativações: {len(df_reativacoes)} linhas.")

        # Botão de download
        st.download_button(
            label="📥 Baixar Arquivo Consolidado Completo",
            data=output.getvalue(),
            file_name="CONSOLIDADO_COMPLETO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
