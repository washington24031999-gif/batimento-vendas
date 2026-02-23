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

st.set_page_config(page_title="Netmania Super Optimizer", layout="wide")
st.title("📊 Consolidador de Planilhas (Base + Protocolos)")

# --- ÁREA DE UPLOAD ---
col1, col2 = st.columns(2)
with col1:
    arquivo_base = st.file_uploader("1. Selecione a BASE PRINCIPAL", type=['xlsx', 'csv', 'xlsm'])
with col2:
    arquivo_proto = st.file_uploader("2. Selecione o RELATÓRIO DE PROTOCOLOS", type=['xlsx', 'xlsm'])

if arquivo_base and arquivo_proto:
    try:
        # 1. Leitura da Base Principal
        if arquivo_base.name.lower().endswith('.csv'):
            df_base = pd.read_csv(arquivo_base, sep=None, engine='python', encoding='latin-1')
        else:
            df_base = pd.read_excel(arquivo_base)
        
        df_base.columns = [str(c).strip() for c in df_base.columns]

        # Filtro de Status (Remove cancelados da base)
        if 'Status Contrato' in df_base.columns:
            df_base = df_base[df_base['Status Contrato'].str.lower() != 'cancelado']

        # 2. Leitura do Relatório de Protocolos
        df_proto_raw = pd.read_excel(arquivo_proto)

        # 3. CRIAÇÃO DA PLANILHA ÚNICA CONSOLIDADA
        # Criamos um DataFrame vazio com as colunas de A até N (14 colunas iniciais)
        colunas_letras = [chr(65 + i) for i in range(14)] # A até N
        df_consolidado = pd.DataFrame(columns=colunas_letras, index=range(len(df_proto_raw)))

        # MAPEAMENTO FIXO (Relatório Protocolos -> Colunas da Planilha Final)
        # De (Protocolo) -> Para (Letra na Final)
        mapeamento = {
            'AK': 'B', 'H': 'C', 'I': 'D', 'J': 'E',
            'K': 'F', 'P': 'G', 'AO': 'H', 'AU': 'I',
            'AV': 'J', 'AS': 'K', 'AQ': 'L', 'AL': 'N'
        }

        for de, para in mapeamento.items():
            try:
                idx_origem = column_index_from_string(de) - 1
                idx_destino = column_index_from_string(para) - 1
                
                if idx_origem < len(df_proto_raw.columns):
                    df_consolidado.iloc[:, idx_destino] = df_proto_raw.iloc[:, idx_origem]
            except:
                continue

        # 4. INTEGRAÇÃO: Adicionamos as colunas da Base Principal ao lado do mapeamento
        # Isso gera uma única tabela larga contendo tudo
        df_final = pd.concat([df_consolidado.reset_index(drop=True), df_base.reset_index(drop=True)], axis=1)

        # --- EXPORTAÇÃO ---
        st.subheader("✅ Tudo pronto!")
        st.write(f"Planilha consolidada com {len(df_final)} linhas.")
        
        st.dataframe(df_final.head(10))

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Planilha_Consolidada')
            ws = writer.sheets['Planilha_Consolidada']
            
            # Formatação
            amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            fonte_cabecalho = Font(bold=True, name='Calibri')
            
            for col_idx in range(1, len(df_final.columns) + 1):
                header = ws.cell(row=1, column=col_idx)
                header.fill = amarelo
                header.font = fonte_cabecalho
                header.alignment = Alignment(horizontal='center')
                ws.column_dimensions[header.column_letter].width = 20

        st.download_button(
            label="📥 Baixar Planilha Única Consolidada",
            data=output.getvalue(),
            file_name="CONSOLIDADO_NETMANIA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao consolidar os dados: {e}")
else:
    st.info("💡 Para gerar a planilha única, suba os dois arquivos acima.")
