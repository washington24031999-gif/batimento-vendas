import sys
from types import ModuleType

# Correção técnica para compatibilidade com Python 3.13
if 'imghdr' not in sys.modules:
    sys.modules['imghdr'] = ModuleType('imghdr')

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from io import BytesIO

st.set_page_config(page_title="Netmania Mapper Automático", layout="wide")

st.title("🚀 Processador DE/PARA Automático")
st.markdown("O sistema seguirá rigorosamente o mapeamento das colunas AK, H, I, J, etc., para as colunas B, C, D, E...")

# --- ÁREA DE UPLOAD ---
col1, col2 = st.columns(2)
with col1:
    arquivo_proto = st.file_uploader("1. ORIGEM: Relatório de Protocolos", type=['xlsx', 'xlsm'])
with col2:
    arquivo_result = st.file_uploader("2. DESTINO: Arquivo Resultante (Modelo)", type=['xlsx', 'xlsm'])

if arquivo_proto and arquivo_result:
    if st.button("🔥 Iniciar Mapeamento e Gerar Arquivo", use_container_width=True):
        try:
            # 1. Carregar os DataFrames
            # Usamos o openpyxl para ler as letras das colunas corretamente
            df_origem = pd.read_excel(arquivo_proto)
            df_modelo = pd.read_excel(arquivo_result)
            
            # Criamos um DataFrame novo com a mesma estrutura do modelo, mas vazio
            df_final = pd.DataFrame(columns=df_modelo.columns, index=range(len(df_origem)))

            # 2. Dicionário de Mapeamento (DE: PARA)
            # Formato: 'Letra_Origem': 'Letra_Destino'
            mapeamento_letras = {
                'AK': 'B', 'H': 'C', 'I': 'D', 'J': 'E',
                'K': 'F', 'P': 'G', 'AO': 'H', 'AU': 'I',
                'AV': 'J', 'AS': 'K', 'AQ': 'L', 'AL': 'N'
            }

            def get_col_by_letter(df, letter):
                """Retorna a série da coluna baseada na letra do Excel"""
                idx = column_index_from_string(letter) - 1
                return df.iloc[:, idx] if idx < len(df.columns) else None

            # 3. Execução do De/Para
            for de, para in mapeamento_letras.items():
                try:
                    col_data = get_col_by_letter(df_origem, de)
                    dest_idx = column_index_from_string(para) - 1
                    
                    if col_data is not None and dest_idx < len(df_final.columns):
                        df_final.iloc[:, dest_idx] = col_data.values
                except Exception as e:
                    st.warning(f"Não foi possível mapear {de} -> {para}: {e}")

            # 4. Filtro de Status (Regra original mantida)
            # Busca a coluna 'Status Contrato' no resultado final caso ela exista
            if 'Status Contrato' in df_final.columns:
                df_final = df_final[df_final['Status Contrato'].astype(str).str.lower() != 'cancelado']

            # 5. Exportação
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.success("✅ Mapeamento fixo concluído!")
            st.dataframe(df_final.head(10))

            st.download_button(
                label="📥 Baixar Arquivo Resultante",
                data=output.getvalue(),
                file_name="RESULTADO_MAPEADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro crítico: {e}. Certifique-se de que os arquivos têm as colunas mencionadas.")
else:
    st.info("Aguardando upload dos arquivos para aplicar as regras de colunas (AK->B, H->C, etc).")
