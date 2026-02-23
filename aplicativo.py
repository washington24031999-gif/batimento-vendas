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

arquivo = st.file_uploader("Selecione o arquivo (Excel ou CSV)", type=['xlsx', 'csv', 'xlsm'])

if arquivo:
    try:
        # 1. Leitura
        if arquivo.name.lower().endswith('.csv'):
            df = pd.read_csv(arquivo, sep=None, engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(arquivo)
            
        df.columns = [str(c).strip() for c in df.columns]

        # 2. Filtro de Status
        if 'Status Contrato' in df.columns:
            df = df[df['Status Contrato'].str.lower() != 'cancelado']

        # --- SEÇÃO DE PERSONALIZAÇÃO ---
        st.subheader("⚙️ Personalize sua exportação")
        
        # Ordem padrão sugerida
        ordem_padrao = [
            'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
            'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
            'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
            'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
            'Origem', 'Valor Primeira Mensalidade'
        ]
        
        # Identifica quais colunas da ordem padrão existem no arquivo e quais outras extras existem
        colunas_disponiveis = list(df.columns)
        selecao_inicial = [c for c in ordem_padrao if c in colunas_disponiveis]

        # Caixa de seleção múltipla
        colunas_selecionadas = st.multiselect(
            "Selecione e ordene as colunas que deseja exportar:",
            options=colunas_disponiveis,
            default=selecao_inicial
        )

        if not colunas_selecionadas:
            st.warning("⚠️ Selecione pelo menos uma coluna para exportar.")
        else:
            # 3. Filtrar o DataFrame com a seleção do usuário
            df_final = df[colunas_selecionadas]
            
            # Visualização prévia no site
            st.write("Prévia dos dados (10 primeiras linhas):")
            st.dataframe(df_final.head(10))

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
                    
                    # Mantém a lógica de cores (se a coluna existir na nova planilha)
                    # Amarelo para as primeiras 9 ou Status
                    if col_idx <= 9 or nome == "Status Contrato":
                        header.fill = amarelo
                    # Verde se for uma das últimas 4 colunas exportadas
                    elif col_idx > len(colunas_selecionadas) - 4:
                        header.fill = verde
                    
                    for cell in col_cells:
                        cell.font = fonte
                    ws.column_dimensions[header.column_letter].width = 22

            st.success(f"✅ Planilha com {len(colunas_selecionadas)} colunas pronta!")
            st.download_button(
                label="📥 Baixar Planilha Personalizada",
                data=output.getvalue(),
                file_name="PLANILHA_PERSONALIZADA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
