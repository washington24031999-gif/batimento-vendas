import sys
from types import ModuleType

# Correção para o erro 'imghdr' no Python 3.13
if 'imghdr' not in sys.modules:
    sys.modules['imghdr'] = ModuleType('imghdr')

import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font
from io import BytesIO

st.set_page_config(page_title="Netmania Optimizer", layout="wide")
st.title("📊 Estruturador de Planilhas Netmania")

arquivo = st.file_uploader("Selecione o arquivo (Excel ou CSV)", type=['xlsx', 'csv', 'xlsm'])

if arquivo:
    try:
        if arquivo.name.lower().endswith('.csv'):
            df = pd.read_csv(arquivo, sep=None, engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(arquivo)
            
        df.columns = [str(c).strip() for c in df.columns]

        if 'Status Contrato' in df.columns:
            df = df[df['Status Contrato'].str.lower() != 'cancelado']

        ordem = [
            'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
            'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
            'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
            'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
            'Origem', 'Valor Primeira Mensalidade'
        ]
        
        df = df[[c for c in ordem if c in df.columns]]

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Planilha Organizada')
            ws = writer.sheets['Planilha Organizada']
            
            cor_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cor_verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
            fonte_calibri = Font(name='Calibri', size=11, bold=False)

            for col_idx, col_cells in enumerate(ws.columns, 1):
                header_cell = ws.cell(row=1, column=col_idx)
                nome_coluna = str(header_cell.value).strip()
                
                if col_idx <= 9 or nome_coluna == "Status Contrato":
                    header_cell.fill = cor_amarelo
                elif col_idx >= 16:
                    header_cell.fill = cor_verde
                
                for cell in col_cells:
                    cell.font = fonte_calibri
                
                ws.column_dimensions[header_cell.column_letter].width = 22

        st.success("✅ Planilha processada!")
        st.download_button(
            label="📥 Baixar Planilha Formatada",
            data=
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")


