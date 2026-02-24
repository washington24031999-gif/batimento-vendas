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

# --- SEÇÃO DE UPLOAD ---
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
                return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1', header=None)
            return pd.read_excel(arq, header=None) # Lemos sem header para trabalhar com índices puros

        # Carregamento da Ativação e Protocolos (com headers normais para o merge inicial)
        df_ativ = pd.read_excel(arquivo_ativacao) if arquivo_ativacao.name.endswith('xlsx') else pd.read_csv(arquivo_ativacao, sep=None, engine='python', encoding='latin-1')
        df_prot = pd.read_excel(arquivo_protocolos) if arquivo_protocolos.name.endswith('xlsx') else pd.read_csv(arquivo_protocolos, sep=None, engine='python', encoding='latin-1')
        
        df_ativ.columns = [str(c).strip() for c in df_ativ.columns]
        df_prot.columns = [str(c).strip() for c in df_prot.columns]

        # 1. Filtro e Merge Base (Ativação + Protocolos)
        if 'Status Contrato' in df_ativ.columns:
            df_ativ = df_ativ[df_ativ['Status Contrato'].astype(str).str.lower() != 'cancelado']

        df_ativ['_JOIN_KEY'] = df_ativ['Nome Cliente'].astype(str).str.strip().str.upper()
        df_prot['_JOIN_KEY'] = df_prot['Cliente'].astype(str).str.strip().str.upper()
        
        df_prot_clean = df_prot.drop_duplicates(subset=['_JOIN_KEY'])[['_JOIN_KEY', 'Responsavel']]
        df_base = pd.merge(df_ativ, df_prot_clean, on='_JOIN_KEY', how='left')
        
        if 'Vendedor 1' in df_base.columns:
            df_base['Responsavel'] = df_base['Responsavel'].fillna(df_base['Vendedor 1'])

        # --- PROCESSAMENTO MANUAL DA PLANILHA DE REATIVAÇÃO ---
        if arquivo_reativacao:
            # Lemos a reativação ignorando nomes de colunas para usar índices (A, B, C...)
            df_reat_raw = pd.read_excel(arquivo_reativacao, header=None)
            
            # Criamos um novo DataFrame seguindo sua estrutura exata (A até S)
            # AJ=35, AM=38, I=8, G=6, K=10, P=15, D=3, AO=40, AU=46, AV=47, AS=44, AQ=42, AL=37
            reat_data = []
            
            for i, row in df_reat_raw.iloc[1:].iterrows(): # Pula o cabeçalho original
                linha = {
                    'Codigo Cliente': "REATIVAÇÃO",
                    'Contrato': row[35],                # Coluna AJ
                    'Data Contrato': row[38],           # Coluna AM
                    'Prazo Ativacao Contrato': row[8],   # Coluna I
                    'Ativacao Contrato': row[6],         # Coluna G
                    'Ativacao Conexao': row[10],        # Coluna K
                    'Nome Cliente': row[15],            # Coluna P
                    'Responsavel': row[3],              # Coluna D
                    'Vendedor 1': row[40],              # Coluna AO
                    'Endereco Ativacao': row[46],       # Coluna AU
                    'CEP': row[47],                     # Coluna AV
                    'Cidade': row[44],                  # Coluna AS
                    'Servico Ativado': row[42],         # Coluna AQ
                    'Val Serv Ativado': "REATIVAÇÃO",    # Coluna N
                    'Status Contrato': row[37],         # Coluna AL (O)
                    'Assinatura Contrato': "",           # Coluna P (Em branco)
                    'Vendedor 2': "",                   # Coluna Q (Em branco)
                    'Origem': "REATIVAÇÃO",             # Coluna R
                    'Valor Primeira Mensalidade': "REATIVAÇÃO" # Coluna S
                }
                reat_data.append(linha)
            
            df_reat_final = pd.DataFrame(reat_data)
            
            # Concatenamos a base original com as novas linhas de reativação
            df_final_all = pd.concat([df_base, df_reat_final], ignore_index=True)
        else:
            df_final_all = df_base

        # --- FORMATAÇÃO E DOWNLOAD ---
        ordem_final = [
            'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
            'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
            'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
            'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
            'Origem', 'Valor Primeira Mensalidade'
        ]
        
        df_output = df_final_all[[c for c in ordem_final if c in df_final_all.columns]]
        st.dataframe(df_output, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_output.to_excel(writer, index=False, sheet_name='Planilha_Final')
            ws = writer.sheets['Planilha_Final']
            
            amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
            centralizado = Alignment(horizontal='center', vertical='center')
            sem_bordas = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

            for col_idx, col_cells in enumerate(ws.columns, 1):
                header = ws.cell(row=1, column=col_idx)
                # Aplicar cores conforme regras anteriores
                if header.value == "Status Contrato": header.fill = amarelo
                elif col_idx in [5, 15] or col_idx > 15: header.fill = verde
                elif col_idx <= 9: header.fill = amarelo
                
                for cell in col_cells:
                    cell.alignment = centralizado
                    cell.border = sem_bordas
                ws.column_dimensions[header.column_letter].width = 25

        st.download_button("📥 Baixar Planilha Consolidada", output.getvalue(), "NETMANIA_REATIVACAO_MANUAL.xlsx")

    except Exception as e:
        st.error(f"Erro ao processar mapeamento manual: {e}")

# --- TUTORIAL NO RODAPÉ ---
st.divider()
st.subheader("📖 Guia de Mapeamento Manual")
st.info("""
**Como a Etapa 3 funciona agora:**
1. O sistema processa a Planilha 1 e 2 normalmente.
2. Ele lê a Planilha 3 (Reativações) e extrai os dados das colunas específicas (AJ, AM, I, G, etc.) para criar novas linhas.
3. As colunas P, Q ficam vazias e as colunas R, S e A são preenchidas com o texto 'REATIVAÇÃO'.
4. Tudo é unificado em um único arquivo final centralizado.
""")
