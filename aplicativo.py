import sys
from types import ModuleType

# Correção técnica para compatibilidade
if 'imghdr' not in sys.modules:
    sys.modules['imghdr'] = ModuleType('imghdr')

import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from io import BytesIO

st.set_page_config(page_title="Netmania Optimizer", layout="wide")
st.title("📊 Estruturador de Planilhas Personalizado - Etapa 3")

# --- FUNÇÃO DE CARREGAMENTO INTELIGENTE ---
def carregar_dados_flexivel(arq, sem_header=False):
    if arq is None: return None
    header_val = None if sem_header else 0
    try:
        if arq.name.lower().endswith('.csv'):
            return pd.read_csv(arq, sep=None, engine='python', encoding='latin-1', header=header_val)
        return pd.read_excel(arq, header=header_val)
    except Exception as e:
        st.error(f"Erro ao ler arquivo {arq.name}: {e}")
        return None

# --- FUNÇÃO PARA LIMPAR DATAS (REMOVE HORÁRIO) ---
def formatar_apenas_data(valor):
    try:
        if pd.isna(valor) or valor == "": return ""
        # Converte para datetime e retorna apenas a data formatada
        dt = pd.to_datetime(valor)
        return dt.strftime('%d/%m/%Y')
    except:
        return str(valor)

# --- SEÇÃO DE UPLOAD ---
col1, col2, col3 = st.columns(3)
with col1: arquivo_ativacao = st.file_uploader("1. Planilha de Ativação", type=['xlsx', 'csv', 'xlsm'])
with col2: arquivo_protocolos = st.file_uploader("2. Planilha de Protocolos", type=['xlsx', 'csv', 'xlsm'])
with col3: arquivo_reativacao = st.file_uploader("3. Relatório de Reativações", type=['xlsx', 'csv', 'xlsm'])

if arquivo_ativacao and arquivo_protocolos:
    try:
        df_ativ = carregar_dados_flexivel(arquivo_ativacao)
        df_prot = carregar_dados_flexivel(arquivo_protocolos)

        if df_ativ is not None and df_prot is not None:
            df_ativ.columns = [str(c).strip() for c in df_ativ.columns]
            df_prot.columns = [str(c).strip() for c in df_prot.columns]

            # Filtro de cancelados
            if 'Status Contrato' in df_ativ.columns:
                df_ativ = df_ativ[df_ativ['Status Contrato'].astype(str).str.lower() != 'cancelado']

            # Vínculo Ativação + Protocolos
            col_vinc_ativ = 'Nome Cliente' if 'Nome Cliente' in df_ativ.columns else df_ativ.columns[6]
            col_vinc_prot = 'Cliente' if 'Cliente' in df_prot.columns else df_prot.columns[15]

            df_ativ['_JOIN'] = df_ativ[col_vinc_ativ].astype(str).str.strip().str.upper()
            df_prot['_JOIN'] = df_prot[col_vinc_prot].astype(str).str.strip().str.upper()

            col_resp = 'Responsavel' if 'Responsavel' in df_prot.columns else df_prot.columns[4]
            df_prot_min = df_prot.drop_duplicates(subset=['_JOIN'])[['_JOIN', col_resp]]
            
            df_base = pd.merge(df_ativ, df_prot_min, on='_JOIN', how='left')
            
            if 'Responsavel' in df_base.columns and 'Vendedor 1' in df_base.columns:
                df_base['Responsavel'] = df_base['Responsavel'].fillna(df_base['Vendedor 1'])

            # --- PROCESSAMENTO MANUAL DA REATIVAÇÃO ---
            if arquivo_reativacao:
                df_reat_raw = carregar_dados_flexivel(arquivo_reativacao, sem_header=True)
                
                if df_reat_raw is not None:
                    reat_rows = []
                    for _, row in df_reat_raw.iloc[1:].iterrows():
                        try:
                            # Aplicando formatar_apenas_data nas colunas de data (AJ=35, AM=38, I=8, G=6, K=10)
                            reat_rows.append({
                                'Codigo Cliente': "REATIVAÇÃO",
                                'Contrato': row[35],
                                'Data Contrato': formatar_apenas_data(row[38]),
                                'Prazo Ativacao Contrato': formatar_apenas_data(row[8]),
                                'Ativacao Contrato': formatar_apenas_data(row[6]),
                                'Ativacao Conexao': formatar_apenas_data(row[10]),
                                'Nome Cliente': row[15],
                                'Responsavel': row[3],
                                'Vendedor 1': row[40],
                                'Endereco Ativacao': row[46],
                                'CEP': row[47],
                                'Cidade': row[44],
                                'Servico Ativado': row[42],
                                'Val Serv Ativado': "REATIVAÇÃO",
                                'Status Contrato': row[37],
                                'Assinatura Contrato': "",
                                'Vendedor 2': "",
                                'Origem': "REATIVAÇÃO",
                                'Valor Primeira Mensalidade': "REATIVAÇÃO"
                            })
                        except: continue
                    
                    df_reat_final = pd.DataFrame(reat_rows)
                    df_final_consolidado = pd.concat([df_base, df_reat_final], ignore_index=True)
                else: df_final_consolidado = df_base
            else: df_final_consolidado = df_base

            # --- FORMATAÇÃO FINAL ---
            colunas_finais = [
                'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
                'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
                'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
                'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
                'Origem', 'Valor Primeira Mensalidade'
            ]
            
            df_export = df_final_consolidado[[c for c in colunas_finais if c in df_final_consolidado.columns]]
            
            # Garante que as datas da planilha 1 também fiquem sem horário
            datas_col = ['Data Contrato', 'Prazo Ativacao Contrato', 'Ativacao Contrato', 'Ativacao Conexao']
            for col in datas_col:
                if col in df_export.columns:
                    df_export[col] = df_export[col].apply(formatar_apenas_data)

            st.dataframe(df_export, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Netmania')
                ws = writer.sheets['Netmania']
                
                # Estilos de Cores e Alinhamento
                amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
                centro = Alignment(horizontal='center', vertical='center')
                vazio = Border(left=Side(style=None), right=Side(style=None), top=Side(style=None), bottom=Side(style=None))

                for col_idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=col_idx)
                    # Regras de cores
                    if header.value == "Status Contrato": header.fill = amarelo
                    elif col_idx in [5, 15] or col_idx > 15: header.fill = verde
                    elif col_idx <= 9: header.fill = amarelo
                    
                    for cell in col_cells:
                        cell.alignment = centro
                        cell.border = vazio
                    ws.column_dimensions[header.column_letter].width = 25

            st.success("✅ Processamento concluído: Datas formatadas sem horário.")
            st.download_button("📥 Baixar Planilha Final", output.getvalue(), "NETMANIA_CONSOLIDADO.xlsx")

    except Exception as e:
        st.error(f"Erro geral: {e}")

# --- TUTORIAL ---
st.divider()
st.subheader("📖 Guia Etapa 3")
st.markdown("""
* **Datas:** Todas as colunas de data foram configuradas para exibir apenas **DIA/MÊS/ANO**.
* **Mapeamento:** Os dados de REATIVAÇÃO são inseridos abaixo dos dados de ATIVAÇÃO.
* **Formatos:** O sistema aceita CSV e Excel simultaneamente.
""")
