import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font
from io import BytesIO

# Configurações da página para ocupar a tela toda
st.set_page_config(page_title="Netmania Optimizer", layout="wide")
st.title("📊 Estruturador de Planilhas Netmania")
st.markdown("Arraste sua planilha para formatar conforme o padrão visual.")

# Seletor de arquivos no navegador
arquivo = st.file_uploader("Selecione o arquivo (Excel ou CSV)", type=['xlsx', 'csv', 'xlsm'])

if arquivo:
    try:
        # 1. Leitura do arquivo (identifica se é CSV ou Excel)
        if arquivo.name.lower().endswith('.csv'):
            df = pd.read_csv(arquivo, sep=None, engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(arquivo)
            
        df.columns = [str(c).strip() for c in df.columns]

        # 2. Remover linhas onde o Status Contrato é "cancelado"
        if 'Status Contrato' in df.columns:
            df = df[df['Status Contrato'].str.lower() != 'cancelado']

        # 3. Organizar colunas na ordem exata das imagens
        ordem = [
            'Codigo Cliente', 'Contrato', 'Data Contrato', 'Prazo Ativacao Contrato', 
            'Ativacao Contrato', 'Ativacao Conexao', 'Nome Cliente', 'Responsavel', 
            'Vendedor 1', 'Endereco Ativacao', 'CEP', 'Cidade', 'Servico Ativado', 
            'Val Serv Ativado', 'Status Contrato', 'Assinatura Contrato', 'Vendedor 2', 
            'Origem', 'Valor Primeira Mensalidade'
        ]
        
        # Filtrar apenas as colunas que existem no arquivo enviado
        colunas_finais = [c for c in ordem if c in df.columns]
        df = df[colunas_finais]

        # 4. Formatação Visual (Cores e Fonte)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Planilha Organizada')
            ws = writer.sheets['Planilha Organizada']
            
            # Definição dos estilos
            cor_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cor_verde = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
            # Fonte Calibri 11 sem negrito
            fonte_calibri = Font(name='Calibri', size=11, bold=False)

            for col_idx, col_cells in enumerate(ws.columns, 1):
                header_cell = ws.cell(row=1, column=col_idx)
                nome_coluna = str(header_cell.value).strip()
                
                # Aplicar cores no cabeçalho conforme as fotos
                # Amarelo: Colunas 1 a 9 e a coluna Status Contrato
                if col_idx <= 9 or nome_coluna == "Status Contrato":
                    header_cell.fill = cor_amarelo
                # Verde: Colunas finais (da 16 em diante)
                elif col_idx >= 16:
                    header_cell.fill = cor_verde
                
                # Aplicar Calibri 11 em todas as células da coluna e ajustar largura
                for cell in col_cells:
                    cell.font = fonte_calibri
                
                ws.column_dimensions[header_cell.column_letter].width = 22

        st.success("✅ Planilha processada com sucesso!")
        
        # Botão para o usuário baixar o resultado
        st.download_button(
            label="📥 Baixar Planilha Formatada",
            data=output.getvalue(),
            file_name="PLANILHA_NETMANIA_FINALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")

else:
    st.info("Aguardando o envio de uma planilha...")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

