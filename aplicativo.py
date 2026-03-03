# --- SEÇÃO DE UPLOAD ---
col1, col2, col3 = st.columns(3)
with col1: arquivo_ativacao = st.file_uploader("1. Planilha de Ativação", type=['xlsx', 'csv', 'xlsm'])
with col2: arquivo_protocolos = st.file_uploader("2. Planilha de Protocolos", type=['xlsx', 'csv', 'xlsm'])
with col3: arquivo_reativacao = st.file_uploader("3. Relatório de Reativações", type=['xlsx', 'csv', 'xlsm'])

if arquivo_ativacao:
    try:
        df_ativ = carregar_dados_flexivel(arquivo_ativacao)
        
        if df_ativ is not None:
            df_ativ.columns = [str(c).strip() for c in df_ativ.columns]

            # Filtro automático de contratos cancelados
            if 'Status Contrato' in df_ativ.columns:
                df_ativ = df_ativ[df_ativ['Status Contrato'].astype(str).str.lower() != 'cancelado']

            # --- LÓGICA DE ALERTAS E PROCESSAMENTO ---
            
            # 1. Tratamento de Protocolos
            if arquivo_protocolos:
                df_prot = carregar_dados_flexivel(arquivo_protocolos)
                # ... (sua lógica de merge aqui)
                df_base = pd.merge(df_ativ, df_prot_min, on='_JOIN', how='left')
                st.sidebar.success("✔️ Protocolos integrados.")
            else:
                df_base = df_ativ.copy()
                if 'Responsavel' not in df_base.columns and 'Vendedor 1' in df_base.columns:
                    df_base['Responsavel'] = df_base['Vendedor 1']
                st.warning("⚠️ **Aviso:** Planilha de Protocolos ausente. A coluna 'Responsável' será preenchida com o 'Vendedor 1'.")

            # 2. Tratamento de Reativações
            if arquivo_reativacao:
                # ... (sua lógica de processamento de reativação aqui)
                df_final_consolidado = pd.concat([df_base, pd.DataFrame(reat_rows)], ignore_index=True)
                st.sidebar.success("✔️ Reativações somadas.")
            else:
                df_final_consolidado = df_base
                st.info("ℹ️ **Nota:** Relatório de Reativações não enviado. O arquivo conterá apenas novas ativações.")

            # --- EXIBIÇÃO E DOWNLOAD ---
            # (Mantém o restante do seu código de formatação Excel)
