import streamlit as st
import pandas as pd
import numpy as np
import io
from utils.calculations import ForestryCalculator
from utils.statistics import StatisticsAnalyzer
from utils.report_generator import ReportGenerator

def calculate_species_volume_summary(results_df, project_info):
    """
    Calculate volume summary by species.
    
    Args:
        results_df (pandas.DataFrame): Processed data with calculations
        project_info (dict): Project information with area data
        
    Returns:
        pandas.DataFrame: Species volume summary table
    """
    # Identificar coluna de espécie (nome comum ou científico)
    species_column = None
    for col in results_df.columns:
        col_upper = str(col).upper()
        if any(keyword in col_upper for keyword in ['NOME COMUM', 'ESPÉCIE', 'SPECIES', 'NOME CIENTÍFICO']):
            species_column = col
            break
    
    if not species_column or results_df[species_column].isna().all():
        return pd.DataFrame()  # Retorna DataFrame vazio se não encontrar coluna de espécie
    
    # Calcular métricas por espécie
    plot_area_ha = project_info['plot_area']
    total_area_ha = project_info['total_area']
    
    species_groups = results_df.groupby(species_column).agg({
        'DAP (cm)': ['count', 'mean'],
        'HT (m)': 'mean',
        'VT (m³)': 'sum',
        'VT (m³/ha)': 'sum'
    }).reset_index()
    
    # Simplificar nomes das colunas
    species_groups.columns = [
        'Espécie',
        'n_trees_plot',
        'DAP médio',
        'Altura média', 
        'Soma de VT (m³)',
        'VT (m³)/ha'
    ]
    
    # Calcular n/ha usando a fórmula correta: (quantidade da espécie encontrada) ÷ (Área Total Amostrada) × 1 ha
    # Nota: plot_area_ha é a área total amostrada em hectares
    species_groups['n/ha'] = (species_groups['n_trees_plot'] / plot_area_ha) * 1
    
    # Calcular n total (extrapolação para área total)
    species_groups['n total'] = species_groups['n/ha'] * total_area_ha
    
    # Calcular V/ha(m³)/Área total/ha usando a fórmula especificada: VT(m³/ha) × Área Total a ser Suprimida (ha)
    species_groups['V/ha(m³)/Área total/ha'] = species_groups['VT (m³)/ha'] * total_area_ha
    
    # Reordenar colunas conforme solicitado
    final_columns = [
        'Espécie', 'n/ha', 'n total', 'DAP médio', 
        'Altura média', 'Soma de VT (m³)', 'VT (m³)/ha', 
        'V/ha(m³)/Área total/ha'
    ]
    
    return species_groups[final_columns].sort_values('VT (m³)/ha', ascending=False)

def detect_and_map_columns(df):
    """Detecta e mapeia automaticamente as colunas da planilha"""
    df_original = df.copy()
    df = df.copy()
    df.columns = df.columns.str.strip()
    
    # Contar linhas iniciais
    initial_rows = len(df)
    st.write(f"**Processamento iniciado com {initial_rows} linhas**")
    
    mapping_results = []
    new_columns = []
    used_mappings = set()
    
    for i, col in enumerate(df.columns):
        original_col = str(col).strip()
        col_upper = original_col.upper()
        mapped_name = None
        
        # Detectar coluna de numeração das árvores
        if ('N°' in col_upper or col_upper in ['N', 'NO', 'NUM', 'NUMERO', 'NÚMERO']) and 'Nº da árvore' not in used_mappings:
            mapped_name = 'Nº da árvore'
            used_mappings.add('Nº da árvore')
        
        # Detectar nome comum
        elif 'NOME' in col_upper and 'COMUM' in col_upper and 'Nome comum' not in used_mappings:
            mapped_name = 'Nome comum'
            used_mappings.add('Nome comum')
        
        # Detectar nome científico
        elif 'NOME' in col_upper and ('CIENTÍFICO' in col_upper or 'CIENTIFICO' in col_upper) and 'Nome científico' not in used_mappings:
            mapped_name = 'Nome científico'
            used_mappings.add('Nome científico')
        
        # Detectar CAP
        elif 'CAP' in col_upper and 'CAP (cm)' not in used_mappings:
            mapped_name = 'CAP (cm)'
            used_mappings.add('CAP (cm)')
        
        # Detectar altura (HT)
        elif ('HT' in col_upper or 'ALTURA' in col_upper) and 'HT (m)' not in used_mappings:
            mapped_name = 'HT (m)'
            used_mappings.add('HT (m)')
        
        # Se não mapear ou já estiver usado, manter nome original
        if mapped_name is None:
            # Garantir nome único
            base_name = original_col
            counter = 1
            while base_name in new_columns:
                base_name = f"{original_col}_{counter}"
                counter += 1
            mapped_name = base_name
            mapping_results.append(f"• '{original_col}' → mantido como '{mapped_name}'")
        else:
            mapping_results.append(f"✓ '{original_col}' → '{mapped_name}'")
        
        new_columns.append(mapped_name)
    
    # Aplicar novos nomes das colunas
    df.columns = new_columns
    st.write(f"**Após renomear colunas: {len(df)} linhas**")
    
    # Criar coluna combinada de nomes se necessário
    if 'Nome comum' in df.columns and 'Nome científico' in df.columns:
        df['Nome comum/científico'] = df['Nome comum'].astype(str) + ' / ' + df['Nome científico'].astype(str)
        mapping_results.append("✓ Combinadas colunas de nomes")
    elif 'Nome comum' in df.columns:
        df['Nome comum/científico'] = df['Nome comum']
        mapping_results.append("✓ Usando nome comum como identificação")
    elif 'Nome científico' in df.columns:
        df['Nome comum/científico'] = df['Nome científico']
        mapping_results.append("✓ Usando nome científico como identificação")
    
    final_rows = len(df)
    st.write(f"**Processamento finalizado com {final_rows} linhas**")
    
    # Verificar se alguma linha foi perdida
    if final_rows < initial_rows:
        st.error(f"❌ PERDA DE DADOS: {initial_rows - final_rows} linhas foram removidas durante o mapeamento!")
    
    # Mostrar resultados do mapeamento
    st.success("Mapeamento de colunas realizado:")
    for result in mapping_results:
        st.write(result)
    
    # Mostrar colunas finais
    st.info(f"Colunas finais: {list(df.columns)}")
    
    return df

def main():
    st.set_page_config(
        page_title="Sistema de Inventário Florestal",
        page_icon="🌳",
        layout="wide"
    )
    
    st.title("🌳 Sistema de Inventário Florestal")
    st.markdown("Sistema para processamento e análise de inventários florestais")
    
    # Initialize session state
    if 'data_processed' not in st.session_state:
        st.session_state.data_processed = False
    if 'results_df' not in st.session_state:
        st.session_state.results_df = None
    if 'statistics' not in st.session_state:
        st.session_state.statistics = None
    if 'project_info' not in st.session_state:
        st.session_state.project_info = None
    if 'file_uploaded' not in st.session_state:
        st.session_state.file_uploaded = False
    if 'input_data' not in st.session_state:
        st.session_state.input_data = None
    
    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📂 Upload de Dados", "⚙️ Processamento", "📊 Estatísticas", "📑 Relatório"])
    
    with tab1:
        upload_data_tab()
    
    with tab2:
        processing_tab()
    
    with tab3:
        statistics_tab()
    
    with tab4:
        report_tab()

def upload_data_tab():
    st.header("📂 Upload de Dados de Campo")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("Informações do Projeto")
        project_name = st.text_input("Nome do Projeto*", key="project_name")
        num_plots = st.number_input("Quantidade de Parcelas*", min_value=1, value=1, key="num_plots")
        
        st.subheader("Dimensões da Parcela")
        plot_length = st.number_input("Comprimento da Parcela (m)*", min_value=0.1, value=20.0, step=0.1, key="plot_length")
        plot_width = st.number_input("Largura da Parcela (m)*", min_value=0.1, value=20.0, step=0.1, key="plot_width")
        
        total_area = st.number_input("Área Total a ser Suprimida (ha)*", min_value=0.01, value=1.0, step=0.01, key="total_area")
        form_factor = st.number_input("Fator de Forma (FF)*", min_value=0.1, max_value=1.0, value=0.7, step=0.01, key="form_factor")
    
    with col2:
        st.subheader("Upload de Planilha")
        uploaded_file = st.file_uploader(
            "Selecione a planilha com dados de campo",
            type=['csv', 'xlsx'],
            help="A planilha deve conter as colunas: UA, N°, NOME COMUM, NOME CIENTÍFICO, CAP (cm), HT(m)"
        )
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.success(f"Arquivo carregado com sucesso! {len(df)} registros encontrados na planilha original.")
                st.dataframe(df.head())
                
                # Mostrar informações detalhadas sobre os dados originais
                st.info(f"📊 Dados originais: {len(df)} linhas, {len(df.columns)} colunas")
                
                # Verificar especificamente a coluna N° para contar árvores
                tree_count_analysis = []
                for col in df.columns:
                    if 'N°' in str(col).upper() or 'N' == str(col).upper():
                        unique_trees = df[col].nunique()
                        valid_entries = df[col].notna().sum()
                        tree_count_analysis.append(f"Coluna '{col}': {valid_entries} entradas válidas, {unique_trees} árvores únicas")
                        
                        # Verificar se há números duplicados
                        duplicates = df[col].duplicated().sum()
                        if duplicates > 0:
                            tree_count_analysis.append(f"  ⚠️ {duplicates} números duplicados encontrados")
                        
                        # Verificar valores vazios
                        empty_values = df[col].isna().sum()
                        if empty_values > 0:
                            tree_count_analysis.append(f"  ⚠️ {empty_values} valores vazios encontrados")
                
                if tree_count_analysis:
                    st.write("**Análise da contagem de árvores:**")
                    for analysis in tree_count_analysis:
                        st.write(analysis)
                
                # Detectar automaticamente as colunas
                df_processed = detect_and_map_columns(df)
                
                # Verificar se perdemos dados durante o processamento
                if len(df_processed) < len(df):
                    lost_rows = len(df) - len(df_processed)
                    st.warning(f"⚠️ Atenção: {lost_rows} linhas foram removidas durante o processamento (provavelmente linhas vazias ou com dados inválidos)")
                    
                    # Mostrar quais linhas foram removidas
                    st.write("**Análise de dados removidos:**")
                    original_indices = set(df.index)
                    processed_indices = set(df_processed.index)
                    removed_indices = original_indices - processed_indices
                    
                    if removed_indices:
                        st.write(f"Linhas removidas: {sorted(list(removed_indices))}")
                
                # Salvar dados automaticamente
                st.session_state.input_data = df_processed
                st.session_state.file_uploaded = True
                st.success("✅ Planilha carregada e dados salvos automaticamente!")
                st.info(f"📊 {len(df_processed)} árvores válidas detectadas e prontas para processamento")
                
                # Mostrar preview dos dados processados
                st.subheader("Preview dos Dados Processados")
                st.dataframe(df_processed.head(), use_container_width=True)
                
                # Mostrar estatísticas dos dados
                st.subheader("Resumo dos Dados")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Original", len(df))
                with col2:
                    st.metric("Total Processado", len(df_processed))
                with col3:
                    if len(df_processed) < len(df):
                        st.metric("Linhas Removidas", len(df) - len(df_processed), delta=-(len(df) - len(df_processed)))
                    else:
                        st.metric("Linhas Removidas", 0)
                    
            except Exception as e:
                st.error(f"Erro ao carregar o arquivo: {str(e)}")
    
    # Process data button
    if st.button("Processar Dados", type="primary"):
        # Validate all required fields
        errors = []
        if not project_name:
            errors.append("Nome do Projeto é obrigatório")
        if num_plots < 1:
            errors.append("Quantidade de Parcelas deve ser pelo menos 1")
        if plot_length <= 0 or plot_width <= 0:
            errors.append("Dimensões da parcela devem ser positivas")
        if total_area <= 0:
            errors.append("Área total deve ser positiva")
        if not (0.1 <= form_factor <= 1.0):
            errors.append("Fator de Forma deve estar entre 0.1 e 1.0")
        if not st.session_state.file_uploaded or st.session_state.input_data is None:
            errors.append("Nenhum arquivo foi carregado")
        
        if errors:
            for error in errors:
                st.error(f"⚠️ {error}")
        else:
            # Store project information
            plot_area = plot_length * plot_width / 10000  # Convert to hectares
            total_sampled_area = plot_area * num_plots
            sampling_percentage = (total_sampled_area / total_area) * 100
            
            project_info = {
                'project_name': project_name,
                'num_plots': num_plots,
                'plot_length': plot_length,
                'plot_width': plot_width,
                'plot_area': plot_area,
                'total_area': total_area,
                'total_sampled_area': total_sampled_area,
                'sampling_percentage': sampling_percentage,
                'form_factor': form_factor
            }
            
            st.session_state.project_info = project_info
            
            # Process calculations
            calculator = ForestryCalculator()
            results_df = calculator.process_data(st.session_state.input_data, form_factor, plot_area)
            st.session_state.results_df = results_df
            
            # Calculate statistics
            analyzer = StatisticsAnalyzer()
            statistics = analyzer.calculate_statistics(results_df, project_info)
            st.session_state.statistics = statistics
            
            st.session_state.data_processed = True
            st.success("✅ Dados processados com sucesso!")
            st.rerun()

def processing_tab():
    st.header("⚙️ Processamento dos Dados")
    
    if not st.session_state.data_processed:
        st.warning("⚠️ Primeiro faça o upload e processamento dos dados na aba 'Upload de Dados'")
        return
    
    st.subheader("Informações do Projeto")
    project_info = st.session_state.project_info
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Projeto", project_info['project_name'])
        st.metric("Parcelas", project_info['num_plots'])
    with col2:
        st.metric("Área por Parcela", f"{project_info['plot_area']:.4f} ha")
        st.metric("Área Total Amostrada", f"{project_info['total_sampled_area']:.4f} ha")
    with col3:
        st.metric("Área de Supressão", f"{project_info['total_area']:.2f} ha")
        st.metric("% Amostrada", f"{project_info['sampling_percentage']:.2f}%")
    
    st.subheader("Cálculos por Árvore")
    results_df = st.session_state.results_df
    
    # Display results table
    st.dataframe(
        results_df.style.format({
            'CAP (cm)': '{:.2f}',
            'HT (m)': '{:.2f}',
            'DAP (cm)': '{:.4f}',
            'VT (m³)': '{:.4f}',
            'VT (m³/ha)': '{:.4f}',
            'VT (st/ha)': '{:.4f}'
        }),
        use_container_width=True
    )
    
    # Summary statistics
    st.subheader("Resumo dos Cálculos")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total de Árvores", len(results_df))
        st.metric("Volume Total (m³)", f"{results_df['VT (m³)'].sum():.4f}")
    
    with col2:
        st.metric("Volume Médio (m³/ha)", f"{results_df['VT (m³/ha)'].mean():.4f}")
        st.metric("Volume Total (m³/ha)", f"{results_df['VT (m³/ha)'].sum():.4f}")
    
    with col3:
        st.metric("Volume Médio (st/ha)", f"{results_df['VT (st/ha)'].mean():.4f}")
        st.metric("Volume Total (st/ha)", f"{results_df['VT (st/ha)'].sum():.4f}")
    
    with col4:
        st.metric("DAP Médio (cm)", f"{results_df['DAP (cm)'].mean():.4f}")
        st.metric("Altura Média (m)", f"{results_df['HT (m)'].mean():.2f}")

    # Volume médio por espécie
    st.subheader("Volume Médio por Espécie")
    species_summary = calculate_species_volume_summary(results_df, project_info)
    
    if not species_summary.empty:
        st.dataframe(
            species_summary.style.format({
                'n/ha': '{:.2f}',
                'n total': '{:.0f}',
                'DAP médio': '{:.2f}',
                'Altura média': '{:.2f}',
                'Soma de VT (m³)': '{:.4f}',
                'VT (m³)/ha': '{:.4f}',
                'V/ha(m³)/Área total/ha': '{:.3f}'
            }),
            use_container_width=True
        )
    else:
        st.warning("Dados de espécie não disponíveis. Verifique se a planilha possui colunas de nome comum ou científico.")

def statistics_tab():
    st.header("📊 Estatísticas e Precisão")
    
    if not st.session_state.data_processed:
        st.warning("⚠️ Primeiro faça o upload e processamento dos dados na aba 'Upload de Dados'")
        return
    
    statistics = st.session_state.statistics
    
    # Precision alert
    if statistics['sampling_error'] > 20:
        st.error("⚠️ Amostragem não atingiu a precisão desejada (erro > 20%). Recomendado aumentar o número de parcelas.")
    else:
        st.success("✅ Amostragem atingiu a precisão desejada (erro ≤ 20%).")
    
    # Statistical metrics
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Estatísticas Descritivas")
        st.metric("Média (m³/ha)", f"{statistics['mean']:.4f}")
        st.metric("Variância", f"{statistics['variance']:.4f}")
        st.metric("Desvio Padrão", f"{statistics['std_dev']:.4f}")
        st.metric("Coeficiente de Variação", f"{statistics['cv']:.2f}%")
    
    with col2:
        st.subheader("Precisão da Amostragem")
        st.metric("Erro Amostral", f"{statistics['sampling_error']:.2f}%")
        st.metric("Limite Inferior IC 90%", f"{statistics['ci_lower']:.4f}")
        st.metric("Limite Superior IC 90%", f"{statistics['ci_upper']:.4f}")
        st.metric("Erro Padrão", f"{statistics['standard_error']:.4f}")
    
    # Volume estimates
    st.subheader("Estimativas de Volume para Área Total")
    project_info = st.session_state.project_info
    
    col1, col2, col3 = st.columns(3)
    with col1:
        volume_estimate = statistics['mean'] * project_info['total_area']
        st.metric("Volume Estimado (m³)", f"{volume_estimate:.2f}")
    
    with col2:
        volume_lower = statistics['ci_lower'] * project_info['total_area']
        st.metric("Volume Mínimo IC 90% (m³)", f"{volume_lower:.2f}")
    
    with col3:
        volume_upper = statistics['ci_upper'] * project_info['total_area']
        st.metric("Volume Máximo IC 90% (m³)", f"{volume_upper:.2f}")

def report_tab():
    st.header("📑 Relatório Final")
    
    if not st.session_state.data_processed:
        st.warning("⚠️ Primeiro faça o upload e processamento dos dados na aba 'Upload de Dados'")
        return
    
    # Generate report
    report_generator = ReportGenerator()
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Gerar Relatório Excel", type="primary"):
            excel_buffer = report_generator.generate_excel_report(
                st.session_state.results_df,
                st.session_state.statistics,
                st.session_state.project_info
            )
            
            st.download_button(
                label="⬇️ Download Relatório Excel",
                data=excel_buffer,
                file_name=f"relatorio_inventario_{st.session_state.project_info['project_name'].replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with col2:
        if st.button("📄 Gerar Relatório PDF", type="secondary"):
            pdf_buffer = report_generator.generate_pdf_report(
                st.session_state.results_df,
                st.session_state.statistics,
                st.session_state.project_info
            )
            
            st.download_button(
                label="⬇️ Download Relatório PDF",
                data=pdf_buffer,
                file_name=f"relatorio_inventario_{st.session_state.project_info['project_name'].replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
    
    # Display summary report
    st.subheader("Resumo do Relatório")
    
    project_info = st.session_state.project_info
    statistics = st.session_state.statistics
    results_df = st.session_state.results_df
    
    st.markdown(f"""
    ### 📋 Informações Gerais
    - **Projeto:** {project_info['project_name']}
    - **Número de Parcelas:** {project_info['num_plots']}
    - **Área por Parcela:** {project_info['plot_area']:.4f} ha
    - **Área Total Amostrada:** {project_info['total_sampled_area']:.4f} ha
    - **Área de Supressão:** {project_info['total_area']:.2f} ha
    - **Porcentagem Amostrada:** {project_info['sampling_percentage']:.2f}%
    - **Fator de Forma:** {project_info['form_factor']:.2f}
    
    ### 📊 Resultados Estatísticos
    - **Total de Árvores:** {len(results_df)}
    - **Volume Médio:** {statistics['mean']:.4f} m³/ha
    - **Desvio Padrão:** {statistics['std_dev']:.4f} m³/ha
    - **Erro Amostral:** {statistics['sampling_error']:.2f}%
    - **Intervalo de Confiança 90%:** {statistics['ci_lower']:.4f} - {statistics['ci_upper']:.4f} m³/ha
    
    ### 🎯 Avaliação da Precisão
    """)
    
    if statistics['sampling_error'] > 20:
        st.markdown("🔴 **Amostragem não atingiu a precisão desejada (erro > 20%). Recomendado aumentar o número de parcelas.**")
    else:
        st.markdown("🟢 **Amostragem atingiu a precisão desejada (erro ≤ 20%).**")
    
    volume_estimate = statistics['mean'] * project_info['total_area']
    st.markdown(f"""
    ### 📈 Estimativa Final de Volume
    - **Volume Total Estimado:** {volume_estimate:.2f} m³
    - **Volume Mínimo (IC 90%):** {statistics['ci_lower'] * project_info['total_area']:.2f} m³
    - **Volume Máximo (IC 90%):** {statistics['ci_upper'] * project_info['total_area']:.2f} m³
    """)

if __name__ == "__main__":
    main()
