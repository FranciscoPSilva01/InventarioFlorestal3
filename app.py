import streamlit as st
import pandas as pd
import numpy as np
import io
from utils.calculations import ForestryCalculator
from utils.statistics import StatisticsAnalyzer
from utils.report_generator import ReportGenerator

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
                
                st.success(f"Arquivo carregado com sucesso! {len(df)} registros encontrados.")
                st.dataframe(df.head())
                
                # Clean column names by stripping whitespace and normalizing
                df.columns = df.columns.str.strip()
                
                # Debug: show original columns with their exact representation
                st.info(f"Colunas encontradas: {[repr(col) for col in df.columns]}")
                
                # More comprehensive column mapping including variations
                original_columns = list(df.columns)
                
                # Map common variations
                for i, col in enumerate(df.columns):
                    col_clean = col.strip().upper()
                    
                    # Map tree number column variations
                    if col_clean in ['N°', 'N', 'NO', 'NUM', 'NUMERO', 'NÚMERO']:
                        df = df.rename(columns={col: 'Nº da árvore'})
                        st.success(f"Coluna '{col}' mapeada para 'Nº da árvore'")
                    
                    # Map common name variations
                    elif col_clean in ['NOME COMUM', 'NOME_COMUM', 'COMMON NAME']:
                        df = df.rename(columns={col: 'Nome comum'})
                        st.success(f"Coluna '{col}' mapeada para 'Nome comum'")
                    
                    # Map scientific name variations
                    elif col_clean in ['NOME CIENTÍFICO', 'NOME_CIENTIFICO', 'SCIENTIFIC NAME']:
                        df = df.rename(columns={col: 'Nome científico'})
                        st.success(f"Coluna '{col}' mapeada para 'Nome científico'")
                    
                    # Map CAP variations
                    elif 'CAP' in col_clean and 'CM' in col_clean:
                        df = df.rename(columns={col: 'CAP (cm)'})
                        st.success(f"Coluna '{col}' mapeada para 'CAP (cm)'")
                    
                    # Map height variations
                    elif ('HT' in col_clean or 'ALTURA' in col_clean) and ('M' in col_clean or 'METRO' in col_clean):
                        df = df.rename(columns={col: 'HT (m)'})
                        st.success(f"Coluna '{col}' mapeada para 'HT (m)'")
                
                # Create combined name column if separate common and scientific names exist
                if 'Nome comum' in df.columns and 'Nome científico' in df.columns:
                    df['Nome comum/científico'] = df['Nome comum'].astype(str) + ' / ' + df['Nome científico'].astype(str)
                    st.success("Combinadas colunas de nome comum e científico")
                elif 'Nome comum' in df.columns:
                    df['Nome comum/científico'] = df['Nome comum']
                    st.success("Usando nome comum como identificação da espécie")
                elif 'Nome científico' in df.columns:
                    df['Nome comum/científico'] = df['Nome científico']
                    st.success("Usando nome científico como identificação da espécie")
                
                # Show final columns
                st.info(f"Colunas finais: {list(df.columns)}")
                
                # Validate required columns after mapping
                required_columns = ['Nº da árvore', 'Nome comum/científico', 'CAP (cm)', 'HT (m)']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    st.error(f"Colunas obrigatórias não encontradas: {', '.join(missing_columns)}")
                    st.info("Certifique-se de que sua planilha contém as colunas: UA, N°, NOME COMUM, NOME CIENTÍFICO, CAP (cm), HT(m)")
                    st.info("Colunas encontradas na sua planilha: " + ", ".join(df.columns.tolist()))
                else:
                    st.session_state.input_data = df
                    st.session_state.file_uploaded = True
                    st.success("✅ Dados da planilha salvos e prontos para processamento!")
                    
                    # Show data summary
                    st.info(f"📊 Resumo: {len(df)} árvores carregadas com {len(df.columns)} colunas")
                    
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
            'DAP (m)': '{:.4f}',
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
        st.metric("DAP Médio (m)", f"{results_df['DAP (m)'].mean():.4f}")
        st.metric("Altura Média (m)", f"{results_df['HT (m)'].mean():.2f}")

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
