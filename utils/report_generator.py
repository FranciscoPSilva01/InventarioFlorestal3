import pandas as pd
import io
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT

class ReportGenerator:
    """Class for generating Excel and PDF reports from forest inventory analysis."""
    
    def __init__(self):
        pass
    
    def generate_excel_report(self, results_df, statistics, project_info):
        """
        Generate a comprehensive Excel report.
        
        Args:
            results_df (pandas.DataFrame): Processed data with calculations
            statistics (dict): Statistical analysis results
            project_info (dict): Project information
            
        Returns:
            io.BytesIO: Excel file buffer
        """
        buffer = io.BytesIO()
        
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # Sheet 1: Project Information
            self._create_project_info_sheet(writer, project_info, statistics)
            
            # Sheet 2: Detailed calculations
            results_df.to_excel(writer, sheet_name='Cálculos Detalhados', index=False)
            
            # Sheet 3: Statistical Summary
            self._create_statistics_sheet(writer, statistics, project_info)
            
            # Sheet 4: Volume Summary
            self._create_volume_summary_sheet(writer, results_df, project_info, statistics)
        
        buffer.seek(0)
        return buffer.getvalue()
    
    def _create_project_info_sheet(self, writer, project_info, statistics):
        """Create project information sheet."""
        project_data = {
            'Parâmetro': [
                'Nome do Projeto',
                'Data do Relatório',
                'Número de Parcelas',
                'Dimensões da Parcela (m)',
                'Área por Parcela (ha)',
                'Área Total Amostrada (ha)',
                'Área de Supressão (ha)',
                'Porcentagem Amostrada (%)',
                'Fator de Forma',
                'Total de Árvores',
                'Erro Amostral (%)',
                'Precisão Atingida'
            ],
            'Valor': [
                project_info['project_name'],
                datetime.now().strftime('%d/%m/%Y %H:%M'),
                project_info['num_plots'],
                f"{project_info['plot_length']} x {project_info['plot_width']}",
                f"{project_info['plot_area']:.4f}",
                f"{project_info['total_sampled_area']:.4f}",
                f"{project_info['total_area']:.2f}",
                f"{project_info['sampling_percentage']:.2f}",
                f"{project_info['form_factor']:.2f}",
                statistics['n_trees'],
                f"{statistics['sampling_error']:.2f}",
                'Sim' if statistics['sampling_error'] <= 20 else 'Não'
            ]
        }
        
        project_df = pd.DataFrame(project_data)
        project_df.to_excel(writer, sheet_name='Informações do Projeto', index=False)
    
    def _create_statistics_sheet(self, writer, statistics, project_info):
        """Create statistical analysis sheet."""
        stats_data = {
            'Estatística': [
                'Número de Árvores',
                'Média (m³/ha)',
                'Variância',
                'Desvio Padrão',
                'Coeficiente de Variação (%)',
                'Erro Padrão',
                'Limite Inferior IC 90%',
                'Limite Superior IC 90%',
                'Erro Amostral (%)',
                'Valor Mínimo',
                'Valor Máximo',
                'Mediana',
                't crítico',
                'Margem de Erro'
            ],
            'Valor': [
                statistics['n_trees'],
                statistics['mean'],
                statistics['variance'],
                statistics['std_dev'],
                statistics['cv'],
                statistics['standard_error'],
                statistics['ci_lower'],
                statistics['ci_upper'],
                statistics['sampling_error'],
                statistics['minimum'],
                statistics['maximum'],
                statistics['median'],
                statistics['t_critical'],
                statistics['margin_of_error']
            ]
        }
        
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name='Análise Estatística', index=False)
    
    def _create_volume_summary_sheet(self, writer, results_df, project_info, statistics):
        """Create volume summary sheet."""
        # Calculate volume estimates for total area
        volume_estimate = statistics['mean'] * project_info['total_area']
        volume_lower = statistics['ci_lower'] * project_info['total_area']
        volume_upper = statistics['ci_upper'] * project_info['total_area']
        
        # Calculate stereo volumes
        stereo_mean = results_df['VT (st/ha)'].mean()
        stereo_estimate = stereo_mean * project_info['total_area']
        
        volume_data = {
            'Tipo de Volume': [
                'Volume Médio por Hectare (m³/ha)',
                'Volume Total Estimado (m³)',
                'Volume Mínimo IC 90% (m³)',
                'Volume Máximo IC 90% (m³)',
                'Volume Médio Estéreo por Hectare (st/ha)',
                'Volume Total Estéreo Estimado (st)',
                'Volume Total das Árvores Amostradas (m³)',
                'Número Total de Árvores Amostradas'
            ],
            'Valor': [
                f"{statistics['mean']:.4f}",
                f"{volume_estimate:.2f}",
                f"{volume_lower:.2f}",
                f"{volume_upper:.2f}",
                f"{stereo_mean:.4f}",
                f"{stereo_estimate:.2f}",
                f"{results_df['VT (m³)'].sum():.4f}",
                len(results_df)
            ]
        }
        
        volume_df = pd.DataFrame(volume_data)
        volume_df.to_excel(writer, sheet_name='Resumo de Volumes', index=False)
    
    def generate_pdf_report(self, results_df, statistics, project_info):
        """
        Generate a comprehensive PDF report.
        
        Args:
            results_df (pandas.DataFrame): Processed data with calculations
            statistics (dict): Statistical analysis results
            project_info (dict): Project information
            
        Returns:
            io.BytesIO: PDF file buffer
        """
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        story = []
        
        # Styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=18,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading1'],
            fontSize=14,
            spaceAfter=12,
            spaceBefore=20
        )
        
        # Title
        story.append(Paragraph("Relatório de Inventário Florestal", title_style))
        story.append(Spacer(1, 20))
        
        # Project Information
        story.append(Paragraph("Informações do Projeto", heading_style))
        project_table_data = [
            ['Parâmetro', 'Valor'],
            ['Nome do Projeto', project_info['project_name']],
            ['Data do Relatório', datetime.now().strftime('%d/%m/%Y %H:%M')],
            ['Número de Parcelas', str(project_info['num_plots'])],
            ['Dimensões da Parcela', f"{project_info['plot_length']} x {project_info['plot_width']} m"],
            ['Área por Parcela', f"{project_info['plot_area']:.4f} ha"],
            ['Área Total Amostrada', f"{project_info['total_sampled_area']:.4f} ha"],
            ['Área de Supressão', f"{project_info['total_area']:.2f} ha"],
            ['Porcentagem Amostrada', f"{project_info['sampling_percentage']:.2f}%"],
            ['Fator de Forma', f"{project_info['form_factor']:.2f}"],
        ]
        
        project_table = Table(project_table_data)
        project_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(project_table)
        story.append(Spacer(1, 20))
        
        # Statistical Analysis
        story.append(Paragraph("Análise Estatística", heading_style))
        stats_table_data = [
            ['Estatística', 'Valor'],
            ['Número de Árvores', str(statistics['n_trees'])],
            ['Média (m³/ha)', f"{statistics['mean']:.4f}"],
            ['Desvio Padrão', f"{statistics['std_dev']:.4f}"],
            ['Coeficiente de Variação (%)', f"{statistics['cv']:.2f}"],
            ['Erro Amostral (%)', f"{statistics['sampling_error']:.2f}"],
            ['Intervalo de Confiança 90%', f"{statistics['ci_lower']:.4f} - {statistics['ci_upper']:.4f}"],
        ]
        
        stats_table = Table(stats_table_data)
        stats_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(stats_table)
        story.append(Spacer(1, 20))
        
        # Precision Assessment
        story.append(Paragraph("Avaliação da Precisão", heading_style))
        if statistics['sampling_error'] <= 20:
            precision_text = f"✓ Amostragem atingiu a precisão desejada (erro {statistics['sampling_error']:.2f}% ≤ 20%)."
        else:
            precision_text = f"⚠ Amostragem não atingiu a precisão desejada (erro {statistics['sampling_error']:.2f}% > 20%). Recomendado aumentar o número de parcelas."
        
        story.append(Paragraph(precision_text, styles['Normal']))
        story.append(Spacer(1, 20))
        
        # Volume Estimates
        story.append(Paragraph("Estimativas de Volume", heading_style))
        volume_estimate = statistics['mean'] * project_info['total_area']
        volume_lower = statistics['ci_lower'] * project_info['total_area']
        volume_upper = statistics['ci_upper'] * project_info['total_area']
        
        volume_table_data = [
            ['Tipo de Estimativa', 'Valor'],
            ['Volume Total Estimado (m³)', f"{volume_estimate:.2f}"],
            ['Volume Mínimo IC 90% (m³)', f"{volume_lower:.2f}"],
            ['Volume Máximo IC 90% (m³)', f"{volume_upper:.2f}"],
        ]
        
        volume_table = Table(volume_table_data)
        volume_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(volume_table)
        
        # Build PDF
        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()
