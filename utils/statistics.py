import pandas as pd
import numpy as np
from scipy import stats

class StatisticsAnalyzer:
    """Class for performing statistical analysis on forest inventory data."""
    
    def __init__(self):
        pass
    
    def calculate_statistics(self, results_df, project_info):
        """
        Calculate comprehensive statistics for the forest inventory data.
        
        Args:
            results_df (pandas.DataFrame): Processed dataframe with volume calculations
            project_info (dict): Project information including plot details
            
        Returns:
            dict: Dictionary containing all statistical measures
        """
        # Para análise estatística correta em inventário florestal, precisamos agrupar por parcela
        # e calcular estatísticas baseadas na média por parcela, não por árvore individual
        
        # Identificar coluna UA/Parcela
        ua_column = None
        for col in results_df.columns:
            if 'UA' in str(col).upper() or 'PARCELA' in str(col).upper() or 'UNIDADE' in str(col).upper():
                ua_column = col
                break
        
        if ua_column is not None:
            # Calcular volume por hectare médio para cada parcela
            plot_volumes = results_df.groupby(ua_column)['VT (m³/ha)'].sum()
            volume_data = plot_volumes
        else:
            # Fallback: usar distribuição por parcelas
            num_plots = project_info['num_plots'] 
            total_trees = len(results_df)
            trees_per_plot = total_trees // num_plots
            
            plot_volumes = []
            for i in range(num_plots):
                start_idx = i * trees_per_plot
                end_idx = start_idx + trees_per_plot
                if i == num_plots - 1:  # Última parcela pega o resto
                    end_idx = total_trees
                plot_volume = results_df.iloc[start_idx:end_idx]['VT (m³/ha)'].sum()
                plot_volumes.append(plot_volume)
            
            volume_data = pd.Series(plot_volumes)
        
        n = len(volume_data)
        
        # Basic descriptive statistics usando float() para garantir tipos corretos
        mean = float(volume_data.mean())
        variance = float(volume_data.var(ddof=1))  # Sample variance
        std_dev = float(volume_data.std(ddof=1))   # Sample standard deviation
        
        # Coefficient of variation
        cv = (std_dev / mean) * 100 if mean != 0 else 0
        
        # Standard error of the mean
        standard_error = std_dev / np.sqrt(n)
        
        # 90% Confidence interval
        confidence_level = 0.90
        alpha = 1 - confidence_level
        t_critical = stats.t.ppf(1 - alpha/2, df=n-1)
        
        margin_of_error = t_critical * standard_error
        ci_lower = mean - margin_of_error
        ci_upper = mean + margin_of_error
        
        # Sampling error (as percentage of the mean)
        sampling_error = (margin_of_error / mean) * 100 if mean != 0 else 0
        
        # Additional statistics - convertendo para float
        minimum = float(volume_data.min())
        maximum = float(volume_data.max())
        median = float(volume_data.median())
        
        # Calculate expansion factors
        plot_area = project_info['plot_area']
        total_area = project_info['total_area']
        expansion_factor = total_area / (plot_area * project_info['num_plots'])
        
        statistics = {
            'n_trees': n,
            'mean': round(mean, 4),
            'variance': round(variance, 4),
            'std_dev': round(std_dev, 4),
            'cv': round(cv, 2),
            'standard_error': round(standard_error, 4),
            'ci_lower': round(ci_lower, 4),
            'ci_upper': round(ci_upper, 4),
            'sampling_error': round(sampling_error, 2),
            'minimum': round(minimum, 4),
            'maximum': round(maximum, 4),
            'median': round(median, 4),
            't_critical': round(float(t_critical), 4),
            'margin_of_error': round(margin_of_error, 4),
            'expansion_factor': round(expansion_factor, 4),
            'confidence_level': confidence_level
        }
        
        return statistics
    
    def assess_sampling_precision(self, sampling_error, threshold=20.0):
        """
        Assess whether the sampling meets precision requirements.
        
        Args:
            sampling_error (float): Sampling error as percentage
            threshold (float): Precision threshold (default 20%)
            
        Returns:
            dict: Assessment results
        """
        meets_precision = sampling_error <= threshold
        
        assessment = {
            'meets_precision': meets_precision,
            'sampling_error': sampling_error,
            'threshold': threshold,
            'message': self._get_precision_message(meets_precision, sampling_error)
        }
        
        return assessment
    
    def _get_precision_message(self, meets_precision, sampling_error):
        """
        Get appropriate message based on precision assessment.
        
        Args:
            meets_precision (bool): Whether precision requirement is met
            sampling_error (float): Sampling error percentage
            
        Returns:
            str: Appropriate message
        """
        if meets_precision:
            return f"Amostragem atingiu a precisão desejada (erro {sampling_error:.2f}% ≤ 20%)."
        else:
            return f"Amostragem não atingiu a precisão desejada (erro {sampling_error:.2f}% > 20%). Recomendado aumentar o número de parcelas."
    
    def calculate_required_plots(self, current_error, target_error=20.0, current_plots=1):
        """
        Calculate the number of plots required to achieve target precision.
        
        Args:
            current_error (float): Current sampling error percentage
            target_error (float): Target sampling error percentage
            current_plots (int): Current number of plots
            
        Returns:
            int: Required number of plots
        """
        if current_error <= target_error:
            return current_plots
        
        # The relationship between number of plots and error is inversely proportional to sqrt(n)
        error_ratio = current_error / target_error
        required_plots = int(np.ceil(current_plots * (error_ratio ** 2)))
        
        return required_plots
    
    def generate_volume_summary(self, results_df, project_info):
        """
        Generate a comprehensive volume summary for the project.
        
        Args:
            results_df (pandas.DataFrame): Processed dataframe with calculations
            project_info (dict): Project information
            
        Returns:
            dict: Volume summary statistics
        """
        # Volume calculations
        total_tree_volume = results_df['VT (m³)'].sum()
        mean_volume_per_ha = results_df['VT (m³/ha)'].mean()
        total_volume_per_ha = results_df['VT (m³/ha)'].sum()
        
        # Stereo volume calculations
        mean_stereo_volume_per_ha = results_df['VT (st/ha)'].mean()
        total_stereo_volume_per_ha = results_df['VT (st/ha)'].sum()
        
        # Project-wide estimates
        total_area = project_info['total_area']
        estimated_total_volume = mean_volume_per_ha * total_area
        estimated_total_stereo = mean_stereo_volume_per_ha * total_area
        
        summary = {
            'total_tree_volume': round(total_tree_volume, 4),
            'mean_volume_per_ha': round(mean_volume_per_ha, 4),
            'total_volume_per_ha': round(total_volume_per_ha, 4),
            'mean_stereo_volume_per_ha': round(mean_stereo_volume_per_ha, 4),
            'total_stereo_volume_per_ha': round(total_stereo_volume_per_ha, 4),
            'estimated_total_volume': round(estimated_total_volume, 2),
            'estimated_total_stereo': round(estimated_total_stereo, 2)
        }
        
        return summary
