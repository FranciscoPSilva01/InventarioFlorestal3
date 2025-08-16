import pandas as pd
import numpy as np

class ForestryCalculator:
    """Class for performing forestry calculations according to the specified formulas."""
    
    def __init__(self):
        pass
    
    def calculate_dap(self, cap):
        """
        Calculate DAP (Diameter at Breast Height) from CAP (Circumference at Breast Height).
        Formula: DAP = CAP / π
        
        Args:
            cap (float): Circumference at breast height in cm
            
        Returns:
            float: Diameter at breast height in meters (converted from cm)
        """
        return (cap / np.pi) / 100  # Convert cm to meters
    
    def calculate_tree_volume(self, dap_m, ht):
        """
        Calculate tree volume using the specified formula.
        Formula: VT = 0.000094 × (DAP^1.830398) × (HT^0.960913)
        
        Args:
            dap_m (float): Diameter at breast height in meters
            ht (float): Total height in meters
            
        Returns:
            float: Tree volume in cubic meters
        """
        return 0.000094 * (dap_m ** 1.830398) * (ht ** 0.960913)
    
    def calculate_volume_per_hectare(self, tree_volume, form_factor, plot_area_ha):
        """
        Calculate volume per hectare.
        Formula: VT_ha = (VT / FF) / plot_area_ha
        
        Args:
            tree_volume (float): Tree volume in cubic meters
            form_factor (float): Form factor
            plot_area_ha (float): Plot area in hectares
            
        Returns:
            float: Volume per hectare in cubic meters per hectare
        """
        return (tree_volume / form_factor) / plot_area_ha
    
    def calculate_stereo_volume(self, volume_per_ha):
        """
        Calculate stereo volume per hectare.
        Formula: VT_st/ha = VT_ha × 2.65
        
        Args:
            volume_per_ha (float): Volume per hectare in cubic meters
            
        Returns:
            float: Stereo volume per hectare
        """
        return volume_per_ha * 2.65
    
    def process_data(self, df, form_factor, plot_area_ha):
        """
        Process the complete dataset with all forestry calculations.
        
        Args:
            df (pandas.DataFrame): Input dataframe with tree data
            form_factor (float): Form factor for calculations
            plot_area_ha (float): Plot area in hectares
            
        Returns:
            pandas.DataFrame: Processed dataframe with all calculations
        """
        # Apply column mapping first
        results_df = self._apply_column_mapping(df.copy())
        
        # Validate required columns
        required_columns = ['CAP (cm)', 'HT (m)']
        for col in required_columns:
            if col not in results_df.columns:
                raise ValueError(f"Required column '{col}' not found in data")
        
        # Remove rows with missing values in critical columns
        results_df = results_df.dropna(subset=required_columns)
        
        # Ensure numeric data types
        results_df['CAP (cm)'] = pd.to_numeric(results_df['CAP (cm)'], errors='coerce')
        results_df['HT (m)'] = pd.to_numeric(results_df['HT (m)'], errors='coerce')
        
        # Remove rows with invalid numeric values
        results_df = results_df.dropna(subset=['CAP (cm)', 'HT (m)'])
        
        # Calculate DAP in meters
        results_df['DAP (m)'] = results_df['CAP (cm)'].apply(self.calculate_dap)
        
        # Calculate tree volume
        results_df['VT (m³)'] = results_df.apply(
            lambda row: self.calculate_tree_volume(row['DAP (m)'], row['HT (m)']),
            axis=1
        )
        
        # Calculate volume per hectare
        results_df['VT (m³/ha)'] = results_df['VT (m³)'].apply(
            lambda vt: self.calculate_volume_per_hectare(vt, form_factor, plot_area_ha)
        )
        
        # Calculate stereo volume per hectare
        results_df['VT (st/ha)'] = results_df['VT (m³/ha)'].apply(self.calculate_stereo_volume)
        
        # Round all calculated values to 4 decimal places for precision
        calculation_columns = ['DAP (m)', 'VT (m³)', 'VT (m³/ha)', 'VT (st/ha)']
        for col in calculation_columns:
            results_df[col] = results_df[col].round(4)
        
        return results_df
    
    def validate_input_data(self, df):
        """
        Validate the input data for common issues.
        
        Args:
            df (pandas.DataFrame): Input dataframe to validate
            
        Returns:
            list: List of validation errors
        """
        errors = []
        
        # Apply column mapping first
        df = self._apply_column_mapping(df)
        
        # Check for required columns
        required_columns = ['Nº da árvore', 'Nome comum/científico', 'CAP (cm)', 'HT (m)']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            errors.append(f"Colunas obrigatórias não encontradas: {', '.join(missing_columns)}")
        
        # Check for empty dataframe
        if len(df) == 0:
            errors.append("Planilha vazia ou sem dados válidos")
        
        # Check for numeric values in CAP and HT columns
        if 'CAP (cm)' in df.columns:
            non_numeric_cap = df['CAP (cm)'].apply(lambda x: not pd.api.types.is_numeric_dtype(type(x)) and pd.notna(x))
            if non_numeric_cap.any():
                errors.append("Valores não numéricos encontrados na coluna CAP (cm)")
        
        if 'HT (m)' in df.columns:
            non_numeric_ht = df['HT (m)'].apply(lambda x: not pd.api.types.is_numeric_dtype(type(x)) and pd.notna(x))
            if non_numeric_ht.any():
                errors.append("Valores não numéricos encontrados na coluna HT (m)")
        
        # Check for negative values
        if 'CAP (cm)' in df.columns:
            negative_cap = df['CAP (cm)'] <= 0
            if negative_cap.any():
                errors.append("Valores negativos ou zero encontrados na coluna CAP (cm)")
        
        if 'HT (m)' in df.columns:
            negative_ht = df['HT (m)'] <= 0
            if negative_ht.any():
                errors.append("Valores negativos ou zero encontrados na coluna HT (m)")
        
        return errors
    
    def _apply_column_mapping(self, df):
        """
        Apply column name mapping to handle different column name formats.
        
        Args:
            df (pandas.DataFrame): Input dataframe
            
        Returns:
            pandas.DataFrame: Dataframe with mapped column names
        """
        df = df.copy()
        
        # Map column names to expected format
        column_mapping = {
            'N°': 'Nº da árvore',
            'NOME COMUM': 'Nome comum/científico',
            'NOME CIENTÍFICO': 'Nome científico',
            'CAP (cm)': 'CAP (cm)',
            'HT(m)': 'HT (m)',
            'Altura total HT(m)': 'HT (m)'
        }
        
        # Rename columns if they exist with different names
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Combine nome comum and científico if they are separate
        if 'Nome comum/científico' not in df.columns:
            if 'NOME COMUM' in df.columns and 'NOME CIENTÍFICO' in df.columns:
                df['Nome comum/científico'] = df['NOME COMUM'].astype(str) + ' / ' + df['NOME CIENTÍFICO'].astype(str)
            elif 'NOME COMUM' in df.columns:
                df['Nome comum/científico'] = df['NOME COMUM']
            elif 'NOME CIENTÍFICO' in df.columns:
                df['Nome comum/científico'] = df['NOME CIENTÍFICO']
        
        return df
