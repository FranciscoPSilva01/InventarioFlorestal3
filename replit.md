# Sistema de Inventário Florestal

## Overview

This is a forest inventory analysis system built with Streamlit that processes field data to calculate forest volume measurements and generate comprehensive statistical reports. The system allows users to upload field measurement data (tree circumference, height, species) and automatically calculates volumes, statistics, and generates professional reports in Excel and PDF formats. It's designed for forestry professionals conducting forest inventories and suppression assessments.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Framework**: Streamlit web application with a tabbed interface
- **UI Structure**: Four main tabs for data upload, processing, statistics, and reporting
- **State Management**: Session state variables to maintain data persistence across tab interactions
- **Layout**: Wide layout with responsive columns for better data presentation

### Backend Architecture
- **Modular Design**: Separated into utility modules for specific functions:
  - `ForestryCalculator`: Handles all forestry-specific calculations (DAP, tree volume, volume per hectare)
  - `StatisticsAnalyzer`: Performs statistical analysis including confidence intervals and sampling error
  - `ReportGenerator`: Creates Excel and PDF reports with multiple sheets/sections
- **Data Processing**: Pandas-based data manipulation for handling uploaded CSV/Excel files
- **Mathematical Engine**: NumPy for numerical calculations and SciPy for statistical computations

### Calculation Engine
- **Forestry Formulas**: Implements specific forestry equations:
  - DAP calculation from circumference (DAP = CAP/π)
  - Tree volume using allometric equation: VT = 0.000094 × (DAP^1.830398) × (HT^0.960913)
  - Volume per hectare calculations with form factor adjustments
- **Statistical Analysis**: Comprehensive statistics including mean, variance, coefficient of variation, confidence intervals, and sampling error calculations

### Report Generation
- **Excel Reports**: Multi-sheet workbooks using openpyxl engine with project info, detailed calculations, statistics, and volume summaries
- **PDF Reports**: Professional reports using ReportLab with structured layouts, tables, and formatting
- **Data Export**: In-memory buffer generation for file downloads without server storage

## External Dependencies

### Core Libraries
- **Streamlit**: Web application framework for the user interface
- **Pandas**: Data manipulation and analysis for handling forest inventory datasets
- **NumPy**: Numerical computing for mathematical calculations
- **SciPy**: Statistical functions for confidence interval calculations

### Report Generation
- **openpyxl**: Excel file creation and manipulation for detailed spreadsheet reports
- **ReportLab**: PDF generation library for professional report formatting

### Data Processing
- **io**: In-memory file handling for upload/download operations without disk storage
- **datetime**: Timestamp generation for report metadata

### Statistical Analysis
- **scipy.stats**: Statistical distributions and hypothesis testing for forestry sample analysis

The system follows a clean separation of concerns with dedicated modules for calculations, statistics, and reporting, making it maintainable and extensible for additional forestry analysis features.