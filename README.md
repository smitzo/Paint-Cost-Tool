# Paint Cost Analysis Tool (Freelance Project)
## Overview
The Paint Cost Analysis Tool (PDP Processor) is a Python application designed to analyze and highlight cost variations in paint production data from Excel spreadsheets. The tool identifies cost outliers, applies color-coding to visualize data patterns, and generates marked Excel files for further analysis.

## Features
- **Excel File Processing**: Analyzes paint production cost data from uploaded Excel files. Refer the book1.xlsx file template for excel files.
- **Dynamic Threshold Calculation**: Automatically determines cost thresholds for each product group
- **Color-Coded Highlighting**: 
  - Red: Highest cost in group with significant gap (>1.5)
  - Blue: Small cost gaps (<1.5) between adjacent products
  - Green: Lowest cost or below average cost
  - Orange: Above average cost but below 75th percentile
- **Marking System**: Flags products with cost variations for easy identification
- **User Authentication**: Secure login system with admin/developer privileges
- **Customizable Settings**: Adjustable cost/quantity columns and sheet names
- **Theme Customization**: Multiple color themes and light/dark mode

## System Requirements
- Python 3.7+
- Required Python packages:
  - `openpyxl`
  - `pandas`
  - `customtkinter`
  - `Pillow`

## Installation
1. Clone or download the repository
2. Install required packages:
   ```
   pip install openpyxl pandas customtkinter Pillow
   ```
3. Run the application:
   ```
   python pdp_processor.py
   ```

## Usage


1. **Main Functions**:
   - **Upload**: Select an Excel file for analysis
   - **Highlight Cells**: Process the file to identify cost variations
   - **Download**: Save the analyzed file with color coding
   - **Reset**: Clear current file selection

2. **Settings**:
   - Customize theme colors
   - Switch between light/dark mode
   - Configure default Excel columns and sheet names

3. **Admin Features** (when logged in as admin/developer):
   - Edit About section content
   - Manage user accounts

## File Format Requirements
The tool expects Excel files with:
- Cost data in column B (default, configurable)
- Quantity data in column C (default, configurable)
- Product codes starting with 'F0' or 'N0' for main products
- Bold formatting used for section headers

## Output
The processed file includes:
- Color-coded cells based on cost analysis
- 'Main Product' column showing product hierarchy
- 'Color Code' column indicating the analysis result
- 'Marking' column with 'X' for products requiring review

## License
This software is proprietary. All rights reserved.

## Version
PDP Processor v1.0
