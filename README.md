# COMP4037-CW2
Coursework 2: NHS Hospital Admissions Data Visualisation

## Data Source
The data file provided in the assignment document

## Data Setup
Download all NHS England HES Excel files and place them in a folder
named hospital_data/ with no subfolders. If any files are in the legacy
.xls format, convert them to .xlsx first (Excel: File -> Save As ->
Excel Workbook *.xlsx). This ensures all files are readable by openpyxl.

## File Structure

    COMP4037-CW2/
    ├── hospital_data/
    │   ├── hosp-epis-stat-admi-prim-diag-sum-98-99-tab.xlsx
    │   ├── hosp-epis-stat-admi-prim-diag-sum-99-00-tab.xlsx
    │   ├── ... (all HES Excel files)
    │   └── hosp-epis-stat-admi-diag-2023-24-tab.xlsx
    ├── scan_all_en.py
    ├── nhs_data_cleaning_v2.py
    └── nhs_visualization_v4_en.py

## How to Run

Install dependencies first:

    pip install pandas openpyxl matplotlib numpy

Then run the scripts in order:

    1. python scan_all.py             — detects sheet names across all files
    2. python nhs_data_cleaning_v2.py    — merges and cleans the data
    3. python nhs_visualization_v4.py — generates the visualisation figures

Output figures will be saved to the hospital_data/ folder.

## Requirements
- Python 3.10+
- pandas, openpyxl, matplotlib, numpy
