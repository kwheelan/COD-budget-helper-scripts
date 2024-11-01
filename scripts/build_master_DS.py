"""
K. Wheelan October 2024

Iterate through each file in detail_sheets,
 copy to the master DS file and save it
"""

from openpyxl import load_workbook
import os
import shutil
# import custom functions
from utils.excel_utils import copy_cols, last_data_row, adjust_formula

# ===================== Constants ===============================

# File paths
DATA = 'input_data'
OUTPUT = 'output'
template_file = f'{DATA}/DS_templates/FY26 Detail Sheet Template Fast.xlsx'
source_folder = f'{DATA}/detail_sheets'
# source_folder = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/03. Form Development/Detail Sheets/Clean Sheets (pre-send)'
dest_file = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'

# Sheet name, columns to copy, start row
SHEETS = {
    'FTE, Salary-Wage, & Benefits' : { 
        'value_cols' : list('ABCDEFGHIJLMNPQRSTUVZ') + ['AE', 'AH'],
        'formula_cols' : list('KOWXY') + ['AA', 'AB', 'AC', 'AD', 'AF', 'AG'],
        'start_row' : 15 },
    'Overtime & Other Personnel' : {
        'value_cols' : list('ABCDEFGHIJKLNOPQRX'),
        'formula_cols' : list('MSTUVW'),
        'start_row' : 15 },
    'Non-Personnel' : {
        'value_cols': list('ABCDEFGHIJKLMOPQRSTUVWXYZ'),
        'formula_cols' : 'N',
        'start_row' : 19 },
    'Revenue' : {
        'value_cols': list('ABCDEFGHIJKLMNPQRSTUVW'),
        'formula_cols' : 'O',
        'start_row' : 15 }
}

# ================== Script functions ===================================

def move_data(detail_sheet, destination_file):
    # Load the workbooks
    source_wb_values = load_workbook(f'{source_folder}/{detail_sheet}', data_only=True)
    source_wb_formulas = load_workbook(f'{source_folder}/{detail_sheet}', data_only=False)
    destination_wb = load_workbook(destination_file, data_only=False)

    for sheet in SHEETS:
        source_ws_values = source_wb_values[sheet]
        source_ws_formulas = source_wb_formulas[sheet]
        dest_ws = destination_wb[sheet]
        config = SHEETS[sheet]
        
        # Determine the starting row in the destination sheet for pasting data
        dest_start_row = last_data_row(dest_ws, config['start_row']) + 1
        source_row_end = last_data_row(source_ws_values, n_header_rows=config['start_row'])
        
        # Copy value columns
        copy_cols(source_ws_values, dest_ws, config['value_cols'], 
                  source_row_start=config['start_row'], 
                  destination_row_start=dest_start_row, 
                  source_row_end=source_row_end)
        
        # Copy formula columns
        copy_cols(source_ws_formulas, dest_ws, config['formula_cols'], 
                  source_row_start=config['start_row'], 
                  destination_row_start=dest_start_row,
                  source_row_end=source_row_end)
        
        print(f'Copied {sheet} data from {detail_sheet}.')

    destination_wb.save(destination_file)

#================== Program on run ============================================

def main():
    # copy the template
    shutil.copy(template_file, dest_file)
    for detail_sheet in os.listdir(source_folder)[0:2]:

        # only attempt on excel sheets (exlude folder, etc)
        if '.xlsx' not in detail_sheet:
            continue
        move_data(detail_sheet, dest_file)


if __name__ == '__main__':
    main()
