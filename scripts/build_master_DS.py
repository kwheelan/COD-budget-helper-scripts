"""
K. Wheelan October 2024

Iterate through each file in detail_sheets,
 copy to the master DS file and save it
"""

from openpyxl import load_workbook
import os
import shutil
# import custom functions
from utils.excel_utils import copy_cols, last_data_row

# ===================== Constants ===============================

# File paths
DATA = 'input_data'
OUTPUT = 'output'
template_file = f'{DATA}/DS_templates/FY26 Detail Sheet Template Fast.xlsx'
source_folder = f'{DATA}/detail_sheets_analyst_review'
# source_folder = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/03. Form Development/Detail Sheets/Clean Sheets (pre-send)'
dest_file = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'

# Sheet name, columns to copy, start row
SHEETS = {
    'FTE, Salary-Wage, & Benefits' : { 
        'value_cols' : list('ABCDGILMNPQRSTUVZ') + ['AE', 'AH'],
        'formula_cols' : list('EFHJKOWXY') + ['AA', 'AB', 'AC', 'AD', 'AF', 'AG'],
        'start_row' : 15 },
    'Overtime & Other Personnel' : {
        'value_cols' : list('ABCDGIKNOPQRX'),
        'formula_cols' : list('EFHJLMSTUVW'),
        'start_row' : 15 },
    'Non-Personnel' : {
        'value_cols': list('ABCDGIKOPQRSTUVWXYZ'),
        'formula_cols' : 'EFHJLMN',
        'start_row' : 19 },
    'Revenue' : {
        'value_cols': list('ABCDGIKPQRSTUVW'),
        'formula_cols' : 'EFHJLMNO',
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
        source_row_end = last_data_row(source_ws_values, n_header_rows=config['start_row']) + 1
        
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

def create_summary(destination_file):

    fund_col = 'D'
    approp_col = 'G'

    # load workbook
    destination_wb = load_workbook(destination_file, data_only=True)
    # placeholder for funds
    funds = {}

    # go through each sheet and collect all unique funds
    for sheet in SHEETS:
        fund_data = destination_wb[sheet][fund_col]
        approp_data = destination_wb[sheet][approp_col]
        # start serach after header
        for i in range(SHEETS[sheet]['start_row'], len(fund_data)):
            fund = fund_data[i].value
            approp = approp_data[i].value
            if fund and (fund not in funds):
                funds[fund] = []
            if approp and (approp not in funds[fund]) and str(approp) not in funds[fund]:
                funds[fund].append(approp)

    # Transform the dictionary into a 2-column list
    rows = []
    for key, values in funds.items():
        for value in values:
            rows.append([key, value])
        rows.append([key, 'Total'])

    # Copy to the summary tab
    summary = destination_wb['Summary']

    # Write the transformed data to the sheet
    for i in range(len(rows)):
        summary.append(rows[i])
        fund_cell = summary.cell(row=i+1, column=1)
        approp_cell = summary.cell(row=i+1, column=2)
        baseline_formula_cell  = summary.cell(row=i+3, column=3)
        supplemental_formula_cell = summary.cell(row=i+3, column=4)
        total_formula_cell = summary.cell(row=i+3, column=5)

        # Creating the SUMIFS formulas
        if approp_cell.value == 'Total':
            baseline_formula = f'=SUMIFS(\'FTE, Salary-Wage, & Benefits\'!$U:$U, \'FTE, Salary-Wage, & Benefits\'!$L:$L, "Baseline", \'FTE, Salary-Wage, & Benefits\'!$D:$D, {fund_cell.coordinate})'
            supplemental_formula = f'=SUMIFS(\'FTE, Salary-Wage, & Benefits\'!$U:$U, \'FTE, Salary-Wage, & Benefits\'!$L:$L, "Supplemental", \'FTE, Salary-Wage, & Benefits\'!$D:$D, {fund_cell.coordinate})'
        else:
            baseline_formula = f'=SUMIFS(\'FTE, Salary-Wage, & Benefits\'!$U:$U, \'FTE, Salary-Wage, & Benefits\'!$L:$L, "Baseline", \'FTE, Salary-Wage, & Benefits\'!$D:$D, {fund_cell.coordinate}, \'FTE, Salary-Wage, & Benefits\'!$G:$G, {approp_cell.coordinate})'
            supplemental_formula = f'=SUMIFS(\'FTE, Salary-Wage, & Benefits\'!$U:$U, \'FTE, Salary-Wage, & Benefits\'!$L:$L, "Supplemental", \'FTE, Salary-Wage, & Benefits\'!$D:$D, {fund_cell.coordinate}, \'FTE, Salary-Wage, & Benefits\'!$G:$G, {approp_cell.coordinate})'
        
        baseline_formula_cell.value = baseline_formula
        supplemental_formula_cell.value = supplemental_formula
        total_formula_cell.value = f'={baseline_formula_cell.coordinate} + {supplemental_formula_cell.coordinate}'

    # Save the workbook
    destination_wb.save(filename=destination_file)


#================== Program on run ============================================

def main():
    # copy the template
    shutil.copy(template_file, dest_file)
    for detail_sheet in os.listdir(source_folder):

        # only attempt on excel sheets (exlude folder, etc)
        if '.xlsx' not in detail_sheet:
            continue
        move_data(detail_sheet, dest_file)
    create_summary(dest_file)


if __name__ == '__main__':
    main()
