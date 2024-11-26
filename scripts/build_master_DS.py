"""
K. Wheelan October 2024

Iterate through each file in detail_sheets,
 copy to the master DS file and save it
"""

from openpyxl import load_workbook
from openpyxl.styles import Font
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

# Dictionary to help with summary tab
SUMMARY_SECTIONS = {
    'FTE' : {
        'tab' : 'FTE, Salary-Wage, & Benefits',
        'count_col' : 'U',
        'baseline_col' : 'L'
    }, 
    'Salary & Benefits' : {
        'tab' : 'FTE, Salary-Wage, & Benefits',
        'count_col': 'AG',
        'baseline_col' : 'L'
    }, 
    'Non-Personnel' : {
        'tab' : 'Non-Personnel',
        'count_col': 'Y',
        'baseline_col' : 'O'
    }, 
    'Overtime' : {
        'tab' : 'Overtime & Other Personnel',
        'count_col': 'W',
        'baseline_col' : 'N'
    }, 
    'Revenue' : {
        'tab' : 'Revenue',
        'count_col': 'V',
        'baseline_col' : 'P'
    }}

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

    # load workbook and formulas
    destination_wb = load_workbook(destination_file, data_only=False)
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

    def summary_formula(section, baseOrSupp, fund_cell, approp_cell):
        """ Creating the SUMIFS formulas for summary tab """
        # get sheet for relevant data
        tab = SUMMARY_SECTIONS[section]['tab']
        # column to aggregate in that sheet
        col = SUMMARY_SECTIONS[section]['count_col']
        # baseline/supplemental column
        b_s_col = SUMMARY_SECTIONS[section]['baseline_col']
        
        if approp_cell.value == 'Total':
            return f'=SUMIFS(\'{tab}\'!${col}:${col}, \'{tab}\'!${b_s_col}:${b_s_col}, "{baseOrSupp}", \'{tab}\'!$D:$D, {fund_cell.coordinate})'
        return f'=SUMIFS(\'{tab}\'!${col}:${col}, \'{tab}\'!${b_s_col}:${b_s_col}, "{baseOrSupp}", \'{tab}\'!$D:$D, {fund_cell.coordinate}, \'{tab}\'!$G:$G, {approp_cell.coordinate})'

    header_rows = 2
    # Write the transformed data to the sheet
    for row_ix in range(len(rows)):
        active_row = row_ix+(header_rows+1)
        # for each row (ie. fund/approp combo)
        fund_cell = summary.cell(row=active_row, column=1)
        approp_cell = summary.cell(row=active_row, column=2)
        fund_cell.value, approp_cell.value = rows[row_ix]

        # for each section: FTE, Salary and Benefits, NP, OT, Revenue
        for section_ix in range(len(SUMMARY_SECTIONS)):
            # get section name
            section = list(SUMMARY_SECTIONS.keys())[section_ix]
            # fetch cell locations and fill with formulas
            baseline_cell  = summary.cell(row=active_row, column=3+(section_ix*3))
            baseline_cell.value = summary_formula(section, 'Baseline', fund_cell, approp_cell)
            supplemental_cell = summary.cell(row=active_row, column=4+(section_ix*3))
            supplemental_cell.value = summary_formula(section, 'Supplemental', fund_cell, approp_cell)
            total_cell = summary.cell(row=active_row, column=5+(section_ix*3))
            total_cell.value = f'={baseline_cell.coordinate} + {supplemental_cell.coordinate}'

        # add total column to sum all subtotals
        total_col_cell = summary.cell(row=active_row, column=18)
        total_col_cell.value = f'=SUM(H{active_row}, K{active_row}, N{active_row})'

        # format as USD
        for col in range(6,19):
            cell = summary.cell(row=active_row, column=col)
            cell.number_format = cell.number_format = '"$"#,##0' 

        # bold row if total row
        if approp_cell.value == 'Total':
            for col in range(1, 19):
                cell = summary.cell(row=active_row, column=col)
                cell.font = Font(bold=True)

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
    print("Created summary tab")


if __name__ == '__main__':
    main()
