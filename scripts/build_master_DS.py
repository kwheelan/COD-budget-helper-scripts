"""
K. Wheelan October 2024

Iterate through each file in detail_sheets,
 copy to the master DS file and save it
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import os
import shutil
# import custom functions
from utils.excel_utils import copy_cols, last_data_row, col_range

# ===================== Constants ===============================

# File paths
DATA = 'input_data'
OUTPUT = 'output'
template_file = f'{DATA}/DS_templates/FY26 Detail Sheet Template Fast.xlsx'
#source_folder = f'{DATA}/detail_sheets_analyst_review'
SOURCE_FOLDER = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/08A. Deputy Budget Director Recommend/'
dest_file = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'

# Sheet name, columns to copy, start row
SHEETS = {
    'FTE, Salary-Wage, & Benefits' : { 
        'cols' : col_range('A', 'BE'), # extra column for analyst notes (at least in HR DS)
        'start_row' : 15 },
    'Overtime & Other Personnel' : {
        'cols' : col_range('A', 'AN'),
        'start_row' : 15 },
    'Non-Personnel' : {
        'cols' : col_range('A', 'AH'),
        'start_row' : 19 },
    'Revenue' : {
        'cols' : col_range('A', 'AE'),
        'start_row' : 15 }
}

# Dictionary to help with summary tab
SUMMARY_SECTIONS = {
    'FTE' : {
        'tab' : 'FTE, Salary-Wage, & Benefits',
        'count_col' : 'AZ',
        'baseline_col' : 'L'
    }, 
    'Salary & Benefits' : {
        'tab' : 'FTE, Salary-Wage, & Benefits',
        'count_col': 'BC',
        'baseline_col' : 'L'
    }, 
    'Non-Personnel' : {
        'tab' : 'Non-Personnel',
        'count_col': 'AG',
        'baseline_col' : 'O'
    }, 
    'Overtime' : {
        'tab' : 'Overtime & Other Personnel',
        'count_col': 'AM',
        'baseline_col' : 'N'
    }, 
    'Revenue' : {
        'tab' : 'Revenue',
        'count_col': 'AD',
        'baseline_col' : 'P'
    }}

# Set the gray fill for background color on Excel cells
gray_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

# ================== Script functions ===================================


def move_data(detail_sheet, destination_file):
    # Load the workbooks
    #source_wb_values = load_workbook(f'{source_folder}/{detail_sheet}', data_only=True)
    source_wb_formulas = load_workbook(detail_sheet, data_only=False)
    destination_wb = load_workbook(destination_file, data_only=False)

    for sheet in SHEETS:
        if sheet in source_wb_formulas.sheetnames:
            source_ws_formulas = source_wb_formulas[sheet]
            dest_ws = destination_wb[sheet]
            config = SHEETS[sheet]

            # Determine the starting row in the destination sheet for pasting data
            dest_start_row = last_data_row(dest_ws, config['start_row']) + 1
            source_row_end = last_data_row(source_ws_formulas, n_header_rows=config['start_row']) + 1
            
            # Copy formula columns
            copy_cols(source_ws_formulas, dest_ws, config['cols'], 
                    source_row_start=config['start_row'], 
                    destination_row_start=dest_start_row,
                    source_row_end=source_row_end)
            
            name = detail_sheet.split('\\')[-1]
            print(f'Copied {sheet} data from {name}.')

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
        rows.append([key, 'Subtotal'])

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
        
        if approp_cell.value == 'Subtotal':
            return f'=SUMIFS(\'{tab}\'!${col}:${col}, \'{tab}\'!${b_s_col}:${b_s_col}, "{baseOrSupp}", \'{tab}\'!${fund_col}:${fund_col}, {fund_cell.coordinate})'
        return f'=SUMIFS(\'{tab}\'!${col}:${col}, \'{tab}\'!${b_s_col}:${b_s_col}, "{baseOrSupp}", \'{tab}\'!${fund_col}:${fund_col}, {fund_cell.coordinate}, \'{tab}\'!${approp_col}:${approp_col}, {approp_cell.coordinate})'

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

        # REPLACED WITH CONDITIONAL FORMATTING IN WB
        # # bold row if total row
        # if approp_cell.value == 'Subtotal':
        #     for col in range(1, 19):
        #         cell = summary.cell(row=active_row, column=col)
        #         # turn row gray with bold
        #         cell.font = Font(bold=True)
        #         cell.fill = gray_fill

    # Save the workbook
    destination_wb.save(filename=destination_file)

# MOVE TO UTILS

def find_DS(folder, verbose = False):
    # Get full file path
    folder_fp = os.path.join(SOURCE_FOLDER, folder)

    # grab list of all files in that folder
    files = os.listdir(folder_fp)

    # generate list of reviewed detail sheets
    reviewed_DS = []
    for file in files:
        if ('(Deputy Budget Director)' in file) and '.xlsx' in file:
            reviewed_DS.append(os.path.join(SOURCE_FOLDER, folder, file))

    # return message
    if len(reviewed_DS) > 1:
        message = f'Multiple potential reviewed detail sheets: {reviewed_DS}'
    elif len(reviewed_DS) == 0:
        message = f'No reviewed detail sheet found in {folder}'
    if verbose:
        print(message)

    return reviewed_DS

def create_file_list(verbose = False):
    # all folders in analyst review section
    DS_FOLDERS = os.listdir(SOURCE_FOLDER)

    # initialize list and add viable files to it
    DS_list = []
    for folder in DS_FOLDERS:
        folder_fp = os.path.join(SOURCE_FOLDER, folder)
        if(os.path.isdir(folder_fp)):
            DS_list += find_DS(folder, verbose)

    return DS_list

def include_dept(file):
    for dept in INCLUDE:
        if dept in file:
            return dept
    return False

def exclude_dept(file):
    for dept in EXCLUDE:
        if dept in file:
            return dept
    return False



#================== Program on run ============================================

def main():
    # copy the template
    shutil.copy(template_file, dest_file)

    # get list of DS files
    DS_list = [file for file in create_file_list() if not exclude_dept(file)]
    for detail_sheet in DS_list:
        load_workbook(detail_sheet)
        move_data(detail_sheet, dest_file)
    create_summary(dest_file)
    print("Created summary tab")

INCLUDE = [
    'Airport',
    'BSEED',
    '16 CDD',
    '18 DSLP',
    '19 DPW',
    '20 DDoT',
    '24 Fire',
    '25 Health',
    '28 HR',
    '29 CRIO'
    '31 DoIT',
    '32 Law',
    '33 Mayor',
    '34 Parking',
    '35 Non-Dept',
    '36 HRD Classic',
    '36 HRD JET Team',
    '38 PLD',
    '43 PDD',
    '45 DAH',
    '47 GSD',
    '50 OAG',
    '51 BZA',
    '52 Council',
    '53 Ombuds',
    '54 OIG',
    '60 36D',
    '70 Clerk',
    '71 Elections'
]

EXCLUDE = [
    '23 OCFO',
    '72 Library'
]

if __name__ == '__main__':
    main()
