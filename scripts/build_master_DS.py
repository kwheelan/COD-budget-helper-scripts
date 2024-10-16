"""
K. Wheelan October 2024

Iterate through each file in detail_sheets,
 copy to the master DS file and save it
"""

from openpyxl import load_workbook
import os
import shutil
# import custom functions
from utils.excel_utils import copy_cols

# ===================== Constants ===============================

# File paths
DATA = 'input_data'
OUTPUT = 'output'
template_file = f'{DATA}/DS_templates/FY26 Detail Sheet Template Fast.xlsx'
source_folder = os.listdir(f'{DATA}/detail_sheets')
dest_file = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'

# Sheet name, columns to copy, start row
SHEETS = {
    'FTE, Salary-Wage, & Benefits' : { 
        'cols' : list('ABCDEFGHIJLMNPQRSTUVZ') + ['AE', 'AH'],
        'start_row' : 15 },
    'Overtime & Other Personnel' : {
        'cols' : list('ABCDEFGHIJKLNOPQRX'),
        'start_row' : 15 },
    'Non-Personnel' : {
        'cols': list('ABCDEFGHIJKLMOPQRSTUVWXYZ'),
        'start_row' : 19 },
    'Revenue' : {
        'cols': list('ABCDEFGHIJKLMNPQRSTUVW'),
        'start_row' : 15 }
}

# ================== Script functions ===================================

def move_data(detail_sheet, destination_file):
    # Load DS worksheet
    source_wb = load_workbook(f'{source_folder}/{detail_sheet}', data_only=False)

    # iterate through the tabs
    for sheet in SHEETS.keys:
        source_ws = source_wb[sheet]

    # TODO
    # copy over filtered data
    copy_cols(columns_to_move, column_destinations, filtered_data, 
            destination_ws, destination_row_start)

    # Save the destination workbook
    destination_wb.save(f'{dest_folder}/{detail_sheet}')
    print(f'saved {detail_sheet}')

#================== Program on run ============================================

def main():
    # copy the template
    shutil.copy(template_file, dest_file)
    for detail_sheet in os.listdir(source_folder):
        # only attempt on excel sheets (exlude folder, etc)
        if '.xlsx' not in detail_sheet:
            continue
        try: 
            move_data(detail_sheet)
        except:
            print(f'failed to copy data from {detail_sheet}')


if __name__ == '__main__':
    main()
