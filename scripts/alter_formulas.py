""" 
K. Wheelan December 2024

A script to change all summary formulas to pull from the approved budget recommendation
"""

import re
import os
from openpyxl import load_workbook

#======================================================================
# Replacer class
#======================================================================

class Replacer:
    def __init__(self, sheet, old_col, new_col):
        self.sheet = sheet
        self.old_col = old_col
        self.new_col = new_col

    def sheet(self):
        return self.sheet
    
    def replace_function(self, formula):
        # Regular expression to match the pattern
        pattern = rf"('{self.sheet}'!\$){self.old_col}(.+:\$){self.old_col}(.+)"

        # Replacement pattern with backreferences
        replacement = rf"\1{self.new_col}\2{self.new_col}\3"

        # Replace the pattern in the text
        result = re.sub(pattern, replacement, formula)

        return(result) 

#======================================================================
# Constants
#======================================================================

SOURCE_FOLDER = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/08. Analyst Review/'

REPLACEMENTS = [
    Replacer('FTE, Salary-Wage, & Benefits', 'AG', 'AR'),
    Replacer('FTE, Salary-Wage, & Benefits', 'U', 'AO'),
    Replacer('Overtime & Other Personnel', 'W', 'AC'),
    Replacer('Non-Personnel', 'Y', 'AC')
]

SHEETS = ['Dept Summary', 'Initiatives Summary']

#======================================================================
# Helpers
#======================================================================

def replace_all(formula):
    for replacer in REPLACEMENTS:
        formula = replacer.replace_function(formula)
    return formula

def find_DS(folder, verbose = False):
    # Get full file path
    folder_fp = os.path.join(SOURCE_FOLDER, folder)

    # grab list of all files in that folder
    files = os.listdir(folder_fp)

    # generate list of reviewed detail sheets
    reviewed_DS = []
    for file in files:
        if ('(Analyst Review)' in file or 'Library' in file) and '.xlsx' in file:
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

def edit_formulas(file, verbose=False):
    # open file 
    wb = load_workbook(file, data_only=False)
    for sheet in SHEETS:
        ws = wb[sheet]
        
        # Iterate through all cells in the sheet
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('=') and 'SUMIFS' in cell.value:
                    # replace formula
                    original_formula = cell.value
                    new_formula = replace_all(original_formula)
                    if original_formula != new_formula:
                        # print optional message
                        if verbose:
                            print(f"Replacing {original_formula} with {new_formula} at cell {cell.coordinate}")
                        cell.value = new_formula

    # save workbook
    name = file.split('\\')[-1]
    wb.save(f'output/formula_replacements/{name}')
    print(f'saved as {name}')

#======================================================================
# Main
#======================================================================

def main():
    files = create_file_list()
    for file in files[0:1]:
        #print(file)
        edit_formulas(file)

if __name__ == '__main__':
    main()