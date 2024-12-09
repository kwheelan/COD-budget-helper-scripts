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
        pattern = rf"('{self.sheet}'!\$){self.old_col}(\$[\d]+)?(:\$)?({self.old_col})?(\$[\d]+)?"
        
        def repl(match):
            sheet_part = match.group(1)
            first_reference_part = match.group(2) or ''
            range_separator = match.group(3) or ''
            second_col_part = match.group(4) or ''
            second_reference_part = match.group(5) or ''
            
            if range_separator:
                # If range_separator is present, handle whole column and regular ranges
                return f"{sheet_part}{self.new_col}{first_reference_part}{range_separator}{self.new_col}{second_reference_part}"
            else:
                # Single cell or whole column reference
                return f"{sheet_part}{self.new_col}{first_reference_part}"

        # Replace the pattern in the text
        result = re.sub(pattern, repl, formula)
        return(result) 

#======================================================================
# Constants
#======================================================================

SOURCE_FOLDER = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/08. Analyst Review/'

REPLACEMENTS = [
    Replacer('FTE, Salary-Wage, & Benefits', 'AG', 'AR'),
    Replacer('FTE, Salary-Wage, & Benefits', 'U', 'AO'),
    Replacer('FTE, Salary-Wage, & Benefits', 'AA', 'AP'),
    Replacer('FTE, Salary-Wage, & Benefits', 'AF', 'AQ'),
    Replacer('Overtime & Other Personnel', 'R', 'Y'),
    Replacer('Overtime & Other Personnel', 'V', 'Z'),
    Replacer('Overtime & Other Personnel', 'W', 'AC'),
    Replacer('Non-Personnel', 'Y', 'AC')
]

POLICE_REPLACEMENTS = REPLACEMENTS.copy()
POLICE_REPLACEMENTS[0] = Replacer('FTE, Salary-Wage, & Benefits', 'AH', 'AS')
POLICE_REPLACEMENTS[1] = Replacer('FTE, Salary-Wage, & Benefits', 'V', 'AP')
POLICE_REPLACEMENTS[2] = Replacer('FTE, Salary-Wage, & Benefits', 'AB', 'AQ')
POLICE_REPLACEMENTS[3] = Replacer('FTE, Salary-Wage, & Benefits', 'AG', 'AR')

SHEETS = ['Dept Summary', 'Initiatives Summary']

#======================================================================
# Helpers
#======================================================================

def replace_all(formula, police = False):
    if police:
        replacements = POLICE_REPLACEMENTS
    else: 
        replacements = REPLACEMENTS
    for replacer in replacements:
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
    # check if it's the police DS
    police = 'Police' in file

    # open file 
    wb = load_workbook(file, data_only=False)

    # make replacements on both dept summary and init summary
    for sheet in SHEETS:
        ws = wb[sheet]
        
        # Iterate through all cells in the sheet
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='): # and 'SUMIFS' in cell.value:
                    # replace formula
                    original_formula = cell.value
                    new_formula = replace_all(original_formula, police)
                    if original_formula != new_formula:
                        # print optional message
                        if verbose:
                            print(f"Replacing {original_formula} with {new_formula} at cell {cell.coordinate}")
                        cell.value = new_formula

    # save workbook
    name = file.split('\\')[-1]
    wb.save(f'output/formula_replacements/{name}')
    print(f'Saved {name}')

#======================================================================
# Main
#======================================================================

def test():
    #text = """'FTE, Salary-Wage, & Benefits'!$AG$2 + SUM('FTE, Salary-Wage, & Benefits'!$AG$15:$AG$1000)"""
    text = """=SUMIFS('FTE, Salary-Wage, & Benefits'!$AG:$AG,'FTE, Salary-Wage, & Benefits'!$L:$L,"Baseline",'FTE, Salary-Wage, & Benefits'!$G:$G,$W7)
+SUMIFS('Overtime & Other Personnel'!$W:$W,'Overtime & Other Personnel'!$N:$N,"Baseline",'Overtime & Other Personnel'!$G:$G,$W7)
+SUMIFS('Non-Personnel'!$Y:$Y,'Non-Personnel'!$O:$O,"Baseline*",'Non-Personnel'!$G:$G,$W7)
"""
    print(replace_all(text))

def main():
    files = create_file_list()
    for file in files:
        edit_formulas(file)

if __name__ == '__main__':
    #test()
    main()