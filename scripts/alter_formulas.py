""" 
K. Wheelan December 2024

A script to change all summary formulas to pull from the approved budget recommendation
"""

import re
import os
from openpyxl import load_workbook, worksheet
from openpyxl.worksheet.datavalidation import DataValidation

#======================================================================
# Constants
#======================================================================

analyst_review_folder = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/08. Analyst Review/'
budget_director_review_folder = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/08A. Deputy Budget Director Recommend/'
SOURCE_FOLDER = budget_director_review_folder

analyst_review_replacements = {
    'FTE, Salary-Wage, & Benefits' : {
        'AG' : 'AR',
        'U': 'AO',
        'AA': 'AP',
        'AF': 'AQ'},
    'Overtime & Other Personnel' : {
        'R' : 'AC',
        'V' : 'AD',
        'W' : 'AE'},
    'Revenue' : {'V' : 'Z'},
    'Non-Personnel' : {'Y' : 'AC'}
}

analyst_review_replacements = {
    'FTE, Salary-Wage, & Benefits' : {
        'AG' : 'AR',
        'U': 'AO',
        'AA': 'AP',
        'AF': 'AQ'},
    'Overtime & Other Personnel' : {
        'R' : 'AC',
        'V' : 'AD',
        'W' : 'AE'},
    'Revenue' : {'V' : 'Z'},
    'Non-Personnel' : {'Y' : 'AC'}
}

police_replacements = analyst_review_replacements.copy()
police_replacements['FTE, Salary-Wage, & Benefits'] = {
        'AH' : 'AS',
        'V': 'AP',
        'AB': 'AQ',
        'AG': 'AR'}

director_review_replacements = {
    'FTE, Salary-Wage, & Benefits' : {
        'AO' : 'AZ',
        'AP' : 'BA',
        'AQ' : 'BB',
        'AR' : 'BC'},
    'Overtime & Other Personnel' : {
        'AC' : 'AK',
        'AD' : 'AL',
        'AE' : 'AM'},
    'Revenue' : {'Z' : 'AD'},
    'Non-Personnel' : {'AC' : 'AG'}
}

#SHEETS = ['Dept Summary', 'Initiatives Summary']
SHEETS = ['Budget Director Summary', 'Initiatives Summary']

REPLACEMENT_DICT = director_review_replacements

SAVE_TO = 'output/formula_replacements/backups'

KEY = '(Deputy Budget Director)' #'(Analyst Review)

EXCLUDE = ['Police', 'DSLP']


def dd_ref(range) : return f"'Drop-Down Menus'!{range}"

DROPDOWN_OPTIONS = {
    'Baseline' : dd_ref('$A$7:$A$8'),
    'Baseline-Capital' : dd_ref('$A$2:$A$5'),
    'Recurring' : dd_ref('$C$2:$C$3'),
    'Employee-Type' : dd_ref('$G$2:$G$8'),
    'Position-Status' : dd_ref('$I$2:$I$4'),
    'Service' : dd_ref('$M$2:$M$11'),
    'Approval' : dd_ref('$E$2:$E$4')
}

DROPDOWNS = {
    'FTE, Salary-Wage, & Benefits' : {
        'L15:L10000' : DROPDOWN_OPTIONS['Baseline'],
        'M15:M10000' : DROPDOWN_OPTIONS['Recurring'],
        'P15:P10000' : DROPDOWN_OPTIONS['Employee-Type'],
        'Q15:Q10000' : DROPDOWN_OPTIONS['Position-Status'],
        'S15:S10000' : DROPDOWN_OPTIONS['Service'],
        'AN15:AN10000' : DROPDOWN_OPTIONS['Approval'],
        'AY15:AY10000' : DROPDOWN_OPTIONS['Approval']
    },
    'Overtime & Other Personnel' : {
        'N15:N10000' : DROPDOWN_OPTIONS['Baseline'],
        'O15:O10000' : DROPDOWN_OPTIONS['Recurring'],
        'Q15:Q10000' : DROPDOWN_OPTIONS['Service'],
        'AB15:AB10000' : DROPDOWN_OPTIONS['Approval'],
        'AJ15:AJ10000' : DROPDOWN_OPTIONS['Approval']
    },
    'Non-Personnel' : {
        'O19:O10000' : DROPDOWN_OPTIONS['Baseline-Capital'],
        'P19:P10000' : DROPDOWN_OPTIONS['Recurring'],
        'X19:X10000' : DROPDOWN_OPTIONS['Service'],
        'AB19:AB10000' : DROPDOWN_OPTIONS['Approval'],
        'AF19:AF10000' : DROPDOWN_OPTIONS['Approval']
    },
    'Revenue' : {
        'P15:P10000' : DROPDOWN_OPTIONS['Baseline-Capital'],
        'Q15:Q10000' : DROPDOWN_OPTIONS['Recurring'],
        'Y15:Y10000' : DROPDOWN_OPTIONS['Approval'],
        'AC15:AC10000' : DROPDOWN_OPTIONS['Approval']
    },
}

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
        pattern = rf"('*{self.sheet}'*!\$){self.old_col}(\$[\d]+)?(:\$)?({self.old_col})?(\$[\d]+)?"
        
        def repl(match):
            sheet_part = match.group(1)
            first_reference_part = match.group(2) or ''
            range_separator = match.group(3) or ''
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


def build_replacements(dict):
    replacer_list = []
    for sheet in dict.keys():
        for old_col in dict[sheet]:
            new_col = dict[sheet][old_col]
            replacer = Replacer(sheet, old_col, new_col)
            replacer_list.append(replacer)
    return replacer_list


POLICE_REPLACEMENTS = []
REPLACEMENTS = build_replacements(REPLACEMENT_DICT)

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

def find_DS(folder, keyword=KEY, exclude=EXCLUDE, verbose=False):
    # Get full file path
    folder_fp = os.path.join(SOURCE_FOLDER, folder)

    # grab list of all files in that folder
    files = os.listdir(folder_fp)

    # generate list of reviewed detail sheets
    reviewed_DS = []
    for file in files:
        if (keyword in file) and ('.xlsx' in file) and not (sum([(word in file) for word in exclude])):
            reviewed_DS.append(os.path.join(SOURCE_FOLDER, folder, file))

    # return message
    message = None
    if len(reviewed_DS) > 1:
        message = f'Multiple potential reviewed detail sheets: {reviewed_DS}'
    elif len(reviewed_DS) == 0:
        message = f'No reviewed detail sheet found in {folder}'
    if verbose and message:
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
            DS_list += find_DS(folder, keyword=KEY, exclude=EXCLUDE, verbose=verbose)

    return DS_list

def edit_formulas(file, verbose=False, save_to=False, backup=False):
    # check if it's the police DS
    police = 'Police' in file

    # open file 
    wb = load_workbook(file, data_only=False)

    if(backup):
        wb.save(backup)

    # make replacements on both dept summary and init summary
    for sheet in SHEETS:
        ws = wb[sheet]
        
        # Iterate through all cells in the sheet
        for row in ws.iter_rows():
            for cell in row:

                if isinstance(cell.value, str) and "=" in cell.value:
                    # store old formula
                    original_formula = cell.value

                    # replace formula
                    new_formula = replace_all(original_formula, police)

                    if original_formula != new_formula:
                        # print optional message
                        if verbose:
                            print(f"Replacing {original_formula} with {new_formula} at cell {cell.coordinate}")
                        cell.value = new_formula

    # Add back drop down menus
    add_dropdowns(wb)

    # save workbook
    name = file.split('\\')[-1]
    if save_to:
        wb.save(f'{SAVE_TO}/{name}')
    else:
        wb.save(file)
    print(f'Saved {name}')

def add_dropdowns(wb):
    for sheet in DROPDOWNS.keys():
        ws = wb[sheet]

        # for each column
        for cell_range in DROPDOWNS[sheet].keys():
                
            # build data validation
            dv = DataValidation(
                type="list",
                formula1=f'={DROPDOWNS[sheet][cell_range]}',
                showDropDown=False
            )

            # Add the data validation to a specific cell, e.g., A1
            ws.add_data_validation(dv)
            dv.add(cell_range)


#======================================================================
# Main
#======================================================================

def test():
    text = """'FTE, Salary-Wage, & Benefits'!$AG$2 + SUM('FTE, Salary-Wage, & Benefits'!$AG$15:$AG$1000)"""
    print(replace_all(text))

def main():
    files = create_file_list()[1:2]
    for file in files:
        edit_formulas(file, backup=SAVE_TO)

if __name__ == '__main__':
    #test()
    main()