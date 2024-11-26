
"""
K. Wheelan 
October 2024

Utility functions for handling Excel sheets
"""

from copy import copy
import re

# default exports
__all__ = ['copy_cell', 'copy_cols']

#----------------------- Functions -----------------------

def copy_cell(src_cell, dest_cell, 
              keep_dest_style = True,
              row_offset = 0):
    """
    Function to copy cell and optionally its style """

    # adjust formula to correct row
    dest_cell.value = adjust_formula( src_cell.value, row_offset)

    # copy over original style if desired
    if (not keep_dest_style) and src_cell.has_style:
        dest_cell.font = copy(src_cell.font)
        dest_cell.border = copy(src_cell.border)
        dest_cell.fill = copy(src_cell.fill)
        dest_cell.number_format = src_cell.number_format
        dest_cell.protection = copy(src_cell.protection)
        dest_cell.alignment = copy(src_cell.alignment)

def letter_to_col(str):
    """
    Convert Excel column name (ex. AA or G) to column number (0-based)
    """
    if len(str) == 1:
        letter = str.upper()
        return ord(letter) - 65
    return 26 * (letter_to_col(str[0]) + 1) + letter_to_col(str[1:]) 

def col_to_letter(n):
    """
    Convert Excel column number (0-based) to letter (ex. A or BD)
    """
    if n < 26:
        return chr(n + 65)
    return col_to_letter( (n // 26) - 1) + col_to_letter(n % 26)

def adjust_formula(formula, row_offset=0):
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    # Remove external workbook references by removing anything inside square brackets
    formula = re.sub(r'\[.*\]', '', formula)

    # Regex to match cell references, ensuring optional sheet names are correct
    cell_ref_pattern = re.compile(r"""
        (
            (?:'[^']+'!)?     # Optional sheet name in single quotes, followed by '!'
            \$?[A-Z]+         # Column reference, optionally with $
            \$?(\d+)          # Row reference, capturing the row number
        )
    """, re.VERBOSE)

    def adjust_cell(match):
        full_match = match.group(0)  # Full match of the cell reference
        col = match.group(1) # column
        row = match.group(2)  # The row number part

        # Adjust the row number only if it is not an absolute reference
        if not full_match.endswith('$') and 'FY' not in col:
            new_row = str(int(row) + row_offset)
            return full_match.replace(row, new_row)
        return full_match

    # Apply adjustments to the formula
    adjusted_formula = cell_ref_pattern.sub(adjust_cell, formula)
    # adjusted_formula = adjusted_formula.replace('_xlfn.', '')
    return adjusted_formula

def last_data_row(sheet, n_header_rows):
    for row in range(n_header_rows+1, sheet.max_row):
        dept = sheet.cell(row=row, column=1).value
        if (not dept) or (dept == '-'):
            return row - 1
    return sheet.max_row

def copy_cols(source_ws, destination_ws, columns_to_move, 
              column_destinations = None, destination_row_start=None, 
              source_row_start=None, keep_style=False, source_row_end = None):
    """
    Copy columns from one worksheet to another

    Returns:
        None
    """
    # Initialize default settings
    if not column_destinations:
        column_destinations = columns_to_move

    # If no specified start row, paste into next open row
    if not source_row_start:
        source_row_start = 1
    if not destination_row_start:
        destination_row_start = last_data_row(destination_ws, source_row_start)+1

    # convert columns to numbers from letters if necessary
    if (type(columns_to_move[0]) is str):
        columns_to_move = [letter_to_col(l) for l in columns_to_move]
    if (type(column_destinations[0]) is str):
        column_destinations =  [letter_to_col(l) for l in column_destinations]

    # Figure out the end of the range of data to copy
    if not source_row_end:
        source_row_end = last_data_row(source_ws, source_row_start)
    # The adjustment factor from the source row to the destination row
    row_offset = destination_row_start - source_row_start

    # Actually copy over the data cell by cell
    for col in columns_to_move:
        for row in range(source_row_start, source_row_end):
            source_cell = source_ws.cell(row = row, column = col + 1)
            dest_cell = destination_ws.cell(row = row + row_offset, 
                                            column = col + 1)
            copy_cell(source_cell, dest_cell, keep_dest_style=keep_style, row_offset=row_offset)


