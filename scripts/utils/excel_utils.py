
"""
K. Wheelan 
October 2024

Utility functions for handling Excel sheets
"""

from copy import copy

# default exports
__all__ = ['copy_cell', 'copy_cols']

#----------------------- Functions -----------------------

def copy_cell(src_cell, dest_cell, keep_dest_style = True):
    """
    Function to copy cell and optionally its style """
    dest_cell.value = src_cell.value
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

def find_next_open_row(sheet):
    """
    This function finds the first completely empty row in the given sheet.
    :param sheet: An openpyxl worksheet object.
    :return: The row number of the first completely empty row.
    """
    for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
        if all(cell is None for cell in row):
            return sheet.iter_rows(min_row=1, max_col=1).index(row) + 1
            
    return sheet.max_row + 1


# function to copy given columns 
def copy_cols(source_data, destination_ws, columns_to_move, column_destinations = None, destination_row_start=None):
    """
    Copy columns from filtered data to a destination worksheet.

    Args:
        columns_to_move (list of int): Column indices from the source data to be moved (0-based or letters).
        column_destinations (list of str): Corresponding column indices in the destination worksheet (0-based or letters).
        source_data (list of list): Source data as a list of rows.
        destination_ws (worksheet): Destination worksheet.
        destination_row_start (int): Row number in destination worksheet to start copying (1-based index).

    Returns:
        None
    """
    # Initialize default settings
    if not column_destinations:
        column_destinations = columns_to_move

    # If no specified start row, paste into next open row
    if not destination_row_start:
        find_next_open_row(destination_ws)

    # convert columns if necessary
    columns_to_move = [col_to_letter(n) for n in columns_to_move if type(n) is int]
    column_destinations = [letter_to_col(l) for l in column_destinations if type(l) is str]

    for col_ix in range(len(columns_to_move)):
        for row_ix, row_data in enumerate(source_data):
            # Get the right column value from the source data
            try:
                source_col_value = row_data[columns_to_move[col_ix]]
            except:
                print(f'row_data: {row_data}')
                print(f'columns_to_move: {columns_to_move}')
                print(f'columns_to_move[col_ix]: {columns_to_move[col_ix]}')

            # get column to paste to
            new_col = column_destinations[col_ix]
            # Determine the destination cell
            dest_cell = destination_ws[f'{new_col}{row_ix + destination_row_start}']
            
            # Copy the value to the destination cell
            dest_cell.value = source_col_value