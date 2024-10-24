
"""
K. Wheelan 
October 2024

Utility functions for handling Excel sheets
"""

from copy import copy

# default exports
__all__ = ['copy_cell', 'copy_cols']

#----------------------- Functions -----------------------

def copy_cell(src_cell, dest_cell, 
              keep_dest_style = True):
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

# def empty_row(sheet, row):
#     for col_ix in range(1, sheet.max_column):
#         value = sheet.cell(row=row, column=col_ix).value
#         if(value not in (None, 0, '', '-', '---', '$ -')):
#             return False

# def next_open_row(sheet):
#     """
#     This function finds the first completely empty row and ensures that the next row is 
#     also empty in the given sheet.
#     :param sheet: An openpyxl worksheet object.
#     :return: The row number of the first completely empty row.
#     """ 
#     max_row = sheet.max_row

#     for row_index in range(1, max_row + 1):
#         current_row_empty = empty_row(sheet, row_index)

#         next_row_empty = (row_index == max_row) or empty_row(sheet, row_index+1)
        
#         if current_row_empty and next_row_empty:
#             return row_index

#     return max_row + 1

def last_data_row(sheet, n_header_rows):
    for row in range(n_header_rows+1, sheet.max_row):
        dept = sheet.cell(row=row, column=1).value
        if (not dept) or (dept == '-'):
            return row - 1
    return sheet.max_row

# function to copy given columns 
def copy_cols(source_ws, destination_ws, columns_to_move, 
              column_destinations = None, destination_row_start=None, 
              source_row_start=None, keep_style=False):
    """
    Copy columns one worksheet to another

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

    # convert columns if necessary
    if (type(columns_to_move[0]) is str):
        columns_to_move = [letter_to_col(l) for l in columns_to_move]
    if (type(column_destinations[0]) is str):
        column_destinations =  [letter_to_col(l) for l in column_destinations]

    last_row_to_copy = last_data_row(source_ws, source_row_start)

    for col in columns_to_move:
        for row in range(source_row_start, last_row_to_copy):
            source_cell = source_ws.cell(row = row, column = col + 1)
            dest_cell = destination_ws.cell(row = destination_row_start + (row - source_row_start), 
                                            column = col + 1)
            copy_cell(source_cell, dest_cell, keep_dest_style=keep_style)