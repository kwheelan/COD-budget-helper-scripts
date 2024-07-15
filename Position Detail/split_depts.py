import pandas as pd
import os
# from openpyxl import load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill

# fiscal year
FY = 25

# columns to keep
COLUMNS = [
    'Department',
    'Fund',
    'Fund Name',
    'Appropriation',
    'Appropriation Name',
    'Cost Center',
    'Cost Center Name',
    'Account String',
    'Job Code',
    'Job Title',
    'FY25 Adopted FTE'
]

# months of the fiscal year
MONTHS = [
    'Apr',
    'May',
    'June',
    'Jul',
    'Aug',
    'Sept',
    'Oct',
    'Nov',
    'Dec',
    'Jan',
    'Feb',
    'March',
    'April',
    'May',
    'June',
    'Placeholder for sum'
]

NEW_CAL_YEAR_IX = MONTHS.index('Jan')

def adjustColWidth(sheet):
    # Auto-adjust column widths
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        # give room for table arrow
        adjusted_width = (max_length + 5)
        sheet.column_dimensions[column].width = adjusted_width

def selectCols(df):
    return df[COLUMNS]

def remove_numbers_dashes(df, columns):
    for col in columns:
        df.loc[:, col] = df[col].astype(str).str.replace(r'[\d-]', '', regex=True).str.strip()
    return df

def addCols(df):
    for i in range(len(MONTHS)):
        year = FY - (i < NEW_CAL_YEAR_IX)
        label = f"{MONTHS[i]}'{year} PAs"
        df = df.assign(**{label: pd.NA})
    return df

def addFormulaCol(sheet):
    # Add a new column 'FY24 Amended FTE's' that sums the previous 13 columns
    last_col = sheet.max_column
    formula_col_letter = get_column_letter(last_col)
    sheet.cell(row=1, column=last_col, value=f"FY{FY} Amended FTE")
    
    for row in range(2, sheet.max_row + 1):
        sum_range = f"{get_column_letter(last_col-len(MONTHS))}{row}:{get_column_letter(last_col-1)}{row}"
        sheet[f"{formula_col_letter}{row}"] = f"=SUM({sum_range})"

def aggregateByJob(df):
    # add an ID column
    return df.groupby(['Account String', 'Job Code']).agg({
        'Department' : 'first',
        'Fund' : 'first',
        'Fund Name' : 'first',
        'Appropriation' : 'first',
        'Appropriation Name' : 'first',
        'Cost Center' : 'first',
        'Cost Center Name' : 'first',
        'Account String' : 'first',
        'Job Code' : 'first',
        'Job Title' : 'first',
        'FY25 Adopted FTE': 'sum'
    })

def addSubtotalRow(sheet, n_rows, n_cols):
    # Add a sum row at the bottom for the last 14 columns
    sum_row = n_rows + 1
    sheet[f"A{sum_row}"] = "Total"
    for col in range(n_cols - len(MONTHS), n_cols + 1):
        col_letter = get_column_letter(col)
        sheet[f"{col_letter}{sum_row}"] = f"=SUM({col_letter}2:{col_letter}{n_rows})"

def format_columns_decimals(sheet, columns, n_row, decimal_places=2):
    """
    Apply number formatting to columns to have exactly two decimal places.
    """
    number_format = f"0.{''.join(['0'] * decimal_places)}"

    for col_idx in columns:
        for row in range(1, n_row + 1):
            sheet.cell(row=row, column=col_idx).number_format = number_format

def addStyling(sheet, n_rows, n_cols):
    # adjust col widths to fit content
    adjustColWidth(sheet)

    # Style the header row: bold font and white font
    header_font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.font = header_font

    # Style the total row: dark blue fill and white font
    total_fill = PatternFill(start_color="00008B", end_color="00008B", fill_type="solid")
    for cell in sheet[n_rows]:
        cell.font = header_font
        cell.fill = total_fill
    
    # add decimal styling to all FTE cols
    cols = range(n_cols - len(MONTHS), n_cols + 1) 
    format_columns_decimals(sheet, cols, n_rows)


# Function to add DataFrame to an openpyxl workbook and convert to table
def df_to_workbook(df, table_name, output_file):
    # trim data to delete unneccessary cols
    df = selectCols(df)
    df = remove_numbers_dashes(df, ['Appropriation Name', 'Cost Center Name', 'Fund Name'])
    df = aggregateByJob(df)
    df = addCols(df)
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

            wb = writer.book
            sheet = wb.active

            addFormulaCol(sheet)

            # get dimensions
            n_cols = df.shape[1]
            n_rows = df.shape[0] + 1  # including the header row

            # add row for substotals
            addSubtotalRow(sheet, n_rows, n_cols)
            n_rows += 1

            # create table with styling
            tab = Table(displayName=table_name, ref=f"A1:{get_column_letter(n_cols)}{n_rows}")
            style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            sheet.add_table(tab)

            # ADD SYLING
            addStyling(sheet, n_rows, n_cols)

            wb.save(output_file)  # Save the workbook

    except Exception as e:
        print(f"Failed to write workbook {output_file}: {e}")


def create_workbooks(input_file, sheet_name, split_column):

    # The actual data to be split
    df = pd.read_excel(input_file, sheet_name=sheet_name, skiprows=11)

    # Get unique values from the column
    unique_values = df[split_column].unique()

    # Create a directory to save the new workbooks
    output_dir = 'Position Detail by Department'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Loop through unique values and create a new DataFrame for each
    for value in unique_values:
        subset_df = df[df[split_column] == value]

        # deal with stray values
        if value == '-':
            value = '0 - No Department Given'

        output_file = f"{output_dir}/{value}.xlsx"

        dept_num = value.split(' ')[0]

        # Write each subset to a new workbook with formatting and table conversion
        df_to_workbook(subset_df, table_name=f"Table_{dept_num}", output_file=output_file)

    print("Workbooks created successfully in the 'All Departments' directory.")

if __name__ == "__main__":
    # Define input parameters
    input_file = "Full_Position_Data.xlsx"
    sheet_name = f"FY{FY} Position Detail"
    split_column = "Department Name"
    
    create_workbooks(input_file, sheet_name, split_column)