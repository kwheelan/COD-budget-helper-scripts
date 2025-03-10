"""
Automate the well-formatted HARI report
"""

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import PatternFill, Alignment, numbers


# File paths
DATA = 'input_data/HARI'
file = os.listdir(DATA)[0]
filepath = os.path.join(DATA, file)
OUTPUT = 'output/HARI/'
output_file = os.path.join(OUTPUT, 'converted_HARI.xlsx')

# Read in data
df = pd.read_excel(filepath, sheet_name='Sheet4', 
                   dtype=str,
                   skiprows=9#, nrows=100
                   )
# print(df)

# Delete columns
to_delete = ['FAC_Numbers',
             'FAC_Names',
             'PARENT2_VALUE',
             'PARENT2_DESCRIPTION',
             'PARENT4_VALUE',
             'PARENT4_DESCRIPTION',
             'JournalStatus',
             'FundsReservedStatus']
df.drop(to_delete, axis=1, inplace=True)

# Rename headers
headers = ['Fund #', 'Fund Name',
           'Appropriation #', 'Appropriation Name',
           'Cost Center #', 'Cost Center Name',
           'Account Type',
           'Object #', 'Object Name',
           'Project #', 'Project Name',
           'Activity #', 'Activity Name',
           'Interfund #', 'Interfund Name',
           'Adopted Budget', 'Amended Budget', 
           'Commitment', 'Obligation', 'Actual', 'Funds Available'
           ]
df.columns = headers

# Add new columns
df['Future #'] = '000000'
df['Future Name'] = 'Undefined Future'
df['Account String'] = df.apply(lambda row: f"{row['Fund #']}-{row['Appropriation #']}-{row['Cost Center #']}-{row['Object #']}-{row['Project #']}-{row['Activity #']}-{row['Interfund #']}-{row['Future #']}", axis=1)
df['Dept #'] =  df['Cost Center #'].str[:2]

# Adjust column order
headers = ['Account Type', 'Dept #',
           'Fund #', 'Fund Name',
           'Appropriation #', 'Appropriation Name',
           'Cost Center #', 'Cost Center Name',
           'Object #', 'Object Name',
           'Project #', 'Project Name',
           'Activity #', 'Activity Name',
           'Interfund #', 'Interfund Name',
           'Future #', 'Future Name',
           'Account String',
           'Adopted Budget', 'Amended Budget', 
           'Commitment', 'Obligation', 'Actual', 'Funds Available'
           ]
df = df[headers]

# Data "compressed" to put repeat account strings on a single line
group_columns = headers[0:18]
value_cols = headers[19:]
df[value_cols] = df[value_cols].astype(float)
df = df.groupby(group_columns, as_index=False).sum()

# Rows with $0.00 in columns T thru Y deleted
df = df.loc[~(df[value_cols].sum(axis=1) == 0)]
# Reset the index of the DataFrame
df.reset_index(drop=True, inplace=True)

# Save DataFrame to an Excel file with table formatting
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Write DataFrame to Excel, starting from the first row
    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=2)
    worksheet = writer.sheets['Sheet1']

    # Get the dimensions of the DataFrame
    rows, cols = df.shape

    # Adjust column widths to fit the header and contents
    for col in worksheet.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        # add extra width for table arrows
        worksheet.column_dimensions[col[0].column_letter].width = max_length + 6

    # Define the range for the table
    table_range = f"A3:{chr(65 + cols - 1)}{rows + 3}"

    # Create and style the table (with a style that has gray headers)
    table = Table(displayName="Table1", ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium2",  # Choose a style that has gray headers
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,  # Set to True if you want striped rows
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    worksheet.add_table(table)

    # Set the header color to gray
    gray_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    for cell in worksheet["3:3"]:
        cell.fill = gray_fill

    # Group name columns for a collapsible section
    for col in 'DFHJLNPR':
        worksheet.column_dimensions.group(col, col, hidden=False)

    # Total & Subtotal formulas above columns T thru Y
    last_row_number = rows + 4  # accounting for two inserted rows and one header row
    # Add labels
    worksheet['S1'] = 'Total:'
    worksheet['S2'] = 'Subtotal:'
    # Right-align every cell
    worksheet['S1'].alignment = Alignment(horizontal='right')
    worksheet['S2'].alignment = Alignment(horizontal='right')
    # 
    for idx, col_letter in enumerate('tuvwxy', start=1):  
        worksheet[f"{col_letter}{1}"] = f"=SUBTOTAL(9,{col_letter}4:{col_letter}{last_row_number - 1})"
        worksheet[f"{col_letter}{2}"] = f"=SUM({col_letter}4:{col_letter}{last_row_number - 1})"

    # Apply right alignment and currency format to specified columns
    currency_columns = ['Adopted Budget', 'Amended Budget', 
           'Commitment', 'Obligation', 'Actual', 'Funds Available']

    for col in worksheet.iter_cols(min_col=1, max_col=cols, min_row=1):
        header = col[2].value
        for cell in col[:2] + col[3:]:  # Skip the header cell
            # Right-align cells
            cell.alignment = Alignment(horizontal='right')
            # Apply currency format to specified columns
            if header in currency_columns:
                cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE