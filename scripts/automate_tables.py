"""
K. Wheelan February 2025

"""

import os
import pandas as pd
from pylatex import Document, Section, Subsection, Table, NoEscape, Package

#====================== Constants =========================================

# File paths
DATA = 'input_data/model_copy'
file = os.listdir(DATA)[0]
filepath = os.path.join(DATA, file)
OUTPUT = 'output/budget_book/'

SHEETS_TO_LOAD = [
    'Section A FTE'
]

# ======================= Classes ========================================

class Budget_Table:
    def __init__(self, df):
        # Extract data
        self.table_data = df.iloc[6:44, 1:10]
        self.adjust_headers()
        # Get headers
        self.topline_header = df.iloc[0, 1]
        self.other_header_lines = [df.iloc[i, 1] for i in range(2, 5)]
        self.clean_table_data()

    def clean_nas(self):
        self.table_data.fillna('', inplace=True)

    def clean_table_data(self):
        self.clean_nas()
        # Round all float values to 2 decimal places
        self.table_data = self.table_data.applymap(self.round_floats)
        # Apply the escaping & function to each element in the DataFrame
        self.table_data = self.table_data.applymap(self.escape_ampersands)
        self.table_data = self.table_data.applymap(self.remove_new_lines)

    def adjust_headers(self):
        # Adjust column names
        self.table_data.columns = self.table_data.iloc[0]
        columns_list = list(self.table_data.columns)
        columns_list[1] = ''
        columns_list[2] ='Department'
        # remove extra spaces or new lines
        for i in range(len(columns_list)): 
            columns_list[i] = columns_list[i].strip().replace('\n','')
        # Drop column names row from data and reset
        self.table_data.columns = columns_list
        self.table_data = self.table_data[1:].reset_index(drop=True) 

    @staticmethod
    def escape_ampersands(text):
        # Replace & with \& in all string elements
        if isinstance(text, str):
            return text.replace('&', r'\&')
        return text
    
    @staticmethod
    def remove_new_lines(text):
        # Replace & with \& in all string elements
        if isinstance(text, str):
            return text.replace('\n', ' ').replace('  ', ' ')
        return text
    
    @staticmethod
    def round_floats(value):
        if isinstance(value, float) or isinstance(value, int):
            # avoid float imprecision errors
            value = round(value, 3)
            if value % 1 == 0:
                # then just add comma
                return f'{value:,.0f}'
            if round(value,1) == value:
                # then round to one decimal
                return f'{value:,.1f}'
            # otherwise, round to 2 places
            return f'{value:,.2f}'
        return value

    def get_table_data(self):
        return self.table_data

    def main(self):
        return self.topline_header

    def subheaders(self):
        return self.other_header_lines

# ======================= Functions =======================================

def dataframe_to_latex_custom(df):
    # Convert the DataFrame to LaTeX with custom formatting for bold headers
    
    # Define custom column widths
    column_format = r'|>{\centering\arraybackslash}p{2.5cm}|' + 'p{0.5cm}|' + 'p{3cm}|' + r'>{\centering\arraybackslash}p{1.5cm}|' * (len(df.columns) - 3)
    
    latex_code = df.to_latex(
        index=False, 
        escape=False, 
        column_format=column_format, 
        header=True,  # Automatically handles headers
        bold_rows=False,
        )
    
    # get just table data
    lines = latex_code.split('\n')
    header = '\n'.join(lines[0:2])
    footer = '\n'.join(lines[-2:])
    rows = lines[2:len(lines)-2]
    rows = '\n'.join(rows)

    # Split into table data
    rows = rows.replace('midrule', r'midrule\\')
    rows = rows.split(r'\\')

    # Bold columns
    bold_cols = [0, 6]
    for i in range(2,len(rows)):
        cells = rows[i].split(' & ')
        if len(cells) > max(bold_cols):
            for col in bold_cols:
                if cells[col].strip() != '':
                    cells[col] = rf'\textbf{{{cells[col].strip()}}}'
        rows[i] = ' & '.join(cells)

    # Bold and color rows
    bold_rows = [0, 30, 37, 38]
    colors = ['orange!50', 'orange!25', 'orange!25', 'orange!50']
    for i in range(len(bold_rows)):
        ix = bold_rows[i]
        line = r'\textbf{' + rows[ix].strip().replace(r' & ', r'} & \textbf{') + r'} '
        rows[ix] = rf'\rowcolor{{{colors[i]}}}' + line

    # =========== TEST CODE STARTS HERE

    # Prepare to process multirows
    multirow_blocks = []
    current_block = []

    for row in rows:
        if not row.strip():  # Skip empty rows
            continue

        cells = row.split(' & ')
        current_value = cells[0].strip()

        if current_value != '':
            if current_block:  # Close previous block
                multirow_blocks.append(current_block)
            current_block = [cells]
        else:
            current_block.append(cells)

    # Finalize the last block
    if current_block:
        multirow_blocks.append(current_block)

    # Construct lines with multirow
    formatted_rows = []

    for block in multirow_blocks:
        rowspan = len(block)
        if rowspan > 1:
            # Prepare the multirow line by including the entire first line's content
            # Ensure the line fits
            first_cell = block[0][0] #'\n'.join(block[0][0].split(' '))
            first_line = [
                r'\multirow[c]{' + str(rowspan) + '}{=}{\centering ' + first_cell + '} & ' + ' & '.join(block[0][1:])
            ]
            formatted_rows.append(''.join(first_line) + r' \\ \cline{2-9}')  # Add cline to the first row
            for line in block[1:]:
                # Make all subsequent lines
                formatted_rows.append(' & ' + ' & '.join(line[1:]) + r' \\ \cline{2-9}')
            # Change the final line's \\ \cline{2-9} to \\ \hline
            if formatted_rows[-1].endswith(r'\cline{2-9}'):
                formatted_rows[-1] = formatted_rows[-1][:-len(r'\cline{2-9}')] + r'\hline'
        else:
            # Single row, add hline
            formatted_rows.append(' & '.join(block[0]) + r' \\ \hline')

    rows = formatted_rows
    # ++++++++++++ ENDS +++++++++++++++++++++

    # Put it all together
    table = header + '\n'.join(rows) + footer
    table = table.replace(r'rule \\ \hline', r'rule')
    return table


def create_latex_doc(table):
    doc = Document()
    doc.preamble.append(Package('geometry', options='margin=1in'))
    doc.preamble.append(Package('xcolor', options='table'))
    doc.preamble.append(Package('colortbl'))
    doc.preamble.append(Package('booktabs'))
    doc.preamble.append(Package('multirow'))
    # doc.preamble.append(Package('array'))

    # Set page style to empty to remove page numbers
    doc.preamble.append(NoEscape(r'\pagestyle{empty}'))

    with doc.create(Table(position='htbp!')):
        # Main header
        doc.append(NoEscape(r'\begin{center}'))
        doc.append(NoEscape(rf'\large \underline{{{table.main()}}} \smallskip \normalsize'))
        for line in table.subheaders():
            doc.append(NoEscape(rf'\\ {{{line}}}'))
        doc.append(NoEscape(r'\end{center}'))

        # Smaller font to fit on one page
        doc.append(NoEscape(r'\footnotesize'))

        # Table with no extra borders, smaller font, and bold headers
        doc.append(NoEscape(r"""
            \renewcommand{\arraystretch}{1.3} % Increase the row height
            \setlength{\tabcolsep}{4pt}       % Add padding
        """))

        # Manually append the DataFrame-to-LaTeX converted table data
        doc.append(NoEscape(dataframe_to_latex_custom(table.get_table_data())))

    return doc

def convert_to_latex(table):
    print('converting to LaTex...')

    # Create the document
    doc = create_latex_doc(table)

    # Save the document to a .tex file
    doc_file = os.path.join(OUTPUT, 'table1')
    doc.generate_tex(doc_file)
    return doc

def convert_to_pdf(doc):
    
    print('converting to PDF...')

    # Optionally, compile to PDF (requires LaTeX installed on your system)
    doc.generate_pdf(os.path.join(OUTPUT, 'table1'), clean_tex=False, compiler='pdflatex')

def read_data(filepath, sheets = SHEETS_TO_LOAD):
    print('Reading data from Excel...')
    sectionsAB = pd.read_excel(filepath, sheet_name=sheets, header=None)
    return sectionsAB[SHEETS_TO_LOAD[0]]

# ======================= Main ===============================================

def main():
    table1 = Budget_Table(read_data(filepath))
    doc = convert_to_latex(table1)
    convert_to_pdf(doc)

def test():
    table1 = Budget_Table(read_data(filepath))
    output = table1.get_table_data().iloc[2,0]
    print(output)
    print(type(output))


if __name__ == '__main__':
    # test()
    main()