"""
K. Wheelan February 2025

"""

from openpyxl import load_workbook
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
        # Adjust columns
        self.table_data.columns = self.table_data.iloc[0] 
        columns_list = list(self.table_data.columns)
        columns_list[1] = ''
        columns_list[2] ='Department'
        self.table_data.columns = columns_list
        self.table_data = self.table_data[1:].reset_index(drop=True) 
        # Get headers
        self.topline_header = df.iloc[0, 1]
        self.other_header_lines = [df.iloc[i, 1] for i in range(2, 5)]
        self.clean_nas()
        # Round all float values to 2 decimal places
        self.table_data = self.table_data.applymap(self.round_floats)
        self.escape_ampersands_in_table_data()

    def clean_nas(self):
        self.table_data.fillna('', inplace=True)

    def escape_ampersands_in_table_data(self):
        # Apply the escaping function to each element in the DataFrame
        self.table_data = self.table_data.applymap(self.escape_ampersands)

    @staticmethod
    def escape_ampersands(text):
        # Replace & with \& in all string elements
        if isinstance(text, str):
            return text.replace("&", r"\&")
        return text
    
    @staticmethod
    def round_floats(value):
        # Round the floats to 2 decimal places
        if isinstance(value, float):
            return str(round(value, 2))
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
    column_format = '|p{1.5cm}|' + 'p{1cm}|' + 'p{3.5cm}|' + 'p{1.5cm}|' * (len(df.columns) - 3)
    
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
    rows = rows.split(r'\\')

    # Modify the header line to bold it
    rows[0] = r'\textbf{' + rows[0].replace(r'&', r'} & \textbf{') + r'}'
    print(rows[0])

    # Return the modified LaTeX code
    return header + r'\\ \hline'.join(rows) + footer

def create_latex_doc(table):
    doc = Document()
    doc.preamble.append(Package('geometry', options='margin=1in'))
    doc.preamble.append(Package('xcolor', options='table'))
    doc.preamble.append(Package('colortbl'))
    doc.preamble.append(Package('booktabs'))

    with doc.create(Section('Section A')):
        with doc.create(Subsection('FTE')):
            with doc.create(Table(position='htbp!')):
                # Main header
                doc.append(NoEscape(r'\begin{center}'))
                doc.append(NoEscape(rf'\underline{{{table.main()}}}'))
                doc.append(NoEscape(r'\end{center}'))

                # Smaller font to fit on one page
                doc.append(NoEscape(r'\small'))

                # Table with no extra borders, smaller font, and bold headers
                doc.append(NoEscape(r"""
                    \renewcommand{\arraystretch}{1.3} % Increase the row height
                    \setlength{\tabcolsep}{4pt}       % Add padding
                """))

                # Manually append the DataFrame-to-LaTeX converted table data
                doc.append(NoEscape(dataframe_to_latex_custom(table.get_table_data())))

    return doc

def convert_to_latex(table):
    print("converting to LaTex...")

    # Create the document
    doc = create_latex_doc(table)

    # Save the document to a .tex file
    doc_file = os.path.join(OUTPUT, 'table1')
    doc.generate_tex(doc_file)
    return doc

def convert_to_pdf(doc):
    
    print("converting to PDF...")

    # Optionally, compile to PDF (requires LaTeX installed on your system)
    doc.generate_pdf(os.path.join(OUTPUT, 'table1'), clean_tex=False, compiler='pdflatex')

def read_data(filepath, sheets = SHEETS_TO_LOAD):
    print("Reading data from Excel...")
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