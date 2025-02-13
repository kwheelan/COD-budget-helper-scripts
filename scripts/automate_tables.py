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

class Budget_Table():
    def __init__(self, df):
        # fix hard-coding
        # Extract data
        self.table_data = df.iloc[6:44, 1:10]
        # Get headers
        self.topline_header = df.iloc[0,1]
        self.other_header_lines = [df.iloc[i,1] for i in range(2,5)]
        self.clean_nas()

    def clean_nas(self):
        self.table_data.fillna('', inplace=True)

    def get_table_data(self):
        return self.table_data
    
    def main(self):
        return self.topline_header
    
    def subheaders(self):
        return self.other_header_lines

# ======================= Functions =======================================

def dataframe_to_latex_custom(df):
    # Convert the DataFrame to LaTeX
    latex_code = df.to_latex(index=False, escape=False, column_format= '|c|' + 'c|' * (len(df.columns)), header=False)
    # Clean up or modify LaTeX code as needed
    return latex_code

# def underline(text):
#     return NoEscape('\underline{') + text + '}'

def create_latex_doc(table):
    doc = Document()
    doc.preamble.append(Package('geometry', options='margin=1in'))
    doc.preamble.append(Package('xcolor', options='table'))
    doc.preamble.append(Package('colortbl'))
    doc.preamble.append(Package('booktabs'))

    with doc.create(Section('Section A')):
        with doc.create(Subsection('FTE')):
            with doc.create(Table(position='htbp!')):  # More flexible positioning
                doc.append(table.main())
                doc.append(NoEscape(r"""
                    \renewcommand{\arraystretch}{1.3}  % Increase the row height
                    \setlength{\tabcolsep}{8pt}        % Add padding
                    \begin{tabular}{|c|c|c|}           % Define the table format
                    \hline                             % Top Border
                """))

                    # \rowcolor{lightgray}                % Example cell color
                    # \hline

                # Manually append the DataFrame-to-LaTeX converted table
                doc.append(NoEscape(dataframe_to_latex_custom(table.get_table_data())))  # Convert dataframe to desired LaTeX

                doc.append(NoEscape(r"""
                    \hline                             % Bottom Border
                    \end{tabular}
                """))

    return doc

def convert(table):
    print("converting to LaTex...")

    # Create the document
    doc = create_latex_doc(table)

    # Save the document to a .tex file
    doc_file = os.path.join(OUTPUT, 'table1')
    doc.generate_tex(doc_file)
    
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
    convert(table1)

def test():
    table1 = Budget_Table(read_data(filepath))
    output = table1.get_table_data().iloc[2,0]
    print(output)
    print(type(output))


if __name__ == '__main__':
    # test()
    main()