"""
K. Wheelan February 2025

"""

from openpyxl import load_workbook
import os
import pandas as pd
from pylatex import Document, Section, Subsection, Table, Tabular, Math, Alignat, NoEscape

# File paths
DATA = 'input_data/model_copy'
file = os.listdir(DATA)[0]
filepath = os.path.join(DATA, file)
OUTPUT = 'output/budget_book/'

SHEETS_TO_LOAD = [
    'Section A FTE'
]

print("Reading data from Excel...")

# full_data = load_workbook(os.path.join(DATA, file), data_only=True)
sectionsAB = pd.read_excel(filepath, sheet_name=SHEETS_TO_LOAD)
fte_tables = sectionsAB[SHEETS_TO_LOAD[0]]

table1 = fte_tables.iloc[0:43, 1:10]

def dataframe_to_latex(df):
    return df.to_latex(index=False)

def create_latex_doc(df):
    # Create a new LaTeX document
    doc = Document()

    # Add sections with text and tables
    with doc.create(Section('Main Section')):
        doc.append('This is a section with a table derived from Excel data.')

        with doc.create(Subsection('Excel Data Table')):
            with doc.create(Table()):
                # Convert the DataFrame to a LaTeX table and add the table to the document
                doc.append(NoEscape(dataframe_to_latex(df)))

    return doc

print("converting to LaTex...")

# Create the document
doc = create_latex_doc(table1)

# Save the document to a .tex file
doc_file = os.path.join(OUTPUT, 'table1')
doc.generate_tex(doc_file)
 
print("converting to PDF...")

# Optionally, compile to PDF (requires LaTeX installed on your system)
doc.generate_pdf(os.path.join(OUTPUT, 'table1'), clean_tex=False, compiler='pdflatex')

