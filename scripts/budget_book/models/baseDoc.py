from pylatex import Document, Section, Subsection, Table, NoEscape, Package
# from . import baseTable
import os
from constants import OUTPUT

class BaseDoc:
    """
    A class to contain a default LaTeX document
    """
    def __init__(self, table_list, save_as):
        self.doc = None
        self.table_list = table_list
        self.save_as = os.path.join(OUTPUT, save_as)

    def create_doc(self, table_list):
        print('writing latex....')
        doc = Document()
        doc.preamble.append(Package('geometry', options='margin=1in'))
        doc.preamble.append(Package('xcolor', options='table'))
        doc.preamble.append(Package('colortbl'))
        doc.preamble.append(Package('booktabs'))
        doc.preamble.append(Package('multirow'))
        # doc.preamble.append(Package('array'))

        # Set page style to empty to remove page numbers
        doc.preamble.append(NoEscape(r'\pagestyle{empty}'))
        
        for table in table_list:
            table.export_latex_table(doc)

        self.doc = doc

    def save_as_latex(self):
        self.create_doc(self.table_list)
        # Save the document to a .tex file
        self.doc.generate_tex(self.save_as)

    def convert_to_pdf(self):
        print('converting to pdf...')
        self.doc.generate_pdf(self.save_as, clean_tex=False, compiler='pdflatex')


    

    