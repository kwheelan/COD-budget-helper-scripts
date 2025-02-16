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
            self.latex_table(doc, table)

        self.doc = doc

    def latex_table(self, doc, table):
        """ Create table from doc and table object """
        with doc.create(Table(position='htbp!')):
            # Main header
            doc.append(NoEscape(r'\begin{center}'))
            doc.append(NoEscape(rf'\large \underline{{{table.main()}}} \smallskip \normalsize'))
            for line in table.subheaders():
                doc.append(NoEscape(rf'\\ {{{line}}}'))
            doc.append(NoEscape(r'\end{center}'))

            # Smaller font to fit on one page
            doc.append(NoEscape(r'\scriptsize'))

            # Table with no extra borders, smaller font, and bold headers
            doc.append(NoEscape(r"""
                \renewcommand{\arraystretch}{1.3} % Increase the row height
                \setlength{\tabcolsep}{4pt}       % Add padding
            """))

            # Manually append the DataFrame-to-LaTeX converted table data
            doc.append(NoEscape(table.process_latex()))

    def save_as_latex(self):
        self.create_doc(self.table_list)
        # Save the document to a .tex file
        self.doc.generate_tex(self.save_as)

    def convert_to_pdf(self):
        print('converting to pdf...')
        self.doc.generate_pdf(self.save_as, clean_tex=False, compiler='pdflatex')


    

    