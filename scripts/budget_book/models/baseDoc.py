from pylatex import Document, Section, Subsection, Table, NoEscape, Package
# from . import baseTable
import os
from constants import OUTPUT
import subprocess
from .SectionB import SummaryTable


class BaseDoc:
    """
    A class to contain a default LaTeX document
    """
    def __init__(self, table_list, save_as):
        self.doc = Document()
        self.table_list = table_list
        self.save_as = os.path.join(OUTPUT, save_as)

    def add_preamble(self):
        """ Create preamble for latex doc: packages, custom commands, etc """
        doc = self.doc
        doc.preamble.append(Package('geometry', options='margin=1.5cm'))
        doc.preamble.append(Package('xcolor', options='table'))
        doc.preamble.append(Package('colortbl'))
        doc.preamble.append(Package('booktabs'))
        doc.preamble.append(Package('multirow'))
        doc.preamble.append(Package('pdflscape'))
        doc.preamble.append(Package('array'))
        doc.preamble.append(Package('roboto', options=['sfdefault', 'condensed']))
        doc.preamble.append(Package('longtable'))

        # Set page style to empty to remove page numbers
        doc.preamble.append(NoEscape(r'\pagestyle{empty}'))
        
        # define colors
        self.define_all_colors()

        self.doc = doc

    def create_doc(self, table_list):
        """ Acutally generate document """
        # add packages, colors, other custom definitions
        self.add_preamble()
        # create tables
        for table in table_list:
            if type(table) is list:
                self.summary_tables(self.doc, table)

            elif not table.isEmpty():
                self.latex_table(self.doc, table)
                
    def define_color(self, name, rgb: tuple[int, int, int]):
        r, g, b = rgb
        self.doc.append(NoEscape(rf'\definecolor{{{name}}}{{RGB}}{{{r},{g},{b}}}'))

    def define_all_colors(self):
        self.define_color('blue1', (155, 194, 230))
        self.define_color('blue2', (189, 215, 238))
        self.define_color('blue3', (221, 235, 247))
        self.define_color('lineblue', (47, 117, 181))
        self.define_color('green1', (169, 208, 142))
        self.define_color('green2', (198, 224, 180))
        self.define_color('green3', (226, 239, 218))
        self.define_color('linegreen', (84, 130, 53))
        self.define_color('orange1', (244, 176, 132))
        self.define_color('orange2', (252, 228, 214))
        self.define_color('lineorange', (237, 125, 49))
        self.define_color('detroitgreen', (159, 213, 179))

    @staticmethod
    def header(doc, table, bold = True, special_main = False):
        """ create header for table """
        # center it
        doc.append(NoEscape(r'\begin{center}'))
        if bold:
            doc.append(NoEscape(r'\textbf{'))
        # some tables need a larger underlined header
        if special_main:
            doc.append(NoEscape(rf'\large \underline{{{table.main()}}} \smallskip \normalsize'))
        else:
            doc.append(NoEscape(table.main()))
        # add a line for each subheader
        for line in table.subheaders():
            doc.append(NoEscape(rf'\\ {{{line}}}'))
        if bold:
            doc.append(NoEscape(r'}'))
        doc.append(NoEscape(r'\end{center}'))

    def latex_table(self, doc, table):
        """ Create table from doc and table object """
        
        # make page landscape
        self.doc.append(NoEscape(r'\begin{landscape}'))
        # table header
        self.header(doc, table)

        # Table with no extra borders, smaller font, and bold headers
        doc.append(NoEscape(r"""
            \renewcommand{\arraystretch}{1.3} % Increase the row height
            \setlength{\tabcolsep}{4pt}       % Add padding
        """))

        # Manually append the DataFrame-to-LaTeX converted table data
        doc.append(NoEscape(table.process_latex()))
        self.doc.append(NoEscape(r'\end{landscape}'))
        self.doc.append(NoEscape(r'\newpage'))

    def summary_tables(self, doc, table_list : list[SummaryTable]):
        """ Create table from doc and table object """
        
        # make page landscape
        self.doc.append(NoEscape(r'\begin{landscape}'))
        # page header
        table1 = table_list[0]
        doc.append(NoEscape(r'\begin{center}'))
        doc.append(NoEscape(rf'\huge \textbf{{{table1.title()}}} \normalsize'))
        doc.append(NoEscape(r'\end{center}'))
        doc.append(NoEscape(r'\begin{flushleft}'))

        # Manually append the DataFrame-to-LaTeX converted table data
        for table in table_list:
            if table.main():
                doc.append(NoEscape('\n' + rf'\hspace{{0.7cm}}\textbf{{{table.main()}}}'))
            doc.append(NoEscape(table.process_latex()))
        doc.append(NoEscape(r'\end{flushleft}'))
        self.doc.append(NoEscape(r'\end{landscape}'))
        self.doc.append(NoEscape(r'\newpage'))

    def save_as_latex(self):
        self.create_doc(self.table_list)
        # Save the document to a .tex file
        self.doc.generate_tex(self.save_as)

    def compile_latex(self):
        # Compile the LaTeX file using subprocess; disable clean to keep auxiliary files
        # subprocess.run(['pdflatex', '-interaction=nonstopmode', self.save_as])
        subprocess.run([
            'pdflatex',
            '-interaction=nonstopmode',
            '-output-directory', os.path.dirname(self.save_as),  # Specify the output directory
            self.save_as
            ], 
            stdout=subprocess.DEVNULL,  # Redirect stdout to prevent terminal output
            stderr=subprocess.DEVNULL,  # Redirect stderr to prevent terminal output
            check=True)

    def convert_to_pdf(self):
        # compile to get aux files and generate PDF
        self.compile_latex()
        self.doc.generate_pdf(self.save_as, clean_tex=True, compiler='pdflatex')


    

    