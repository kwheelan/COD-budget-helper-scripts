"""
Base table class for the budget book
K. Wheelan 
Feb 2025
"""

from pylatex import Table, NoEscape

class BaseTable:
    """
    A class to hold a basic table and convert to LaTeX
    """
    def __init__(self, custom_df, main_header, subheaders):
        # read data
        self.table_df = custom_df # type BaseDF
        self.topline_header = main_header
        self.other_header_lines = subheaders
        self.latex = ''

    def table_data(self):
        return self.table_df.latex_ready_data()

    def main(self):
        return self.topline_header

    def subheaders(self):
        return self.other_header_lines
    
    def column_format(self):
        """ Determines column width and centering for table """
        n_cols = len(self.table_data().columns)
        # default to center with auto-width
        return '|c|' + 'c|' * (n_cols - 1)

    def default_latex(self):
        latex = self.table_data().to_latex(
            index=False, 
            escape=False, 
            column_format=self.column_format(), 
            header=True,  # Automatically handles headers
            bold_rows=False,
        )
        # add horizontal lines
        return latex.replace(r'\\', r'\\ \hline')

    def latex_table_rows(self):
        # Split into table data
        rows = self.latex.replace('midrule', r'midrule\\')
        return rows.split(r'\\')
    
    def columns(self):
        return list(self.table_data().columns)
    
    def bold_cols(self, col_list):
        """ Add bolding to columns in data """
        col_ix = [self.columns().index(col) for col in col_list]
        rows = self.latex_table_rows()
        # Start with row 2 to skip latex table preamble, headers and \midrule
        # skip last 2 lines as well for end of tabular 
        for i in range(2, len(rows)-2):
            cells = rows[i].split(' & ')
            # if cell isn't empty, add the bold tag
            for col in col_ix:
                if cells[col].strip() != '':
                    cells[col] = rf'\textbf{{{cells[col].strip()}}}'
            rows[i] = ' & '.join(cells)
        self.update_latex(rows)
        
    def update_latex(self, rows):
        latex = r'\\'.join(rows).replace(r'rule\\', r'rule')
        self.latex = latex.replace(r'rule \\ \hline', r'rule')
    
    def process_latex(self):
        self.latex = self.default_latex()
        self.bold_cols(['Department'])
        return self.latex

    def export_latex_table(self, doc):
        with doc.create(Table(position='htbp!')):
            # Main header
            doc.append(NoEscape(r'\begin{center}'))
            doc.append(NoEscape(rf'\large \underline{{{self.main()}}} \smallskip \normalsize'))
            for line in self.subheaders():
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
            doc.append(NoEscape(self.process_latex()))


