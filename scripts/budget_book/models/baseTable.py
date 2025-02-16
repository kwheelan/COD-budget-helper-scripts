"""
Base table class for the budget book
K. Wheelan 
Feb 2025
"""

from models import BaseDF
import pandas as pd
import typing

class BaseTable:
    """
    A class to hold a basic table and convert to LaTeX
    """
    def __init__(self, custom_df : BaseDF, main_header : str, subheaders : str):
        # read data
        self.table_df = custom_df # type BaseDF
        self.topline_header = main_header
        self.other_header_lines = subheaders
        self.latex = ''

    def table_data(self) -> pd.DataFrame:
        return self.table_df.latex_ready_data()

    def main(self) -> str:
        return self.topline_header

    def subheaders(self) -> list[str]:
        return self.other_header_lines
    
    def column_format(self, format=None):
        """ Determines column width and centering for table """
        n_cols = len(self.table_data().columns)
        if format is None:
            # default to center with auto-width
            return '|c|' + 'c|' * (n_cols - 1)
    
    def rename_cols(self, new_cols):
        """ rename headers """
        self.table_df.adjust_col_names(new_cols)

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
        """ Split into table data list by row """
        lines = self.latex.split('\n')
        # Remove header and footer
        rows = lines[2:len(lines)-2]
        rows = '\n'.join(rows)
        rows = rows.replace('midrule', r'midrule\\')
        divider = r'\\ \hline' + '\n'
        return rows.split(divider)
    
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

    def bold_rows(self, row_nums):
        """ Bold rows in the table by row number (row 0 is header)"""
        for ix in row_nums:
            rows = self.latex_table_rows()
            rows[ix] = r'\textbf{' + rows[ix].strip().replace(r' & ', r'} & \textbf{') + r'} '
            self.update_latex(rows)

    def highlight_row(self, row_num, color):
        """ Add color to a row in the table """
        rows = self.latex_table_rows()
        rows[row_num] = rf'\rowcolor{{{color}}}' + rows[row_num]
        self.update_latex(rows)

    def highlight_rows(self, row_list, color_list):
        """ color multiple rows at once """
        for row, color in zip(row_list, color_list):
            self.highlight_row(row, color)

    def update_latex(self, rows):
        """ convert list of rows to full latex string """
        divider = r'\\ \hline' + '\n'
        latex = divider.join(rows).replace(r'rule\\', r'rule')
        latex = latex.replace(r'rule \\ \hline', r'rule')
        header = r'\begin{tabular}' + rf'{{{self.column_format()}}}' + '\n' + r'\toprule' + '\n'
        footer = '\n\n' + r'\bottomrule' + '\n' +  r'\end{tabular}'
        self.latex = header + latex + footer
    
    def process_latex(self):
        """ Add any formatting and return latex """
        self.latex = self.default_latex()
        return self.latex        


