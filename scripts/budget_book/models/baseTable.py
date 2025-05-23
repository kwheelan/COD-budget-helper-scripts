"""
Base table class for the budget book
K. Wheelan 
Feb 2025
"""

from models import BaseDF
import pandas as pd

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
    
    def isEmpty(self):
        try:
            data = self.table_data()
        except:
            return True
        
        # Check if data is None or if it's an empty DataFrame
        return not (isinstance(data, pd.DataFrame) and not data.empty)

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
        if self.isEmpty():
            return ''
        latex = self.table_data().to_latex(
            index=False, 
            escape=False, 
            column_format=self.column_format()
        )
        # remove dollar signs
        latex = latex.replace('$', r'\$')
        # remove underscores
        latex = latex.replace('_', r' ')
        # replace zeroes
        latex = latex.replace('& 0 ', '& - ')
         # Remove midrule, etc
        latex = latex.replace("\\toprule\n", "")
        latex = latex.replace("\\midrule\n", "")
        latex = latex.replace("\\bottomrule\n", "")
        # add horizontal lines
        return latex.replace(r'\\' + '\n', self.divider())
    
    @staticmethod
    def divider():
        return r'\\ \hline' + '\n'

    def latex_table_rows(self):
        """ Split into table data list by row """
        lines = self.latex.split('\n')
        # remove empty trailing space
        while lines[-1] == '':
            lines = lines[:-1]
        # Remove header and footer
        rows = lines[1:len(lines)-1]
        # add final new line to ensure \hline disappears
        rows = '\n'.join(rows) + '\n'
        return rows.split(self.divider())
    
    def columns(self):
        return list(self.table_data().columns)
    
    def bold_cols(self, col_list):
        """ Add bolding to columns in data """
        col_ix = [self.columns().index(col) for col in col_list]
        rows = self.latex_table_rows()
        # Start with row 2 to skip latex table preamble and headers
        # skip last line as well for end of tabular 
        for i in range(2, len(rows)-1):
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

    def highlight_rows(self, row_list, color):
        """ color multiple rows at once """
        for row in row_list:
            self.highlight_row(row, color)

    def update_latex(self, rows):
        """ convert list of rows to full latex string """
        latex = self.divider().join(rows)
        footer = r'\end{longtable}'
        header = r'\arrayrulecolor{black}\begin{longtable}' + rf'{{{self.column_format()}}}' + '\n'
        self.latex = header + latex + footer
    
    def process_latex(self):
        """ Add any formatting and return latex """
        if not self.table_data():
            return None
        self.latex = self.default_latex()
        return self.latex  

    def merge_rows(self, col_name):
        """ merge rows in a single column (all empty cells merged with value above)"""
        # pass over headers for merging
        headers = self.latex_table_rows()[:2]
        rows = self.latex_table_rows()[2:]
        col_ix = self.columns().index(col_name)
        rows = headers + RowMerger().merge_rows(rows, col_ix)
        self.update_latex(rows)

    def replace_header(self, new_header):
        """ replace header for table """
        rows = self.latex_table_rows()
        rows[0] = new_header
        header = r'\arrayrulecolor{black}\begin{longtable}' + rf'{{{self.column_format()}}}' + '\n'
        footer = r'\end{longtable}'
        self.latex = header + rows[0] + '\n' + self.divider().join(rows[1:]) + footer

    def n_rows(self):
        return len(self.latex_table_rows())

    def add_tab(self, col, row_list, length = '0.75cm'):
        """ Add space to each column """
        if type(col) is int:
            col_ix = col
        else:
            col_ix = self.columns().index(col)
        rows = self.latex_table_rows()
        for i in row_list:
            cells = rows[i].split(' & ')
            # if cell isn't empty, add the tab
            if cells[col_ix].strip() != '':
                cells[col_ix] = rf'\hspace{{{length}}}' + cells[col_ix].strip()
            rows[i] = ' & '.join(cells)
        self.update_latex(rows)

    @staticmethod
    def count_digits_before_dash(input_string):
        """
        Counts the number of numerical digits before the first dash in a string.

        Parameters:
        input_string (str): The input string from which to extract digits.

        Returns:
        int: The number of numerical digits before the dash.
        """
        if '- ' not in input_string:
            return 0

        # split
        parts = input_string.split('-')
        
        if parts:
            # Get the part before the dash and strip any leading/trailing spaces
            number_part = parts[0].strip()
            
            # Count the digits in this part
            digit_count = sum(char.isdigit() for char in number_part)
            
            return digit_count
        else:
            return 0
    
class RowMerger:
    def merge_rows(self, rows: list[str], col_ix: int):
        """Merge rows vertically based on the specified column."""
        multirow_blocks = self.create_multirow_blocks(rows, col_ix)
        formatted_rows = self.format_multirow_blocks(multirow_blocks, col_ix)
        return formatted_rows

    def create_multirow_blocks(self, rows: list[str], col_ix: int):
        """Identify blocks of rows that should be merged as multirows."""
        multirow_blocks = []
        current_block = []

        for row in rows:
            if not row.strip():
                continue

            cells = row.split(' & ')
            current_value = cells[col_ix].strip()

            if current_value != '':
                if current_block:
                    multirow_blocks.append(current_block)
                current_block = [cells]
            else:
                current_block.append(cells)

        if current_block:
            multirow_blocks.append(current_block)

        return multirow_blocks
    
    @staticmethod
    def cline(line, col_ix):
        """ Create horizontal line of correct length """
        ncols = len(line)
        if col_ix == 0:
            return rf'\cline{{2-{ncols}}}'
        elif col_ix == ncols-1:
            return rf'\cline{{1-{ncols-1}}}'
        return rf'\cline{{1-{col_ix}}} \cline{{{col_ix+1}-{ncols}}}'

    def format_multirow_blocks(self, multirow_blocks, col_ix):
        """Format each block with multirow LaTeX commands."""
        formatted_rows = []

        for block in multirow_blocks:
            rowspan = len(block)
            if rowspan > 1:
                pre = ' & '.join(block[0][:col_ix])
                post = ' & '.join(block[0][col_ix + 1:])
                target_cell = block[0][col_ix]

                first_line = (
                    (pre + ' & ' if pre else '') +
                    r'\multirow{{{}}}{{*}}{{{}}}'.format(rowspan, target_cell) +
                    (f' & {post}' if post else '') +
                    self.cline(block, col_ix) +
                    r' \\'
                )
                formatted_rows.append(first_line)

                # Add the rest of the lines
                for line in block[1:]:
                    formatted_row = ' & '.join(
                        line[:col_ix] + line[col_ix + 1:]
                    ) 
                    formatted_rows.append(formatted_row)
            else:
                formatted_rows.append(' & '.join(block[0]))

        return formatted_rows


