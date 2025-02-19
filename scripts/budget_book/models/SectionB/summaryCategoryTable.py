
from models import BaseTable
from models.modelDF import Expenditures

class SummaryCategoryTable(BaseTable):
    """
    Represents Section B Department tables
    Expenditures and Revenues by Summary Category
    """

    def __init__(self, custom_df, primary_color, line_color, main_header, subheaders):
        super().__init__(custom_df, main_header, subheaders)
        self.primary_color = primary_color
        self.line_color = line_color

    def column_format(self, format=None):
        n_cols = len(self.table_data().columns)
        return r'>{\arraybackslash}p{10.5cm}' + r'>{\centering\arraybackslash}p{2.5cm}' * (n_cols - 1) 

    @staticmethod
    def header():
        return r"""
        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{\rule{0pt}{1.5cm}\textbf{\shortstack{Department \# - Department Name\\ \hspace{-0.5cm}Summary Category}}\rule[-0.25cm]{0pt}{0.5cm}} &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{\rule{0pt}{1cm}\textbf{\shortstack{FY2029 \\ Forecast}}\rule[-0.75cm]{0pt}{1cm}} \\
        \specialrule{1.5pt}{0pt}{0pt}
        """
    
    def divider(self):
        return rf'\\ \arrayrulecolor{{{self.line_color}}}\hline' + '\n'
    
    def last_row(self):
        return self.n_rows() - 2
    
    def thick_line(self, color):
        return rf'\arrayrulecolor{{{color}}}\specialrule{{1.5pt}}{{0pt}}{{0pt}}' + '\n'
    
    def outline_last_row(self, color):
        rows = self.latex_table_rows()
        ix = self.last_row()
        rows[ix] = self.thick_line(color) + rows[ix] + r'\\' + '\n' + self.thick_line(color)
        self.update_latex(rows[:-1])

    def process_latex(self):
        # self.rename_cols()
        self.latex = self.default_latex()
        #self.bold_cols(['Category', 'Variance FY25 vs FY26'])
        self.bold_rows([0, 1, self.last_row()])
        self.highlight_rows([1, self.last_row()], [self.primary_color, self.primary_color])
        self.add_tab(col=0, row_list=range(2, self.last_row()))
        self.outline_last_row(self.line_color)
        self.replace_header(self.header())
        return self.latex
    
class ExpenditureCategories(SummaryCategoryTable):

    def __init__(self, filepath, dept):
        custom_df = Expenditures(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      'EXPENDITURES BY SUMMARY CATEGORY - ALL FUNDS',
                      dept_name.upper()]
        super().__init__(custom_df, 'blue1', 'lineblue', main_header, subheaders)
        
    def table_data(self):
        return self.table_df.group_by_category('10')


class RevenueCategories(SummaryCategoryTable):

    def __init__(self, filepath, dept):
        custom_df = Expenditures(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      'EXPENDITURES BY SUMMARY CATEGORY - ALL FUNDS',
                      dept_name.upper()]
        super().__init__(custom_df, 'green1', 'linegreen', main_header, subheaders)