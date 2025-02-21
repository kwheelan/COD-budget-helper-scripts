
from models.baseTable import BaseTable
from models.BudgetData.summary_tables import Summary
from .tableHeader import Header

class SummaryTable(BaseTable):
    """
    Section B first summary tables
    """

    def __init__(self, custom_df : Summary, primary_color, dept):
        super().__init__(custom_df, 'Budget Summary', [])
        self.primary_color = primary_color
        self.dept = dept

    def column_format(self):
        n_cols = len(self.table_data().columns)
        return r'| >{\raggedleft\arraybackslash}p{4cm}' + r'| >{\raggedleft\arraybackslash}p{2.5cm}' * (n_cols - 1) + '|' 

    def header(self):
        return Header.summary_table()
    
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
        self.bold_rows([0, self.last_row()])
        self.highlight_rows([0, self.last_row()], self.primary_color)
        self.replace_header(self.header())
        self.outline_last_row('black')
        return self.latex
    
    def divider(self):
        return r'\\ \hline' + '\n'
    
    @staticmethod
    def round_to_dollar(value):
        # true strings should be returned
        if type(value) is str and value[0] not in '0123456789':
            return value
        # all else rounded to the nearest dollar
        elif type(value) is str:
            value = float(value.replace(',',''))
        return f'{value:,.0f}'
    
class SummaryTable1(SummaryTable):
    """
    Section B first summary tables
    """

    def __init__(self, custom_df : Summary, primary_color, dept):
        super().__init__(custom_df, 'Budget Summary', [])
        self.primary_color = primary_color
        self.dept = dept

    def table_data(self):
        df = self.table_df.table1_part1(self.dept)
        df = df.map(self.round_to_dollar)
        return df
    
    def title(self):
        return rf'Department {self.table_df.exp.dept_name(self.dept)}\\'

class SummaryTable2(SummaryTable):
    def __init__(self, custom_df : Summary, primary_color, dept):
        super().__init__(custom_df, '', [])
        self.primary_color = primary_color
        self.dept = dept

    def main(self):
        return ''

    def table_data(self):
        df = self.table_df.table1_part2(self.dept)
        df = df.map(self.round_to_dollar)
        return df
    
    def header(self):
        return Header.summary_table2()
    
class SummaryTable3(SummaryTable):

    def __init__(self, custom_df : Summary, primary_color, dept):
        super().__init__(custom_df, 'General Fund Recurring vs One-Time Expenditures', [])
        self.primary_color = primary_color
        self.dept = dept

    def table_data(self):
        df = self.table_df.table3(self.dept)
        df = df.map(self.round_to_dollar)
        return df

    def main(self):
        return 'General Fund Recurring vs One-Time Expenditures'
    
    def header(self):
        return Header.summary_table3()

    def column_format(self):
        n_cols = len(self.table_data().columns)
        return r'| >{\raggedleft\arraybackslash}p{4cm}' + r'| >{\raggedleft\arraybackslash}p{2.5cm}' * (n_cols - 1) + '|' 