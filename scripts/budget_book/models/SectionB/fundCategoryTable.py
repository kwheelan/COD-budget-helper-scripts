from . import SummaryCategoryTable
from models.BudgetData import *

class FundCategoryTable(SummaryCategoryTable):

    def __init__(self, custom_df, primary_color, line_color, main_header, subheaders):
        super().__init__(custom_df, primary_color, line_color, main_header, subheaders)

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
    
class ExpFundCatTable(FundCategoryTable):

    def __init__(self, filepath, dept):
        self.dept = dept
        custom_df = Expenditures(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      'EXPENDITURES BY SUMMARY CATEGORY - ALL FUNDS',
                      'DEPARTMENT' + dept_name.upper()]
        super().__init__(custom_df, 'blue1', 'lineblue', main_header, subheaders)
        
    def table_data(self):
        return self.table_df.group_by_category_and_fund(self.dept)


class RevFundCatTable(FundCategoryTable):

    def __init__(self, filepath, dept):
        self.dept = dept
        custom_df = Revenues(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      'REVENUES BY SUMMARY CATEGORY - ALL FUNDS',
                      'DEPARTMENT' + dept_name.upper()]
        super().__init__(custom_df, 'green1', 'linegreen', main_header, subheaders)

    def table_data(self):
        return self.table_df.group_by_category_and_fund(self.dept)