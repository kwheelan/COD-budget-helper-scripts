from . import SummaryCategoryTable
from models.BudgetData import *
from .tableHeader import Header

class FundCategoryTable(SummaryCategoryTable):

    def __init__(self, custom_df, primary_color, 
                 color2, color3,
                 line_color, main_header, subheaders):
        self.color2 = color2
        self.color3 = color3
        super().__init__(custom_df, primary_color, 
                         line_color, main_header, subheaders)
        
    def header(self):
        return Header.fund_categories(self.main(), self.subheaders())

    def process_latex(self):
        # self.rename_cols()
        self.latex = self.default_latex()
        #self.bold_cols(['Category', 'Variance FY25 vs FY26'])
        self.add_tab(col=0, row_list=range(2, self.last_row()), length='0.5cm')
        self.add_tab(col=0, row_list=self.data_rows(), length='0.5cm')
        self.bold_rows([0, 1, self.last_row()] + self.fund_rows())
        self.highlight_rows([1, self.last_row()], self.primary_color)
        self.highlight_rows(self.fund_rows(), self.color2)
        self.outline_last_row(self.line_color)
        self.replace_header(self.header())
        return self.latex
    
    def fund_rows(self):
        df = self.table_data()
        fund_rows = []
        # skip first total row
        for i in range(1, df.shape[0]):
            first_char = df.iloc[i, 0][0]
            if first_char in '1234567890':
                # add one because latex counts headers as row 0
                fund_rows.append(i+1)
        return fund_rows

    def data_rows(self):
        rows = []
        for i in range(2, self.n_rows()-1):
            if i not in self.fund_rows():
                rows.append(i)
        return rows
    
class ExpFundCatTable(FundCategoryTable):

    def __init__(self, filepath, dept, custom_df):
        self.dept = dept
        if custom_df is None:
            custom_df = Expenditures(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      'EXPENDITURES BY SUMMARY CATEGORY - FUND DETAIL',
                      'DEPARTMENT ' + dept_name.upper()]
        super().__init__(custom_df, 'blue1', 
                         'blue2', 'blue3',
                         'lineblue', main_header, subheaders)
        
    def table_data(self):
        return self.table_df.group_by_category_and_fund(self.dept)


class RevFundCatTable(FundCategoryTable):

    def __init__(self, filepath, dept, custom_df):
        self.dept = dept
        if custom_df is None:
            custom_df = Revenues(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        if dept_name is not None:
            subheaders = ['BUDGET DEVELOPMENT',
                      'EXPENDITURES BY SUMMARY CATEGORY - FUND DETAIL',
                      'DEPARTMENT ' + dept_name.upper()]
        else: 
            subheaders = []
        super().__init__(custom_df, 'green1', 
                         'green2', 'green3', 
                         'linegreen', main_header, subheaders)

    def table_data(self):
        return self.table_df.group_by_category_and_fund(self.dept)