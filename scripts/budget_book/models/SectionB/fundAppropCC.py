from .fundCategoryTable import *
from .tableHeader import *

class FundAppropCCTable(FundCategoryTable):

    def __init__(self, custom_df, primary_color, 
                 color2, color3,
                 line_color, main_header, subheaders):
        super().__init__(custom_df, primary_color, 
                        color2, color3,
                         line_color, main_header, subheaders)
        
    def header(self):
        return Header.fund_approp_cc(self.main(), self.subheaders())
    
    def fund_rows(self):
        df = self.table_data()
        fund_rows = []
        # skip first total row
        for i in range(1, df.shape[0]):
            if self.count_digits_before_dash(df.iloc[i, 0]) == 4:
                fund_rows.append(i+1)
        return fund_rows
    
    def approp_rows(self):
        df = self.table_data()
        fund_rows = []
        # skip first total row
        for i in range(1, df.shape[0]):
            if self.count_digits_before_dash(df.iloc[i, 0]) == 5:
                fund_rows.append(i+1)
        return fund_rows
    
    def cc_rows(self):
        df = self.table_data()
        fund_rows = []
        # skip first total row
        for i in range(1, df.shape[0]):
            if self.count_digits_before_dash(df.iloc[i, 0]) == 6:
                fund_rows.append(i+1)
        return fund_rows
    
    def process_latex(self):
        if self.table_data() is None:
            return ''
        # self.rename_cols()
        self.latex = self.default_latex()
        # indent cc x 3, approp x2, fund x1
        self.add_tab(col=0, row_list=self.fund_rows(), length='0.5cm')
        self.add_tab(col=0, row_list=self.approp_rows(), length='1cm')
        self.add_tab(col=0, row_list=self.cc_rows(), length='1.5cm')
        self.bold_rows([0, 1, self.last_row()] + self.fund_rows()  + self.approp_rows())
        self.highlight_rows([1, self.last_row()], self.primary_color)
        self.highlight_rows(self.fund_rows(), self.color2)
        self.highlight_rows(self.approp_rows(), self.color3)
        self.outline_last_row(self.line_color)
        self.replace_header(self.header())
        return self.latex
    

class ExpFullTable(FundAppropCCTable):

    def __init__(self, filepath, dept, custom_df):
        self.dept = dept
        if custom_df is None:
            custom_df = Expenditures(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      r'FINANCIAL DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER - EXPENDITURES',
                      'DEPARTMENT ' + dept_name.upper()]
        super().__init__(custom_df, 'blue1', 
                         'blue2', 'blue3',
                         'lineblue', main_header, subheaders)
        
    def table_data(self):
        return self.table_df.group_by_fund_approp_cc(self.dept)


class RevFullTable(FundAppropCCTable):

    def __init__(self, filepath, dept, custom_df):
        self.dept = dept
        if custom_df is None:
            custom_df = Revenues(filepath)
        dept_name = custom_df.dept_name(dept)
        main_header = 'CITY OF DETROIT'
        subheaders = ['BUDGET DEVELOPMENT',
                      r'FINANCIAL DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER - REVENUES',
                      'DEPARTMENT ' + dept_name.upper()]
        super().__init__(custom_df, 'green1', 
                         'green2', 'green3', 
                         'linegreen', main_header, subheaders)

    def table_data(self):
        return self.table_df.group_by_fund_approp_cc(self.dept)
    
class FTEFullTable(FundAppropCCTable):

    def __init__(self, filepath, dept, custom_df):
        self.dept = dept
        if custom_df is None:
            custom_df = FTEs(filepath)
        dept_name = custom_df.dept_name(dept)
        # deal with empty table case
        if dept_name is None:
            main_header = ''
            subheaders = []
        else:
            main_header = 'CITY OF DETROIT'
            subheaders = ['BUDGET DEVELOPMENT',
                        r'POSITION DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER',
                        'DEPARTMENT ' + dept_name.upper()]
        super().__init__(custom_df, 'orange1', 
                         'orange2',  None,
                         'lineorange', main_header, subheaders)
        
    def header(self):
        return Header.fte(self.main(), self.subheaders())

    def table_data(self):
        if not self.table_df.dept_name(self.dept):
            return None
        return self.table_df.group_by_fund_approp_cc(self.dept)

    def job_rows(self):
        df = self.table_data()
        fund_rows = []
        jobs = list(set(self.table_df.filter_by_dept(self.dept)['Job Title']))
        # skip first total row
        for i in range(1, df.shape[0]):
            if df.iloc[i, 0] in jobs:
                fund_rows.append(i+1)
        return fund_rows
    
    def cc_rows(self):
        job_rows = self.job_rows()
        return [i for i in super().cc_rows() if i not in job_rows]
    
    def process_latex(self):
        """ change table to formatted latex """
        if self.table_data() is None:
            return ''
        # self.rename_cols()
        self.latex = self.default_latex()
        # indent cc x 3, approp x2, fund x1
        self.add_tab(col=0, row_list=self.fund_rows(), length='0.5cm')
        self.add_tab(col=0, row_list=self.approp_rows(), length='1cm')
        self.add_tab(col=0, row_list=self.cc_rows(), length='1.5cm')
        self.add_tab(col=0, row_list=self.job_rows(), length='2cm')
        self.bold_rows([0, 1, self.last_row()] + self.fund_rows()  + self.approp_rows() + self.cc_rows())
        self.highlight_rows([1, self.last_row()], self.primary_color)
        self.highlight_rows(self.fund_rows(), self.color2)
        self.highlight_rows(self.cc_rows(), self.color2)
        self.outline_last_row(self.line_color)
        self.replace_header(self.header())
        return self.latex