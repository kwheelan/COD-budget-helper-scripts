
from models import ExcelTable

class SummaryCategoryTable(ExcelTable):
    """
    Represents Section B Department tables
    Expenditures and Revenues by Summary Category
    """

    def __init__(self, custom_df):
        super().__init__(custom_df)

    def column_format(self, format=None):
        n_cols = len(self.table_data().columns)
        return r'|>{\centering\arraybackslash}p{4cm}|' + r'>{\centering\arraybackslash}p{2cm}|' * (n_cols - 1)

    def process_latex(self):
        # self.rename_cols()
        self.latex = self.default_latex()
        #self.bold_cols(['Category', 'Variance FY25 vs FY26'])
        self.bold_rows([0, 1, 8])
        self.highlight_rows([1, 8], ['lightblue', 'lightblue'])
        #self.merge_rows(col_name='Category')
        return self.latex