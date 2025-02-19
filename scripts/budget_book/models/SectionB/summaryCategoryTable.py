
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
    
    @staticmethod
    def divider():
        return r'\\ \arrayrulecolor{lineblue}\hline' + '\n'

    def process_latex(self):
        # self.rename_cols()
        self.latex = self.default_latex()
        #self.bold_cols(['Category', 'Variance FY25 vs FY26'])
        self.bold_rows([0, 1, 8])
        self.highlight_rows([1, 8], ['lightblue', 'lightblue'])
        #self.merge_rows(col_name='Category')
        self.replace_header(self.header())
        return self.latex