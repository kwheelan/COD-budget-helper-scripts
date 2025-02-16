
from models import ExcelTable

class FteTable(ExcelTable):
    """
    Represents Section A FTE tables:
        - Total positions by FTE - All funds
        - Total Positions by FTE - General fund
        - Total Positions by FTE - Non-general fund
    """

    def __init__(self, custom_df):
        # main_header = "FY2025 - FY2029 Budgeted Positions by Department"
        super().__init__(custom_df)

    def column_format(self, format=None):
        n_cols = len(self.table_data().columns)
        return r'|>{\centering\arraybackslash}p{2.5cm}|' + 'p{0.5cm}|' + 'p{2.75cm}|' + r'>{\centering\arraybackslash}p{1.25cm}|' * (n_cols - 3)

    def process_latex(self):
        # self.rename_cols()
        self.latex = self.default_latex()
        self.bold_cols(['Category', 'Variance FY25 vs FY26'])
        self.bold_rows([0, 29, 36, 37])
        self.highlight_rows([0, 29, 36, 37], ['orange!50', 'orange!25', 'orange!25', 'orange!50'])
        return self.latex