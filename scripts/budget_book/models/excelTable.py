
from models import BaseDF, BaseTable
from constants import SHEETS_TO_LOAD
import pandas as pd

class ExcelDF(BaseDF):

    def __init__(self, filepath, sheet, start_row=6, end_row=44, start_col=1, end_col=11):
        print('Reading data from Excel...')
        tabs = pd.read_excel(filepath, sheet_name=SHEETS_TO_LOAD, header=None)
        df = tabs[sheet]
        # save header data
        self.header = df.iloc[0, 1]
        self.other_header_lines = [df.iloc[i, start_col] for i in range(1, 5) if type(df.iloc[i, start_col]) is str]
        # then trim to table data
        df = df.iloc[start_row:end_row, start_col:end_col]
        super().__init__(df)

class ExcelTable(BaseTable):
    def __init__(self, custom_df):
        main_header = custom_df.header
        subheaders = custom_df.other_header_lines
        super().__init__(custom_df, main_header, subheaders)
