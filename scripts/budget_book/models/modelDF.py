
from models import BaseDF
import pandas as pd

class ModelDF(BaseDF):
    """ Import from budget model """

    def read_model(filepath, sheet):
        tabs = pd.read_excel(filepath, sheet_name=[sheet], header=None)
        df = tabs[sheet]
        # then trim to table data 
        # TODO: remove hard coding
        df = df.iloc[11:22, 0:50]
        return df

    def __init__(self, filepath, sheet):
        df = self.read_model(filepath, sheet)
        super().__init__(df)