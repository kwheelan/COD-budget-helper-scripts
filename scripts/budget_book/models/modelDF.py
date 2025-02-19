
from models import BaseDF
import pandas as pd

class Expenditures(BaseDF):
    """ Import from budget model """

    def read_model(self, filepath, sheet):
        tabs = pd.read_excel(filepath, sheet_name=[sheet], header=None)
        df = tabs[sheet]
        # then trim to table data 
        # TODO: remove hard coding
        df = df.iloc[11:, 0:49]
        return df
    
    @staticmethod
    def value_columns():
        return ['FY25 Adopted', 
                'FY26 Adopted',
                'FY27 Forecast',
                'FY28 Forecast',
                'FY29 Forecast']
    
    def dept_name(self, dept):
        """ get department name from number """
        df = self.processed[(self.processed['Department #'] == dept)]
        return df['Department Name'][0]

    def group_by_category(self, dept):
        """ create a DF for expeditures by category """
        df = self.processed

        # filter by dept
        df = df[(df['Department #'] == dept)]
        dept_name = df['Department Name'][0]
        # Select specific columns
        df = df[self.value_columns() + ['Object Classification']]
        # convert to numeric
        df[self.value_columns()] = df[self.value_columns()].replace(
            ',', '', regex=True).apply(
            pd.to_numeric, errors='coerce')

        # Group by 'Major Classification' and aggregate
        df = df.groupby(
            'Object Classification').agg({
                'FY25 Adopted' : 'sum', 
                'FY26 Adopted': 'sum', 
                'FY27 Forecast': 'sum', 
                'FY28 Forecast': 'sum', 
                'FY29 Forecast': 'sum'
            }).reset_index()
        
        # Remove rows where all specific columns have zero
        df = df[~(df[self.value_columns()].eq(0).all(axis=1))]

        # Define the custom order for the 'Major Classification' column
        custom_order = [
            'Salaries \& Wages', 
            'Employee Benefits', 
            'Professional \& Contractual Services', 
            'Operating Supplies', 
            'Operating Services', 
            'Fixed Charges'
        ]

        # Convert the 'Major Classification' column to a categorical type
        df['Object Classification'] = pd.Categorical(df['Object Classification'], categories=custom_order, ordered=True)

        # Sort the DataFrame by 'Obj Classification' according to the custom order
        df = df.sort_values(by='Object Classification')

        # Reset index if necessary
        df.reset_index(drop=True, inplace=True)

        # Calculate the sum for each numeric column
        sum_row = df[self.value_columns()].sum()

        # add sums to top and bottom of table
        bottom = {'Object Classification': 'Grand Total'}
        bottom.update(sum_row)
        bottom = pd.DataFrame([bottom])

        top = {'Object Classification': dept_name}
        top.update(sum_row)
        top = pd.DataFrame([top])

        # Add the sum rows to the top and bottom of the DataFrame
        df_with_sums_top = pd.concat([top, df], ignore_index=True)
        df = pd.concat([df_with_sums_top, bottom], ignore_index=True)
        # Round all float values to 2 decimal places
        df = df.map(self.round_floats)

        return(df)

    def __init__(self, filepath):
        df = self.read_model(filepath, 'Expenditures')
        super().__init__(df)