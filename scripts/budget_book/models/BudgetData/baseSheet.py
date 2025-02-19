
from models import BaseDF
import pandas as pd

class Sheet(BaseDF):
    """ Import from budget model """

    def __init__(self, filepath, sheet):
        df = self.read_model(filepath, sheet)
        super().__init__(df)

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
    
    @staticmethod
    def custom_order():
        # Define the custom order for the 'Major Classification' column
        return [
            'Salaries \& Wages', 
            'Employee Benefits', 
            'Professional \& Contractual Services', 
            'Operating Supplies', 
            'Operating Services', 
            'Fixed Charges'
        ]

    def group_df_by_category(self, df):
        """ create a DF for expeditures by category """
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

        # Convert the 'Major Classification' column to a categorical type
        df['Object Classification'] = pd.Categorical(df['Object Classification'], 
                        categories=self.custom_order(), ordered=True)

        # Sort the DataFrame by 'Obj Classification' according to the custom order
        df = df.sort_values(by='Object Classification')

        # Reset index if necessary
        df.reset_index(drop=True, inplace=True)
        return(df)
    
    def total_row(self, df, name):
        # Calculate the sum for each numeric column
        sum_row = df[self.value_columns()].sum()
        row_dict = {'Object Classification': name}
        row_dict.update(sum_row)
        return pd.DataFrame([row_dict])
    
    def total_row_without_double_counting(self, df, name, n):
        # Calculate the sum for each numeric column divided by n
        sum_row = df[self.value_columns()].sum() / n
        row_dict = {'Object Classification': name}
        row_dict.update(sum_row)
        return pd.DataFrame([row_dict])
    
    def filter_by_dept(self, dept):
        # filter by dept
        df = self.processed
        return df[(df['Department #'] == dept)]

    def group_by_category(self, dept):
        # filter by dept
        df = self.group_df_by_category(df.filter_by_dept(dept))

        # add sums to top and bottom of table
        bottom = self.total_row(df, 'Grand Total')
        top =  self.total_row(df, self.dept_name(dept))
        {'Object Classification': self.dept_name(dept)}

        # Add the sum rows to the top and bottom of the DataFrame
        df_with_sums_top = pd.concat([top, df], ignore_index=True)
        df = pd.concat([df_with_sums_top, bottom], ignore_index=True)
        # Round all float values to 2 decimal places
        df = df.map(self.round_floats)

    def fund_name(self, fund):
        """ get fund name from number """
        df = self.processed[(self.processed['Fund #'] == fund)]
        return list(df['Fund Name'])[0]

    def group_by_category_and_fund(self, dept):
        df = self.filter_by_dept(dept)
        funds = list(set(self.processed['Fund #']))
        final_df = pd.DataFrame()
        for fund in funds:
            # for each fund, get sums by obj category
            fund_df = df[df['Fund #'] == fund]
            fund_df = self.group_df_by_category(fund_df)
            # add the fund total on top
            if fund_df.shape[0] > 0:
                fund_total = self.total_row(fund_df, self.fund_name(fund))
                fund_df = pd.concat([fund_total, fund_df])
                final_df = pd.concat([final_df, fund_df])
        # add total to top and bottom
        # add sums to top and bottom of table
        bottom = self.total_row_without_double_counting(final_df, 'Grand Total', 2)
        top = self.total_row_without_double_counting(final_df, self.dept_name(dept), 2)

        # Add the sum rows to the top and bottom of the DataFrame
        df_with_sums_top = pd.concat([top, final_df], ignore_index=True)
        df = pd.concat([df_with_sums_top, bottom], ignore_index=True)
        # Round all float values to 2 decimal places
        df = df.map(self.round_floats)
        return df
        
        

    