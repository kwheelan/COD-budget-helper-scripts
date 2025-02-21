
from . import Sheet, Revenues, Expenditures, FTEs
import pandas as pd

class Summary():
    """ Create dataframes for the summary tables """

    def __init__(self, rev : Revenues, exp : Expenditures, positions : FTEs):
        self.rev = rev
        self.exp = exp
        self.ftes = positions

    @staticmethod
    def value_columns():
        return ['FY24 Actual',
                'FY25 Adopted', 
                'FY26 Adopted',
                'FY27 Forecast',
                'FY28 Forecast',
                'FY29 Forecast']

    def collapse(self, dept : str, data : Sheet):
        df = data.filter_by_dept(dept)
        general_fund = df[df['Fund #'] == '1000']

        # convert to numeric
        df[self.value_columns()] = df[self.value_columns()].replace(
            ',', '', regex=True).apply(
            pd.to_numeric, errors='coerce')

        # Build aggregation dictionary using list comprehension
        agg_dict = {col: 'sum' for col in self.value_columns()}

        # sum for every year
        general_fund = general_fund.agg(agg_dict).reset_index().set_index('index').transpose()
        all_funds = df.agg(agg_dict).reset_index().set_index('index').transpose()

        # alternate columns
        columns_alternated = []
        for col1, col2 in zip(general_fund.columns, all_funds.columns):
            columns_alternated.append(general_fund[col1])
            columns_alternated.append(all_funds[col2])

        df = pd.concat(columns_alternated, axis=1)
        return df
    
    def rev_rows(self, dept):
        return self.collapse(dept, self.rev)
    
    def exp_rows(self, dept):
        return self.collapse(dept, self.exp)
    
    # def total_row(self, df, name, col = 'title'):
    #     # Calculate the sum for each numeric column
    #     sum_row = df[self.value_columns()].sum()
    #     row_dict = {col: name}
    #     row_dict.update(sum_row)
    #     return pd.DataFrame([row_dict])

    def table1(self, dept):
        rev = self.rev_rows(dept)
        exp = self.exp_rows(dept)
        df = pd.concat([rev, exp], axis=0, ignore_index=True).reset_index(drop=True)

        # Calculate the difference between the second and first rows
        new_row = df.iloc[1] - df.iloc[0]

        # Convert `new_row` to a DataFrame
        new_row_df = pd.DataFrame(new_row)

        # Append the new row DataFrame to the original DataFrame
        df = pd.concat([df, new_row_df.T], ignore_index=True)

        # Add first row
        row_names = pd.DataFrame({'title': ['Total Revenues', 'Total Expenditures', 'Net Tax Cost']})
        df = pd.concat([row_names, df], axis=1)
        return df

    def table1_part1(self, dept):
        table = self.table1(dept)
        return table.iloc[:,0:7]
    
    def table1_part2(self, dept):
        table = self.table1(dept)
        return table.iloc[:, [0, 7, 8, 9, 10, 11, 12]]