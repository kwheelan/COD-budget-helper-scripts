
from . import Sheet, Revenues, Expenditures, FTEs
from models import BaseDF
import pandas as pd

class Summary(BaseDF):
    """ Create dataframes for the summary tables """

    def __init__(self, rev : Revenues, exp : Expenditures, positions : FTEs):
        self.rev = rev
        self.exp = exp
        self.ftes = positions
        # placeholder
        df = rev.raw_df
        super().__init__(df)

    @staticmethod
    def value_columns():
        return ['FY24 Actual',
                'FY25 Adopted', 
                'FY26 Adopted',
                'FY27 Forecast',
                'FY28 Forecast',
                'FY29 Forecast']

    def collapse(self, dept : str, data : Sheet):

        # get department data
        df = data.filter_by_dept(dept)

        # convert to numeric
        df[self.value_columns()] = df[self.value_columns()].replace(
            ',', '', regex=True).apply(
            pd.to_numeric, errors='coerce')

        # separate GF from all funds
        df['Fund #'] = df['Fund #'].replace(',', '', regex=True)
        general_fund = df[df['Fund #'] == '1000']

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

    def total_row(self, df, name, cols, col = 'title'):
        # Calculate the sum for each numeric column
        sum_row = df[cols].sum()
        row_dict = {col: name}
        row_dict.update(sum_row)
        return pd.DataFrame([row_dict])

    def table3(self, dept):
        df = self.exp.filter_by_dept(dept)

        # columns
        cols = ['FY25 Adopted', 'FY26 Adopted']

        # convert to numeric
        df[cols] = df[cols].replace(
            ',', '', regex=True).apply(
            pd.to_numeric, errors='coerce')
        
        # separate GF from all funds
        df['Fund #'] = df['Fund #'].replace(',', '', regex=True)
        df = df[df['Fund #'] == '1000']

        # two segments
        recurring = df[df['Rec vs 1-Time'] == 'Recurring']
        onetime = df[df['Rec vs 1-Time'] == 'One-Time']

        # Build aggregation dictionary using list comprehension
        agg_dict = {col: 'sum' for col in cols}

        # sum for every year
        recurring = recurring.agg(agg_dict).reset_index().set_index('index').transpose()
        onetime = onetime.agg(agg_dict).reset_index().set_index('index').transpose()

        # combine rows
        df = pd.concat([recurring, onetime], axis=0, ignore_index=True).reset_index(drop=True)

        # Add first column
        row_names = pd.DataFrame({'title': ['Recurring Expenditures', 'One-Time Expenditures']})
        df = pd.concat([row_names, df], axis=1)

        # add total row on bottom
        total = self.total_row(df, 'Total Expenditures', cols=cols)
        df = pd.concat([df, total])

        return df
    
    def table4(self, dept, filepath):
        df = self.ftes.filter_by_dept(dept)

        cols = ['FY25 Adopted FTE', 
                'FY26 Adopted FTE',
                'FY27 Forecast FTE',
                'FY28 Forecast FTE',
                'FY29 Forecast FTE']

        # convert to numeric
        df[cols] = df[cols].replace(
            ',', '', regex=True).apply(
            pd.to_numeric, errors='coerce')
        
        # separate GF from all funds
        df['Fund #'] = df['Fund #'].replace(',', '', regex=True)
        general_fund = df[df['Fund #'] == '1000']
        non_gf = df[df['Fund #'] != '1000']
        # Note: I guess no ARPA positions for FY25+

        # Build aggregation dictionary using list comprehension
        agg_dict = {col: 'sum' for col in cols}

        # sum for every year
        general_fund = general_fund.agg(agg_dict).reset_index().set_index('index').transpose()
        non_gf = non_gf.agg(agg_dict).reset_index().set_index('index').transpose()

        # combine rows
        df = pd.concat([general_fund, non_gf], axis=0, ignore_index=True).reset_index(drop=True)

        # add ARPA row on bottom
        arpa = pd.DataFrame({col: [0] for col in cols})
        df = pd.concat([df, arpa]).reset_index(drop=True)

        # Add first column
        row_names = pd.DataFrame({'title': ['General Fund', 'Non-General Fund', 'ARPA']})
        # replace 2024 actual with current counts
        current = self.current_positions(dept, filepath)
        df = pd.concat([row_names, current, df], axis=1)

        # add total row on bottom
        total = self.total_row(df, 'Total Positions', cols + ['Actual'])
        df = pd.concat([df, total]).reset_index(drop=True)

        return df
    
    def current_positions(self, dept, filepath):
        """ Get counts of actuals right now """

        sheet = 'Section B Dept Summary'
        tabs = pd.read_excel(filepath, sheet_name=[sheet], header=None)
        df = tabs[sheet]
        # then trim to table data 
        df = df.iloc[3:38, 9:14]
        df = BaseDF.adjust_headers(df)
        # Filter to department
        dept_df = df[df['Dept'] == int(dept)]
        data = [list(dept_df['GF'])[0], 
                list(dept_df['NonGF'])[0], 
                list(dept_df['ARPA'])[0]]
        # build and return df
        df = pd.DataFrame({'Actual' : data})
        df['Actual'] = df['Actual'].apply(pd.to_numeric, errors='coerce')
        return df





        
