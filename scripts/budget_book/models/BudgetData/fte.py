from .baseSheet import Sheet
import pandas as pd

class FTEs(Sheet):
    """ Import from budget model """

    def __init__(self, filepath):
        super().__init__(filepath, 'Positions')

    @staticmethod
    def value_columns():
        return ['FY25 Adopted FTE', 
                'FY26 Adopted FTE',
                'FY27 Forecast FTE',
                'FY28 Forecast FTE',
                'FY29 Forecast FTE']
    
    def group_by_fund_approp_cc(self, dept):
        df = self.filter_by_dept(dept)
        # convert to numeric
        df[self.value_columns()] = df[self.value_columns()].replace(
            ',', '', regex=True).apply(
            pd.to_numeric, errors='coerce')
        
        funds = list(set(self.processed['Fund #']))
        final_df = pd.DataFrame()
        for fund in funds:
            # for each fund, get sums by obj category
            fund_df = df[df['Fund #'] == fund]
            fund_df = fund_df[self.value_columns() + 
                              ['Appropriation #', 'Cost Center Name', 'Job Title']]
            # get approps
            approps = list(set(fund_df['Appropriation #']))
            approp_df = pd.DataFrame()
            for approp in approps:
                approp_df = fund_df[fund_df['Appropriation #'] == approp]
                # get cost centers
                cc_list = list(set(fund_df['Cost Center Name']))
                cc_df = pd.DataFrame()
                for cc in cc_list:
                    cc_df = approp_df[approp_df['Cost Center Name'] == cc]
                    # get job codes
                    cc_df = self.group_df_by_col(approp_df, col='Job Title')
                    # add the cc total on top
                    if cc_df.shape[0] > 0:
                        cc_total = self.total_row(cc_df, cc, 'Job Title')
                        cc_df = pd.concat([cc_total, cc_df])
                        final_df = pd.concat([final_df, cc_df])
                # add the approp total on top
                if approp_df.shape[0] > 0:
                    approp_total = self.total_row_without_double_counting(approp_df, self.approp_name(approp), 2, 'Job Title')
                    approp_df = pd.concat([approp_total, approp_df])
                    final_df = pd.concat([final_df, approp_df])
            # add the fund total on top
            if fund_df.shape[0] > 0:
                fund_total = self.total_row_without_double_counting(fund_df, self.fund_name(fund), 3, 'Job Title')
                final_df = pd.concat([fund_total, final_df])
        # add total to top and bottom
        # add sums to top and bottom of table
        bottom = self.total_row_without_double_counting(final_df, 'Grand Total', 4, 'Job Title')
        top = self.total_row_without_double_counting(final_df, self.dept_name(dept), 4, 'Job Title')

        # Add the sum rows to the top and bottom of the DataFrame
        df_with_sums_top = pd.concat([top, final_df], ignore_index=True)
        df = pd.concat([df_with_sums_top, bottom], ignore_index=True)

        # Remove rows where all specific columns have zero
        df = df[~(df[self.value_columns()].eq(0).all(axis=1))]
        df = df[['Job Title'] + self.value_columns()]

        # Round all float values to 2 decimal places
        df = df.map(self.round_floats)
        return df