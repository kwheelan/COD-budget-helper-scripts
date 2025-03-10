class BaseDF:
    """ A class to hold the table data for a budget book table """
    def __init__(self, df):
        self.raw_df = df
        self.processed = self.clean_table_data(df)

    def adjust_col_names(self, new_cols : list[str]) -> list[str]:
        """ Placeholder for renaming columns explicitly """
        columns_list = list(self.processed)
        if len(columns_list) == len(new_cols):
            self.processed.columns = new_cols
        else: 
            print ('Error: attempt to replace col names with list of wrong length')

    @staticmethod
    def adjust_headers(df):
        # Adjust column names
        df.columns = df.iloc[0]
        columns_list = list(df.columns)
        # remove extra spaces or new lines
        for i in range(len(columns_list)): 
            if type(columns_list[i]) is not str:
                columns_list[i] = ''
            columns_list[i] = columns_list[i].strip().replace('\n','')
        # Drop column names row from data and reset headers
        df.columns = columns_list
        df = df[1:].reset_index(drop=True)
        return df 

    @staticmethod
    def escape_ampersands(text):
        # Replace & with \& in all string elements
        if isinstance(text, str):
            return text.replace('&', r'\&')
        if isinstance(text, str):
            return text.replace('$', r'\$')
        return text
    
    @staticmethod
    def remove_new_lines(text):
        # Replace & with \& in all string elements
        if isinstance(text, str):
            return text.replace('\n', ' ').replace('  ', ' ')
        return text
    
    @staticmethod
    def round_floats(value):
        if isinstance(value, float) or isinstance(value, int):
            # avoid float imprecision errors
            value = round(value, 3)
            if value % 1 == 0:
                # then just add comma
                return f'{value:,.0f}'
            if round(value,1) == value:
                # then round to one decimal
                return f'{value:,.1f}'
            # otherwise, round to 2 places
            return f'{value:,.2f}'
        return value
    
    def clean_table_data(self, df):
        # Adjust columns
        df = self.adjust_headers(df)
        # Remove NAs
        df.fillna('', inplace=True)
        # Round all float values to 2 decimal places
        df = df.map(self.round_floats)
        # Apply the escaping & function to each element in the DataFrame
        df = df.map(self.escape_ampersands)
        df = df.map(self.remove_new_lines)
        return df

    def latex_ready_data(self):
        return self.processed