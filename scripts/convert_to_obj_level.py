import pandas as pd
import re
import os

# Constants
YEAR = 26
FRINGE_TAB_NAME = f'FY{YEAR} Fringe Rates'

# File paths
DATA = 'input_data'
OUTPUT = 'output'

source_folder = f'{DATA}/detail_sheets'
# master_DS_filepath = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'
master_DS_filepath = f'{source_folder}/{os.listdir(source_folder)[0]}'


# columns to be retained from all sheets
cols_to_keep =  ['Fund',
                 'Appropriation', 
                 'Cost Center',
                 'Recurring or One-Time', 
                 'Baseline or Supplemental' ]

# Sheet names; columns to keep
SHEETS = {
    'FTE, Salary-Wage, & Benefits': {
        'cols': cols_to_keep + ['Salary/Wage Total Request (incl. GWI & Merit)',
                  'Fringe Benefits Package'],
        'header': 13
    },  
    'Overtime & Other Personnel': {
        'cols': cols_to_keep + ['Departmental Request OT/SP/Hol',
           'FICA Object',
           'OT/SP/Hol Object'],
        'header': 13
    },
    'Non-Personnel': {
        'cols': cols_to_keep + ['FY26 Departmental Request Total', 'Object'],
        'header': 17
    },
    'Revenue': {
        'cols': cols_to_keep + ['FY26 Departmental Estimate Total', 'Object'],
        'header': 13
    },
    FRINGE_TAB_NAME: {
        'header': 2
    }
}

def read_master_detail(sheet_filename):
    # read Excel using pandas
    xls = pd.ExcelFile(sheet_filename)

    # Read individual sheets into DataFrames
    dfs = {sheet: pd.read_excel(sheet_filename, sheet_name=sheet, header=SHEETS[sheet]['header']) for sheet in SHEETS}
    
    # Separate FY26 Fringe for lookup
    lookup_df = dfs.pop(FRINGE_TAB_NAME)
    
    return dfs, lookup_df

def clean_column_names(column_names):
    cleaned_columns = []
    for col in column_names:
        # Remove newline characters
        col = col.replace("\n", " ")
        
        # Replace multiple spaces with a single space and strip leading/trailing whitespace
        col = re.sub(' +', ' ', col).strip()
        
        cleaned_columns.append(col)
    
    return cleaned_columns

def clean_dataframe_columns(df):
    df.columns = clean_column_names(df.columns)
    return df

# ====================== Fringe package lookup =================================

class FringeLookup:
    def __init__(self, df):
        self.df = df
        self.process_fringe_lookup()

    def convert_to_long_format(self, df):
        # Identify id_vars and value_vars for melting
        id_vars = ['Fringe_Package']
        value_vars = [col for col in df.columns if col != 'Fringe_Package']
        
        # Use pd.melt to go from wide to long format
        df_long = pd.melt(df, id_vars=id_vars, value_vars=value_vars, 
                        var_name='Type', value_name='Value')
        
        # Extract the type and variable (object/rate) from the Type column
        df_long[['Type_Name', 'Variable']] = df_long['Type'].str.rsplit('_', n=1, expand=True)
        
        # Pivot the DataFrame so that object and rate are columns again
        df_long = df_long.pivot_table(index=['Fringe_Package', 'Type_Name'], 
                                    columns='Variable', values='Value', 
                                    aggfunc='first').reset_index()
        
        # Rename columns for clarity
        df_long.columns.name = None  # Remove the name of the columns axis
        df_long = df_long.rename(columns={'object': 'Object_Number', 'rate': 'Rate'})

        return df_long

    def process_fringe_lookup(self):
        self.df = self.df.iloc[:, 1:22]

        # Define the new column names based on the original ones
        original_columns = self.df.columns.tolist()
        new_columns = ['Fringe_Package']
        
        for i in range(1, len(original_columns), 2):
            base_name = original_columns[i].replace("\n", " ").replace("  ", " ").strip() 
            if pd.isna(base_name):
                base_name = original_columns[i-2]
            object_col = f'{base_name.strip()}_object'
            rate_col = f'{base_name.strip()}_rate'
            new_columns.extend([object_col, rate_col])
        
        # Rename the columns
        self.df.columns = new_columns

        # drop first row
        self.df = self.df.iloc[1:].reset_index(drop=True)

        # convert to long
        self.df = self.convert_to_long_format(self.df)

    def lookup_fringe(self, package, obj):
        df = self.df
        rate = df.loc[(df['Type_Name'] == package) & (df['Object_Number'] == obj), 'Rate']
        if not rate.empty:
            return int(rate.iloc[0])  # Ensure it returns an integer value
        return 0

    def obj_list(self):
        return self.df['Object_Number'].unique()
    
    def __repr__(self):
        return repr(self.df)

# Function to create an ID column by concatenating specified columns
def create_id_column(df, include_columns=None):
    if include_columns is not None:
        df = df.copy()  # Avoid modifying the original DataFrame
        df['ID'] = df[include_columns].apply(lambda row: '-'.join(row.values.astype(str)), axis=1)
    return df

def process_sheet(dfs, sheet_name, fringe_lookup):
    # clean column names
    df = clean_dataframe_columns(dfs[sheet_name])

    # filter to only relevant columns
    cols = SHEETS[sheet_name]['cols']

    # Ensure the columns being selected are present in the DataFrame
    missing_cols = [col for col in cols if col not in df.columns]
    if missing_cols:
        raise KeyError(f"Columns missing in DataFrame: {missing_cols}")
    # do the subsetting
    df = df.loc[:, cols]

    # process personnel
    if sheet_name == 'FTE, Salary-Wage, & Benefits':
        df = process_personnel(df, fringe_lookup)
        print(df)

    # ID column
    df = create_id_column(df, ['Fund', 'Appropriation', 'Cost Center', 
                               #'Object',
                              'Recurring or One-Time', 'Baseline or Supplemental'])
    return df

#================= Process personnel ========================================

# Function to process personnel data
def process_personnel(df, fringe_lookup):
    # Expand dataframe so each row becomes a row for each object in FringeLookup.obj_list()
    expanded_rows = []
    for _, row in df.iterrows():
        for obj in fringe_lookup.obj_list():
            new_row = row.copy()
            new_row['Object'] = obj
            rate = fringe_lookup.lookup_fringe(new_row['Fringe Benefits Package'], obj)
            new_row['Fringe Rate'] = rate
            new_row['Fringe Request'] = rate * new_row['Salary/Wage Total Request (incl. GWI & Merit)']
            expanded_rows.append(new_row)
    
    expanded_df = pd.DataFrame(expanded_rows)
    return expanded_df

# ==================== Main ==========================================

def main():
    dfs, lookup_df = read_master_detail(master_DS_filepath)
    # create lookup table
    fringeLookup = FringeLookup(lookup_df)
    # process all the dataframes
    for df_key in dfs:
        df_processed = process_sheet(dfs, df_key, fringeLookup)

if __name__ == '__main__':
    main()