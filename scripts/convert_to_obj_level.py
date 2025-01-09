import pandas as pd
import re
import os

# Constants
YEAR = 26
FRINGE_TAB_NAME = f'FY{YEAR} Fringe Rates'

# File paths
DATA = 'input_data'
OUTPUT = 'output'

master_DS_filepath = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'
#source_folder = f'{DATA}/detail_sheets'
#master_DS_filepath = f'{source_folder}/{os.listdir(source_folder)[0]}'
output_file = f'{OUTPUT}/obj_level.xlsx'

# columns to be retained from all sheets
cols_to_keep =  ['Fund', 'Fund Name',
                 'Appropriation', 'Appropriation Name',
                 'Cost Center', 'Cost Center Name',
                 'Recurring or One-Time', 
                 'Baseline or Supplemental' ]

# columns to be retained from all sheets
full_cols_to_keep =  ['Fund', 'Fund Name',
                    'Appropriation', 'Appropriation Name',
                    'Cost Center', 'Cost Center Name',
                    'Object', #'Object Name',
                    'Recurring or One-Time', 
                    'Baseline or Supplemental' ]

totals = {
    'FTE, Salary-Wage, & Benefits': 'Budget Recommend Salary/Wage',
    'Overtime & Other Personnel': ['Budget Recommend OT/SP/Hol', 'Budget Recommend FICA'],
    'Non-Personnel': 'Budget Recommend Total',
    'Revenue': 'Budget Recommend Total',
}

# Sheet names; columns to keep
SHEETS = {
    'FTE, Salary-Wage, & Benefits': {
        'cols': cols_to_keep + ['Fringe Benefits Package'] + [totals['FTE, Salary-Wage, & Benefits']],
        'header': 13
    },  
    'Overtime & Other Personnel': {
        'cols': cols_to_keep + 
            ['FICA Object', 'OT/SP/Hol Object', 'Object Name', 'FICA Object Name'] + 
            totals['Overtime & Other Personnel'],
        'header': 13
    },
    'Non-Personnel': {
        'cols': full_cols_to_keep + [totals['Non-Personnel']],
        'header': 17
    },
    'Revenue': {
        'cols': full_cols_to_keep + [totals['Revenue']],
        'header': 13
    },
    FRINGE_TAB_NAME: {
        'header': 2
    }
}

def read_master_detail(sheet_filename):

    # Read individual sheets (tabs) into DataFrames
    dfs = {sheet: pd.read_excel(sheet_filename, 
                                sheet_name=sheet, 
                                header=SHEETS[sheet]['header']) 
            for sheet in SHEETS}
        
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
        id_vars = ['Fringe_Package']
        value_vars = [col for col in df.columns if col != 'Fringe_Package']
        
        df_long = pd.melt(df, id_vars=id_vars, value_vars=value_vars, 
                         var_name='Type', value_name='Value')
        
        df_long[['Type_Name', 'Variable']] = df_long['Type'].str.rsplit('_', n=1, expand=True)
        
        df_long = df_long.pivot_table(index=['Fringe_Package', 'Type_Name'], 
                                     columns='Variable', values='Value', 
                                     aggfunc='first').reset_index()
        
        df_long.columns.name = None
        df_long = df_long.rename(columns={'object': 'Object_Number', 'rate': 'Rate'})
        df_long['Object_Number'] = df_long['Object_Number'].astype(str)

        return df_long

    def process_fringe_lookup(self):
        self.df = self.df.iloc[:, 1:22]

        original_columns = self.df.columns.tolist()
        new_columns = ['Fringe_Package']
        
        for i in range(1, len(original_columns), 2):
            base_name = original_columns[i].replace("\n", " ").replace("  ", " ").strip()
            if pd.isna(base_name):
                base_name = original_columns[i-2]
            object_col = f'{base_name.strip()}_object'
            rate_col = f'{base_name.strip()}_rate'
            new_columns.extend([object_col, rate_col])
        
        self.df.columns = new_columns
        self.df = self.df.iloc[1:].reset_index(drop=True)
        self.df = self.convert_to_long_format(self.df)

    def lookup_fringe(self, package, obj):
        df = self.df
        lookup_result = df.loc[(df['Type_Name'] == package) & (df['Object_Number'] == str(obj))]
        rate = lookup_result['Rate']
        if not rate.empty:
            return float(rate.iloc[0])  # Ensure it returns a float value
        return 0.0

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
    print(f'Processing {sheet_name}')
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

    # process OT
    if sheet_name == 'Overtime & Other Personnel':
        df = process_OT(df)

    # Rename total to Amount
    if ('Amount' not in df.columns):
        df = df.rename(columns={totals[sheet_name] :'Amount'})

    # ID column
    df = create_id_column(df, ['Fund', 'Appropriation', 'Cost Center', 
                               'Object',
                              'Recurring or One-Time', 'Baseline or Supplemental'])
    
    # Create the aggregation dictionary
    agg_dict = {col: 'first' for col in full_cols_to_keep}
    agg_dict['Amount'] = 'sum'
    # Group by the ID column and aggregate using the aggregation dictionary
    grouped_df = df.groupby('ID', as_index=False).agg(agg_dict)
    
    return grouped_df

#================= Process personnel ========================================

# Function to process personnel data
def process_personnel(df, fringe_lookup):
    # Expand dataframe so each row becomes a row for each object in FringeLookup.obj_list()
    expanded_rows = []
    for _, row in df.iterrows():
        for obj in fringe_lookup.obj_list():
            new_row = row.copy()
            new_row['Object'] = obj
            package = new_row['Fringe Benefits Package']
            rate = fringe_lookup.lookup_fringe(package, obj)
            new_row['Fringe Rate'] = rate
            new_row['Amount'] = rate * new_row[totals['FTE, Salary-Wage, & Benefits']]
            expanded_rows.append(new_row)
    
    expanded_df = pd.DataFrame(expanded_rows)
    return expanded_df


def process_OT(df):
    df.rename(columns={'OT/SP/Hol Object' :'Object'})

    # Expand dataframe so each row becomes a row for each object in FringeLookup.obj_list()
    expanded_rows = []
    for _, row in df.iterrows():
        for obj in ['OT/SP/Hol', 'FICA']:
            new_row = row.copy()
            new_row['Object'] = new_row[f'{obj} Object']
            obj_name_col = f'{'FICA '*(obj == 'FICA')}Object Name'
            new_row['Object Name'] = new_row[obj_name_col]
            new_row['Amount'] = new_row[f'Budget Recommend {obj}']
            expanded_rows.append(new_row)
    
    expanded_df = pd.DataFrame(expanded_rows)
    return expanded_df

# ==================== Main ==========================================


def test():
    dfs, lookup_df = read_master_detail(master_DS_filepath)
    # create lookup table
    fringeLookup = FringeLookup(lookup_df)
    for obj in fringeLookup.obj_list():
        rate = fringeLookup.lookup_fringe('General City Civilian', obj)
        print(f'Obj: {obj}, rate: {rate*100}%') 

def main():
    dfs, lookup_df = read_master_detail(master_DS_filepath)
    # create lookup table
    fringeLookup = FringeLookup(lookup_df)
    dataframes = []

    sheets = list(dfs.keys())
    # process all the dataframes and add them to an array
    for i in range(len(sheets)):
        dataframes.append(process_sheet(dfs, sheets[i], fringeLookup))
    
    # row bind them
    df_processed = pd.concat(dataframes, ignore_index=True)

    # delete emty rows
    cleaned_df = df_processed[df_processed['Amount'] > 0]
    
    # save as excel
    cleaned_df.to_excel(output_file, index=False)

if __name__ == '__main__':
    main()