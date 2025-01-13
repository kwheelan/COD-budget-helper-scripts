import pandas as pd
import re
import os
import warnings

# Constants
YEAR = 26
FRINGE_TAB_NAME = f'FY{YEAR} Fringe Rates'

# File paths
DATA = 'input_data'
OUTPUT = 'output'

master_DS_filepath = f'{OUTPUT}/master_DS/master_detail_sheet_FY26.xlsx'
output_file = f'{OUTPUT}/obj_level.xlsx'

SOURCE_FOLDER = 'C:/Users/katrina.wheelan/OneDrive - City of Detroit/Documents - M365-OCFO-Budget/BPA Team/FY 2026/1. Budget Development/08A. Deputy Budget Director Recommend/'


# columns to be retained from all sheets
cols_to_keep =  ['Fund', #'Fund Name',
                 'Appropriation', #'Appropriation Name',
                 'Cost Center', #'Cost Center Name',
                 'Recurring or One-Time', 
                 'Baseline or Supplemental']

# columns to be retained from all sheets
full_cols_to_keep =  ['Fund', #'Fund Name',
                    'Appropriation', #'Appropriation Name',
                    'Cost Center', #'Cost Center Name',
                    'Object', #'Object Name',
                    'Recurring or One-Time', 
                    'Baseline or Supplemental',
                    'Service']

totals = {
    'FTE, Salary-Wage, & Benefits': 'Budget Director Salary/Wage',
    'Overtime & Other Personnel': ['Budget Director OT/SP/Hol', 'Budget Director FICA'],
    'Non-Personnel': 'Budget Director Total',
    'Revenue': 'Budget Director Total',
}

# Sheet names; columns to keep
SHEETS = {
    'FTE, Salary-Wage, & Benefits': {
        'cols': cols_to_keep + ['Fringe Benefits Package', 'Service'] + [totals['FTE, Salary-Wage, & Benefits']],
        'header': 13
    },  
    'Overtime & Other Personnel': {
        'cols': cols_to_keep + 
            ['FICA Object', 'OT/SP/Hol Object', 'Object Name', 'FICA Object Name', 'Service'] + 
            totals['Overtime & Other Personnel'],
        'header': 13
    },
    'Non-Personnel': {
        'cols': full_cols_to_keep + [totals['Non-Personnel']],
        'header': 17
    },
    # 'Revenue': {
    #     'cols': full_cols_to_keep + [totals['Revenue']],
    #     'header': 13
    # },
    FRINGE_TAB_NAME: {
        'header': 2
    }
}

def find_DS(folder, keyword='Detail Sheet', exclude=[], verbose=False):
    # Get full file path
    folder_fp = os.path.join(SOURCE_FOLDER, folder)

    # grab list of all files in that folder
    files = os.listdir(folder_fp)

    # generate list of reviewed detail sheets
    reviewed_DS = []
    for file in files:
        if (keyword in file) and ('.xlsx' in file) and not (sum([(word in file) for word in exclude])):
            reviewed_DS.append(os.path.join(SOURCE_FOLDER, folder, file))

    # return message
    message = None
    if len(reviewed_DS) > 1:
        message = f'Multiple potential reviewed detail sheets: {reviewed_DS}'
    elif len(reviewed_DS) == 0:
        message = f'No reviewed detail sheet found in {folder}'
    if verbose and message:
        print(message)

    return reviewed_DS

def read_DS(sheet_filename):

    # Read individual sheets (tabs) into DataFrames
    dfs = {sheet: pd.read_excel(sheet_filename, 
                                sheet_name=sheet, 
                                header=SHEETS[sheet]['header']) 
            for sheet in SHEETS}
    
    dfs = clean_rows(dfs)
        
    # Separate FY26 Fringe for lookup
    lookup_df = dfs.pop(FRINGE_TAB_NAME)
    return dfs, lookup_df

def clean_rows(dfs):
    # Process each DataFrame to remove rows with all NaNs or zeros
    dfs_cleaned = {}
    for sheet, df in dfs.items():
        # Drop rows where the 'Fund' column is NaN
        if 'Fund' in df.columns:
            cleaned_df = df.dropna(subset=['Fund'])
        else:
            print(f"Warning: 'Fund' column not found in {sheet}")
            cleaned_df = df  # Optionally retain the DataFrame as-is if the 'Fund' column is not found
        dfs_cleaned[sheet] = cleaned_df
    return dfs_cleaned


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
        self._raw = df
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
        self.add_salary_wage()

    def add_salary_wage(self):
        sal_wag = self._raw.iloc[:10, [29,31]]
        sal_wage_df = pd.DataFrame( {
            'Fringe_Package' : 'Salary/Wage',
            'Type_Name' : sal_wag.iloc[:, 0],
            'Object_Number' : [str(obj) for obj in sal_wag.iloc[:, 1]],
            'Rate' : 1.0
        } )
        self.df = pd.concat([self.df, sal_wage_df], ignore_index=True)


    def lookup_fringe(self, package, obj):
        # Return fringe rate given package and object number
        df = self.df
        lookup_result = df.loc[(df['Type_Name'] == package) & (df['Object_Number'] == obj)]
        rate = lookup_result['Rate']
        if not rate.empty:
            return float(rate.iloc[0])  # Ensure it returns a float value
        #print(f'Could not find rate for package {package} and object {obj}')
        return 0.0

    def obj_list(self):
        return [str(obj) for obj in self.df['Object_Number'].unique()]

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
        raise KeyError(f"Columns missing in {sheet_name}: {missing_cols}")
    # do the subsetting
    df = df.loc[:, cols]

    if df.empty:
        return df

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
                              'Recurring or One-Time', 'Baseline or Supplemental',
                              'Service'])
    
    # Create the aggregation dictionary
    agg_dict = {col: 'first' for col in full_cols_to_keep}
    agg_dict['Amount'] = 'sum'

    # Group by the ID column and aggregate using the aggregation dictionary
    try:
        grouped_df = df.groupby('ID', as_index=False).agg(agg_dict)
    except Exception as e:
        print(df)
        print("Error during group-by or aggregate operation:", e)
        raise
    
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

    # Expand dataframe so each row becomes a row for each object in OT/FICA
    expanded_rows = []
    for _, row in df.iterrows():
        for obj in ['OT/SP/Hol', 'FICA']:
            new_row = row.copy()
            new_row['Object'] = new_row[f'{obj} Object']
            obj_name_col = f'{'FICA '*(obj == 'FICA')}Object Name'
            new_row['Object Name'] = new_row[obj_name_col]
            new_row['Amount'] = new_row[f'Budget Director {obj}']
            expanded_rows.append(new_row)
    
    expanded_df = pd.DataFrame(expanded_rows)
    return expanded_df

def convert(filename):
    dfs, lookup_df = read_DS(filename)
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
    return cleaned_df

def save_to_excel(df, output_file):
    # save as excel
    df.to_excel(output_file, index=False)

def create_file_list(verbose = False):
    # all folders in analyst review section
    DS_FOLDERS = os.listdir(SOURCE_FOLDER)

    # initialize list and add viable files to it
    DS_list = []
    for folder in DS_FOLDERS:
        folder_fp = os.path.join(SOURCE_FOLDER, folder)
        if(os.path.isdir(folder_fp)):
            DS_list += find_DS(folder, verbose=verbose)

    return DS_list

def include_dept(file):
    for dept in INCLUDE:
        if dept in file:
            return dept
    return False

# ==================== Main ==========================================


def test():
    _, lookup_df = read_DS(master_DS_filepath)
    # create lookup table
    fringeLookup = FringeLookup(lookup_df)
    for obj in fringeLookup.obj_list():
        rate = fringeLookup.lookup_fringe('Uniform Fire', obj)
        print(f'Obj: {obj}, rate: {rate*100}%') 

INCLUDE = ['BSEED',
        #    'DPW',
        #    'DDoT',
        #    'Fire',
        #    'Health',
           '28 HR',
        #    'CRIO',
        #    'DoIT',
        #    'Mayor',
        #    'Parking',
        #    'HRD Classic',
        #    'HRD JET',
        #    'DAH',
        #    'GSD',
        #    'Elections',
        #    'Law',
        #    'Police'
]

def main():
    warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")
    warnings.filterwarnings("ignore", message=" Print area cannot be set to Defined name")

    #initialize DF
    df = pd.DataFrame()

    # get list of DS files
    DS_list = [file for file in create_file_list() if include_dept(file)]
    for detail_sheet in DS_list:
        print(f'Processing {detail_sheet}')
        df = pd.concat([df, convert(detail_sheet)], ignore_index=True)
    save_to_excel(df, output_file)

if __name__ == '__main__':
    main()