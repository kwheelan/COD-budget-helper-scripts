import os

# File paths
DATA = 'input_data/model_copy'
file = os.listdir(DATA)[0]
filepath = os.path.join(DATA, file)
OUTPUT = 'output/budget_book/'

SHEETS_TO_LOAD = [
    'Section A FTE',
    'Section B Dept Packet'
]

MODEL_DATA = [
    'Expenditures',
    'Revenue',
    'Positions',
    'FY24-FY25'
]