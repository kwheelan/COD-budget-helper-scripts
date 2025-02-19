import os

# File paths
filepath = r'c:\Users\katrina.wheelan\OneDrive - City of Detroit\Documents - M365-OCFO-Budget\BPA Team\FY 2026\1. Budget Development\08B. THE MODEL\FY26 Budget Model - Budget Director Recommendation, 2025.02.15 - Updated Fringe Rates.xlsx'
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