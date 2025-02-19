
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath, SHEETS_TO_LOAD

def main():
    table1 = ExpenditureCategories(filepath, '10')
    # table2 = RevenueCategories(ExcelDF(filepath, sheet = SHEETS_TO_LOAD[1],
    #                           start_row = 9, end_row=15, 
    #                           start_col=10, end_col=16))
    doc = BaseDoc([table1], 'sample_tables')
    doc.save_as_latex()
    doc.convert_to_pdf()

def test():
    Expenditures(filepath, 'Expenditures')

if __name__ == '__main__':
    # test()
    main()