
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath, SHEETS_TO_LOAD

def main():
    table1 = ExpenditureCategories(ExcelDF(filepath, sheet = SHEETS_TO_LOAD[1],
                              start_row = 9, end_row=18, 
                              start_col=1, end_col=7))
    table2 = RevenueCategories(ExcelDF(filepath, sheet = SHEETS_TO_LOAD[1],
                              start_row = 9, end_row=15, 
                              start_col=10, end_col=16))
    doc = BaseDoc([table1, table2], 'sample_tables')
    doc.save_as_latex()
    doc.convert_to_pdf()

def test():
    pass
    # out1 = cline([1,2,3,4,5], 0)
    # out2 = cline([1,2,3,4,5], 4)
    # out3 = cline([1,2,3,4,5], 2)
    # print(out1)
    # print(out2)
    # print(out3)

if __name__ == '__main__':
    # test()
    main()