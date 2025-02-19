
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath, SHEETS_TO_LOAD

def main():
    table1 = ExpenditureCategories(filepath, '10')
    table2 = RevenueCategories(filepath, '10')
    doc = BaseDoc([table1, table2], 'sample_tables')
    doc.save_as_latex()
    doc.convert_to_pdf()

def test():
    # Expenditures(filepath, 'Expenditures')
    rev = Revenues(filepath)
    print(rev.group_by_category('10'))

if __name__ == '__main__':
    # test()
    main()