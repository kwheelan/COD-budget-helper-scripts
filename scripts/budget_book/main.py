
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath, SHEETS_TO_LOAD

def main():
    table1 = FteTable(ExcelDF(filepath, sheet = SHEETS_TO_LOAD[0]))
    doc = BaseDoc([table1], 'sample_tables')
    doc.save_as_latex()
    doc.convert_to_pdf()

def test():
    pass

if __name__ == '__main__':
    # test()
    main()