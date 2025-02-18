
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
    # out1 = cline([1,2,3,4,5], 0)
    # out2 = cline([1,2,3,4,5], 4)
    # out3 = cline([1,2,3,4,5], 2)
    # print(out1)
    # print(out2)
    # print(out3)

if __name__ == '__main__':
    # test()
    main()