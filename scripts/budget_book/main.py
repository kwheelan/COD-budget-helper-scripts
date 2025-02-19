
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath
import warnings

def main():
    warnings.filterwarnings("ignore")
    # tables
    cat_exp = ExpenditureCategories(filepath, '10')
    cat_rev = RevenueCategories(filepath, '10')
    cat_fund_exp = ExpFundCatTable(filepath, '10')
    cat_fund_rev = RevFundCatTable(filepath, '10')

    # build doc
    doc = BaseDoc([cat_exp, cat_rev, cat_fund_exp, cat_fund_rev], 'sample_tables')
    doc.save_as_latex()
    doc.convert_to_pdf()

def test():
    # Expenditures(filepath, 'Expenditures')
    rev = Revenues(filepath)
    print(rev.group_by_category_and_fund('10'))

if __name__ == '__main__':
    # test()
    main()