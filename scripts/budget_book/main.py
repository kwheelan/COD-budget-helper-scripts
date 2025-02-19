
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath
import warnings

def main():
    # tables
    cat_exp = ExpenditureCategories(filepath, '10')
    cat_rev = RevenueCategories(filepath, '10')
    cat_fund_exp = ExpFundCatTable(filepath, '10')
    cat_fund_rev = RevFundCatTable(filepath, '10')
    all_exp = ExpFullTable(filepath, '10')
    all_rev = RevFullTable(filepath, '10')
    # all_fte = FTEFullTable(filepath, '10')

    # build doc
    doc = BaseDoc([
        cat_exp, 
        cat_rev, 
        cat_fund_exp, 
        cat_fund_rev,
        all_exp,
        all_rev,
        # all_fte
        ], 'sample_tables')
    doc.save_as_latex()
    doc.convert_to_pdf()

def test():
    # Expenditures(filepath, 'Expenditures')
    rev = FTEs(filepath)
    print(rev.group_by_fund_approp_cc('10'))

if __name__ == '__main__':
    warnings.filterwarnings("ignore")
    test()
    # main()