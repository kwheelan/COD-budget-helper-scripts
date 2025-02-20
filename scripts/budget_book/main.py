
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath
import warnings

exp = Expenditures(filepath)
rev = Revenues(filepath)
positions = FTEs(filepath)

def all_tables(filepath, dept):
    cat_exp = ExpenditureCategories(filepath, dept, exp)
    cat_rev = RevenueCategories(filepath, dept, rev)
    cat_fund_exp = ExpFundCatTable(filepath, dept, exp)
    cat_fund_rev = RevFundCatTable(filepath, dept, rev)
    all_exp = ExpFullTable(filepath, dept, exp)
    all_rev = RevFullTable(filepath, dept, rev)
    all_fte = FTEFullTable(filepath, dept, positions)
    return [
        cat_exp, 
        cat_rev, 
        cat_fund_exp, 
        cat_fund_rev,
        all_exp,
        all_rev,
        all_fte
        ]

def main(depts):
    for i in depts:
        print(f'processing dept {i}:')
        # tables
        tables = all_tables(filepath, dept=i)

        # build doc
        folder = os.path.join(OUTPUT, f'Dept{i}')
        if not os.path.exists(folder):
            os.mkdir(folder)
        doc = BaseDoc(tables, f'Dept{i}/Department {i}')
        doc.save_as_latex()
        doc.convert_to_pdf()

def test():
    # Expenditures(filepath, 'Expenditures')
    rev = FTEs(filepath)
    print(rev.group_by_fund_approp_cc('10'))

depts = ['13']

if __name__ == '__main__':
    warnings.filterwarnings("ignore")
    # test()
    main(depts)