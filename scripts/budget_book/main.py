
"""
Main script to automate the creation of the budget book
"""

from models import *
from constants import filepath, WORD_FP
import warnings

exp = Expenditures(filepath)
rev = Revenues(filepath)
positions = FTEs(filepath)

def all_tables(filepath, dept):
    # cat_exp = ExpenditureCategories(filepath, dept, exp)
    # cat_rev = RevenueCategories(filepath, dept, rev)
    # cat_fund_exp = ExpFundCatTable(filepath, dept, exp)
    # cat_fund_rev = RevFundCatTable(filepath, dept, rev)
    all_exp = ExpFullTable(filepath, dept, exp)
    # all_rev = RevFullTable(filepath, dept, rev)
    # all_fte = FTEFullTable(filepath, dept, positions)
    return [
        # cat_exp, 
        # cat_rev, 
        # cat_fund_exp, 
        # cat_fund_rev,
        all_exp,
        # all_rev,
        # all_fte
        ]

def main(depts):
    file_list = []
    for i in depts:
        print(f'Processing Dept {i}:')
        # tables
        tables = all_tables(filepath, dept=i)

        # build tables doc
        folder = os.path.join(OUTPUT, f'Dept{i}')
        if not os.path.exists(folder):
            os.mkdir(folder)
        filename = f'Dept{i}/Department {i} Tables'
        doc = BaseDoc(tables, filename)
        print('writing latex....')
        doc.save_as_latex()
        
        print('converting tables to pdf...')
        doc.convert_to_pdf()

        # convert Word narrative to PDF
        print(f'converting narrative Word doc to PDF')
        output_file = os.path.join(folder, f'Department {i} Narrative.pdf')
        PDFTool.find_and_convert(WORD_FP, dept=i, pdf_output_path=output_file)

        file_list.append(output_file)
        file_list.append(os.path.join(OUTPUT, f'{filename}.pdf'))
    
    # merge all PDFs
    print('merging final PDF...')
    final_pdf = os.path.join(OUTPUT, 'Merged Section B.pdf')
    PDFTool.combine_pdfs(file_list, final_pdf)
    print('Done!')

def test():
    # Expenditures(filepath, 'Expenditures')
    rev = FTEs(filepath)
    print(rev.group_by_fund_approp_cc('10'))

depts = [
    # '10', 
    # '13', 
    # '18', 
    '19', 
    # '20'
    ]

if __name__ == '__main__':
    warnings.filterwarnings("ignore")
    # test()
    main(depts)