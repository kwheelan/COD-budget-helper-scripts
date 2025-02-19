
from .baseSheet import Sheet

class Expenditures(Sheet):
    """ Import from budget model """

    def __init__(self, filepath):
        super().__init__(filepath, 'Expenditures')