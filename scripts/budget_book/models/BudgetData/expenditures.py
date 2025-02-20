
from .baseSheet import Sheet

class Expenditures(Sheet):
    """ Import from budget model """

    def __init__(self, filepath):
        super().__init__(filepath, 'Expenditures')

    @staticmethod
    def custom_order():
        # Define the custom order for the 'Major Classification' column
        return [
            'Salaries \& Wages', 
            'Employee Benefits', 
            'Professional \& Contractual Services', 
            'Operating Supplies', 
            'Operating Services', 
            'Fixed Charges',
            'Equipment Acquisition',
            'Other Expenses'
        ]