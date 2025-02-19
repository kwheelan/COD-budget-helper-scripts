

from . import Sheet

class Revenues(Sheet):
    """ Import from budget model """

    def __init__(self, filepath):
        super().__init__(filepath, 'Revenues')

    @staticmethod
    def custom_order():
        # Define the custom order for the 'Major Classification' column
        return [
            'Grants, Shared Taxes, \& Revenues', 
            'Revenues from Use of Assets', 
            'Sales of Assets \& Compensation for Losses', 
            'Sales \& Charges for Services', 
            'Fines, Forfeits, \& Penalties', 
            'Contributions \& Transfers'
        ]