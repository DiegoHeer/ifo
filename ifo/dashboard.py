import pandas


class Dashboard:

    def __init__(self):
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def fill_in_data_validation_cells(self):
        # Takes the specific variable list and updates the data validation list of the specific named range cell
        pass

    def get_data_validation_list(self):
        # Gets a list without duplicates of specific variables required to fill in the data validation cells.
        # One to three filters can be applied to get the correct list
        pass

    def general_data_validation_update(self):
        # Updates all data based on general filters, like currency, year or month
        pass

    def specific_data_validation_update(self):
        # Updates only a named range cell based on specific filters, like bank accounts
        pass

    def get_all_current_data_validation_selections(self):
        # Gets all current data validation selections of the Dashboard and returns it as a dictionary
        pass

    def dashboard_processing_startup(self):
        # Checks if parent function has valid entry parameters. If not, it created the valid entry parameters required.
        pass


def tester():
    test = Dashboard()
