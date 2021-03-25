import pandas


class Database:

    def __init__(self):
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def excel_to_dataframe(self):
        # Extracts data from excel tables of a sheet and converts it into a dataframe
        pass

    def save_database_json(self):
        # Saves the dictionary containing the database into a json file
        # The database file already in the data folder will not be overwritten,
        # but moved to the backup folder
        pass

    def load_database_json(self):
        # Loads the json file containing the database into a usable dictionary
        # If there is no database available, it will use the most recent backup database file
        pass

    def dataframe_to_dict(self):
        # Converts a standard database dataframe into a dictionary
        pass

    def dict_to_dataframe(self):
        # Converts a standard database dictionary into a dataframe
        pass

    def remove_transaction_from_dataframe(self):
        # Removes the row from dataframe containing the transaction based on index
        pass

    def new_transaction_to_dataframe(self):
        # Enters a new row in the dataframe, containing the new transaction
        pass

    def update_transactions_in_dataframe(self):
        # Compares all indexes of updated filtered dataframe
        pass

    def filter_data_from_dataframe(self):
        # Returns a filtered dataframe based on the filters applied to the original database
        pass

    def backup_old_database(self):
        # Moves the current database json file to the backup folder and changes it name with a suffix timestamp
        pass

    def restore_old_database(self):
        # Searches for the most recent database json file in the backup folder,
        # moves it to the data folder and changes it name to database.json
        pass

    def get_old_ifo_data(self):
        # Extracts database data from the excel table of an old excel sheet and converts it into a dataframe
        pass

    def get_filtered_excel_data(self):
        # Extracts the filtered excel data which has been updated by user and converts it into a dataframe
        pass

    def database_processing_startup(self):
        # Checks if parent function has valid entry parameters. If not, it created the valid entry parameters required.
        pass


def tester():
    test = Database()
