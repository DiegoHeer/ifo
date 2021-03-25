import pandas as pd
from os.path import join
import pathlib
import json


class Database:

    def __init__(self):
        # General paths and sheet definitions
        self.wb_path = join(pathlib.Path(__file__).parent.absolute(), "IFO.xlsm")
        self.database_path = join(pathlib.Path(__file__).parent.absolute(), 'data', 'database.json')
        self.database_sheet_name = "Database"
        self.temporary_sheet_name = "Filtered Data"

        # Parameters to be used by functions
        self.excel_df = None
        self.database_df = None
        self.database_dict = None
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def get_legacy_database_from_ifo(self, wb_path=None):

        excel_df = self.excel_to_dataframe(wb_path=wb_path, sheet_name='Database')
        self.convert_datetime_to_str(excel_df)
        self.dataframe_to_dict()
        self.save_database_json()

    def excel_to_dataframe(self, wb_path=None, sheet_name=None):
        # Extracts data from excel tables of a sheet and converts it into a dataframe

        # Check if a sheet and workbook path has been chosen. If not, select the temporary sheet to extract data
        if sheet_name is None:
            sheet_name = self.temporary_sheet_name
        if wb_path is None:
            wb_path = self.wb_path

        # Extract data into dataframe using pandas
        self.excel_df = pd.read_excel(wb_path, sheet_name=sheet_name, engine='openpyxl')

        return self.excel_df

    def save_database_json(self, dictionary=None):
        # Saves the dictionary containing the database into a json file
        # The database file already in the data folder will not be overwritten,
        # but moved to the backup folder

        # if input is None, it selects the database_dict parameter
        if dictionary is None:
            dictionary = self.database_dict

        # Saves the dictionary into a json file format
        if dictionary is not None:
            with open(self.database_path, 'w', encoding='utf8') as file:
                json.dump(dictionary, file, indent=4, sort_keys=False, default=str, ensure_ascii=False)

    def load_database_json(self, database_path=None):
        # Loads the json file containing the database into a usable dictionary
        # If there is no database available, it will use the most recent backup database file

        # Loads class database path if no input is provided
        if database_path is None:
            database_path = self.database_path

        # Extracts the data from the database path
        with open(database_path, 'r', encoding='utf8') as file:
            self.database_dict = json.load(file)

        return self.database_dict

    def dataframe_to_dict(self, df=None):
        # Converts a standard database dataframe into a dictionary

        # Checks if there is a dataframe as input. If not, it uses the database_df parameter
        if df is None:
            df = self.database_df

        # Converts dataframe to dictionary
        self.database_dict = df.to_dict(orient='index')

        return self.database_dict

    def dict_to_dataframe(self, dictionary=None):
        # Converts a standard database dictionary into a dataframe

        # if input is None, it selects the database_dict parameter
        if dictionary is None:
            dictionary = self.database_dict

        # Converts the dictionary to a dataframe
        self.database_df = pd.DataFrame(dictionary).transpose()

        return self.database_df

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

    def convert_datetime_to_str(self, df):
        # Converts a dataframe column that contains dates in datetime format into str, for easier usability

        # Checks if there is a dataframe as input. If not, it uses the database_df parameter
        if df is None:
            df = self.database_df

        # Checks the Date column of the dataframe and converts it
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
        self.database_df = df

        return self.database_df


def tester():
    test = Database()
    test.get_legacy_database_from_ifo(r"D:\Google Drive\Financial organizer\International Financial Organizer.xlsm")


if __name__ == '__main__':
    tester()
