import pandas as pd
import os
from os.path import join
from datetime import datetime
import pathlib
import json
import shutil
import xlwings as xw


class Database:

    def __init__(self):
        # General paths and sheet definitions
        self.wb_path = join(pathlib.Path(__file__).parent.absolute(), "IFO.xlsm")
        self.database_dir = join(pathlib.Path(__file__).parent.absolute(), 'data')
        self.database_path = join(self.database_dir, 'database.json')
        self.backup_dir = join(pathlib.Path(__file__).parent.absolute(), 'data', 'database backup')
        self.database_sheet_name = "Database"
        self.temporary_sheet_name = "Filtered Data"

        # Parameters to be used by functions
        self.excel_df = None
        self.database_df = None
        self.filtered_df = None
        self.database_dict = None

        # xlwings parameters
        self.wb = xw.Book(self.wb_path)
        self.ws = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def get_legacy_database_from_ifo(self, wb_path=None):

        excel_df = self.excel_to_dataframe(wb_path=wb_path, sheet_name='Database')
        self.convert_datetime_to_str(excel_df)
        self.dataframe_to_dict()
        self.save_database_json()

    def get_current_database_dataframe(self):

        self.load_database_json()
        self.dict_to_dataframe()

        return self.database_df

    def excel_to_dataframe(self, wb_path=None, sheet_name=None):
        # Extracts data from excel tables of a sheet and converts it into a dataframe

        # Check if a sheet and workbook path has been chosen. If not, select the temporary sheet to extract data
        if sheet_name is None:
            sheet_name = self.temporary_sheet_name
        if wb_path is None:
            wb_path = self.wb_path

        # Save workbook before extracting it
        self.wb.save(self.wb_path)

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

    def remove_transaction_from_dataframe(self, index_list, df=None):
        # Removes the rows from dataframe containing the transaction based on index

        # TODO: test this function

        # Checks if there is a dataframe as input. If not, it uses the database_df parameter
        if df is None:
            df = self.database_df

        # Deleting the rows with a specific indexes
        for index in index_list:
            df = df.drop(df.iloc[index])

        self.database_df = df
        return self.database_df

    def new_transaction_to_dataframe(self, new_trn_dict, df=None):
        # Enters a new row in the dataframe, containing the new transaction

        # TODO: test this function

        # Checks if there is a dataframe as input. If not, it uses the database_df parameter
        if df is None:
            df = self.database_df

        # Converts dataframe into a dictionary
        self.dataframe_to_dict(df)

        # Converts dictionary into a list with dictionaries
        db_list = list()
        for key, value in self.database_dict.items():
            db_list.append(value)

        # Adds new transaction to list
        db_list.append(new_trn_dict)

        # Converts list into dataframe
        self.database_df = pd.DataFrame(db_list)

        return self.database_df

    def update_transactions_in_dataframe(self, filtered_df=None):
        # Compares all indexes of updated filtered dataframe

        # If input parameter is None, use standard filtered dataframe
        if filtered_df is None:
            filtered_df = self.filtered_df

        # get current dataframe from database file
        self.get_current_database_dataframe()

        # Iterate through each row of the filtered data to substitute it in the database
        for index, row in filtered_df.iterrows():
            self.database_df.iloc[index] = row

        return self.database_df

    def filter_data_from_dataframe(self, filter_dict, df=None):
        # Returns a filtered dataframe based on the filters applied to the original database

        # Checks if there is a dataframe as input. If not, it uses the database_df parameter
        if df is None:
            df = self.database_df

        # Converts column with dates as string in dataframe to datetime format
        df['Date'] = pd.to_datetime(df['Date'])

        # Filter out per column type
        for key, value in filter_dict.items():
            if key == "Start Date":
                value = datetime.strptime(value, "%Y-%m-%d")
                df = df.loc[df[key] >= value]

            elif key == "End Date":
                value = datetime.strptime(value, "%Y-%m-%d")
                df = df.loc[df[key] <= value]

            elif key == "Minimum Input Value" or key == "Minimum Output Value":
                df = df.loc[df[key] >= value]

            elif key == "Maximum Input Value" or key == "Maximum Output Value":
                df = df.loc[df[key] <= value]

            elif key == "Description":
                df = df.loc[df[key].str.contains(value)]

            else:
                df = df.loc[df[key] == value]

        # Convert index of dataframe into column
        df['Index'] = df.index

        self.filtered_df = df
        return self.filtered_df

    def backup_old_database(self):
        # Moves the current database json file to the backup folder and changes it name with a suffix timestamp

        # Checks if there exists a database file
        if not os.path.exists(self.database_path):
            return

        # Creates a new name for old database file
        backup_file_name = 'database_' + datetime.today().strftime("%Y-%m-%d") + '.json'
        backup_file_path = join(self.backup_dir, backup_file_name)

        # Checks if the file exists. If true, delete it
        if os.path.exists(backup_file_path):
            os.remove(backup_file_path)

        # Copy database file and rename it to the backup folder
        shutil.copy(self.database_path, backup_file_path)

    def restore_old_database(self):
        # Searches for the most recent database json file in the backup folder,
        # moves it to the data folder and changes it name to database.json

        # Check if there exists database files in the backup folder
        if not os.listdir(self.backup_dir):
            return

        # Loop through each file in directory, applying filters to find the correct ones
        date_list = list()
        date_dict = dict()
        for file_path in os.listdir(self.backup_dir):
            filename = os.path.split(file_path)[1]
            filetype = os.path.splitext(file_path)[1]

            if 'database' in filename and filetype == 'json':
                date_str = filename.replace('database_', '')
                date = datetime.strptime(date_str, '%Y-%m-%d')

                # Add parameters to list and dictionary
                date_list.append(date)
                date_dict[date] = file_path

        # Obtain the latest file path
        latest_file_path = date_dict[max(date_list)]

        # If there exists a current database file in data directory, delete it
        if os.path.exists(self.database_path):
            os.remove(self.database_path)

        # replace the latest backup file to the data folder
        shutil.move(latest_file_path, self.database_path)

    def get_filtered_excel_data(self):
        # Extracts the filtered excel data which has been updated by user and converts it into a dataframe

        # check if temporary sheet exists in excel workbook
        sheet_list = [sh.name for sh in self.wb.sheets]
        if self.temporary_sheet_name not in sheet_list:
            return

        # Extracts updated/filtered dataframe from temporary excel sheet
        self.filtered_df = self.excel_to_dataframe(sheet_name=self.temporary_sheet_name)

        # Change index column to main index
        self.filtered_df = self.filtered_df.set_index('Index', inplace=True)

        return self.filtered_df

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
