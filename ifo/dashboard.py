import pandas as pd
import xlwings as xw
from xlwings import constants as xw_constants
from os.path import join
import pathlib
from datetime import datetime

import database


class Dashboard:

    def __init__(self):
        # Main data validation cell names in list format
        self.validation_type_list = ["CurrencyValidation",
                                     "YearValidation",
                                     "MonthValidation",
                                     "MainBankValidation",
                                     "SavingBankValidation"]

        # Dataframe used for extracting data
        self.df = None

        # validation list parameter
        self.validation_list = None

        # Excel file path
        self.wb_path = join(pathlib.Path(__file__).parent.absolute(), "IFO.xlsm")

        # xlwings parameters
        self.wb = xw.Book(self.wb_path)
        self.ws = self.wb.sheets["Dashboard"].api

        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def fill_in_data_validation_cells(self):
        # Takes the specific variable list and updates the data validation list of the specific named range cell
        pass

    def get_database_dataframe(self):
        # Gets the dataframe by using the Database module functions
        with database.Database() as data:
            self.df = data.get_current_database_dataframe()

        return self.df

    def get_data_validation_list(self, validation_type, df=None):
        # Gets a list without duplicates of specific variables required to fill in the data validation cells.
        # One to three filters can be applied to get the correct list

        # Checks first if there is dataframe input to be used if not, get dataframe
        if df is None:
            df = self.get_database_dataframe()

        # Get column of dataframe and change it into a list
        filtered_validation_list = list()
        if validation_type == "YearValidation":
            validation_list = df['Date'].tolist()

            for value in list(set(validation_list)):
                year = datetime.strptime(value, "%Y=%m-%d").year
                filtered_validation_list.append(year)

        elif validation_type == "MonthValidation":
            validation_list = df['Date'].tolist()

            for value in list(set(validation_list)):
                month = datetime.strptime(value, "%Y=%m-%d").month
                filtered_validation_list.append(month)

        elif validation_type == "MainBankValidation":
            validation_list = df['Input Bank'].tolist() + df['Output Bank'].tolist()

            for value in list(set(validation_list)):
                if not value.contains("saving"):
                    filtered_validation_list.append(value)

        elif validation_type == "SavingBankValidation":
            validation_list = df['Input Bank'].tolist() + df['Output Bank'].tolist()

            for value in list(set(validation_list)):
                if value.contains("saving"):
                    filtered_validation_list.append(value)

        else:
            # The last option would be currency
            filtered_validation_list = list(set(df["Currency"].tolist()))

        # Sort the list
        filtered_validation_list.sort()
        self.validation_list = filtered_validation_list

        # Return the list
        return self.validation_list

    def data_validation_update(self, validation_type, validation_list=None):
        # Updates all data based on general filters, like currency, year or month

        # If there is no validation list as input, it gets a list
        if validation_list is None:
            validation_list = self.get_data_validation_list(validation_type)

        # Obtain the constants required to modify the data validation cells
        dv_type = xw_constants.DVType.xlValidateList
        dv_alert_style = xw_constants.DVAlertStyle.xlValidAlertStop
        dv_operator = xw_constants.FormatConditionOperator.xlEqual

        # Get validation named cell range
        validation_range = self.ws.Range(validation_type)

        # Get the current value on display of the data validation cell, before modifying it
        current_display_value = validation_range.Value

        # Delete current validation present in cell
        validation_range.Delete()

        # Add new validation list to validation cell
        validation_range.Add(dv_type, dv_alert_style, dv_operator, ",".join(validation_list))

        # Set the original value back as current selection
        validation_range.Value = current_display_value

    def get_all_current_data_validation_selections(self):
        # Gets all current data validation selections of the Dashboard and returns it as a dictionary

        # Creates the dictionary and extracts the values from the Dashboard
        current_validation_values_dict = dict()
        for validation_type in self.validation_type_list:
            current_validation_values_dict[validation_type] = self.ws.Range(validation_type).Value

        return current_validation_values_dict

    def update_last_transaction_entry(self, df):
        # Searches the database for the last transaction made in het specific currency

        # if database input parameter is none, get database dataframe
        if df is None:
            df = self.get_database_dataframe()

        # Get currency selection
        validation_dict = self.get_all_current_data_validation_selections()
        currency = validation_dict["CurrencyValidation"]

        # Apply filter to database to obtain all dates related to the currency
        df = df.loc[df["Currency"] == currency]

        # Search for last transaction date
        df['Date'] = pd.to_datetime(df['Date'])
        last_date = df['Date'].max()

        # Update the last entry date in the dashboard
        self.ws.Range("LastTransactionEntry").Value = last_date.strftime('%d-%m-%Y')


def tester():
    test = Dashboard()
