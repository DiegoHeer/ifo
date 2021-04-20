import calendar

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
                                     "CheckingAccountValidation",
                                     "CheckingAccountValidation2",
                                     "SavingAccountValidation",
                                     "SavingAccountValidation2"]

        # Dataframe used for extracting data
        self.df = None

        # validation list parameter
        self.validation_list = None

        # Excel file path
        self.wb_path = join(pathlib.Path(__file__).parent.absolute(), "IFO.xlsm")

        # xlwings parameters
        self.wb = xw.Book(self.wb_path)
        self.ws = self.wb.sheets["Dashboard"].api
        self.currency_selection = self.ws.Range("CurrencyValidation").Value

        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
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

        # Function for obtaining account validation
        def get_account_validation(dataframe, account_type):
            # Filter first the dataframe based on currency
            dataframe = dataframe.loc[dataframe["Currency"] == self.currency_selection]

            # Get all the accounts from all the input and output transactions
            account_list = dataframe['Input Account'].tolist() + dataframe['Output Account'].tolist()

            # Remove duplicates from list
            account_list = list(set(account_list))

            # Remove blank items from list
            account_list = list(filter(None, account_list))

            filtered_account_list = list()
            for account in account_list:
                if account_type == "saving":
                    if "saving" in account:
                        filtered_account_list.append(account)
                else:
                    if "saving" not in account:
                        filtered_account_list.append(account)

            # Sort the filtered account list
            filtered_account_list.sort()

            return filtered_account_list

        # Get column of dataframe and change it into a list
        filtered_validation_list = list()
        if validation_type == "YearValidation":
            validation_list = df['Date'].tolist()

            # Filter out the years
            for value in validation_list:
                year = datetime.strptime(value, "%Y-%m-%d").year

                # Convert it all to strings
                year = str(year)

                # Add to list
                filtered_validation_list.append(year)

            # Remove the duplicates from the filtered validation list and sort it
            filtered_validation_list = list(set(filtered_validation_list))

            # Sort the validation list
            filtered_validation_list.sort()

            # Add one year extra to the list
            filtered_validation_list.append(str(int(filtered_validation_list[-1]) + 1))

        elif validation_type == "MonthValidation":
            filtered_validation_list = [month for index, month in enumerate(calendar.month_name)][1:]

        elif "CheckingAccountValidation" in validation_type:
            filtered_validation_list = get_account_validation(df, "checking")

        elif "SavingAccountValidation" in validation_type:
            filtered_validation_list = get_account_validation(df, "saving")

        else:
            # The last option would be currency
            filtered_validation_list = list(set(df["Currency"].tolist()))
            filtered_validation_list.sort()

        # Return the list
        self.validation_list = filtered_validation_list
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

        # Current display value is only valid if it is in the validation list
        if current_display_value not in validation_list:
            current_display_value = validation_list[0]

        # Delete current validation present in cell
        validation_range.Validation.Delete()

        # Add new validation list to validation cell
        validation_range.Validation.Add(dv_type, dv_alert_style, dv_operator, ";".join(validation_list))

        # Set the original value back as current selection
        validation_range.Value = current_display_value

    def get_all_current_data_validation_selections(self):
        # Gets all current data validation selections of the Dashboard and returns it as a dictionary

        # Creates the dictionary and extracts the values from the Dashboard
        current_validation_values_dict = dict()
        for validation_type in self.validation_type_list:
            validation_value = self.ws.Range(validation_type).Value

            # For validation values that are float, convert them into integer
            if type(validation_value) is float:
                validation_value = int(validation_value)

            current_validation_values_dict[validation_type] = validation_value

        return current_validation_values_dict

    def update_last_transaction_entry(self, df=None):
        # Searches the database for the last transaction made in het specific currency

        # if database input parameter is none, get database dataframe
        if df is None:
            df = self.get_database_dataframe()

        # Apply filter to database to obtain all dates related to the currency
        df = df.loc[df["Currency"] == self.currency_selection]

        # Search for last transaction date
        df['Date'] = pd.to_datetime(df['Date'])
        last_date = df['Date'].max()

        # Update the last entry date in the dashboard
        self.ws.Range("LastTransactionEntry").Value = last_date.strftime('%d-%m-%Y')

    def fill_in_most_used_account(self, account_type, df=None):

        # Checks first if there is dataframe input to be used if not, get dataframe
        if df is None:
            df = self.get_database_dataframe()

        # Filter first the dataframe based on currency
        df = df.loc[df["Currency"] == self.currency_selection]

        # Get all the accounts from all the input and output transactions
        account_list = df['Input Account'].tolist() + df['Output Account'].tolist()

        # Remove blank items from list
        account_list = list(filter(None, account_list))

        # Filter based on account type
        filtered_account_list = list()
        for account in account_list:
            if account_type == "saving":
                if "saving" in account:
                    filtered_account_list.append(account)
            else:
                if "saving" not in account:
                    filtered_account_list.append(account)

        if len(filtered_account_list) == 0:
            most_used_account = ""
        else:
            # Get the most frequent account in list
            most_used_account = max(set(filtered_account_list), key=filtered_account_list.count)

        # get the correct named range
        if account_type == "saving":
            account_named_range = "MostUsedSavingAccount"
        else:
            account_named_range = "MostUsedCheckingAccount"

        # Fill in the data into the excel dashboard
        self.ws.Range(account_named_range).Value = most_used_account

        return most_used_account


def tester():
    # test = Dashboard()
    # print(test.get_data_validation_list("YearValidation"))
    # print(test.get_all_current_data_validation_selections())

    # test.data_validation_update("CheckingAccountValidation")
    # test.data_validation_update("SavingAccountValidation")
    # print(test.get_all_current_data_validation_selections())

    # test.fill_in_most_used_account("checking")
    # test.fill_in_most_used_account("saving")

    # test.update_last_transaction_entry()
    pass


if __name__ == '__main__':
    tester()
