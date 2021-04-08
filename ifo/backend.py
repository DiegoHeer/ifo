from os.path import join
import pathlib
import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar

import database


def first_day_this_month():
    return datetime.today().replace(day=1).date()


def last_day_this_month():
    today = datetime.today().date()
    last_day_month = calendar.monthrange(today.year, today.month)[1]
    return datetime(year=today.year, month=today.month, day=last_day_month).date()


def first_day_last_month():
    return first_day_this_month() - relativedelta(months=1)


def last_day_last_month():
    return first_day_this_month() - relativedelta(days=1)


def get_unfiltered_database(df):
    # Function gets the original database if none is provided in the parent method/function
    if df is None:
        with database.Database() as data:
            df = data.get_current_database_dataframe()

    return df


def filter_dataframe(unfiltered_df, filter_dict):
    # Get dataframe for this calculation
    with database.Database() as data:
        filtered_df = data.filter_data_from_dataframe(filter_dict, df=unfiltered_df)

    return filtered_df


class Backend:

    def __init__(self):
        # Excel file path
        self.wb_path = join(pathlib.Path(__file__).parent.absolute(), "IFO.xlsm")

        # xlwings parameters
        self.wb = xw.Book(self.wb_path)
        self.ws = self.wb.sheets["Backend"].api

        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def create_filter_dict(self, start_date, end_date, transaction_type=None, category=None, input_account=None,
                           output_account=None, input_account_type=None, output_account_type=None):
        # Creates a dictionary that serves a filter to gather a defined dataframe to extract information
        filter_dict = dict()
        filter_dict['Currency'] = self.ws.Range("CurrencyValidation").Value
        filter_dict['Start Date'] = start_date
        filter_dict['End Date'] = end_date

        if transaction_type is not None:
            filter_dict['Type'] = transaction_type
        if category is not None:
            filter_dict['Category'] = category
        if input_account is not None:
            filter_dict['Input Account'] = input_account
        if output_account is not None:
            filter_dict['Output Account'] = output_account
        if input_account_type is not None:
            filter_dict['Input Account Type'] = input_account_type
        if output_account_type is not None:
            filter_dict['Output Account Type'] = output_account_type

        return filter_dict

    def get_sum_value_filtered_df(self, unfiltered_df, sum_column, start_date, end_date, transaction_type=None,
                                  category=None, input_account=None, output_account=None, input_account_type=None,
                                  output_account_type=None):
        # This function provides the sum value of a column of a filtered dataframe

        # Create a dictionary of filters required
        filter_dict = self.create_filter_dict(start_date, end_date, transaction_type, category, input_account,
                                              output_account, input_account_type, output_account_type)

        # Filters the dataframe with the defined filters
        filtered_df = filter_dataframe(unfiltered_df, filter_dict)

        # Sums the entire column and returns the value
        return filtered_df[sum_column].sum()

    def get_account_balance(self, unfiltered_df, month_selection, account):
        if month_selection == "this month":
            start_date = first_day_this_month()
            end_date = last_day_this_month()
        else:
            start_date = first_day_last_month()
            end_date = last_day_last_month()

        input_value = self.get_sum_value_filtered_df(unfiltered_df, "Input Value", start_date=start_date,
                                                     end_date=end_date, input_account=account)

        output_value = self.get_sum_value_filtered_df(unfiltered_df, "Output Value", start_date=start_date,
                                                      end_date=end_date, output_account=account)

        return input_value - output_value

    def get_total_balance(self, unfiltered_df, month_selection, account_type):
        if month_selection == "this month":
            start_date = first_day_this_month()
            end_date = last_day_this_month()
        else:
            start_date = first_day_last_month()
            end_date = last_day_last_month()

        input_value = self.get_sum_value_filtered_df(unfiltered_df, "Input Value", start_date=start_date,
                                                     end_date=end_date, input_account_type=account_type)

        output_value = self.get_sum_value_filtered_df(unfiltered_df, "Output Value", start_date=start_date,
                                                      end_date=end_date, output_account_type=account_type)

        return input_value - output_value

    def monthly_spending_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Calculate the spending value of this month
        spending_this_month = self.get_sum_value_filtered_df(unfiltered_df, "Output Value",
                                                             start_date=first_day_this_month(),
                                                             end_date=last_day_this_month(),
                                                             transaction_type='spending')

        # Calculate the spending from last month
        spending_last_month = self.get_sum_value_filtered_df(unfiltered_df, "Output Value",
                                                             start_date=first_day_last_month(),
                                                             end_date=last_day_last_month(),
                                                             transaction_type='spending')

        # Fill in the backend sheet with calculations
        self.ws.Range('ThisMonthSpend').Value = spending_this_month
        self.ws.Range('LastMonthSpend').Value = spending_last_month

    def monthly_earning_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Calculate the spending value of this month
        earning_this_month = self.get_sum_value_filtered_df(unfiltered_df, "Input Value",
                                                            start_date=first_day_this_month(),
                                                            end_date=last_day_this_month(),
                                                            transaction_type='earning')

        # Calculate the spending from last month
        earning_last_month = self.get_sum_value_filtered_df(unfiltered_df, "Input Value",
                                                            start_date=first_day_last_month(),
                                                            end_date=last_day_last_month(),
                                                            transaction_type='earning')

        # Fill in the backend sheet with calculations
        self.ws.Range('ThisMonthEarned').Value = earning_this_month
        self.ws.Range('LastMonthEarned').Value = earning_last_month

    def monthly_balance_and_saving_block(self, unfiltered_df=None, saving_bool=False):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Defines parameters based on the type of block that will be updated: checking or saving
        if saving_bool is True:
            id_parameter = "Saving"
            account_type = 'saving accounts'
        else:
            id_parameter = "Balance"
            account_type = 'checking accounts'

        # First calculation is the total balance of the month, excluded saving accounts
        # Get the total balance for this month
        balance_value_this_month = self.get_total_balance(unfiltered_df, "this month", account_type=account_type)

        # Get the total balance for last month
        balance_value_last_month = self.get_total_balance(unfiltered_df, "last month", account_type=account_type)

        # Fill in the backend sheet with calculations for total balance, except savings
        self.ws.Range('ThisMonthTotal' + id_parameter).Value = balance_value_this_month
        self.ws.Range('LastMonthTotal' + id_parameter).Value = balance_value_last_month

        # Find out the most used bank account in the last 6 months
        # Define first the start and end date for the filter
        end_date = datetime.today().date()
        start_date = end_date - relativedelta(months=6)

        # Get a filtered dataframe from this 6 months (in the right currency)
        filter_dict = self.create_filter_dict(start_date=start_date, end_date=end_date)
        filtered_df = filter_dataframe(unfiltered_df, filter_dict)

        # Calculate the frequency of all the accounts
        account_frequency_dict = filtered_df['Output Account'].value_counts().to_dict()

        # Get the most used account
        most_used_account = max(account_frequency_dict, key=account_frequency_dict.get)

        # Now get the balance for this account
        balance_most_used_account_this_month = self.get_account_balance(unfiltered_df, "this month",
                                                                        account=most_used_account)
        balance_most_used_account_last_month = self.get_account_balance(unfiltered_df, "last month",
                                                                        account=most_used_account)

        # Fill in the backend sheet with the balances of the most used account
        self.ws.Range(f'MostUsed{id_parameter}Account').Value = most_used_account
        self.ws.Range(f'ThisMonth{id_parameter}1').Value = balance_most_used_account_this_month
        self.ws.Range(f'LastMonth{id_parameter}1').Value = balance_most_used_account_last_month

        # Fill in for the selected accounts
        for i in range(1, 3):
            selected_account = self.ws.Range(f"Selected{id_parameter}Account{i}").Value

            balance_selected_account_this_month = self.get_account_balance(unfiltered_df, "this month",
                                                                           account=selected_account)
            balance_selected_account_last_month = self.get_account_balance(unfiltered_df, "last month",
                                                                           account=selected_account)

            self.ws.Range(f'ThisMonth{id_parameter}{i + 1}').Value = balance_selected_account_this_month
            self.ws.Range(f'LastMonth{id_parameter}{i + 1}').Value = balance_selected_account_last_month

    def week_quarter_spending_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def month_quarter_investments_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def recent_transactions_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def average_day_spending_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def spending_per_category_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def investment_portfolio_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def spending_per_type_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass


def tester():
    # test = Backend()
    print(first_day_this_month())
    print(last_day_this_month())
    print(first_day_last_month())
    print(last_day_last_month())


if __name__ == '__main__':
    tester()
