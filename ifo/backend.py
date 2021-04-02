import pandas as pd
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

    def monthly_spending_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Create a dictionary of filters required to make the calculation of this month spending
        filter_dict = dict()
        filter_dict['Currency'] = self.ws.Range("CurrencyValidation").Value
        filter_dict['Start Date'] = first_day_this_month()
        filter_dict['End Date'] = last_day_this_month()
        filter_dict['Type'] = 'spending'

        # Get dataframe for this calculation
        with database.Database() as data:
            filtered_df = data.filter_data_from_dataframe(filter_dict, df=unfiltered_df)

        # Calculate the spending value of this month
        spending_this_month = filtered_df['Output Value'].sum()

        # Change filter dict to get data from last month
        filter_dict['Start Date'] = first_day_last_month()
        filter_dict['End Date'] = last_day_last_month()

        # Get dataframe for last month
        with database.Database() as data:
            filtered_df = data.filter_data_from_dataframe(filter_dict, df=unfiltered_df)

        # Calculate the spending from last month
        spending_last_month = filtered_df['Output Value'].sum()

        # Fill in the backend sheet with calculations
        self.ws.Range('ThisMonthSpend').Value = spending_this_month
        self.ws.Range('LastMonthSpend').Value = spending_last_month

    def monthly_earning_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def monthly_balance_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

    def monthly_savings_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        pass

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
