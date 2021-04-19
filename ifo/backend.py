from os.path import join
import pathlib
import xlwings as xw
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import calendar

import database


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

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def get_validation_end_date(self):
        # Extracts the date selected through cell validation in the Dashboard sheet
        ws = self.wb.sheets["Dashboard"].api

        month_validation_name = ws.Range("MonthValidation").Value
        month_validation_num = datetime.strptime(month_validation_name, "%B").month
        year_validation = int(ws.Range("YearValidation").Value)

        last_day_month = calendar.monthrange(year_validation, month_validation_num)[1]

        validation_end_date = datetime(year=year_validation, month=month_validation_num, day=last_day_month).date()
        return validation_end_date

    def get_validation_start_date(self):
        validation_end_date = self.get_validation_end_date()
        return validation_end_date.replace(day=1)

    def get_validation_last_month_start_date(self):
        validation_start_date = self.get_validation_start_date()
        return validation_start_date - relativedelta(months=1)

    def get_validation_last_month_end_date(self):
        validation_start_date = self.get_validation_start_date()
        return validation_start_date - relativedelta(days=1)

    def create_filter_dict(self, start_date, end_date, transaction_type=None, category=None, input_account=None,
                           output_account=None, input_account_type=None, output_account_type=None):
        # Creates a dictionary that serves a filter to gather a defined dataframe to extract information
        filter_dict = dict()
        filter_dict['Currency'] = self.wb.sheets["Dashboard"].api.Range("CurrencyValidation").Value
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
            start_date = self.get_validation_start_date()
            end_date = self.get_validation_end_date()
        else:
            start_date = self.get_validation_last_month_start_date()
            end_date = self.get_validation_last_month_end_date()

        input_value = self.get_sum_value_filtered_df(unfiltered_df, "Input Value", start_date=start_date,
                                                     end_date=end_date, input_account=account)

        output_value = self.get_sum_value_filtered_df(unfiltered_df, "Output Value", start_date=start_date,
                                                      end_date=end_date, output_account=account)

        return input_value - output_value

    def get_total_balance(self, unfiltered_df, month_selection, account_type):
        if month_selection == "this month":
            start_date = self.get_validation_start_date()
            end_date = self.get_validation_end_date()
        else:
            start_date = self.get_validation_last_month_start_date()
            end_date = self.get_validation_last_month_end_date()

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
                                                             start_date=self.get_validation_start_date(),
                                                             end_date=self.get_validation_end_date(),
                                                             transaction_type='spending')

        # Calculate the spending from last month
        spending_last_month = self.get_sum_value_filtered_df(unfiltered_df, "Output Value",
                                                             start_date=self.get_validation_last_month_start_date(),
                                                             end_date=self.get_validation_last_month_end_date(),
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
                                                            start_date=self.get_validation_start_date(),
                                                            end_date=self.get_validation_end_date(),
                                                            transaction_type='earning')

        # Calculate the spending from last month
        earning_last_month = self.get_sum_value_filtered_df(unfiltered_df, "Input Value",
                                                            start_date=self.get_validation_last_month_start_date(),
                                                            end_date=self.get_validation_last_month_end_date(),
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

    def week_quarter_spending_and_investment_block(self, unfiltered_df=None, investment_bool=False):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Defines parameters based on the type of block that will be updated: checking or saving
        if investment_bool is True:
            focus_parameter = "MonthInvestment"
            transaction_type = "investment"
            sum_column = "Input Value"
        else:
            focus_parameter = "WeekSpending"
            transaction_type = "spending"
            sum_column = "Output Value"

        # Calculate current week spending or month investment
        today = datetime.today().date()
        start_date = today - timedelta(days=today.weekday())
        end_date = start_date + timedelta(days=6)

        period_value = self.get_sum_value_filtered_df(unfiltered_df, sum_column=sum_column,
                                                      start_date=start_date,
                                                      end_date=end_date,
                                                      transaction_type=transaction_type)

        self.ws.Range(focus_parameter).Value = period_value

        # Calculate and fill in quarter spending or investment this year
        for i in range(1, 5):
            start_date = datetime(year=datetime.today().year, month=(3 * i - 2), day=1).date()
            end_date = start_date + relativedelta(months=3) - timedelta(days=1)

            quarter_value = self.get_sum_value_filtered_df(unfiltered_df, sum_column=sum_column,
                                                           start_date=start_date,
                                                           end_date=end_date,
                                                           transaction_type=transaction_type)

            self.ws.Range(f"Quarter{i}{transaction_type.capitalize()}").Value = quarter_value

    def recent_transactions_block(self, unfiltered_df=None):
        # Updates the values in the cells related to the specific function named topic

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Apply a filter to obtain the dataframe from the last 365 days in current currency
        today = datetime.today().date()
        start_date = today - relativedelta(years=1)
        end_date = today

        filter_dict = self.create_filter_dict(start_date=start_date, end_date=end_date)
        filtered_df = filter_dataframe(unfiltered_df, filter_dict)

        # Get the last 10 rows from the dataframe
        df_tail = filtered_df.tail(10)

        # Fill in the values of the Recent Transactions Table
        for i in range(1, 11):
            self.ws.Range(f"RecentDate{i}").Value = df_tail['Date'].iloc[i]
            self.ws.Range(f"RecentType{i}").Value = df_tail['Type'].iloc[i]
            self.ws.Range(f"RecentCategory{i}").Value = df_tail['Category'].iloc[i]
            self.ws.Range(f"RecentCurrency{i}").Value = df_tail['Currency'].iloc[i]
            self.ws.Range(f"RecentInputValue{i}").Value = df_tail['Input Value'].iloc[i]
            self.ws.Range(f"RecentInputAccount{i}").Value = df_tail['Input Account'].iloc[i]
            self.ws.Range(f"RecentOutputValue{i}").Value = df_tail['Output Value'].iloc[i]
            self.ws.Range(f"RecentOutputAccount{i}").Value = df_tail['Output Account'].iloc[i]
            self.ws.Range(f"RecentDescription{i}").Value = df_tail['Description'].iloc[i]

    def average_day_spending_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Average spending
        # Get the value of this month spending
        this_month_spending = self.ws.Range("ThisMonthSpend").Value

        # Get the maximum amount of days in this month
        maximum_days_in_month = self.get_validation_end_date().day

        # Return the result to the backend sheet
        self.ws.Range("AverageSpending").Value = round(this_month_spending / maximum_days_in_month, 2)

        # Maximal spending
        # Create a filter dictionary for the spending this month
        start_date = self.get_validation_start_date()
        end_date = self.get_validation_end_date()
        filter_dict = self.create_filter_dict(start_date, end_date, transaction_type="spending")
        this_month_spending_df = filter_dataframe(unfiltered_df, filter_dict)

        # Get the maximum spending value
        maximum_spending = this_month_spending_df["Output Value"].max()
        self.ws.Range("MaximalSpending").Value = maximum_spending

        # Get the minimal spending
        minimal_spending = this_month_spending_df['Output Value'].min()
        self.ws.Range("MinimalSpending").Value = minimal_spending

        # Get today's spending
        today = datetime.today().date()
        today_spending_df = this_month_spending_df.loc[this_month_spending_df["Date"] == today]
        today_spending = today_spending_df["Output Value"].sum()
        self.ws.Range("TodaySpending").Value = today_spending

        # Last month average spending
        last_month_spending = self.ws.Range("LastMonthSpend").Value
        maximum_days_last_month = self.get_validation_last_month_end_date().day
        self.ws.Range("LastMonthAvSpending").Value = last_month_spending / maximum_days_last_month

    def spending_per_category_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Get the list of categories available
        category_tuple = self.ws.Range("ListedCategories").Value
        category_list = [item for t in category_tuple for item in t]

        # The calculation is performed for every listed category for all 12 months
        # Obtain first a filtered dataframe for spending based on the complete year of display
        end_year_number = int(self.ws.Range("EndYearNumber").Value)
        end_month_number = int(self.ws.Range("EndMonthNumber").Value)

        last_day = calendar.monthrange(end_year_number, end_month_number)[1]
        year_end_date = datetime(year=end_year_number, month=end_month_number, day=last_day).date()
        year_start_date = year_end_date + relativedelta(days=1) - relativedelta(years=1)

        filter_dict = self.create_filter_dict(year_start_date, year_end_date, transaction_type="spending")
        year_spending_df = filter_dataframe(unfiltered_df, filter_dict)

        for category in category_list:
            for month_num in range(0, 12):
                month_start_date = year_start_date + relativedelta(months=month_num)
                month_end_date = month_start_date + relativedelta(months=1) - relativedelta(days=1)

                category_spending_sum = self.get_sum_value_filtered_df(year_spending_df, sum_column="Output Value",
                                                                       start_date=month_start_date,
                                                                       end_date=month_end_date, category=category)

                named_range = f"Spending{category.capitalize()}MonthNum{month_num + 1}"
                self.ws.Range(named_range).Value = category_spending_sum

    def investment_portfolio_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Calculate the total value of investments made for stocks and bonds until month and year validation
        start_date = unfiltered_df['Date'].min()
        end_date = self.get_validation_end_date()

        # General function for obtaining the total invested value for bonds and stocks
        def total_invested_value(ws, dataframe, start_date, end_date, category, output_range_name):
            output_value = self.get_sum_value_filtered_df(dataframe, "Output Value", start_date, end_date,
                                                          transaction_type="investment", category=category)
            input_value = self.get_sum_value_filtered_df(dataframe, "Output Value", start_date, end_date,
                                                         transaction_type="investment", category=category)

            ws.Range(output_range_name).Value = output_value - input_value

        # Bonds total invested value
        total_invested_value(self.ws, unfiltered_df, start_date, end_date, "bonds", "TotalInvestedBonds")

        # Stocks total invested value
        total_invested_value(self.ws, unfiltered_df, start_date, end_date, "stocks", "TotalInvestedStocks")

        # Obtain the total investments for one month before the validation end date
        end_date = end_date - relativedelta(months=1)
        total_invested_value(self.ws, unfiltered_df, start_date, end_date, "bonds", "TotalInvestedBonds")
        total_invested_value(self.ws, unfiltered_df, start_date, end_date, "stocks", "TotalInvestedStocks")

    def spending_per_type_chart(self, unfiltered_df=None):
        # Updates the values of this topic, which updates the related chart displayed in the Dashboard

        # Check if database dataframe is provided. If not, gets it
        unfiltered_df = get_unfiltered_database(unfiltered_df)

        # Get a list of the possible transactions
        transaction_list = ["spending", "earning", "change", "investment"]

        # The calculation is performed for every listed category for all 12 months
        # Obtain first a filtered dataframe for spending based on the complete year of display
        end_year_number = int(self.ws.Range("EndYearNumber").Value)
        end_month_number = int(self.ws.Range("EndMonthNumber").Value)

        last_day = calendar.monthrange(end_year_number, end_month_number)[1]
        year_end_date = datetime(year=end_year_number, month=end_month_number, day=last_day).date()
        year_start_date = year_end_date + relativedelta(days=1) - relativedelta(years=1)

        for transaction in transaction_list:
            for month_num in range(0, 12):
                month_start_date = year_start_date + relativedelta(months=month_num)
                month_end_date = month_start_date + relativedelta(months=1) - relativedelta(days=1)

                if transaction == "earning":
                    sum_column = "Input Value"
                else:
                    sum_column = "Output Value"

                transaction_sum = self.get_sum_value_filtered_df(unfiltered_df, sum_column=sum_column,
                                                                 start_date=month_start_date,
                                                                 end_date=month_end_date, transaction_type=transaction)

                named_range = f"{transaction.capitalize()}MonthNum{month_num + 1}"
                self.ws.Range(named_range).Value = transaction_sum


def tester():
    test = Backend()
    test.monthly_balance_and_saving_block()


if __name__ == '__main__':
    tester()
