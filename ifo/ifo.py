import pymsgbox

import backend
import dashboard
import database


def update_ifo():
    # TODO
    # Updates all data of backend based on database
    pass


def currency_update():
    # TODO
    # Updates all data based on the currency selected in the Transaction Block of the Dashboard
    pass


def entry():
    # TODO
    # Provide a user form for transaction entry and selectively updates backend
    pass


def manual_update():
    # TODO
    # Provide a user form to filter out data, exporting the filtered data to a table in a new sheet,
    # which than can be manually updated by the user. After the update is complete the user can refresh
    # the database with the updated data using the refresh database button
    pass


def manual_remove():
    # TODO
    # Provides a user form to filter out data, exporting the filtered data to a table in a new sheet.
    # The user can than delete lines of data which are not required anymore.
    # The database can than be updated by the user using the refresh database button
    pass


def refresh_database():
    # TODO
    # Works in conjunction with the manual_update or manual_remove button.
    # After the user updated the data from the sheet,
    # the table is exported to a dataframe which is then used to update the database
    pass


def tester():
    # Temporary function to test main functions of project
    # dashboard.tester()
    # backend.tester()
    # database.tester()
    pass
