"""
CURRENT VERSION: 0.9*************************************************************************

Program is being created for the sake of automizing a workflow of a regional flower logistics
company. For the past years, everything has been done manually using Word, Excel and a local
accounting program (1C), which is not up-to-the-date. 

The main task is to get data from various resources (partner's logistics program, 
web services, pdf files, xls files, emails), clean and normalize it, and upload it to our SQLite 
database for further use in analytics or billing customers.

For privacy reasons, I cannot reveal any information about logins, passwords, and
leave it empty, though it is required to be filled for the proper functioning of the code.

Libraries that were used to create an early version of the program include:

- Standard libraries (os, sys, io, datetime, calendar, time, fitz, requests, lxml)
- Pandas
- Sqlalchemy
- Pyautogui
- Pywinauto
- Selenium
- Webdriver

COMMAND TO CHECK THE CONTROL IDENTIFIERS: 
app.window(title = window_title).print_control_identifiers()

PRINT CONTROL IDS:
app.window(title = 'AutoCargo - ООО "Флауэрс Лайф" - ЕВГЕНИЯ').print_control_identifiers()
    
"""

import os
from datetime import datetime
import calendar
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

import autocargo_functions
import report_prep
import database_classes
import web_scrap_functions

def main_loop (time_list, operation_codes, balance_day, truck_day):

    """
    This is the main loop that checks the current time and compares it with a time passed as an argument
    to a 'time_list' where 0 index represents hours, and the 1 index represents minutes. Once it matches,
    the loop starts according to the 'operation_codes' list and date options passed as a coefficient, which can be
    a number from zero to a negative number of days from today until the beginning of the 
    previous month.

    Description for operation codes:

    - 1 OPERATION CODE. Starts the Autocargo program and logs in;

    - 2 OPERATION CODE. Goes to the Balance menu, checks the preffered day for 
    the option (calculated using today's date + date coefficient stored in 
    balance_day), and saves the balance file for each of the 12 balance 
    types in the XLSX format;

    - 3 OPERATION CODE. Goes to the Truck menu, checks the preffered day for the
    option (calculated using today's date + truck coefficient stored in
    truck_day), and saves the truck file in the XLSX format.;

    - 4 OPERATION CODE. Closes the Autocargo program;

    - 5 OPERATION CODE. Cleans the data from the previously saved
    balance files. Imports newly created balance files into the database;

    - 5.1 OPERATION CODE. Creates sale files based on the balance files from operation code 5.
    Imports newly created sale files into the database;

    - 6 OPERATION CODE. Cleans the data from the previously saved 
    truck files and save them. Imports newly created truck files into the database;

    - 7 OPERATION CODE. Downloads the history and saldo files from the IpHandlers web service.
    Cleans the data and saves it. Imports newly created iphandlers files into the database;

    - 7.1 OPERATION CODE. Creates sale files based on the iphandlers files from operation code 7.
    Imports newly created sale files into the database;

    """

    # Conditions to check if argument format is correct
    correct_codes = [1, 2, 3, 4, 5, 5.1, 6, 7, 7.1]

    if (not isinstance(time_list[0], int) or
        not isinstance(time_list[1], int)):

        print('Time number is not an integer')

        exit()

    elif (time_list[0] < 0 or
          time_list[0] > 23 or
          time_list[1] < 0 or
          time_list[1] > 59):

        print('Time number is out of bounds')

        exit()

    for code in operation_codes:

        if code not in correct_codes:

            print('Operation code format is not correct')

            exit()

    current_date = datetime.now()

    if current_date.month != 1:

        previous_month = current_date.month - 1

        previous_month_range = calendar.monthrange(current_date.year, previous_month)

    elif current_date.month == 1:

        previous_year = current_date.year - 1

        previous_month = current_date.month - 1

        previous_month_range = calendar.monthrange(previous_year, previous_month)

    for b_day in balance_day:

        if (b_day > 0 or
            b_day < -int((previous_month_range[1] + current_date.day - 1))):

            print('Balance date is out of bounds')

            exit()

    for t_day in truck_day:

        if (t_day > 0 or
            t_day < -int((previous_month_range[1] + current_date.day - 1))):

            print('Truck date is out of bounds')

            exit()

    # Database's path confirmation
    if os.path.exists('C:/sqlite/'):

        directory_name = 'C:/sqlite/'

    else:

        directory_name = 'C:/'

    database_path = f'{directory_name}office_cargo.db'

    # Autocargo's path
    executable_path = "C:/AutoCargo/AutoCargo.exe"

    # Check if AutoCargo path is valid
    if not os.path.exists(executable_path):

        print('AutoCargo program was not found! Related operation codes are removed.')
        print('You can continue working with database and existing files.')

        autocargo_operations = [1, 2, 3, 4]

        for code in operation_codes:

            if code in autocargo_operations:

                operation_codes.remove(code)

        error_react = input('Do you want to exit? yes/no')

        while (error_react != 'yes' and
               error_react != 'no'):

            print('Answer should be yes or no!')
            error_react = input('Do you want to exit? yes/no')

        if error_react == 'yes':

            exit()

    # Autocargo's window title
    window_title = 'AutoCargo - ООО "Флауэрс Лайф" - ЕВГЕНИЯ'

    # Password for the AutoCargo
    password = ""

    # Total count of the balances available to check in the Autocargo
    balances_count = 12

    # Dictionary of the time and operation codes
    time_schedule = {f'{time_list[0]}:{time_list[1]}':operation_codes}

    # Wait for the time matching the schedule
    check_time = True

    while check_time:

        time_now = datetime.now()

        hour = time_now.hour

        minute = time_now.minute

        if f'{hour}:{minute}' in time_schedule:

            check_time = False

    # List of operation's codes for the main loop
    operations_list = time_schedule[f'{hour}:{minute}'][:]

    # Index to add to the current date for balances check
    balance_date_option = balance_day

    # Index to add to the current date for trucks check
    truck_date_option = truck_day

    # Create database if it doesn't exist
    if not os.path.exists(database_path):

        database_classes.database_create(database_path)

    # Create database session
    engine = create_engine(f"sqlite:///{database_path}", echo = False)

    session = sessionmaker(bind=engine)

    session_current = session()

    # Iterate based on the operations_list codes
    for code in operations_list:

        operation_code = code

        # Open Autocargo program
        if operation_code == 1:

            print('Opening AutoCargo program')

            app = autocargo_functions.autocargo_start(executable_path, password)

        # Download balances from Autocargo and start import to the database
        elif operation_code == 2:

            print('Downloading balance reports from AutoCargo program')

            autocargo_functions.autocargo_balance(app, balance_date_option,
                                                    balances_count, window_title)

        # Download truck's list from Autocargo and start import to the database
        elif operation_code == 3:

            print('Downloading truck reports from AutoCargo program')

            autocargo_functions.autocargo_truck(app, truck_date_option, window_title)

            print(truck_date_option)

        # Close AutoCargo
        elif operation_code == 4:

            print('Closing AutoCargo program')

            autocargo_functions.autocargo_close(app, window_title)

        # Clean balance report and import data to the database
        elif operation_code == 5:

            for balance_date in balance_date_option:

                print('Cleaning balance reports')

                # Creates date variables in case if reports are downloaded
                # and only need to be cleaned and imported
                #if 2 not in operations_list:

                balance_dates = autocargo_functions.balance_date_calc(balance_date)

                balance_df_list = report_prep.prepare_balance(balance_dates[0],
                                                              balance_dates[1],
                                                              balance_dates[2])

                if balance_df_list:

                    print('Creating balance records')

                    for b_df in balance_df_list:

                        database_classes.create_table(b_df,
                                                        database_classes.OperationsForever,
                                                        session_current)

                        database_classes.empty_customer_id(session_current)
                        database_classes.empty_shipment_id(session_current)

                else:

                    print("Balance file is empty!")
                    print("Operations Forever tables weren't created")

                # Create sale report and import data to the database
                if 5.1 in operations_list:

                    if balance_df_list:

                        print('Balance list is not empty')

                        for b_df in balance_df_list:

                            sale_forever_df_list = report_prep.create_forever_sales(b_df)

                            if sale_forever_df_list:

                                print('Sale list is not empty')

                                for s_df in sale_forever_df_list:

                                    print('Creating sales records')

                                    database_classes.create_table(s_df,
                                                                  database_classes.OperationsSales,
                                                                  session_current)

                                    database_classes.empty_customer_id(session_current)
                                    database_classes.empty_shipment_id(session_current)

                            else:

                                print("Sale file is empty!")
                                print("Operations Sales tables for Forever weren't created")

                    else:

                        print("Balance file is empty!")
                        print("Operations Sales tables weren't created")

        # Clean truck report and import data to the database
        elif operation_code == 6:

            for truck_date in truck_date_option:

                # Creates date variables in case if reports are downloaded
                # and only need to be cleaned and imported
                #if 3 not in operations_list:

                truck_dates = autocargo_functions.truck_date_calc(truck_date)

                truck_df_list = report_prep.prepare_truck(truck_dates[0],
                                                            truck_dates[1],
                                                            truck_dates[2])

                if truck_df_list:

                    for t_df in truck_df_list:

                        database_classes.create_table(t_df,
                                                        database_classes.Shipments,
                                                        session_current)

                        database_classes.empty_customer_id(session_current)

                else:

                    print("Truck file is empty!")
                    print("Shipments tables weren't created")

        # Create IpHandlers report and import data to the database
        elif operation_code == 7:

            ip_df_list = web_scrap_functions.precool_extract_data(
                                        web_scrap_functions.parse_precooling())

            if ip_df_list:

                for ip_df in ip_df_list:

                    database_classes.create_table(ip_df,
                                                    database_classes.OperationsIphandlers,
                                                    session_current)

                    database_classes.empty_customer_id(session_current)
                    database_classes.empty_shipment_id(session_current)

            else:

                print("Ip file is empty!")
                print("Operations Iphandlers tables weren't created")

            # Create sale report and import data to the database
            if 7.1 in operations_list:

                if ip_df_list:

                    for ip_df in ip_df_list:

                        sale_ip_df_list = web_scrap_functions.create_ip_sales(ip_df)

                        if sale_ip_df_list:

                            for s_i_df in sale_ip_df_list:

                                database_classes.create_table(s_i_df,
                                                              database_classes.OperationsSales,
                                                              session_current)

                                database_classes.empty_customer_id(session_current)
                                database_classes.empty_shipment_id(session_current)

                        else:

                            print("Sale file is empty!")
                            print("Operations Sales tables for Iphandlers weren't created")

                else:

                    print("Ip file is empty!")
                    print("Operations Iphandlers tables weren't created")

# Example of using the main function
main_loop([14, 40],
          [1, 2, 3, 4, 6, 5, 5.2, 7, 7.1],
          [-2, -1],
          [-2, -1])
