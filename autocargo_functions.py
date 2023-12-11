"""
This module allows to run and download reports from the AutoCargo program by 
Forever company. Flow of the program is controlled by the operation codes 
preset in the list named standard_operations.

"""

import sys
from io import StringIO
import calendar
import time
from datetime import datetime
import pyautogui

from pywinauto.application import Application

def autocargo_start(path, passw):

    """
    This function starts the Autocargo program and connects to it

    """

    application = Application(backend='uia').start(path).connect(title = 'AutoCargo',
                                                                 timeout = 100)

    start_button = application.AutoCargo.child_window(title="Старт")

    time.sleep(5)

    start_button.click_input()

    pyautogui.press('up')

    time.sleep(5)

    pyautogui.press('enter')

    time.sleep(5)

    start_button.click_input()

    time.sleep(5)

    pyautogui.typewrite(passw)

    time.sleep(5)

    pyautogui.press('enter')

    time.sleep(20)

    pyautogui.press('enter')

    time.sleep(5)

    return application

def autocargo_balance(app, balance_date_option, balances_count, window_title):

    """
    This function downloads Autocargo balance reports

    """

    # Define button's names
    file_type_button = app.window(title = window_title) \
                          .child_window(title="Save as type:",
                                        auto_id="FileTypeControlHost",
                                        control_type="ComboBox")

    save_as_button = app.window(title = window_title) \
                        .child_window(title="Сохранить",
                                      control_type="Button")

    save_file_button = app.window(title = window_title) \
                          .child_window(title="Save",
                                        auto_id="1",
                                        control_type="Button")

    # Full information about today's day, month, year
    today_full_date = datetime.now()
    today_month_num = today_full_date.strftime('%m')
    today_month_full = calendar.month_name[int(today_month_num)]
    today_day = today_full_date.strftime("%d")
    today_year = today_full_date.strftime('%Y')

    for balance_date in balance_date_option:

        balances_dropdown = app.window(title = window_title)['MenuItem7']

        balances_dropdown.select()

        time.sleep(5)

        all_dropdown = app.window(title = window_title)['MenuItem1']

        all_dropdown.select()

        time.sleep(5)

        customers_dropdown = app.window(title = window_title)['MenuItem1']

        customers_dropdown.select()

        time.sleep(5)

        # Repeat the same order of actions for each balance type
        for balance_code in range(balances_count):

            # If it is the first balance type we start with default set up
            if balance_code != 0:

                pyautogui.press('down')

            time.sleep(5)

            # Check if the balance day is correct after adding balance_date to today_day
            if ((int(today_day) + balance_date) > 0 and
                balance_code == 0):

                balance_day = int(today_day) + balance_date

                first_day_weekday = calendar.weekday(
                    int(today_year),
                    int(today_month_num),
                    1)

                balance_date_button_name = f'DataItem{first_day_weekday + 2 + (int(balance_day) - 1)}'

                balance_month_full = today_month_full.lower()

                balance_year = today_year

                balance_date_button = app.window(title =
                                                window_title)[f'{balance_date_button_name}']

                balance_date_button.click_input()

                time.sleep(5)

            elif ((int(today_day) + balance_date) <= 0 and
                  int(today_month_num) != 1 and
                  balance_code == 0):

                previous_month_full = calendar.month_name[int(today_month_num)-1]

                previous_month_num = int(today_month_num) - 1

                balance_date = calendar.monthrange(int(today_year), previous_month_num)[1] + int(balance_date) + int(today_day)

                first_day_weekday = calendar.weekday(int(today_year), previous_month_num, 1)

                balance_date_button_name = f'DataItem{first_day_weekday + 2 + (balance_date - 1)}'

                balance_previous_month_button = app.window(title = window_title) \
                                                .child_window(title="Previous Button",
                                                                control_type="Button")

                balance_previous_month_button.click_input()

                time.sleep(5)

                balance_day = balance_date

                balance_month_full = previous_month_full.lower()

                balance_year = today_year

                balance_date_button = app.window(title = window_title)[f'{balance_date_button_name}']

                balance_date_button.click_input()

                time.sleep(5)

            elif ((int(today_day) + balance_date) <= 0 and
                  int(today_month_num) == 1 and
                  balance_code == 1):

                previous_month_full = calendar.month_name[12]

                previous_month_num = 12

                previous_year = int(today_year) - 1

                balance_date = calendar.monthrange(int(previous_year), previous_month_num)[1] + int(balance_date) + int(today_day)

                first_day_weekday = calendar.weekday(int(previous_year), previous_month_num, 1)

                balance_date_button_name = f'DataItem{first_day_weekday + 2 + (balance_date - 1)}'

                balance_previous_month_button = app.window(title = window_title) \
                                                .child_window(title="Previous Button",
                                                                control_type="Button")

                balance_previous_month_button.click_input()

                time.sleep(5)

                balance_date_button = app.window(title = window_title)[f'{balance_date_button_name}']

                balance_date_button.click_input()

                time.sleep(5)

                balance_day = balance_date

                balance_month_full = previous_month_full.lower()

                balance_year = previous_year

            file_name = f'balance_{balance_code+1}_{balance_day}_{balance_month_full}_{balance_year}'

            report_button = app.window(title = window_title).child_window(title="Отчет")

            report_button.click_input()

            time.sleep(5)

            save_as_button.click_input()

            time.sleep(5)

            pyautogui.typewrite(file_name)

            time.sleep(5)

            file_type_button.click_input()

            time.sleep(5)

            pyautogui.press('up')

            time.sleep(5)

            pyautogui.press('up')

            time.sleep(5)

            pyautogui.press('enter')

            time.sleep(5)

            save_file_button.click_input()

            time.sleep(5)

            pyautogui.press('enter')

            time.sleep(5)

            pyautogui.press('enter')

            time.sleep(5)

        close_balance_button = app.window(title = window_title)['Button21']

        close_balance_button.click()

def autocargo_truck(app, truck_date_option, window_title):

    """
    This function downloads Autocargo truck reports

    """

    # Define button's names
    file_type_button = app.window(title = window_title) \
                            .child_window(title="Save as type:",
                                        auto_id="FileTypeControlHost",
                                        control_type="ComboBox")

    save_as_button = app.window(title = window_title) \
                        .child_window(title="Сохранить",
                                      control_type="Button")

    save_file_button = app.window(title = window_title) \
                          .child_window(title="Save",
                                        auto_id="1",
                                        control_type="Button")

    # List with balance dates that have no truck records found
    remove_list = []

    # Triggers to identify were the buttons next or previous month used
    prev_month_button_used = 0
    next_month_button_used = 0

    # Full information about today's day, month, year
    today_full_date = datetime.now()
    today_month_num = today_full_date.strftime('%m')
    today_month_full = calendar.month_name[int(today_month_num)]
    today_day = today_full_date.strftime("%d")
    today_year = today_full_date.strftime('%Y')

    for truck_date in truck_date_option:

        reports_dropdown = app.window(title = window_title)['MenuItem8']

        reports_dropdown.select()

        time.sleep(5)

        transport_dropdown = app.window(title = window_title)['MenuItem1']

        transport_dropdown.select()

        time.sleep(5)

        auto_dropdown = app.window(title = window_title)['MenuItem5']

        auto_dropdown.select()

        time.sleep(5)

        loading_trucks_dropdown = app.window(title = window_title)['MenuItem7']

        loading_trucks_dropdown.select()

        time.sleep(5)

        # Check if the truck day is correct after adding BALANCE_DATE_OPTION to today_day
        if (int(today_day) + truck_date) > 0:

            truck_day = int(today_day) + truck_date

            current_month_first_day_weekday = calendar.weekday(int(today_year), int(today_month_num), 1)

            truck_date_button_name = f'DataItem{current_month_first_day_weekday + 2 + (int(truck_day) - 1)}'

            if prev_month_button_used == 1:

                truck_next_month_button = app.window(title = window_title) \
                                                .child_window(title="Next Button",
                                                              control_type="Button")

                truck_next_month_button.click_input()

                next_month_button_used += 1
                prev_month_button_used -= 1

                time.sleep(5)

            truck_month_full = today_month_full.lower()

            truck_year = today_year

            truck_date_button = app.window(title = window_title)[f'{truck_date_button_name}']

            truck_date_button.click_input()

            time.sleep(5)

        elif ((int(today_day) + truck_date) <= 0 and
              int(today_month_num) != 1):

            previous_month_full = calendar.month_name[int(today_month_num)-1]

            previous_month_num = int(today_month_num) - 1

            previous_month_truck_date = calendar.monthrange(int(today_year), previous_month_num)[1] + int(truck_date) + int(today_day)

            previous_month_first_day_weekday = calendar.weekday(int(today_year), previous_month_num, 1)

            truck_date_button_name = f'DataItem{previous_month_first_day_weekday + 2 + (previous_month_truck_date - 1)}'

            if prev_month_button_used == 0:

                truck_previous_month_button = app.window(title = window_title) \
                                                .child_window(title="Previous Button",
                                                            control_type="Button")

                truck_previous_month_button.click_input()

                if next_month_button_used == 1:

                    next_month_button_used -= 1

                prev_month_button_used += 1

                time.sleep(5)

            time.sleep(5)

            truck_day = previous_month_truck_date

            truck_month_full = previous_month_full.lower()

            truck_year = today_year

            truck_date_button = app.window(title = window_title)[f'{truck_date_button_name}']

            truck_date_button.click_input()

            time.sleep(5)

        elif ((int(today_day) + truck_date) <= 0 and
              int(today_month_num) == 1):

            previous_month_full = calendar.month_name[12]

            previous_month_num = 12

            previous_year = int(today_year) - 1

            previous_month_previous_year_truck_date = calendar.monthrange(int(previous_year), previous_month_num)[1] + int(truck_date) + int(today_day)

            previous_month_previous_year_first_day_weekday = calendar.weekday(int(previous_year), previous_month_num, 1)

            truck_date_button_name = f'DataItem{previous_month_previous_year_first_day_weekday + 2 + (previous_month_previous_year_truck_date - 1)}'

            if prev_month_button_used == 0:

                truck_previous_month_button = app.window(title = window_title) \
                                                .child_window(title="Previous Button",
                                                            control_type="Button")

                truck_previous_month_button.click_input()

                if next_month_button_used == 1:

                    next_month_button_used -= 1

                prev_month_button_used += 1

                time.sleep(5)

            time.sleep(5)

            truck_day = previous_month_previous_year_truck_date

            truck_month_full = previous_month_full.lower()

            truck_year = previous_year

            truck_date_button = app.window(title = window_title)[f'{truck_date_button_name}']

            truck_date_button.click_input()

            time.sleep(5)

            prev_month_button_used += 1

        file_name = f'truck_{truck_day}_{truck_month_full}_{truck_year}'

        confirm_button = app.window(title = window_title).child_window(title="Да",
                                                                    control_type="Button")

        confirm_button.click()

        time.sleep(5)

        # Check if there is no info about the truck and we should quit the section
        # Redirect the standard output to a variable
        original_stdout = sys.stdout
        sys.stdout = StringIO()

        # Dump the tree structure of the window
        app.window(title = 'AutoCargo - ООО "Флауэрс Лайф" - ЕВГЕНИЯ').dump_tree()

        # Get the dumped tree as a string
        window_info = sys.stdout.getvalue()

        # Reset the standard output
        sys.stdout = original_stdout

        # Add truck date in the list to be removed and continue if no records found
        if 'Информация не найдена' in window_info:

            remove_list.append(truck_date)

            pyautogui.press('enter')

            time.sleep(5)

            continue

        confirm_button.click()

        time.sleep(5)

        save_as_button.click_input()

        time.sleep(5)

        pyautogui.typewrite(f'{file_name}')

        time.sleep(5)

        file_type_button.click_input()

        time.sleep(5)

        pyautogui.press('up')

        time.sleep(5)

        pyautogui.press('up')

        time.sleep(5)

        pyautogui.press('enter')

        time.sleep(5)

        save_file_button.click_input()

        time.sleep(10)

        pyautogui.press('enter')

        time.sleep(5)

    if remove_list:

        for date in remove_list:

            truck_date_option.remove(date)

def autocargo_close(app, window_title):

    """
    This function closes the Autocargo program

    """

    time.sleep(5)

    close_button = app.window(title = window_title).child_window(
        title="Close",
        control_type="Button")

    close_button.click()

def balance_date_calc(balance_date):

    """
    This function calculates the balance date to be used as arguments,
    if balance file were downloaded earlier

    """

    # Full information about today's day, month, year
    today_full_date = datetime.now()
    today_month_num = today_full_date.strftime('%m')
    today_month_full = calendar.month_name[int(today_month_num)]
    today_day = today_full_date.strftime("%d")
    today_year = today_full_date.strftime('%Y')
    previous_year = int(today_year) - 1

    if (int(today_day) + balance_date) > 0:

        balance_day = int(today_day) + balance_date

        balance_month_full = today_month_full.lower()

        balance_year = today_year

    elif (int(today_day) + balance_date) <= 0 and int(today_month_num) != 1:

        previous_month_num = int(today_month_num) - 1

        balance_day = calendar.monthrange(int(today_year), previous_month_num)[1] + int(balance_date) + int(today_day)

        balance_month_full = (calendar.month_name[int(today_month_num)-1]).lower()

        balance_year = today_year

    elif (int(today_day) + balance_date) <= 0 and int(today_month_num) == 1:

        previous_month_num = 12

        balance_day = calendar.monthrange(int(previous_year), 12)[1] + int(balance_date) + int(today_day)

        balance_month_full = (calendar.month_name[12]).lower()

        balance_year = int(today_year) - 1

    return [balance_day, balance_month_full, balance_year]

def truck_date_calc(truck_date):

    """
    This function calculates the truck date to be used as arguments,
    if truck file were downloaded earlier

    """

    # Full information about today's day, month, year
    today_full_date = datetime.now()
    today_month_num = today_full_date.strftime('%m')
    today_month_full = calendar.month_name[int(today_month_num)]
    today_day = today_full_date.strftime("%d")
    today_year = today_full_date.strftime('%Y')
    previous_year = int(today_year) - 1

    if (int(today_day) + truck_date) > 0:

        truck_day = int(today_day) + truck_date

        truck_month_full = today_month_full.lower()

        truck_year = today_year

    elif (int(today_day) + truck_date) <= 0 and int(today_month_num) != 1:

        previous_month_num = int(today_month_num) - 1

        truck_day = calendar.monthrange(int(today_year), previous_month_num)[1] + int(truck_date) + int(today_day)

        truck_month_full = (calendar.month_name[int(today_month_num)-1]).lower()

        truck_year = today_year

    elif (int(today_day) + truck_date) <= 0 and int(today_month_num) == 1:

        truck_day = calendar.monthrange(int(previous_year), 12)[1] + int(truck_date) + int(today_day)

        truck_month_full = (calendar.month_name[12]).lower()

        truck_year = int(today_year) - 1

    return [truck_day, truck_month_full, truck_year]
