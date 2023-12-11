"""
This module contains functions that clean and normalize the xls files fetched
from AutoCargo program, create sale file based on the expenses from balance files
    
"""

import os
import math
from datetime import datetime
import pandas as pd

from openpyxl.styles.colors import WHITE, RGB
__old_rgb_set__ = RGB.__set__

# Fix the ARGB hex values error by making it white
def __rgb_set_fixed__(self, instance, value):

    try:

        __old_rgb_set__(self, instance, value)

    except ValueError as e:

        if e.args[0] == 'Colors must be aRGB hex values':
            __old_rgb_set__(self, instance, WHITE)

# Clean balance reports
def prepare_balance(balance_day, balance_month_full, balance_year):

    """
    This function prepares balance and truck reports
    for the future import to the database

    """
    # List of df to be returned
    balance_df_list = []

    # Settings for cleaning balance reports
    balance_cols = 'A:C, E, G, J, K, M, N, R, U:W'

    balance_skip_rows = [0, 1, 2, 3, 4]

    balance_rename_cols = ['expense_date', 'expense_content_supplier', 'expense_type',
                           'expense_awb', 'expense_currency_rate', 'expense_total_rub',
                           'expense_currency','expense_total_usd', 'expense_total_eur',
                           'expense_marking','expense_full_box', 'expense_weight',
                           'expense_volume']

    # Fix the ARGB hex values error by making it white
    RGB.__set__ = __rgb_set_fixed__

    # List of usd balances
    usd_balance_codes = [3, 5, 8, 11]

    # List of rub balances
    rub_balance_codes = [1, 2, 4, 6, 7, 9, 10, 12]

    # Total count of the balances available to check in the Autocargo
    balance_count = 12

    # Dictionary translating balance code to the balance name
    balance_codes_translation = {1:'', 2:'', 3:'',
                                 4:'', 5:'', 6:'',
                                 7:'', 8:'', 9:'',
                                 10:'', 11:'', 12:'',
                                 13:'', 14:'', 15:'', 16:'',
                                 17:''}

    # Detect row with precool
    docs_dict = {}

    # Detect Ams rows with the same marking and date
    ams_dict = {}

    # Detect consolidation rows with the same marking and awb
    consolidation_dict = {}

    # Rows to delete after uniting Ams shipments
    extra_ams_rows = []

    # Rows to delete after uniting non Ams shipments
    extra_rows = []

    # Rows to delete
    extra_precool_rows = []

    # Clean the balance reports, import them to the database and create a sale info
    for balance in range(balance_count):

        balance_code = balance + 1

        # Assign file name, path and new path
        balance_name = f'balance_{balance_code}_{balance_day}_{balance_month_full}_{balance_year}'
        balance_path = f'C:/autocargo_reports/{balance_name}.xlsx'

        print(f'Start {balance_name}')

        if os.path.exists(balance_path):

            # Read the xls file into pandas DataFrame
            balance_file = pd.read_excel(balance_path, usecols=balance_cols,
                                         skiprows=balance_skip_rows, engine='openpyxl')

            # Rename columns
            for name in balance_rename_cols:

                balance_file.columns.values[balance_rename_cols.index(name)] = (
                    balance_rename_cols[balance_rename_cols.index(name)]
                )

            # Delete various rows
            balance_file = balance_file.drop(balance_file[balance_file['expense_date'] == 'ДАТА'].index)
            balance_file = balance_file.drop(balance_file[balance_file['expense_type'] == ' >>>'].index)

            for ind in balance_file.index:

                if 'Страница' in str(balance_file['expense_date'][ind]):

                    balance_file = balance_file.drop([ind])

            balance_file.dropna(subset = ['expense_date'], inplace = True)

            # Add columns
            balance_file.insert(loc=13, column='expense_balance_code', value = '')
            balance_file.insert(loc=14, column='expense_balance_currency', value = '')
            balance_file.insert(loc=15, column='supplier_id', value = '')

            # Normalize and sort the data
            for ind in balance_file.index:

                currency = balance_file['expense_total_rub'][ind]
                currency_rate = balance_file['expense_currency'][ind]

                # Fill currency information based on the balance type
                if balance_code in usd_balance_codes:

                    balance_file['expense_balance_currency'][ind] = 'usd'

                    if 'rub' in currency.lower():

                        balance_file['expense_total_rub'][ind] = balance_file['expense_currency_rate'][ind]
                        balance_file['expense_total_eur'][ind] = 0

                    if 'eur' in currency.lower():

                        if balance_file['expense_currency_rate'][ind] != ' ':

                            balance_file['expense_total_eur'][ind] = balance_file['expense_currency_rate'][ind]

                        else:
                            balance_file['expense_total_eur'][ind] = 0

                        balance_file['expense_total_rub'][ind] = 0

                    if 'usd' in str(currency).lower():

                        balance_file['expense_total_rub'][ind] = 0
                        balance_file['expense_total_eur'][ind] = 0

                elif balance_code in rub_balance_codes:

                    balance_file['expense_balance_currency'][ind] = 'rub'
                    balance_file['expense_total_rub'][ind] = balance_file['expense_total_usd'][ind]

                    if 'usd' in str(currency).lower():

                        balance_file['expense_total_usd'][ind] = balance_file['expense_currency_rate'][ind]
                        balance_file['expense_total_eur'][ind] = 0

                    elif 'eur' in str(currency).lower():

                        if balance_file['expense_currency_rate'][ind] != ' ':

                            balance_file['expense_total_eur'][ind] = balance_file['expense_currency_rate'][ind]

                        else:
                            balance_file['expense_total_eur'][ind] = 0

                        balance_file['expense_total_usd'][ind] = 0

                    elif 'rub' in str(currency).lower():

                        balance_file['expense_total_eur'][ind] = 0
                        balance_file['expense_total_usd'][ind] = 0

                balance_file['expense_currency'][ind] = currency
                balance_file['expense_currency_rate'][ind] = currency_rate

                # Fill empty values
                empty_fill = {'expense_currency_rate':0,
                              'expense_volume':0,
                              'expense_weight':0,
                              'expense_full_box':0,
                              'expense_content_supplier':'none',
                              'expense_awb':'none',
                              'expense_total_usd':0,
                              'expense_total_eur':0,
                              'expense_total_rub':0
                              }

                for column, value in empty_fill.items():

                    if (pd.isnull(balance_file.at[ind, column]) or
                        balance_file.at[ind, column] == ' ' or
                        balance_file.at[ind, column] == '' or
                        str(balance_file.at[ind, column]).lower() == 'nan'):

                        balance_file.at[ind, column] = value

                # Find and separate awb number and company name of shipment
                if (balance_file['expense_content_supplier'][ind][1:4].isnumeric() and
                    balance_file['expense_content_supplier'][ind][5:13].isnumeric()):

                    awb = balance_file['expense_content_supplier'][ind][:13].strip()
                    company = balance_file['expense_content_supplier'][ind][13:].strip()

                    balance_file['expense_content_supplier'][ind] = company
                    balance_file['expense_awb'][ind] = awb

                # Identify the box type without name
                if 'тара' in balance_file['expense_type'][ind].lower():

                    box_size = 1/(float(balance_file['expense_volume'][ind])/float(balance_file['expense_full_box'][ind]))
                    box_size = int(math.ceil(box_size))
                    box_type_ind = balance_file["expense_type"][ind].find(']')
                    balance_file["expense_type"][ind] = balance_file["expense_type"][ind][:box_type_ind]
                    balance_file["expense_type"][ind] += f'{str(box_size)}]'

                # Extract shipments markings
                if '{' in balance_file['expense_marking'][ind]:

                    marking = ''

                    for letter in balance_file['expense_marking'][ind]:

                        if letter != '{':

                            marking += letter

                        elif letter == '{':

                            break

                    balance_file['expense_marking'][ind] = marking.strip()

                elif ('отправка' in balance_file['expense_marking'][ind].lower() and
                      'телег' in balance_file['expense_marking'][ind].lower()):

                    balance_file['expense_type'][ind] = balance_file['expense_marking'][ind]

                elif 'штраф' in balance_file['expense_marking'][ind].lower():

                    balance_file['expense_type'][ind] = balance_file['expense_marking'][ind]

                # Find weight and identify the type for pre-cool rows
                elif 'прикулинг' in balance_file['expense_marking'][ind].lower():

                    if (balance_file['expense_total_rub'][ind] != 0 and
                        balance_file['expense_total_usd'][ind] != 0 and
                        balance_file['expense_total_eur'][ind] != 0):

                        balance_file['expense_awb'][ind] = balance_file['expense_type'][ind].strip()
                        balance_file['expense_type'][ind] = 'прикулинг'

                    elif (balance_file['expense_total_rub'][ind] == 0 and
                          balance_file['expense_total_usd'][ind] == 0 and
                          balance_file['expense_total_eur'][ind] == 0):

                        extra_precool_rows.append(ind)

                # Find who sent and who received transfer
                elif 'трансфер' in balance_file['expense_marking'][ind].lower():

                    balance_file['expense_type'][ind] = 'трансфер'
                    balance_file['expense_marking'][ind] = balance_file['expense_marking'][ind][1:(len(balance_file['expense_marking'][ind]))-2]
                    transfer_separator = balance_file['expense_marking'][ind].find('->')

                    # Identify who sends/receives the transfer
                    if balance_codes_translation[balance_code] in balance_file['expense_marking'][ind][:transfer_separator].lower():

                        balance_file['expense_marking'][ind] += ' = отправитель'

                    elif balance_codes_translation[balance_code] in balance_file['expense_marking'][ind][transfer_separator:].lower():

                        balance_file['expense_marking'][ind] += ' = получатель'

                # Rename documents fee row
                elif 'док' in balance_file['expense_type'][ind].lower():

                    balance_file['expense_type'][ind] = 'оформление документов'
                    balance_file['expense_full_box'][ind] = 0
                    balance_file['expense_weight'][ind] = 0
                    balance_file['expense_volume'][ind] = 0

                    # Extract shipments markings and content supplier
                    if '_' in balance_file['expense_marking'][ind]:

                        copy_m = 0
                        marking = ''

                        for letter in str(balance_file['expense_marking'][ind]):

                            if letter == '_':

                                copy_m = 1

                            elif letter != '_' and copy_m == 1:

                                marking += letter

                        company_ind = balance_file['expense_marking'][ind].index('_')
                        company = balance_file['expense_marking'][ind][:company_ind]
                        balance_file['expense_content_supplier'][ind] = company.strip()
                        balance_file['expense_marking'][ind] = marking.strip()

                    doc_marking = balance_file['expense_marking'][ind]
                    doc_date = balance_file['expense_date'][ind]

                    docs_total_usd = float(balance_file['expense_total_usd'][ind])
                    docs_total_eur = float(balance_file['expense_total_eur'][ind])
                    docs_total_rub = float(balance_file['expense_total_rub'][ind])

                    docs_dict[f'{doc_marking}{doc_date}'] = [ind,
                                                             docs_total_usd,
                                                             docs_total_eur,
                                                             docs_total_rub]

                # Identify the type for fee rows
                elif 'комиссия' in balance_file['expense_marking'][ind].lower():

                    balance_file['expense_type'][ind] = 'комиссия'

                # Identify the type for payment rows (positive number)
                elif 'оплата' in balance_file['expense_awb'][ind].lower():

                    balance_file['expense_type'][ind] = balance_file['expense_awb'][ind]

                # Turn the expense number to a positive number for rub balance
                if '-' in str(balance_file['expense_total_rub'][ind]):

                    balance_file['expense_total_rub'][ind] = str(balance_file['expense_total_rub'][ind])[1:]

                # Turn the expense number to a positive number for usd balance
                elif '-' in str(balance_file['expense_total_usd'][ind]):

                    balance_file['expense_total_usd'][ind] = str(balance_file['expense_total_usd'][ind])[1:]

                # Turn the expense number to a positive number for eur balance
                elif '-' in str(balance_file['expense_total_eur'][ind]):

                    balance_file['expense_total_eur'][ind] = str(balance_file['expense_total_eur'][ind])[1:]

                # Unite records with the same marking and date from Ams origin
                if ('срезка в ассортименте' in balance_file['expense_type'][ind].lower() or
                    'срезка гоа' in balance_file['expense_type'][ind].lower() or
                    'зелень декоративная по объему' in balance_file['expense_type'][ind].lower() or
                    'горшечные растения' in balance_file['expense_type'][ind].lower()):

                    ams_box_index = balance_file['expense_type'][ind].find('[')
                    ams_box_amount = int(balance_file['expense_full_box'][ind])
                    ams_marking = balance_file['expense_marking'][ind]
                    ams_date = balance_file['expense_date'][ind]
                    ams_total_usd = float(balance_file['expense_total_usd'][ind])
                    ams_total_eur = float(balance_file['expense_total_eur'][ind])
                    ams_total_rub = float(balance_file['expense_total_rub'][ind])
                    balance_file['expense_type'][ind] = balance_file['expense_type'][ind][:ams_box_index + 1] + str(ams_box_amount) + balance_file['expense_type'][ind][ams_box_index + 1:]
                    ams_type_ind = balance_file['expense_type'][ind].find('[')
                    ams_type = balance_file['expense_type'][ind][ams_type_ind:].lower()

                    if f'{ams_marking}{ams_date}' not in ams_dict:

                        ams_dict[f'{ams_marking}{ams_date}'] = [ind,
                                                                ams_box_amount,
                                                                ams_total_usd,
                                                                ams_total_eur,
                                                                ams_total_rub,
                                                                False,
                                                                balance_file['expense_type'][ind].lower().strip()]

                    else:

                        ams_dict[f'{ams_marking}{ams_date}'][1] += ams_box_amount
                        ams_dict[f'{ams_marking}{ams_date}'][2] += ams_total_usd
                        ams_dict[f'{ams_marking}{ams_date}'][3] += ams_total_eur
                        ams_dict[f'{ams_marking}{ams_date}'][4] += ams_total_rub
                        ams_dict[f'{ams_marking}{ams_date}'][6] += ams_type

                        extra_ams_rows.append(ind)

                # Turn all the strings of the row lowercase
                for column in balance_rename_cols:

                    balance_file[column][ind] = str(balance_file[column][ind]).lower().strip(',.·•').strip()

                balance_file['expense_balance_code'][ind] = balance_code
                balance_file['supplier_id'][ind] = 2

            # Update united rows
            for info_ams in ams_dict.values():

                balance_file['expense_full_box'][info_ams[0]] = info_ams[1]
                balance_file['expense_total_usd'][info_ams[0]] = info_ams[2]
                balance_file['expense_total_eur'][info_ams[0]] = info_ams[3]
                balance_file['expense_total_rub'][info_ams[0]] = info_ams[4]
                balance_file['expense_type'][info_ams[0]] = info_ams[6]

            balance_file = balance_file.drop(extra_ams_rows)
            extra_ams_rows.clear()

            # Unite records with the same flight number, marking and date of non Ams origin
            for ind in balance_file.index:

                if 'консолидат' in balance_file['expense_type'][ind]:

                    fl_marking = balance_file['expense_marking'][ind]
                    fl_weight = float(balance_file['expense_weight'][ind])
                    fl_awb = balance_file['expense_awb'][ind]
                    fl_date = balance_file['expense_date'][ind]
                    fl_total_usd = float(balance_file['expense_total_usd'][ind])
                    fl_total_eur = float(balance_file['expense_total_eur'][ind])
                    fl_total_rub = float(balance_file['expense_total_rub'][ind])

                    if f'{fl_marking}{fl_awb}{fl_date}' not in consolidation_dict:

                        consolidation_dict[f'{fl_marking}{fl_awb}{fl_date}'] = [ind,
                                                                                fl_weight,
                                                                                fl_marking,
                                                                                fl_awb,
                                                                                fl_date,
                                                                                fl_total_usd,
                                                                                fl_total_eur,
                                                                                fl_total_rub,
                                                                                'no']

                    else:
                        consolidation_dict[f'{fl_marking}{fl_awb}{fl_date}'][1] += fl_weight
                        consolidation_dict[f'{fl_marking}{fl_awb}{fl_date}'][5] += fl_total_usd
                        consolidation_dict[f'{fl_marking}{fl_awb}{fl_date}'][6] += fl_total_eur
                        consolidation_dict[f'{fl_marking}{fl_awb}{fl_date}'][7] += fl_total_rub
                        extra_rows.append(ind)

            # Update united rows
            for info_cons in consolidation_dict.values():

                balance_file['expense_weight'][info_cons[0]] = info_cons[1]
                balance_file['expense_total_usd'][info_cons[0]] = info_cons[5]
                balance_file['expense_total_eur'][info_cons[0]] = info_cons[6]
                balance_file['expense_total_rub'][info_cons[0]] = info_cons[7]

            # Remove rows that were added to it's match
            consolidation_dict.clear()

            # Remove empty pre-cooling rows
            balance_file = balance_file.drop(extra_precool_rows)
            extra_precool_rows.clear()            

            # Remove rows that were added to it's match
            balance_file = balance_file.drop(extra_rows)
            extra_rows.clear()

            # Unite ams shipments and docs cost with the same marking and date
            for ams_name, info_ams in ams_dict.items():

                for doc_name, info_doc in docs_dict.items():

                    if (ams_name == doc_name and
                        info_ams[5] is False):

                        info_ams[5] = True
                        info_ams[2] += info_doc[1]
                        info_ams[3] += info_doc[2]
                        info_ams[4] += info_doc[3]

                        balance_file['expense_total_usd'][info_ams[0]] = info_ams[2]
                        balance_file['expense_total_eur'][info_ams[0]] = info_ams[3]
                        balance_file['expense_total_rub'][info_ams[0]] = info_ams[4]

                        balance_file = balance_file.drop(info_doc[0])

            ams_dict.clear()
            docs_dict.clear()

            print(f'Save and return {balance_name}')

            if os.path.exists('C:/autocargo_reports/export/'):

                directory_name = 'C:/autocargo_reports/export/'

            else:

                directory_name = 'C:/'

                print('Couldnt find a path to save the file, saving to the disc C')

            balance_xls_path = f'{directory_name}balance_{balance_code}_{balance_day}_{balance_month_full}_{balance_year}.xlsx'

            balance_file.to_excel(balance_xls_path)

            if os.path.exists(balance_path):

                os.remove(balance_path)

            # Check if balance_file DataFrame exists
            if not balance_file.empty:

                balance_df_list.append(balance_file)

        else:

            print(f'Path {balance_path} is not valid!')

    return balance_df_list

# Clean truck reports
def prepare_truck(truck_day, truck_month_full, truck_year):

    """
    This function prepares truck reports for import to the database

    """

    # List of df to be returned
    truck_df_list = []

    # Settings for cleaning truck reports
    truck_cols = 'A, B, C, D:F, H:K, M, N, P, Q'

    truck_skip_rows = [0, 1]

    truck_rename_cols = ['shipment_truck_name', 'shipment_country', 'shipment_ams_arrival',
                         'shipment_awb', 'shipment_supplier', 'shipment_forever_balance',
                         'shipment_marking', 'shipment_box_full', 'shipment_box_amount', 
                         'shipment_weight_fact', 'shipment_weight_vol', 'shipment_date',
                         'shipment_volume', 'shipment_comment']

    # Fix the ARGB hex values error by making it white
    RGB.__set__ = __rgb_set_fixed__

    # Clean the truck reports and import them to the database
    truck_name_df = f'truck_{truck_day}_{truck_month_full}_{truck_year}'
    truck_path = f'C:/autocargo_reports/{truck_name_df}.xlsx'

    print(f'Start {truck_name_df}')

    if os.path.exists(truck_path):

        # Read the CSV file into a pandas DataFrame
        truck_file = pd.read_excel(truck_path, usecols=truck_cols,
                                    skiprows=truck_skip_rows, engine='openpyxl')

        # Rename columns
        for name in truck_rename_cols:

            truck_file.columns.values[truck_rename_cols.index(name)] = truck_rename_cols[truck_rename_cols.index(name)]

        # Delete various row
        truck_file.dropna(subset=['shipment_marking'], inplace=True)

        truck_file['shipment_truck_name'].fillna('unknown', inplace=True)

        truck_file = truck_file.drop(truck_file[truck_file['shipment_truck_name'] == 'ПРИКУЛИНГ'].index)

        truck_file = truck_file.drop(truck_file[truck_file['shipment_marking'] == 'forevercargo.ru - личный кабинет клиента'].index)

        truck_file.insert(loc=13, column='shipment_truck_balance', value = '')

        truck_file.insert(loc=14, column='shipment_status', value = '')

        # Normalize and sort
        for ind in truck_file.index:

            awb = ''

            truck_file['shipment_truck_balance'][ind] = 'forever'

            truck_file['shipment_status'][ind] = 'на границе'

            # Rename truck
            if ('БЕЗ ПРИКУЛИНГА' not in truck_file['shipment_truck_name'][ind] and
                'IPHANDLERS' not in truck_file['shipment_truck_name'][ind]):

                truck_name = truck_file['shipment_truck_name'][ind]

                ams_date = truck_file['shipment_date'][ind]
                ams_year = ams_date[-4:]
                ams_day = ams_date[:2]
                ams_month = (ams_date[3:len(ams_date)-5]).lower().strip()
                ams_month_number = datetime.strptime(ams_month, '%B').month
                ams_correct_date = f'{ams_year}-{ams_month_number}-{ams_day}'

                truck_file = truck_file.drop([ind])

                continue

            if ('БЕЗ ПРИКУЛИНГА' in truck_file['shipment_truck_name'][ind] or
                'IPHANDLERS' in truck_file['shipment_truck_name'][ind]):

                truck_file['shipment_truck_name'][ind] = truck_name
                truck_file['shipment_date'][ind] = ams_correct_date
                truck_file['shipment_ams_arrival'][ind] = ams_correct_date

            # Sort awb
            if '-' in str(truck_file['shipment_awb'][ind]):

                awb = str(truck_file['shipment_awb'][ind])[:12]

                truck_file['shipment_awb'][ind] = awb

            elif '/' in str(truck_file['shipment_awb'][ind]):

                truck_file['shipment_awb'][ind] = ''

            truck_file['shipment_marking'][ind] = str(truck_file['shipment_marking'][ind]).strip()

            truck_file['shipment_box_full'][ind] = float(truck_file['shipment_box_full'][ind])

            # Turn all the strings lowercase
            for column in truck_rename_cols:

                truck_file[column][ind] = str(truck_file[column][ind]).lower()
                truck_file[column][ind] = str(truck_file[column][ind]).strip(',.·').strip()

        # Check if balance_file DataFrame exists
        if not truck_file.empty:

            print(f'Save and return truck_{truck_day}_{truck_month_full}_{truck_year}')

            if os.path.exists('C:/autocargo_reports/export/'):

                directory_name = 'C:/autocargo_reports/export/'

            else:

                directory_name = 'C:/'

                print('Couldnt find a path to save the file, saving to the disc C')

            truck_file_path = f'{directory_name}truck_{truck_day}_{truck_month_full}_{truck_year}.xlsx'

            truck_file.to_excel(truck_file_path)

            if os.path.exists(truck_path):

                os.remove(truck_path)

            truck_df_list.append(truck_file)

    else:

        print(f'Path {truck_path} is not valid!')

    return truck_df_list

# Create sale and other operations reports based on the balance reports
def create_forever_sales(balance_report):

    """
    This function prepares creates a sales report for database
    based on the cleaned forever balance report

    """

    # List of df to be returned
    sale_df_list = []

    print('Start Forever sale')

    balance_rename_cols = ['expense_date', 'expense_content_supplier', 'expense_type',
                           'expense_awb', 'expense_currency_rate', 'expense_total_rub',
                           'expense_currency','expense_total_usd', 'expense_total_eur',
                           'expense_marking','expense_full_box', 'expense_weight',
                           'expense_volume']

    # Dictionary of balance report columns indexes and connected tp them columns in sales
    sales_cols = {0:'sale_date', 1:'sale_content_supplier', 2:'sale_type',
                  3:'sale_awb', 7:'sale_total_usd', 8:'sale_total_eur',
                  6:'sale_currency', 9:'sale_marking', 10:'sale_full_box',
                  11:'sale_weight', 12:'sale_volume'}

    sales_file = pd.DataFrame()

    for balance_col_ind, sales_col in sales_cols.items():

        row_content = []

        for row in balance_report.index:

            if ('оплата' not in balance_report['expense_type'][row] and
                'трансфер' not in balance_report['expense_type'][row] and
                'комиссия' not in balance_report['expense_type'][row] and
                'штраф' not in balance_report['expense_type'][row] and
                'отправка' not in balance_report['expense_type'][row]):

                balance_col_name = balance_rename_cols[balance_col_ind]

                row_content.append(str(balance_report[balance_col_name][row]))

        sales_file[sales_col] = row_content

        sales_file['supplier_id'] = 2

    if not sales_file.empty:

        print('Save and return Forever sale')

        if os.path.exists('C:/autocargo_reports/export/'):

            directory_name = 'C:/autocargo_reports/export/'

        else:

            directory_name = 'C:/'

            print('Couldnt find a path to save the file, saving to the disc C')

        sale_code = 1

        sale_path = f'{directory_name}sale_forever{sale_code}.xlsx'

        while os.path.exists(sale_path):

            sale_code += 1

            sale_path = f'{directory_name}sale_forever{sale_code}.xlsx'

        sales_file.to_excel(sale_path)

        sale_df_list.append(sales_file)

        return sale_df_list

    else:

        print('Empty Sale Forever file!')
