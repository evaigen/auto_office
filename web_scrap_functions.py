"""
This module contains web scrapping and extracting data 
functions for the information about:
- Currency rates from Russian Central Bank
- Pre-cooling information from IP Handlers
- Flights information from Logiztik Alliance
    
"""

import os
import time
import fitz
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
import pyautogui
from lxml import etree

def parse_currency(day, month_num, year, code):

    """
    This function does web-scrapping of currency rate from Russian Central Bank

    """

    # URL of the currency website
    url = 'https://www.cbr.ru/currency_base/daily/'
    url += f'?UniDbQuery.Posted=True&UniDbQuery.To={day}.{month_num}.{year}'

    # Send a GET request to the website
    response = requests.get(url, timeout=15)

    # Check if the request was successful
    if response.status_code == 200:

        # Parse the HTML content
        html_content = response.text
        parser = etree.HTMLParser()
        tree = etree.fromstring(html_content, parser)

        # Define the XPath expression to extract the currency rate
        xpath_dollar = '//*[@id="content"]/div/div/div/div[3]/div/table/tbody/tr[15]/td[5]'
        xpath_euro = '//*[@id="content"]/div/div/div/div[3]/div/table/tbody/tr[16]/td[5]'
        xpath_date = '//*[@id="UniDbQuery_form"]/div/div/div/div/button'

        # Find the currency rate element using XPath
        dollar_element = tree.xpath(xpath_dollar)
        euro_element = tree.xpath(xpath_euro)
        date_element = tree.xpath(xpath_date)

        # Check if the element was found
        if dollar_element and euro_element and date_element:

            # Get the text content of the element
            dollar_rate = dollar_element[0].text
            euro_rate = euro_element[0].text
            date_info = f'{year}-{month_num}-{day}'

            dollar_rate = dollar_rate.replace(",", ".")
            euro_rate = euro_rate.replace(",", ".")

            dollar_rate = round(float(dollar_rate), 2)
            euro_rate = round(float(euro_rate), 2)

            euro_data = {'currency_rate':[euro_rate],
                         'currency_date':[date_info],
                         'currency_type':['eur']}

            dollar_data = {'currency_rate':[dollar_rate],
                           'currency_date':[date_info],
                           'currency_type':['usd']}

            eur_file = pd.DataFrame(euro_data)

            usd_file = pd.DataFrame(dollar_data)

            # Return dataframes with information about currency rate and date
            if code == 'usd':

                return usd_file

            if code == 'eur':

                return eur_file

            if code == 'usdeur':

                return usd_file, eur_file

        else:
            print("Currency rate element not found.")

    else:
        print("Failed to fetch the webpage.")

def parse_precooling():

    """
    This function does web-scrapping of pre-cooling services from IP Handlers

    """

    history_path = 'C:/Users/USER/Downloads/UFOTRUCK_ShipmentsHistory.xlsx'

    saldo_path = 'C:/Users/USER/Downloads/UFOTRUCK_Saldocard.xlsx'

    if not os.path.exists('C:/Users/USER/Downloads'):

        print('Invalid downloads path, empty strings returned')

        path_list = ['', '']

        return path_list

    if not os.path.exists(history_path) or not os.path.exists(saldo_path):

        login_url = 'https://tnt.iphandlers.nl/'
        login_payload = {
            'username': '',
            'password': ''
        }

        # Log in using Selenium
        driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
        driver.get(login_url)

        # Find the login form elements and submit the login details
        username_input = driver.find_element(By.XPATH,
                        '//*[@id="MainContent_UserName"]')
        password_input = driver.find_element(By.XPATH,
                        '//*[@id="MainContent_Password"]')
        submit_button = driver.find_element(By.XPATH,
                        '//*[@id="loginForm"]/div/div[4]/div/input')

        time.sleep(10)

        username_input.send_keys(login_payload['username'])

        time.sleep(10)

        password_input.send_keys(login_payload['password'])

        submit_button.click()

        time.sleep(10)

        if not os.path.exists(history_path):

            # Navigate to the history section
            history_button = driver.find_element(By.XPATH,
                            '//*[@id="mHistory"]/a/span')
            history_button.click()

            time.sleep(10)

            # Download the history report
            download_button = driver.find_element(By.XPATH,
                            '//*[@id="tnthistoryhdr"]/table/tbody/tr/td[2]/span')

            download_attempt = 0

            while not os.path.exists(history_path) and download_attempt < 5:

                download_attempt += 1

                time.sleep(10)

                download_button.click()

                time.sleep(10)

            if download_attempt == 5:

                print('History was not downloaded!')

        if not os.path.exists(saldo_path):

            # Navigate to the invoice history section
            invoice_button = driver.find_element(By.XPATH,
                            '//*[@id="mSaldo"]/a/span')
            invoice_button.click()

            time.sleep(10)

            # Download the saldocard report
            download_button = driver.find_element(By.XPATH,
                            '//*[@id="tntfinancehdr"]/table/tbody/tr/td[2]/span')

            download_attempt = 0

            while not os.path.exists(saldo_path) and download_attempt < 5:

                download_attempt += 1

                time.sleep(10)

                download_button.click()

                time.sleep(10)

            if download_attempt == 5:

                print('Saldocard was not downloaded!')

        # Close the browser window
        driver.quit()

        if os.path.exists(saldo_path) and os.path.exists(history_path):

            path_list = [history_path, saldo_path]

            return path_list

    else:

        print('Files from previous download exist, delete them first, empty strings returned')

        path_list = ['', '']

        return path_list

def get_downloaded_file_name():

    """
    This function downloads the latest downloaded PDF file name

    """
    # Specify the download directory
    download_directory = "C:/Users/USER/Downloads/"

    if not os.path.exists(download_directory):

        print('Invalid downloads path, empty string returned')

        latest_xml_file = ''

        return latest_xml_file

    # List files in the download directory
    files = os.listdir(download_directory)

    # Filter files based on your criteria (e.g., file extension)
    pdf_files = [file for file in files if file.endswith(".PDF")]

    if pdf_files:
        # Assuming you want the latest downloaded file
        latest_xml_file = max(pdf_files,
                              key=lambda f: os.path.getmtime(os.path.join(download_directory,
                              f)))

    else:
        print("No PDF files found in the download directory")

    return latest_xml_file

def parse_zikrach_markings(awb):

    """
    This function does web-scrapping of avia services from Logiztik Alliance

    """

    awb_formatted = awb[:3] + '-' + awb[3:7] + ' ' + awb[7:]

    login_url = 'https://cloud.logiztikalliance.com/logCloud/'
    login_url += 'DocumentacionOperativa/DocumentosExternos'
    login_payload = {
        'username': '',
        'password': ''
    }

    # Log in using Selenium
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

    driver.get(login_url)

    # Find the login form elements and submit the login details
    username_input = driver.find_element(By.XPATH,
                    '//*[@id="UserName"]')

    password_input = driver.find_element(By.XPATH,
                    '//*[@id="Password"]')

    login_button = driver.find_element(By.XPATH,
                    '//*[@id="loginForm"]/form/div/div[6]/div/button')

    time.sleep(10)

    username_input.send_keys(login_payload['username'])

    time.sleep(10)

    password_input.send_keys(login_payload['password'])

    time.sleep(10)

    login_button.click()

    time.sleep(20)

    # Find the awb form and filter documents
    awb_input = driver.find_element(By.XPATH,
                '//*[@id="gridGuias"]/table/thead/tr[2]/th[8]/span/span/span/input')

    awb_input.send_keys(awb_formatted)

    time.sleep(5)

    pyautogui.press('enter')

    time.sleep(5)

    download_docs_button = driver.find_element(By.XPATH,
                '//*[@id="gridGuias"]/table/tbody/tr/td[6]/div/button[1]/span')

    download_docs_button.click()

    time.sleep(10)

    awb_pdf_button = driver.find_element(By.XPATH,
                '//*[@id="gridDocumentoExterno"]/table/tbody/tr[3]/td[3]/div/a/span')

    awb_pdf_button.click()

    time.sleep(10)

    pdf_name = get_downloaded_file_name()

    if os.path.exists('C:/Users/USER/Downloads/'):

        directory_name = 'C:/Users/USER/Downloads/'

    else:

        directory_name = 'C:/'

        print('Couldnt find a path to save the file, saving to the disc C')

    if pdf_name:

        pdf_path = f'{directory_name}{pdf_name}'

    else:
        print('Empty pdf name, empty path returned')

        pdf_path = ''

    # Close the browser window
    driver.quit()

    return pdf_path

def parse_zikrach_price(awb):

    """
    This function does web-scrapping of avia services from Logiztik Alliance

    """

    awb_formatted = awb[:3] + '-' + awb[3:7] + ' ' + awb[7:]

    login_url = 'https://cloud.logiztikalliance.com/logCloud/'
    login_url += 'DocumentacionOperativa/DocumentosExternos'
    login_payload = {
        'username': '',
        'password': ''
    }

    # Log in using Selenium
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

    driver.get(login_url)

    # Find the login form elements and submit the login details
    username_input = driver.find_element(By.XPATH,
                    '//*[@id="UserName"]')

    password_input = driver.find_element(By.XPATH,
                    '//*[@id="Password"]')

    login_button = driver.find_element(By.XPATH,
                    '//*[@id="loginForm"]/form/div/div[6]/div/button')

    time.sleep(10)

    username_input.send_keys(login_payload['username'])

    time.sleep(10)

    password_input.send_keys(login_payload['password'])

    time.sleep(10)

    login_button.click()

    time.sleep(20)

    # Find the awb form and filter documents
    awb_input = driver.find_element(By.XPATH,
                '//*[@id="gridGuias"]/table/thead/tr[2]/th[8]/span/span/span/input')

    awb_input.send_keys(awb_formatted)

    time.sleep(5)

    pyautogui.press('enter')

    time.sleep(5)

    download_docs_button = driver.find_element(By.XPATH,
                           '//*[@id="gridGuias"]/table/tbody/tr/td[6]/div/button[1]/span')

    download_docs_button.click()

    time.sleep(10)

    master_awb_pdf_button = driver.find_element(By.XPATH,
                            '//*[@id="gridDocumentoExterno"]/table/tbody/tr[1]/td[3]/div/a')

    master_awb_pdf_button.click()

    time.sleep(10)

    pdf_name = get_downloaded_file_name()

    if pdf_name:

        pdf_path = f'C:/Users/USER/Downloads/{pdf_name}'

    else:

        print('Empty pdf name, empty path returned')

        pdf_path = ''

    # Close the browser window
    driver.quit()

    return pdf_path

def pdf_extract_data(pdf_path, code_type):

    """
    This function extracts the strings from PDF file

    """

    markings = ['Z-ARM', 'Z-SHEV', 'Z-KRAS', 'Z-VOLG', 'Z-RUSALIM']

    awb_details = {}

    if os.path.exists(pdf_path):

        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Initialize an empty string to store text
        text = ""

        # Iterate through each page in the PDF
        for page_num in range(pdf_document.page_count):
            # Get the page
            page = pdf_document[page_num]

            # Extract text from the page
            text = page.get_text()

            if code_type == 'markings':

                box_ind_st = text.find(' Weight') + 8
                box_ind_end = box_ind_st + text[box_ind_st:box_ind_st + 5].find('\n')

                weight_st = text.find('KG') + 3
                weight_end =  weight_st + text[weight_st:weight_st + 5].find('\n')

                full_box_st = text.find('FULL BOXES: ') + 12
                full_box_end = full_box_st + text[full_box_st:full_box_st + 10].find('\n')

                for marking in markings:

                    if marking in text:
                        awb_details[marking] = [text[box_ind_st:box_ind_end],
                                                text[full_box_st:full_box_end],
                                                text[weight_st:weight_end]]

            if code_type == 'price':
                sum_ind_st = text.find('Total Collect\n') + 14
                sum_ind_end = sum_ind_st + text[sum_ind_st:sum_ind_st + 10].find('\n')

                awb_sum = text[sum_ind_st:sum_ind_end]

        # Close the PDF file
        pdf_document.close()

        if code_type == 'markings':
            return awb_details

        if code_type == 'price':
            return awb_sum

    else:

        print('Invalid pdf path, empty string returned')

        empty_return = ''

        return empty_return

def test_alliance():

    """
    TEST

    """
    test_data = {'36987872956':{'Z-ARM':['27', '9.250', '449'], 'Z-SHEV':['14', '6.750', '320']},

             '36987877941':{'Z-ARM':['37', '14.000', '672'], 'Z-KRAS':['3', '1.500', '76'], 
                            'Z-SHEV':['16', '7.000', '365'], 'Z-VOLG':['8', '3.750', '191']},

             '15789019884':{'Z-ARM':['24', '8.000', '417'], 'Z-SHEV':['18', '8.750', '423']}}

    fin_test_data = {'36987872956':'2544.25',

                    '36987877941':'4283.00',

                    '15789019884':'2641.30'} 

    mark_test = True
    fin_test = True

    if mark_test:

        for awb_number, res_dict in test_data.items():

            test_dict = pdf_extract_data(parse_zikrach_markings(awb_number), 'markings')

            if test_dict == res_dict:
                print('Test completed')
                print(f'Test: {test_dict}')
                print(f'Answer: {res_dict}')

            else:
                print('Something went wrong')
                print(f'Test: {test_dict}')
                print(f'Answer: {res_dict}')

    if fin_test:

        for awb_num, awb_sum in fin_test_data.items():

            fin_test = pdf_extract_data(parse_zikrach_price(awb_num), 'price')

            if fin_test == awb_sum:
                print('TESTING IS COMPLETED:')
                print(f'Test: {fin_test}')
                print(f'Answer: {awb_sum}')

            else:
                print('SOMETHING WENT WRONG:')
                print(f'Test: {fin_test}')
                print(f'Answer: {awb_sum}')

def precool_extract_data(path_list):

    """
    This function prepares pre-cooling reports
    for the future import to the database

    """

    print('Start ip history')

    if (os.path.exists(path_list[0]) and
        os.path.exists(path_list[1])):

        ip_df_list = []

        # Settings for cleaning balance reports
        history_cols = 'C:D, F, H, I, K, L, N, O, Q'

        history_rename_cols = ['expense_eta_date', 'expense_country', 'expense_awb',
                               'expense_box', 'expense_full_box', 'expense_weight',
                               'expense_marking', 'expense_total', 'expense_load_date',
                               'expense_account']

        # Read the CSV file into pandas DataFrame
        history_file = pd.read_excel(path_list[0], usecols=history_cols,
                                     engine='openpyxl')

        # Rename columns
        for name in history_rename_cols:

            history_file.columns.values[history_rename_cols.index(name)] = (
                history_rename_cols[history_rename_cols.index(name)]
            )

        # Normalize and sort the data
        for ind in history_file.index:

            # Remove rows
            if (history_file['expense_account'][ind] != 'UFOTRUCK' or
                history_file['expense_country'][ind] == ''):

                history_file.drop(ind, inplace=True)

                continue

            #Make lower and get rid of whitespaces
            for column in history_rename_cols:

                history_file.at[ind, column] = str(history_file[column][ind]).lower().strip()

            history_file.at[ind, 'expense_total'] = '0.0'

            if '.0' in history_file.at[ind, 'expense_full_box']:

                history_file.at[ind, 'expense_full_box'] = str(history_file.at[ind, 'expense_full_box'])[:-2]

            record_month = history_file['expense_eta_date'][ind][3:5]
            record_year = history_file['expense_eta_date'][ind][6:10]
            record_date = history_file['expense_eta_date'][ind][:2]

            date_reverse = f'{record_year}-{record_month}-{record_date}'

            history_file['expense_eta_date'][ind] = date_reverse

        # Settings for cleaning balance reports
        saldo_cols = 'A:B, E:K'

        saldo_skip_rows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]

        saldo_rename_cols = ['expense_date', 'expense_awb', 'expense_weight',
                             'expense_box', 'expense_full_box', 'expense_marking',
                             'precool_invoice', 'precool_currency', 'expense_total']

        # Read the CSV file into pandas DataFrame
        saldo_file = pd.read_excel(path_list[1], usecols=saldo_cols,
                                skiprows=saldo_skip_rows, engine='openpyxl')

        # Rename columns
        for name in saldo_rename_cols:

            saldo_file.columns.values[saldo_rename_cols.index(name)] = (
                saldo_rename_cols[saldo_rename_cols.index(name)]
            )

        total_rows = len(saldo_file.index)

        saldo_file.drop(saldo_file.loc[(total_rows-10):].index, inplace=True)

        # Normalize and sort the data
        for ind in saldo_file.index:

            #Make lower and get rid of whitespaces
            for column in saldo_rename_cols:

                saldo_file.at[ind, column] = str(saldo_file[column][ind]).lower().strip()

        for ind_s in saldo_file.index:

            for ind_h in history_file.index:

                if (saldo_file['expense_awb'][ind_s] == history_file['expense_awb'][ind_h] and
                    saldo_file['expense_marking'][ind_s] == history_file['expense_marking'][ind_h]):

                    history_file.at[ind_h, 'expense_total'] = saldo_file.at[ind_s, 'expense_total']

        if not history_file.empty:

            for path_ip in path_list:

                if os.path.exists(path_ip):

                    os.remove(path_ip)

            print('Save and return ip history')

            if os.path.exists('C:/autocargo_reports/export/'):

                directory_name = 'C:/autocargo_reports/export/'

            else:

                directory_name = 'C:/'

                print('Couldnt find a path to save the file, saving to the disc C')

            ip_code = 1

            ip_path = f'{directory_name}ip_{ip_code}.xlsx'

            while os.path.exists(ip_path):

                ip_code += 1

                ip_path = f'{directory_name}ip_{ip_code}.xlsx'

            history_file.to_excel(ip_path)

            ip_df_list.append(history_file)

            return ip_df_list

    else:

        print(f'Path {path_list[0]} and {path_list[1]} are not valid!')

def create_ip_sales(balance_report):

    """
    This function prepares creates a sales report for database
    based on the cleaned ip balance report

    """

    print('Start IP Handlers sale')

    sale_ip_df_list = []

    history_cols = ['expense_eta_date', 'expense_country', 'expense_awb',
                    'expense_box', 'expense_full_box', 'expense_weight',
                    'expense_marking', 'expense_total', 'expense_load_date',
                    'expense_account']

    # Dictionary of balance report columns indexes and connected tp them columns in sales
    sales_cols = {0:'sale_date', 2:'sale_awb', 7:'sale_total_eur',
                  6:'sale_marking', 4:'sale_full_box', 5:'sale_weight'}

    sales_file = pd.DataFrame()

    for balance_col_ind, sales_col in sales_cols.items():

        row_content = []

        for row in balance_report.index:

            balance_col_name = history_cols[balance_col_ind]

            row_content.append(str(balance_report[balance_col_name][row]))

        sales_file[sales_col] = row_content

        sales_file['sale_content_supplier'] = ''
        sales_file['sale_type'] = 'прикулинг Iphandlers'
        sales_file['sale_total_usd'] = 0.0
        sales_file['sale_currency'] = 'eur'
        sales_file['sale_volume'] = 0.0
        sales_file['supplier_id'] = 5

    if not sales_file.empty:

        print('Save and return IP Handlers sale')

        if os.path.exists('C:/autocargo_reports/export/'):

            directory_name = 'C:/autocargo_reports/export/'

        else:

            directory_name = 'C:/'

            print('Couldnt find a path to save the file, saving to the disc C')

        sale_code = 1

        sale_path = f'{directory_name}sale_ip{sale_code}.xlsx'

        while os.path.exists(sale_path):

            sale_code += 1

            sale_path = f'{directory_name}sale_ip{sale_code}.xlsx'

        sales_file.to_excel(sale_path)

        sale_ip_df_list.append(sales_file)

        return sale_ip_df_list
