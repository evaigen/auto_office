CURRENT VERSION: 0.9

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
