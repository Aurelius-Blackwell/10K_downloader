import os
from pathlib import Path
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pyperclip
import ast

''' This program uses three identifiers relating to the same company:
        CIK, unique SEC filing identifiers
        Tickers, companies' stockmarket identifiers
        Names, companies' names
    These are drawn from 2 files: a JSON string from the SEC (containing
    CIK and tickers) and an Excel (containing tickers and names).
'''

#CIK data (JSON)
CIK_file = open('CIK_dictionary.txt')
CIK_fileContent = CIK_file.read()

#Russell 3000 constituent list
doc = Path.home() / 'Documents/r3000.xlsx'
wb = openpyxl.load_workbook(doc)
sheet = wb['russell-3000-constituents']

#Files in destination
p = os.listdir('F:\\2024_10K_files')

#Set variables
YEAR = 2024
present = [] #As opposed to missing
names = []
ticker_dict = {}
CIK_dict = ast.literal_eval(CIK_fileContent)

#If the 10K is already present, this list keeps track
for file in p:
    x = file.rstrip('.txt')
    present.append(x)

#Get ticker list from Excell file
for cellObj in sheet['b11':'b2685']:   #range = 11:2685
    for x in cellObj:
        if x.value not in present:
            names.append(x.value)
            # Coordinate of company name in the B file
            b = x.coordinate
            # Coordinate of company ticker in the A file
            a = 'A' + b.lstrip('B')
            # Dict entry is name:ticker
            ticker_dict[x.value] = sheet[a].value

#Shows which companies will be searched for
print(ticker_dict)

#Iterate over each name to get annual report
for name in range(len(names)):
    for x in range(9712):    #length of CIK_dict
        if CIK_dict[str(x)]['ticker'] == ticker_dict[names[name]]:
            cik = str(CIK_dict[str(x)]['cik_str'])
            if len(cik) < 10:
                cik_full = "0"*(10-len(cik))+cik
            else:
                cik_full = cik
            break
    url = 'https://www.sec.gov/edgar/search/#/dateRange=custom&category=custom&entityName='+cik_full+'&startdt='+str(YEAR)+'-01-01&enddt='+str(YEAR)+'-12-31&forms=10-K'
    browser = webdriver.Chrome()
    browser.get(url)
    try:
        elem = browser.find_element(By.CLASS_NAME, 'preview-file')
    except:
        browser.close()
        continue
    url = elem.get_attribute('href')

    #Open 10K filing in its own html browser
    browser.get(url)

    #Copy
    a = ActionChains(browser)
    a.key_down(Keys.CONTROL).send_keys('A').key_up(Keys.CONTROL).perform()
    a.key_down(Keys.CONTROL).send_keys('C').key_up(Keys.CONTROL).perform()
    destination = 'F:\\2024_10K_files\\'+names[name]+'.txt' #destination that 10Ks are written to
    with open(destination, 'w', encoding="utf-8") as file:
        file.write(pyperclip.paste())
    #Close browser to prevent bug in next search
    browser.close()
