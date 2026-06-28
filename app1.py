from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import *
from selenium.webdriver.chrome.service import Service

import time
import os
import sys
import random
import string
import re
import glob
import shutil
import csv
import pandas as pd
import numpy as np
import chromedriver_autoinstaller

from string import *
from random import randint
from os import path
from datetime import date, timedelta, datetime
#import xlrd

from openpyxl import load_workbook
import openpyxl
from openpyxl import Workbook

#chrome_options = Options()
#chrome_options.set_headless()
#driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
#driver = webdriver.Chrome(service=webdriver.chrome.service.Service(executable_path="C:\webdrivers\chromedriver.exe"))

""" app1.py is similar to app.py but in this one it is not about finding errors in files but more on file management.
 There are different entities and once all files are downloaded they are moved to their respective folders.
 This script also saves considerable amount of time enabling to gain more time for analysis and project management. """

service = Service(ChromeDriverManager().install())
#service = Service('C:\webdrivers\chromedriver.exe')
#driver = webdriver.Chrome('C:\webdrivers\chromedriver.exe')
driver = webdriver.Chrome(service = service)

driver.get("xxx")
time.sleep(3)
driver.refresh()
time.sleep(2)
driver.refresh()
driver.get("xxx")

driver.maximize_window()
wait = WebDriverWait(driver, 30)

d_load = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[@data-file='D...']")))
d_load.click()
etl_log = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[@data-file='...']")))
etl_log.click()
time.sleep(15)
search = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id = '...x']")))
search.click()

#action = ActionChains()
#action.send_keys(search)
passe_download = 'C:/_____Passe_____'+'/'
if not os.path.exists(passe_download):
    os.makedirs(passe_download)
elif os.path.exists(passe_download):
    pass

df = pd.read_excel('C:/_____Passe_____/Ql.xlsx', header = 0)

filename = df['File Name'].tolist()
print()
print(len(filename))

def create_fol(country):
    global newpath
    newpath = passe_download + str(country) + '/'

    if not os.path.exists(newpath):
        os.makedirs(newpath)

    elif os.path.exists(newpath):
        for filename in os.listdir(newpath):
            if os.path.isfile(os.path.join(newpath, filename)):
                os.remove(os.path.join(newpath, filename))

create_fol('Sl')
create_fol('Sp')

for i in range(len(filename)):
    FN = filename[i][:-4]
    print(filename[i][:-4], "Count ---> ", i)
    search.send_keys(filename[i][:-4])
    time.sleep(3)
    try:
        search.send_keys(Keys.ENTER)
    except StaleElementReferenceException as e:
        search.send_keys(Keys.ENTER)
        
    fichier = wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody[@id = '..']")))
    try:
        fichier.click()
    except TimeoutException as e:
        driver.refresh()
        time.sleep(5)
        search.send_keys(filename[i][:-4])
        time.sleep(3)
        search.send_keys(Keys.ENTER)
        fichier.click()
        
    except ElementNotInteractableException as e:
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[text() = '.../']"))).click()

    print(filename[i][11:13])
    time.sleep(2)
    country = filename[i][11:13]
    search = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 's..']")))
    search.clear()
    time.sleep(1)
time.sleep(3)

def partition_files():
    files = glob.glob('C:/Users/'+ str(os.getlogin()) +'/Downloads/..._*.log')
    for file in files:
        file_name = 'C:/Users/'+ str(os.getlogin()) +'/Downloads/' + str(filename[i][0:33]) + '.log'
        print(file_name)

        def convertir_xl(pays):
            with open(file, encoding="utf-8") as handle:
                content = handle.readlines()
                info = content[14:]
            with open('C:/_____Passe_____/' + str(pays) + '/' + str(file[28:-3]) + str('csv'), 'w') as x:
                for line in info:
                    x.write(line)

        def move_to_folder(pays):
            xl = pd.read_csv(r'C:/_____Passe_____/' + str(pays) + '/' + str(file[28:-3]) + str('csv'), delimiter = ';', encoding="latin-1", engine='python')
            xl = pd.read_csv(r'C:/_____Passe_____/' + str(pays) + '/' + str(file[28:-3]) + str('csv'), delimiter = '|', encoding="latin-1", engine='python')    

        if file[39:41] == 'SK':
            shutil.move(file,  'C:/_____Passe_____/Sl/' + str(file[28:]))

        elif file[39:41] == 'SP':
            shutil.move(file, 'C:/_____Passe_____/Sp/' + str(file[28:]))

time.sleep(15)
partition_files()
time.sleep(100)
partition_files()

print("Download complete !")
