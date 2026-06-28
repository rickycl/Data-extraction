from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import *

import time
import os
import sys
import random
import string
import re
import pandas as pd

from string import *
from random import randint
from os import path
from datetime import date, timedelta, datetime
import glob
import shutil

#! driver = webdriver.Chrome("C:\webdrivers\chromedriver.exe")

import chromedriver_autoinstaller
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service = service)

#chrome_options = Options()
#chrome_options.set_headless()
#driver =webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)

""" This script shows the navigation in a web application. It inputs the actual date so as to download the most recent data.
It also takes into consideration the data during the weekend if we are on Monday, that is data starting from Friday to Monday.
It downloads an excel file containing the IDs that needs to be dowloaded one by one. It will copy and paste 1 ID at a time and download them.
Once all the individual txt file for each ID are downloaded, it will read them one by one to detect some unique keyword which is associated to an error.
If found, a folder for the day is created and then the error txt file is transferred over there for analysis.
Other functionalities in the script include file renaming and management.
This daily activity is time-consuming and with this automation, the process is done in a few minutes depending on the number of files.
"""

driver.get("https://xxxxx")
driver.maximize_window()

wait = WebDriverWait(driver, 8)

actions = ActionChains(driver)
#actions.move_to_element(log_in).perform()
time.sleep(2)

'''
original_window = driver.current_window_handle
#assert len(driver.window_handles) == 1
wait.until(EC.number_of_windows_to_be(2))

for window_handle in driver.window_handles:
    if window_handle != original_window:
        driver.switch_to.window(window_handle)
        break

cookie = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@id = 'accept-button']")))
cookie.click()
'''
##################################
username = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='uname']")))
username.send_keys('xxx') # <-----------------------

password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='pword']")))
password.send_keys('xx') #<------------------
##################################

submit = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name = 'submit']")))
submit.click()

def change_date(day):
    global iframe
    time.sleep(3)
    iframe = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//iframe[@src='https://.../search']")))
    #iframe = driver.find_element_by_xpath("//iframe[@src='.../search']")
    driver.switch_to.frame(iframe)

    #wait.until(EC.presence_of_element_located((By.XPATH, "//input[@class='...']"))).click()
    date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name = 'sdate']")))
    date_input.click()
    date_input.clear()
    date_input.send_keys(day.strftime('%d-%m-%Y'))

    state = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//select[@name = 'state']")))
    #state = driver.find_element_by_xpath("//select[@name = 'state']")
    state.click()
    time.sleep(3)
    state.send_keys(Keys.END)

    warning = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//option[@value = 'W...s']")))
    warning.click()

    """ActionChains(driver).key_down(Keys.CONTROL).\
        click("//option[@value = 'W...s']").\
        click("//option[@value = 'E...s']").\
        key_up(Keys.CONTROL).perform()
    """
    try:
        ActionChains(driver).send_keys(Keys.SHIFT, Keys.ARROW_DOWN).perform()
    except StaleElementReferenceException as e:
        pass

jour = datetime.now().strftime("%a")
if jour == 'Mon':
    day_3back = date.today() - timedelta(days=3)
    print(day_3back)
    change_date(day_3back)
else:
    yesterday = date.today() - timedelta(days=1)
    #print(yesterday.strftime('%d-%m-%Y'))
    change_date(yesterday)

WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@class = 'buttonR' and @value = 'S']"))).click()
WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@class = 'buttonR' and @value = 'E']"))).click()

time.sleep(2)
transactions = pd.read_csv("C:/Users/"+ str(os.getlogin()) +"/Downloads/transactions.csv")

#transactions1 = transactions1[~transactions1['Message'].str.contains('S...')]
transactions = transactions[transactions['Message'].str.contains('... - 1.0')]
trans_id = transactions['Txxx ID'].tolist()

#print(trans_id)
driver.back()
driver.switch_to.frame(iframe)

directory = 'C:/Users/'+ str(os.getlogin()) +'/Downloads/'
#latest_file = max(glob.glob(directory + 'transactions (*).csv'), key = os.path.getctime)

newpath = 'C:/Errors_'+ str(datetime.now().strftime("%a"))
if not os.path.exists(newpath):
    os.makedirs(newpath)

elif os.path.exists(newpath):
    for filename in os.listdir(newpath):
        if os.path.isfile(os.path.join(newpath, filename)):
            os.remove(os.path.join(newpath, filename))
#i = 1
for i in range(len(trans_id)):
    print(trans_id[i], i)

    tr_id = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name = '...Id']")))
    tr_id.send_keys(trans_id[i])
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@value = 'S']"))).click()
    #wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value = 'Ex']"))).click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '...Txx')]"))).click()
    time.sleep(3)

    table_ev = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@class='...']/tbdy/tr")))
    print(len(table_ev))
    #last_event = driver.find_element_by_xpath("//tr[" + str(len(table_events)) + "]//a")
    last_event = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//tr[" + str(len(table_ev)) + "]//a")))

    print(last_event.text)
    last_event.click()
    
    time.sleep(1)

    ##!!wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr[15]/td/a[contains(@href, '...Transaction')]"))).click()
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//td[3]/a[2][contains(@href, '...Txx')]"))).click()
    
    except TimeoutException as e:
        #driver.find_element_by_xpath()
        print("Status ERROR")
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//a[1][contains(@href, '...Txx')]"))).click()
        except NoSuchElementException as e1:
            print("-----> This is a one-time error with ID: ", trans_id[i])

    except StaleElementReferenceException as e2:
        time.sleep(2)
        pass
        
    time.sleep(3)

    driver.back()
    driver.back()
    driver.back()
    time.sleep(2)

    files = glob.glob('C:/Users/'+ str(os.getlogin()) +'/Downloads/*.*')

    for file in files:
        def move():
            try:
                shutil.move(new_name, newpath + '/' + str(trans_id[i]) + '.txt')
                #shutil.move(new_name1, newpath)
            except Exception as e:
                pass
        
        if file.endswith('dOr.err'):
            new_name = f'C:/Users/'+ str(os.getlogin()) +'/Downloads/'+ str(trans_id[i]) + '.txt'
            # Rename the file using os.rename()
            os.rename(file, new_name)
            file = open(new_name, 'r')
            f = file.read()

            uniqueWords=set(f.split())
        
            if 'Promise' in uniqueWords:
                move()
            else:
                pass
        #elif file.__contains__('.file'):
        elif file.endswith('.file'):
            print("TEST", file)
            new_name = f'C:/Users/'+ str(os.getlogin()) +'/Downloads/'+ str(trans_id[i]) + '.txt'
            print('extension .file has id: ', trans_id[i])
            # Rename the file using os.rename()
            #time.sleep(3)
            try:
                os.rename(file, new_name)
            except PermissionError as e:
                time.sleep(3)
                os.rename(file, new_name)
            except FileNotFoundError as e:
                os.rename(file, new_name)
            file = open(new_name, 'r')
            f = file.read()

            uniqueWords=set(f.split())
            #print(uniqueWords)
        
            if 'Header' in uniqueWords:
                move()
            else:
                pass
        else:
            pass

    driver.switch_to.frame(iframe)
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name = '...Id']"))).clear()
    i += 1

print("Download is complete!")
