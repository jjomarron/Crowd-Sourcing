## General
import time
import shutil, os
from pathlib import Path
from datetime import date, timedelta
import datetime

path=r'C:\Users\Jack\Documents\Misc\Friend Help\Fantasy Soccer'

## Excel
import pandas as pd
from io import StringIO
import numpy as np
import os
import win32com.client

#Selenium/Firefox
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException

profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2) 
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', r'C:\Users\Jack\Documents\Work\Chicago Stars\Python.Line Up')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')

browser = webdriver.Firefox(profile)
browser.implicitly_wait(10)

df = pd.read_csv(r"C:\Users\Jack\Documents\Misc\Friend Help\Fantasy Soccer\stest.csv",encoding='UTF-8')

df['ID'] = pd.Series(df['ID'], dtype="string")

my_cols = [str(i) for i in range(25)]
df1 = pd.DataFrame()


for i in range(len(df)):
    browser.get('https://www.whoscored.com/Players/'+df['ID'][i])
    browser.find_element_by_css_selector("#sub-navigation > ul:nth-child(1) > li:nth-child(2) > a:nth-child(1)").click()
    #table = browser.find_elements_by_css_selector("#top-player-stats-summary-grid")
    row = browser.find_elements_by_xpath("//table[@class='grid with-centered-columns hover']/tbody/tr[1]")  
    for r in row:
        print(r.text)
        list = [i]
        hold = pd.read_csv(StringIO(r.text),sep=' ',names=my_cols,header=None,index_col=list,lineterminator='\t')
        df1 = df1.append(hold)
        browser.find_element_by_css_selector("#player-matches-stats-options > li:nth-child(2) > a:nth-child(1)").click()
    
df1=df1.reset_index()
df1=df1.drop(['level_0','level_1','level_2'],axis=1)
df1=df1.rename(columns={"2": "Minutes", "3": "Goals", "4": "Assists", "5": "Yellow", "6": "Red", "7": "Shots", "8": "PS%", "9": "Aerials Won", "10": "Rating"})

result = pd.concat([df, df1], axis=1, join="inner")

result.to_csv('initial_stat_check.csv', mode='w',index=False)







