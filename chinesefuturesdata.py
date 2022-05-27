import webbrowser
import time
import os
import shutil
import requests
import xlsxwriter
import openpyxl
import xlwings as xw
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, GradientFill
from xlrd import open_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from PIL import ImageGrab
import pandas as pd
from datetime import datetime

#month_dictionary = {'Number':['01','02','03','04','05','06','07','08','09','10','11','12'], 'Month':['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']}
#month_dictionary = dict(zip(month_dictionary['Number'], month_dictionary['Month']))

x = datetime.now()
print(x.strftime("%d%b%y"))
day = x.strftime("%d")
month = x.strftime("%b")
year = x.strftime("%y")
hour = x.strftime ("%I")
AMPM = x.strftime ("%p")

op = webdriver.ChromeOptions()
p = {'download.default_directory': r'C:\Users\asus\Desktop\KCH\Automation\ChineseFutures'}
op.add_experimental_option('prefs', p)

driver = webdriver.Chrome(executable_path = 'C:\\chromedriver.exe')
driver.maximize_window()
driver.get("https://www.cngold.org/qihuo/donglimei.html")


# we put time.sleep to wait the content to appear first, if content has not appeared but we already execute command to find the element location by xpath, it will not work properly
time.sleep(5)

def getimage():
    driver.execute_script("window.scrollTo(0, 270)")

    imagename = 'Chinese Futures %s' %(x.strftime("%d%b%y %I%p")) + '.png'
    pathimage= r'C:\Users\asus\Desktop\KCH\Automation\ChineseFutures\%s' %imagename
    time.sleep(2)
    image = ImageGrab.grab(bbox=(300,200,1300,1000))
    image.save(pathimage)
    print("Successfully saved image under this name: ", imagename)

def extractvalue():
    # find the element of index value from inspect
    index = driver.find_element_by_xpath("//b[@id = 'now_price']")
    time.sleep(1)
    
    
    if (index.text) == "-":
        chinesefutures = "-"
    else:
        chinesefutures = float(index.text)
        
    print ("chinesefutures:", chinesefutures)
    # storing the index in list
    indexlist = []
    indexlist.append(x.strftime("%d %b"))
    indexlist.append(x.strftime("%I %p"))
    indexlist.append(chinesefutures)
    print (indexlist)

    df = pd.DataFrame(np.array([indexlist]),columns=['Date','Time/Session','Index'])
    df['Index'] = df['Index'].astype('float32').apply(lambda x: round(x,2))
    print(df)
    Filename = 'ChineseFutures' + x.strftime("%d%m%y %I%p") + '.xlsx'
    df.to_excel(r'C:\Users\asus\Desktop\KCH\Automation\ChineseFutures\%s' %Filename, sheet_name = "Chinese Futures", index = False)
    driver.quit()

getimage()
extractvalue()