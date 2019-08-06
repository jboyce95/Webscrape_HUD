# -*- coding: utf-8 -*-
"""
Created on Wed Jul 10 09:34:21 2019

@author: jboyce
"""

#imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import ui
import pandas as pd
import time
import datetime
import sqlalchemy
import pyodbc
import smtplib #for sending emails
import requests
import urllib.request, urllib.parse, urllib.error
from bs4 import BeautifulSoup
import os 
#import lxmlm.html as lh
 
#set url and project dir
url = 'https://www.federalreserve.gov/datadownload/Choose.aspx?rel=H15'
proj_dir = r'M:\Capital Markets\Users\Johnathan Boyce\Misc\Programming\Python\Scripts'


#set preferences
chromeOptions = webdriver.ChromeOptions()
#need to set chrome default directory to project folder?
#System.setProperty("webdriver.chrome.driver", r'M:\Capital Markets\Users\Johnathan Boyce\Misc\Programming\Python\Scripts\chromedriver.exe')
prefs = {'download.default_directory' : r'M:\Capital Markets\Users\Johnathan Boyce\Misc\Programming\Python\Scripts'} #failed - path too long, tried one \
chromeOptions.add_experimental_option('prefs', prefs)


#browser/driver setup - ! DO NOT EDIT !
driver = webdriver.Chrome(executable_path=r'M:\Capital Markets\Users\Johnathan Boyce\Misc\Programming\Python\Scripts\chromedriver.exe', chrome_options=chromeOptions)


#open Chrome browser and open the HUD interest rate site
driver.get(url)
button = driver.find_element_by_id('FreqRequest_3') #select Monthly Averages
button.click()
button_download = driver.find_element_by_id('btnToDownload') #locate download button
button_download.click()

time.sleep(1)

button_download2 = driver.find_element_by_id('btnDownloadFile') #locate download button (after redirect resulting from clicking 'btnToDownload')
button_download2.click()


#Use pandas to read the csv downloaded from the HUD url
csv = proj_dir + r'\FRB_H15.csv'
df = pd.read_csv(csv, header=0)
#print(df.head())
print(df.columns.names)



#BeautifulSoup grab
html = urllib.request.urlopen(url).read()
soup = BeautifulSoup(html, 'html.parser') #The soup object contains all of the HTML in the original document.


#print(soup.prettify)