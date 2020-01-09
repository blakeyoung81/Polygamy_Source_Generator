from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from lxml import html
import time
import sys
import requests
from bs4 import BeautifulSoup
from contextlib import suppress
import pickle
from selenium.common.exceptions import NoSuchElementException
import xlrd
i = 1

def get_url(i):
    loc = ("G:\My Drive\Blake Young\Coding\Python\Projects\Polygamy\learningsuite.byu.edu_6th_Aug_2019.xlsx") 
    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 
    # For row 0 and column 0
    
    url = sheet.cell_value(i-1, 0)
    #print(url)
    name = sheet.cell_value(i +35, 0)
    print(name, "", "")
    chrome.get(url)
    site = chrome.get(url)
    #print("Getting the page")
    time.sleep(2)
    get_info()
    
def get_info():
    global i
    body=chrome.find_element_by_tag_name("body")
    body = body.text

    #Marriage Date
    print("1: ", end =" ")
    try:
        print(body[body.index("born") -10:body.index("born") + 200])
    except:
        a = 1
    try:
        print(body[body.index("Born") -200:body.index("Born") + 200])
    except:
        a = 1

    #Relationship
    print("2: ", end =" ")
    try:
        print(body[body.index("met") -200:body.index("met") + 200])
    except:
        a = 1
    try:
        print(body[body.index("Met") -200:body.index("Met") + 200])
    except:
        a = 1
    try:
        print(body[body.index("Joseph") -200:body.index("Joseph") + 200])
    except:
        a = 1
    try:
        print(body[body.index("relationship") -200:body.index("relationship") + 200])
    except:
        a = 1
    try:
        print(body[body.index("first") -200:body.index("first") + 200])
    except:
        a = 1    
    
       
    #3: Marriage ceremony / attendees / etc.
    print("3: ", end =" ")
    try:
        print(body[body.index("seal") -200:body.index("seal") + 200])
    except:
        a = 1
    try:
        print(body[body.index("attend") -200:body.index("attend") + 200])
    except:
        a = 1
    try:
        print(body[body.index("married") -200:body.index("married") + 200])
    except:
        a = 1
    try:
        print(body[body.index("marry") -200:body.index("marry") + 200])
    except:
        a = 1
    try:
        print(body[body.index("ceremony") -200:body.index("ceremony") + 200])
    except:
        a = 1
    #4: Conjugal
        print("4: ", end =" ")
    try:
        print(body[body.index("consumate") -200:body.index("consumate") + 200])
    except:
        a = 1
    try:
        print(body[body.index("evidence") -200:body.index("evidence") + 200])
    except:
        a = 1
    try:
        print(body[body.index("eternity-only") -200:body.index("eternity-only") + 200])
    except:
        a = 1
    try:
        print(body[body.index("physical") -200:body.index("physical") + 200])
    except:
        a = 1
    try:
        print(body[body.index("conjugal") -200:body.index("conjugal") + 200])
    except:
        a = 1
    try:
        print(body[body.index("sex") -200:body.index("sex") + 200])
    except:
        a = 1
    try:
        print(body[body.index("affair") -200:body.index("affair") + 200])
    except:
        a = 1
    try:
        print(body[body.index("sleep") -200:body.index("sleep") + 200])
    except:
        a = 1
    try:
        print(body[body.index("intimacy") -200:body.index("intimacy") + 200])
    except:
        a = 1
    try:
        print(body[body.index("intercourse") -200:body.index("intercourse") + 200])
    except:
        a = 1
    try:
        print(body[body.index("affection") -200:body.index("affection") + 200])
    except:
        a = 1
    try:
        print(body[body.index("rapport") -200:body.index("rapport") + 200])
    except:
        a = 1
    
    #5: Biographical statements
    print("5: ", end =" ")
    try:
        print(body[body.index("I never") -200:body.index("I never") + 200])
    except:
        a = 1
    try:
        print(body[body.index("I always") -200:body.index("I always") + 200])
    except:
        a = 1
    try:
        print(body[body.index(" I ") -200:body.index(" I ") + 200])
    except:
        a = 1
    print("", "")
    i+=1
    get_url(i)
    chrome.close()

if __name__ == "__main__":
    i = 1
    one = []
    two = []
    three = []   
    four = []
    five = []
    
    chrome = webdriver.Chrome(executable_path=r"C:\Users\blake\Google Drive\Blake Young\Coding\Python\Projects\Dr_Chester\chromedriver.exe")
    get_url(i)
    
    
            
    

