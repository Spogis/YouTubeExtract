# -*- coding: utf-8 -*-
"""
Created on Fri Apr  9 15:06:32 2021

@author: Nicolas Spogis
"""

# importar os pacotes necessários
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl import Workbook
import os
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
from selenium import webdriver
import time

global MinChar
global chrome_path
MinChar = 2
global MaxHeight
MaxHeight = 10

chrome_path = r'C:/bin/chromedriver.exe'

def GetURLsFromYoutube(Busca):
    # Open The Output Excel
    global chrome_path
    if os.path.exists("./SiteLists/URL_List.xlsx"):
        os.remove("./SiteLists/URL_List.xlsx")
    
    wb = Workbook()
    wb.save(filename = './SiteLists/URL_List.xlsx')
    workbook = load_workbook(filename="./SiteLists/URL_List.xlsx")
    sheet = workbook.active
    
    for url in Busca:
        # Start the driver
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--mute-audio")
        driver = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
    
        driver.get(url)
        time.sleep(5)
        height = driver.execute_script("return document.documentElement.scrollHeight")
        lastheight = 0
        
        countheight=0
        while True:
            countheight = countheight + 1
            if countheight == MaxHeight:
                break
            if lastheight == height:
                break
            lastheight = height
            driver.execute_script("window.scrollTo(0, " + str(height) + ");")
            time.sleep(2)
            height = driver.execute_script("return document.documentElement.scrollHeight")
        
        user_data = driver.find_elements_by_xpath('//*[@id="video-title"]')
        
        for i in user_data:
            TempList = (i.get_attribute('href'))
            rows = sheet.max_row
            sheet.cell(row=rows+1, column=1).value = TempList
        
        driver.quit()
    
    workbook.save("./SiteLists/URL_List.xlsx")
    workbook.close()
          
def getSiteList(open_file):
    global SitesURLs
    open_file="./SiteLists/"+open_file
    workbook = load_workbook(open_file)
    sheet = workbook.active
    SitesURLs =[]

    for cell in sheet['A']:
        if cell.value is not None:
            SitesURLs.append(cell.value)
            
    workbook.close()

def getTags():
    global tags
    # Open The Output Excel
    if os.path.exists("HData.xlsx"):
        os.remove("HData.xlsx")

    wb = Workbook()
    wb.save(filename = 'HData.xlsx')
    workbook = load_workbook(filename="HData.xlsx")
    sheet = workbook.active

    for url in SitesURLs:
        request = requests.get(url)
        html = BeautifulSoup(request.content, "html.parser")
        tags = html.find_all("meta",property="og:video:tag")
    
        for tag in tags:
            rows = sheet.max_row
            if len(tag['content'])>MinChar: 
                sheet.cell(row=rows+1, column=1).value = tag['content']
            
    workbook.save("HData.xlsx")
    workbook.close()

def getOtherData():
    global tags
    global title
    global UrlData
    global description
    
    # Open The Output Excel
    if os.path.exists("VideoData.xlsx"):
        os.remove("VideoData.xlsx")

    wb = Workbook()
    wb.save(filename = 'VideoData.xlsx')
    workbook = load_workbook(filename="VideoData.xlsx")
    sheet = workbook.active

    for url in SitesURLs:
        request = requests.get(url)
        html = BeautifulSoup(request.content, "html.parser")
        title = html.find("meta", property="og:title")
        UrlData = html.find("meta", property="og:url")
        description = html.find("meta",property="og:description")
            
        rows = sheet.max_row
        sheet.cell(row=rows+1, column=1).value = title.get('content')
        sheet.cell(row=rows+1, column=2).value = UrlData.get('content')
        sheet.cell(row=rows+1, column=3).value = description.get('content')
            
    workbook.save("VideoData.xlsx")
    workbook.close()

def GenerateWordCloud():
    
    # lista de stopword
    STOPWORDS_DATA = []
    workbook = load_workbook(filename="./Others/STOPWORDS.xlsx")
    sheet = workbook.active
    for cell in sheet['A']:
        STOPWORDS_DATA.append(cell.value)
        
    workbook.close()
    
    stopwords = set(STOPWORDS)
    stopwords.update(STOPWORDS_DATA)

    # Start by opening the spreadsheet and selecting the main sheet
    workbook = load_workbook(filename="HData.xlsx")
    sheet = workbook.active
    summary =[]
    
    for cell in sheet['A']:
        summary.append(cell.value)
    
    # concatenar as palavras
    all_summary = ' '.join([str(elem) for elem in summary])
    
    # gerar uma wordcloud
    wc = WordCloud(stopwords=stopwords,
                          width=1520, height=535, max_font_size=100)
    
    # generate word cloud
    wc.generate(all_summary)
     
    # mostrar a imagem final
    fig, ax = plt.subplots(figsize=(100,100))
    ax.imshow(wc, interpolation='bilinear')
    ax.set_axis_off()
     
    plt.imshow(wc, interpolation='bilinear');
    
    # Determine incremented filename
    a= getNextFilePath("./Pictures")
    filename = "./Pictures/WordCloud_" +str(a) + ".png"
    wc.to_file(filename)

def getNextFilePath(output_folder):
    highest_num = 0
    for f in os.listdir(output_folder):
        highest_num = highest_num + 1

    output_file = str(highest_num+1)
    return output_file

def RunAll():
    GetURLsFromYoutube(BuscaSites)
    getSiteList("URL_List.xlsx")
    getTags()
    GenerateWordCloud()


#BuscaSites = ["https://www.youtube.com/results?search_query=ASPEN+Plus&sp=CAMSBAgFEAE%253D",
#              "https://www.youtube.com/results?search_query=DWSIM&sp=CAMSBAgFEAE%253Dhttps://www.youtube.com/results?search_query=DWSIM&sp=CAMSBAgFEAE%253D",
#              "https://www.youtube.com/results?search_query=Hysys&sp=CAMSBAgFEAE%253D"] 
#RunAll()

BuscaSites = ["https://www.youtube.com/results?search_query=OBS&sp=EgIQAQ%253D%253D"]
RunAll()