# -*- coding: utf-8 -*-
"""
Created on Thu Mar 12 16:17:28 2020

@author: amoynihan
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Dec 18 14:56:53 2019

@author: amoynihan
"""

 
import urllib.request
from bs4 import BeautifulSoup as bs

import xlsxwriter 
import pandas as pd
name_of_json = input("What is the name of the All Content Report JSON?")
masterJson = pd.read_json(name_of_json)

#creates an excel workbook 
name_of_excel_sheet = input("Enter The name of the Excel Document")

#adding the .xlsx extension
name_of_excel_sheet += ".xlsx"

workbook = xlsxwriter.Workbook(name_of_excel_sheet)
#
##creates worksheet
worksheet = workbook.add_worksheet("Links")   
#
body_title = masterJson.filter(['body', 'title', 'id', 'category', 'published'] )


links = []

row = 0
col = 0
##making the title cells bold
cell_format = workbook.add_format({'bold': True})
##Creating the  title cells
worksheet.write(0,0, "URL", cell_format)
worksheet.write(0, 1, "TEXT", cell_format)
worksheet.write(0, 2, "Title", cell_format)
worksheet.write(0, 3, 'ID', cell_format)
worksheet.write(0, 4, 'Category', cell_format)
worksheet.write(0, 5, 'Published', cell_format)

cell_format.set_bg_color('#000066')
#setting font color to white
cell_format.set_font_color('#FFFFFF')
#setting font to calibri
cell_format.set_font_name('Calibri')
#making the title row size 30
worksheet.set_row(0, 30)

cell_format.set_align('center')
cell_format.set_align('vcenter')
    
blue = workbook.add_format()
blue.set_bg_color('#ccecff')
#making a set border
blue.set_border(2)

white = workbook.add_format()
white.set_border(2)
allUrls = []

urls = {}
no_dupe_urls= {}

def remove_dupes(listOfATags):
    rawLinks = []
    noDupeLinks = []
    for item in listOfATags:
        rawLinks.append(item.get('href'))
    for link in rawLinks:
        if link not in noDupeLinks:
            noDupeLinks.append(link)
        else:
            print(str(link) + 'Already in List, Skipping')
    return noDupeLinks
#This gets all the links including duplicates
for index, rows in body_title.iterrows():
    
    soup = bs(str(rows['body']), 'html.parser')
    soup.prettify()
    
    link_search = soup.findAll('a')
    
    
  
    
    
    for item in link_search:
       
       if item.get('href') not in allUrls:
           if(row+1) % 2 == 1:
                urls[item.text] = item.get('href')
                worksheet.write(row+1, col, item.get('href'),blue)
                worksheet.write(row+1, col+1, item.text,blue)
                worksheet.write(row+1, col+2, str(body_title.loc[index]['title']),blue)
                worksheet.write(row+1, col+3, str(body_title.loc[index]['id']),blue)
                worksheet.write(row+1, col+4, str(body_title.loc[index]['category']),blue)
                worksheet.write(row+1, col+5, str(body_title.loc[index]['published']),blue)
                row+=1
           else: 
                worksheet.write(row+1, col, item.get('href'),white)
                worksheet.write(row+1, col+1, item.text,white)
                worksheet.write(row+1, col+2, str(body_title.loc[index]['title']),white)
                worksheet.write(row+1, col+3, str(body_title.loc[index]['id']),white)
                worksheet.write(row+1, col+4, str(body_title.loc[index]['category']),white)
                worksheet.write(row+1, col+5, str(body_title.loc[index]['published']),white)
                row+=1
            #comment this line to get all links, leave uncommented to get unique links
           #allUrls.append(item.get('href'))
       else: 
            print( str(item.get('href')) + ' Already in Report, Skipping')
       

#for index, rows in body_title.iterrows():
#    soup = bs(str(rows['body']), 'html.parser')
#    theLinks = soup.findAll('a')
#    for item in theLinks:
#        allUrls.append(item.get('href'))
#        if item.get('href') not in allUrls:
            
     
#for key,value in urls.items():
#    if key not in no_dupe_urls:
#        no_dupe_urls[key] = value
#        
#
#
#        
#for key,value in no_dupe_urls.items():
#    worksheet.write(row+1, col, value)
#    worksheet.write(row+1, col+1, key)
#    row+=1
        
#setting final report to landscape
worksheet.set_landscape() 
#displaying in landscape    
worksheet.set_page_view() 
#setting the header in the center, bold, 25 point font
worksheet.set_header('&C&25&"Bold" Consumer Link Report')

#setting the footer on the left Claibri Italic, on the right theres a picture 
worksheet.set_footer('&L&"Calibri Italic" SilverCloud Inc. - Confidential Proprietary &R&G', {'image_right': 'sclogo.png'})
print('Report Complete!')
workbook.close()
     

