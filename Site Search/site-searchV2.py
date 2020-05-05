# -*- coding: utf-8 -*-
"""
Created on Mon Nov 25 12:30:51 2019

@author: amoynihan

CREATED a new site search that utilizes a dictionary as opposed to a list
This allows for us to remove duplicate links

"""

import requests 
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
from urllib.parse import urlparse



#method created to see if URL is absolute
def is_absolute(url):
    return bool(urlparse(url).netloc)



#creates an excel workbook 
name_of_excel_sheet = input("Enter The name of the Excel Document")

#adding the .xlsx extension
name_of_excel_sheet += ".xlsx"

workbook = xlsxwriter.Workbook(name_of_excel_sheet)

#creates worksheet
worksheet = workbook.add_worksheet("Links")
weird_links_sheet = workbook.add_worksheet('weird or broken Links')

#url we want to scrape
url = input("Enter The URL you want to scrape:") 

#addded so we don't get blocked from sites 
agent= {"User-Agent":'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}

#connecting to URL
response = requests.get(url, headers = agent);


#Parse Html and save in bs object 
soup = BeautifulSoup(response.text, "html.parser")

#makes html readable
soup.prettify()

#finds all the a tags 
acsearch = soup.findAll("a")



#dictionary to collect the urls and there associated text 
url_list = {}

#adding the urls to the above dictionary
for item in acsearch:
    url_list[item.text] = item.get('href')



#list that should get rid of the duplicates
no_dupe_urls ={}

# removing duplicates  from the dictionary 
for key, value in url_list.items():
    if value not in no_dupe_urls.values():
        no_dupe_urls[key] = value
        
   
row = 0
col = 0

row2 = 0;
col2 = 0;     

#making the title cells bold
cell_format = workbook.add_format({'bold': True})

#Creating the  title cells
worksheet.write(0,0, "url", cell_format)
worksheet.write(0, 1, "title", cell_format)
worksheet.write(0, 2, "content_type")
worksheet.write(0,3, "category" )
worksheet.write(0,4,"published")
worksheet.write(0,5,"id")


weird_links_sheet.write(0,0, "url", cell_format)
weird_links_sheet.write(0,1, "title", cell_format)



#itterating through the dictionary
for key, value in no_dupe_urls.items():
    concString = ''
    #checking to see if the url absolute
    if is_absolute(value) == False and value is not None:
        #if url is absolute concatenate with the url given by user
        concString += url + value
        
        #writing to excel
        worksheet.write(row+1,col,concString)
        worksheet.write(row+1, col+1,key)
        worksheet.write(row+1, col+2, 'link')
        worksheet.write(row+1, col+3, 'Web Links')
        worksheet.write(row+1, col+4, 'TRUE')
        
        
    #if a link is none
    elif is_absolute(value) == False and value is None:
        weird_links_sheet.write(row2+1, col2, value)
        weird_links_sheet.write(row2+1, col2+1, key)
        row2+=1
    #if URL is absolute 
    else:
        worksheet.write(row+1,col,value)
        worksheet.write(row+1, col+1, key)
        worksheet.write(row+1, col+2, 'link')
        worksheet.write(row+1, col+3, 'Web Links')
        worksheet.write(row+1, col+4, 'TRUE')
        
        
        
    row+=1

        
worksheet.set_default_row(hide_unused_rows=True)  
weird_links_sheet.set_default_row(hide_unused_rows=True)  
workbook.close()

