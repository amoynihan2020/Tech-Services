# -*- coding: utf-8 -*-
"""
Created on Wed Dec 18 10:18:32 2019

@author: amoynihan


#uses fuzzy wuzzy to compare titles of clients consumer content
Uses two identical dataframes derived from the user submitted all content csv
CSV must contain: 
    Title
    content type 
    category
    id (The Long One)
    published 
If not it will probably break.

The script compares titles via two for loops one, being nested. The first loop grabs the 
title from the first data frame labeled "theList". The second(or Nested) loop then 
iterates through the  the titles from the second dataframe called "theCompareList"
comparing the first title to all the titles in the theCompareList. If the titles have 
a matching percentage greater than or equal to the threshold submitted by the user,
it will be written to the final excel report. If not, it will move onto the next title
Once the first title from "theList" has been compared to all of the titles in "theCompareList" 
it will iterate to the next title and start the process all over again. Once it has been through 
it will return a final report and a message stating that the report is done.

Script should take about 5 min to run

"""

from fuzzywuzzy import fuzz
from fuzzywuzzy import process

import pandas as pd
import xlsxwriter as excel

name_of_csv = input('What is the name of the all content report? no extension, must be csv')
name_of_csv += '.csv'
rat = input("What is the percentage Threshold?")
final_report = input("What do you want the final Report to be named? No Extension")
print('We are getting Started! Should take anywhere from 2-5 minutes to run')
final_report += '.xlsx'


#getting rid of any spaces 
final_name = name_of_csv.strip()

#creating two identical dataframes to compare titles
theList = pd.read_csv(final_name) 
theCompareList = pd.read_csv(final_name)

#creating an excel workbook
workbook = excel.Workbook(final_report)
worksheet = workbook.add_worksheet('Final_Report')


row1 = 0;
col = 0;     

#making the title cells bold
cell_format = workbook.add_format({'bold': True})

#setting title cells to specified color 
cell_format.set_bg_color('#000066')
#setting font color to white
cell_format.set_font_color('#FFFFFF')
#setting font to calibri
cell_format.set_font_name('Calibri')
#making the title row size 30
worksheet.set_row(0, 30)

#alligning the titles
cell_format.set_align('center')
cell_format.set_align('vcenter')

#creating a blue to alternate rows
blue = workbook.add_format()
blue.set_bg_color('#ccecff')
#making a set border
blue.set_border(2)

#creating a border for white cells
white = workbook.add_format()
white.set_border(2)

#creating titles
worksheet.write(0,0, 'Title One', cell_format)
worksheet.write(0,1, 'Content Type', cell_format)
worksheet.write(0,2, 'Category', cell_format)
worksheet.write(0,3, 'ID', cell_format)
worksheet.write(0,4, 'Published', cell_format)
worksheet.write(0,5, "",cell_format)



worksheet.write(0,6, 'Matching Title', cell_format)
worksheet.write(0,7, 'Content Type', cell_format)
worksheet.write(0,8, 'Category', cell_format)
worksheet.write(0,9, 'ID', cell_format)
worksheet.write(0,10, 'Published', cell_format)
worksheet.write(0, 11, 'Matching Percentage', cell_format)

#comparisons start here 
#getting the first title to compare on
for index, rows in theList.iterrows():
    #getting the  second titles from the other df to compare 
    for index1, row in theCompareList.iterrows():
        #comparing the titles and spitting back a ratio
        ratio = fuzz.ratio(str(theList.loc[index]['title']), str(theCompareList.loc[index1]['title']))
        try :
            #if ratio is more than the user specified ratio add to final report 
            if ratio >= int(rat) and ratio < 100:
                masterTitle = str(theList.loc[index]['title'])
                compTitle = str(theCompareList.loc[index1]['title'])
                #if an odd row make it blue
                if ((row1+1) % 2 == 1):
                
                    worksheet.write(row1+1, col, str(masterTitle),blue)
                    worksheet.write(row1+1, col+1, str(theList.loc[index]['content_type']),blue)
                    worksheet.write(row1+1, col+2, str(theList.loc[index]['category']),blue)
                    worksheet.write(row1+1, col+3, str(theList.loc[index]['id']),blue)
                    worksheet.write(row1+1, col+4, str(theList.loc[index]['published']),blue)
                    worksheet.write(row1+1, col+5, '', blue )
                    worksheet.write(row1+1, col+6, str(theCompareList.loc[index1]['title']),blue)
                    worksheet.write(row1+1, col+7, str(theCompareList.loc[index1]['content_type']),blue)
                    worksheet.write(row1+1, col+8, str(theCompareList.loc[index1]['category']),blue)
                    worksheet.write(row1+1, col+9, str(theCompareList.loc[index1]['id']),blue)
                    worksheet.write(row1+1, col+10, str(theCompareList.loc[index1]['published']),blue)
                    worksheet.write(row1+1, col+11, ratio,blue)
                    row1+=1
                    #need to figure out how to get rid of duplicates
                #if even make it white
                else:
                    worksheet.write(row1+1, col, str(masterTitle),white)
                    worksheet.write(row1+1, col+1, str(theList.loc[index]['content_type']),white)
                    worksheet.write(row1+1, col+2, str(theList.loc[index]['category']),white)
                    worksheet.write(row1+1, col+3, str(theList.loc[index]['id']),white)
                    worksheet.write(row1+1, col+4, str(theList.loc[index]['published']),white)
                    worksheet.write(row1+1, col+5, '', white)
                    worksheet.write(row1+1, col+6, str(theCompareList.loc[index1]['title']), white)
                    worksheet.write(row1+1, col+7, str(theCompareList.loc[index1]['content_type']),white)
                    worksheet.write(row1+1, col+8, str(theCompareList.loc[index1]['category']),white)
                    worksheet.write(row1+1, col+9, str(theCompareList.loc[index1]['id']),white)
                    worksheet.write(row1+1, col+10, str(theCompareList.loc[index1]['published']),white)
                    worksheet.write(row1+1, col+11, ratio,white)
                    row1+=1
     
            #catchin errors hopefully no one knows this exists     
        except Exception as e:
            print ("index: " + str(index))
            print("index1: " + str(index1))
            print("Master: " + masterTitle)
            print( type(row))
            print("Comp: " + compTitle)
            print("Ratio "+ str(rat))
            print(str(e))
            print('\n')
            print('If you see this, alert Tech Services, something is up')


#setting final report to landscape
worksheet.set_landscape() 
#displaying in landscape    
worksheet.set_page_view() 
#setting the header in the center, bold, 25 point font
worksheet.set_header('&C&25&"Bold"Consumer Similar Content Report')

#setting the footer on the left Claibri Italic, on the right theres a picture 
worksheet.set_footer('&L&"Calibri Italic" SilverCloud Inc. - Confidential Proprietary &R&G', {'image_right': 'sclogo.png'})

#all done 
print('All Done!, Check Directory for Excel')
workbook.close()

       