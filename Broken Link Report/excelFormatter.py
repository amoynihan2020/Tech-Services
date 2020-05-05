# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 16:28:21 2020

@author: amoynihan
"""

import xlsxwriter as excel
import numpy as np
confirmURL = 'Please confirm this URL is functional. If not, please provide updated URL.'
notFunctional = 'URL not functional. Please provide updated web address [External links only]'
notLoaded = 'Content not yet loaded - link will be updated in next QA phase'
orgo = 'This item was organically referenced in your content, but we do not have it in our inventory. Please send us this content for upload, or let us know if you would prefer the item stay unlinked in your content'
sendUsContent = '''We do not have this content item in our inventory. Please send it to us for upload, or confirm the current URL is functional '''

def formatBLD(dframe,sheetName, workbook):
   
    
    sheetName.write('Z21', confirmURL)
    sheetName.write('Z22', notFunctional)
    sheetName.write('Z23', notLoaded)
    sheetName.write('Z24', orgo)
    sheetName.write('Z25', sendUsContent)
    sheetName.set_column('Z:Z', None, None, {'hidden':True})
    
    sheetName.set_column('A:C', 28.07)
    sheetName.set_column('D:D', 16.61)
    sheetName.set_column('E:E', 28.07)
    sheetName.set_column('F:F', 10.83)
    sheetName.set_column('G:G',28.07)
    sheetName.set_column('H:H',8.67) 
    sheetName.set_column('I:I',33.43)
    sheetName.set_column('J:J',16.61)
    sheetName.set_column('K:K', 24.46)
    sheetName.set_column('L:L', 20.74)    
    dframe = dframe.fillna(0)
    title = workbook.add_format({'bold':True})
    title.set_font_color('#ffffff')
    title.set_bg_color('#263441')
    sheetName.write(0, 0, 'Source', title)
    sheetName.write(0,1, 'Parent Procedure ID', title)
    sheetName.write(0,2, 'Procedure Title', title)
    sheetName.write(0,3, 'Category', title)
    sheetName.write(0,4, 'System ID', title)
    sheetName.write(0,5, 'Content Type', title)
    sheetName.write(0,6, 'Title', title)
    sheetName.write(0,7, 'Published', title)
    sheetName.write(0,8, 'BrokenLink', title)
    sheetName.write(0,9, 'Linked Text', title)
    sheetName.write(0,10, 'Request', title)
    sheetName.write(0,11, 'Customer Response', title)              
    
    blue = workbook.add_format()

    blue.set_bg_color('#d4dee8')
    blue.set_text_wrap()
    blue.set_border(1)
    sheetName.set_header('Broken Link Report Detail')
    
    white = workbook.add_format()
    white.set_bg_color('#ffffff')
    white.set_text_wrap()
    white.set_border(1)
    
    format2 = workbook.add_format({'bg_color': '#C6EFCE'})
    
    row = 0
    col = 0
    for  index, rows in dframe.iterrows():
        if (row+1) % 2 == 1:
            sheetName.write(row+1, 0, rows['Source'],blue)
            sheetName.write(row+1, col+1, rows['Parent Procedure ID'], blue)
            sheetName.write(row+1, col+2, rows['Procedure Title'], blue)
            sheetName.write(row+1 , col+3, rows['Category'],blue)
            sheetName.write(row+1, col+4, rows['System ID'],blue)
            sheetName.write(row+1, col+5, rows['Content Type'],blue)
            sheetName.write(row+1, col+6, rows['Title'],blue)
            sheetName.write(row+1, col+7, rows['Published'],blue)
            sheetName.write(row+1, col+8, rows['BrokenLink'],blue)
            sheetName.write(row+1, col+9, rows['Linked Text'],blue)
            sheetName.write(row+1, col+10, rows['Request'],blue)
            sheetName.data_validation('K'+ str(row+1), {'validate':'list',
                                     'source': '=$Z$21:$Z$25'})
            sheetName.write(row+1, col+10, rows['Customer Response'],blue)
            sheetName.write(row+1, col+11, '', blue)
            row+=1
        else:
            sheetName.write(row+1, 0, rows['Source'], white)
            sheetName.write(row+1, col+1, rows['Parent Procedure ID'],white)
            sheetName.write(row+1, col+2, rows['Procedure Title'],white)
            sheetName.write(row+1 , col+3, rows['Category'], white)
            sheetName.write(row+1, col+4, rows['System ID'], white)
            sheetName.write(row+1, col+5, rows['Content Type'], white)
            sheetName.write(row+1, col+6, rows['Title'], white)
            sheetName.write(row+1, col+7, rows['Published'], white)
            sheetName.write(row+1, col+8, rows['BrokenLink'], white)
            sheetName.write(row+1, col+9, rows['Linked Text'], white)
            sheetName.write(row+1, col+10, rows['Request'], white)
            sheetName.data_validation('K'+ str(row+1), {'validate':'list',
                                     'source': '=$Z$21:$Z$25'})
            
            sheetName.write(row+1, col+10, rows['Customer Response'], white)
            sheetName.write(row+1, col+11, '', white)
            row+=1
    sheetName.conditional_format('H1:H80', {'type': 'cell',
                                            'criteria':'=',
                                            'value': 'TRUE',
                                            'format': format2})   
                           
def formatSum(worksheetName, workbook):
     sumFormat = workbook.add_format()
     sumFormat.set_border(1)
     sumFormat.set_bg_color('#a5a5a5')
     sumFormat.set_align('center')
     sumFormat.set_font_color('#FFFFFF')
                              
     bold = workbook.add_format({'bold':True})
     blank = workbook.add_format()
     blank.set_border(1)
     blank.set_text_wrap()
     bold.set_border(1)
     bold.set_text_wrap()
     
     worksheetName.write('A1', 'Total Broken Links Identified', sumFormat)
     worksheetName.set_column(0, 0, 47.57)
     
     worksheetName.write('A2', 'Total Items Containing Broken Links', sumFormat)
     
     worksheetName.write('A3', 'Unique Broken Links', sumFormat)
     
     worksheetName.write('A6', 'Broken Link Type', sumFormat)
     worksheetName.write('B6', 'Cause', sumFormat)
     worksheetName.write('C6', 'Total', sumFormat)
     worksheetName.write('A7', 'Link Opportunity Identified', bold)
     worksheetName.write('B7', orgo, blank)
     
     worksheetName.write('A8', 'Content not in SilverCloud Inventory', bold )
     worksheetName.write('B8', sendUsContent, blank)
     worksheetName.write('A9', 'SilverCloud Cannot Validate URL(internal links)', bold )
     worksheetName.write('B9', confirmURL,blank)
     worksheetName.write('A10', 'Broken Link (external links)', bold)
     worksheetName.write('B10', notFunctional, blank)
     worksheetName.write('A11', 'Content Not in Knowledgebase', bold)
     worksheetName.write('B11', notLoaded,blank)
      
      
     worksheetName.write('B1','', blank)
     worksheetName.write('B2', '', blank)
     worksheetName.write('B3', '', blank)
     worksheetName.write('C7', '', blank)
     worksheetName.write('C8', '', blank)
     worksheetName.write('C9', '', blank)
     worksheetName.write('C10', '', blank)
     worksheetName.write('C11', '', blank)
      
      
     worksheetName.set_column(1,1, 34)
     worksheetName.set_column(2,2, 10.29)
     worksheetName.set_row(6, 90)
     worksheetName.set_row(7, 60)
     worksheetName.set_row(8,30)
     worksheetName.set_row(15, 36.75)
     worksheetName.merge_range('A17:C17', 'Please note: SilverCloud does not validate email addresses, and as such none are featured in these reports.', blank)
      
     worksheetName.merge_range('A14:C14', 'Instructions', sumFormat)
     worksheetName.merge_range('A15:C16', 'Please Use the Unique URL Validation tab to confirm the functionality of the URLS provided, or to provide additional instruction to SilverCloud. The Broken Links - Detail tab offers as inventory of all broken links and their location within content(for reference)', blank)
     worksheetName.write_formula('C7', '=COUNTIF(\'Broken Links-Detail\'!K:K,Summary!B7)',blank)
     worksheetName.write_formula('C8', '=COUNTIF(\'Broken Links-Detail\'!K:K,Summary!B8)',blank)
     worksheetName.write_formula('C9', '=COUNTIF(\'Broken Links-Detail\'!K:K,Summary!B9)',blank)
     worksheetName.write_formula('C10', '=COUNTIF(\'Broken Links-Detail\'!K:K,Summary!B10)',blank) 
     worksheetName.write_formula('C11', '=COUNTIF(\'Broken Links-Detail\'!K:K,Summary!B11)',blank)
     worksheetName.write_formula('B1', '=COUNTIF(\'Broken Links-Detail\'!I:I,\"*\")',blank)
     worksheetName.write_formula('B3', '=COUNTIF(\'Unique URL Validation\'!A:A, \"*\")', blank)
def formatUnique(worksheetName, workbook, dataframe):
    worksheetName.set_column('A:A', 42.43)
    worksheetName.set_column('B:B', 67.57)
    worksheetName.set_column('C:C', 50.71)
    
    uniqueTitlesFormat = workbook.add_format({'bold': True})
    uniqueTitlesFormat.set_bg_color('#263441')
    uniqueTitlesFormat.set_font_color('#FFFFFF')
                                      
    worksheetName.write('A1', 'Broken Link', uniqueTitlesFormat)
    worksheetName.write('B1', 'Request', uniqueTitlesFormat)
    worksheetName.write('C1', 'Customer Response', uniqueTitlesFormat)
    
    blue = workbook.add_format()
    blue.set_bg_color('#d4dee8')
    blue.set_text_wrap()
    blue.set_border(1)
    
    white = workbook.add_format()
    white.set_bg_color('#ffffff')
    white.set_text_wrap()
    white.set_border(1)
    
    row = 0
    vlook = 1
    for item in dataframe:
        if (row+1) % 2 == 1:
            worksheetName.write(row+1, 0, str(item), blue)
            worksheetName.write_formula('B' + str(vlook+1),'=VLOOKUP(A%s,\'Broken Links-Detail\'!I1:K80, 3, FALSE)' %(str(vlook+1)),blue)
            worksheetName.write(row+1, 2, ' ' , blue)
            row+=1
            vlook+=1
        else:
            worksheetName.write(row+1, 0, str(item), white) 
            worksheetName.write_formula('B' + str(vlook+1),'=VLOOKUP(A%s,\'Broken Links-Detail\'!I1:K80, 3, FALSE)' %(str(vlook+1)),white)
            worksheetName.write(row+1, 2, ' ', white)
            row+=1
            vlook+=1
       
     