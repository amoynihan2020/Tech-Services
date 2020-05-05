# -*- coding: utf-8 -*-
"""
Created on Tue Mar 10 15:52:07 2020

@author: amoynihan
"""
import pandas as pd
date_of_folder = ''
def find_complete(df, category):
    
    
    counter=0
    for index, row in df.iterrows():
        if row['Location/Category'] == category and row['Status'] == 'Uploaded':
            counter+=1
    return counter
def find_latestDate(df, category):
    
    newDf = df.loc[df['Location/Category']==category]
    newDf['Est. Date of Completion'] = pd.to_datetime(newDf['Est. Date of Completion'])
    return newDf['Est. Date of Completion'].max()    
def find_keyNumber(df, key):
    counter = 0
    for index, row in df.iterrows():
        if row['Status'] == key:
            counter +=1
    return counter
        
    
    

#
    
def formatSum(df, sheetname, workbook, clientName):
    pend = 'Line item has not yet been uploaded.'
    inP = 'Line item is in-progress, but not fully loaded.'
    upload = 'Line item has succesfully been uploaded into the system'
    dupey = 'Line item was an exact duplicate with another sent us, and has not been loaded twice'
    awaiting = 'Line item will require a response from customer before it is uploaded' 
    
    
    topRow = workbook.add_format({'bold':True, 'align':'center'})
    topRow.set_bg_color('#5b9ad5')
    topRow.set_font_color('#FFFFFF')
    topRow.set_border(1)
  
    total = workbook.add_format({'bold':True})   
    totalDate = workbook.add_format({'bold':True, 'num_format': 'yyyy/mm/dd'})
    totalDate.set_border(1)
    total.set_border(1)                       
    blueDate = workbook.add_format({'num_format':'yyyy/mm/dd'})
    blueDate.set_bg_color('#d4dee8')
    blueDate.set_text_wrap()
    blueDate.set_border(1)
    
  
    
    whiteDate = workbook.add_format({'num_format':'yyyy/mm/dd'})
    whiteDate.set_bg_color('#ffffff')
    whiteDate.set_text_wrap()
    whiteDate.set_border()
    
   
    blue = workbook.add_format()

    blue.set_bg_color('#d4dee8')
    blue.set_text_wrap()
    blue.set_border(1)
    
    white = workbook.add_format()
    white.set_bg_color('#ffffff')
    white.set_text_wrap()
    white.set_border(1)
    
    blueDate = workbook.add_format({'num_format':'yyyy/mm/dd'})
    blueDate.set_bg_color('#d4dee8')
    blueDate.set_text_wrap()
    blueDate.set_border(1)
    
    whiteDate = workbook.add_format({'num_format':'yyyy/mm/dd'})
    whiteDate.set_bg_color('#ffffff')
    whiteDate.set_text_wrap()
    whiteDate.set_border()
    
    sheetname.write(0, 0, 'Location/Category', topRow)
    sheetname.write(0,1, 'Content Total' , topRow)
    sheetname.write(0,2, 'Complete', topRow)
    sheetname.write(0,3, 'Remaining', topRow)
    sheetname.write(0,4, 'Est. Date Of Completion', topRow)
    sheetname.write(0,6, 'Key', topRow)
    sheetname.write(0,7, 'Definition', topRow)
    sheetname.write(0,8, 'Count' , topRow)
    
    sheetname.set_column('A:A', 40.71)
    sheetname.set_column('B:D', 21.86)
    sheetname.set_column('E:E', 26.43)
    sheetname.set_column('G:G', 28.57)
    sheetname.set_column('H:H', 34.86)
    
    category = df['Location/Category'].unique()
    catDF = pd.DataFrame(category, columns = ['category'])
    rows = 0
    myLen = len(catDF) +1
    categories = df.pivot_table(index=['Location/Category'], aggfunc = 'size')
    
    
   # print(catDF)
   
    for index, row in catDF.iterrows():
        
       
        if rows % 2 == 0:
            sheetname.write(rows+1, 0, row['category'],blue)
            sheetname.write(rows+1, 1, categories[row['category']], blue)
            sheetname.write(rows+1, 2, find_complete(df, row['category']),blue )
            sheetname.write(rows+1, 3, '=B%s-C%s' %(rows+2, rows+2),blue)
            sheetname.write(rows+1, 4, find_latestDate(df, row['category']),blueDate)
            rows+=1
        else:
            sheetname.write(rows+1, 0, row['category'],white)
            sheetname.write(rows+1, 1, categories[row['category']], white)
            sheetname.write(rows+1, 2, find_complete(df, row['category']),white )
            sheetname.write(rows+1, 3, '=B%s-C%s' %(rows+2, rows+2),white)
            sheetname.write(rows+1, 4, find_latestDate(df, row['category']),whiteDate)
            rows+=1
   
        
           
    
    sheetname.write(myLen, 0, 'Totals', total)
    sheetname.write(myLen, 1, '=SUM(B2:B%s)' % (myLen), total)
    sheetname.write(myLen, 2, '=SUM(C2:C%s)' %(myLen), total)
    sheetname.write(myLen, 3, '=SUM(D2:D%s)' %(myLen),total)
    sheetname.write(myLen, 4, '=MAX(E2:D%s)' %(myLen), totalDate)
    
    
    
    sheetname.write('G2' , 'Pending', blue)
    sheetname.write('G3', 'In-Progress', white)
    sheetname.write('G4', 'Uploaded', blue)
    sheetname.write('G5', 'Duplicate - Double-Categorized',white)
    sheetname.write('G6', 'Awaiting Question Response', blue)
    sheetname.write('H2', pend, blue)
    sheetname.write('H3', inP, white)
    sheetname.write('H4', upload, blue)
    sheetname.write('H5', dupey, white)
    sheetname.write('H6' , awaiting, blue)
    
    sheetname.write('I2', find_keyNumber(df, 'Pending'),blue)
    sheetname.write('I3', find_keyNumber(df, 'In-Progress'), white)
    sheetname.write('I4', find_keyNumber(df, 'Uploaded'),blue)
    sheetname.write('I5', find_keyNumber(df ,'Duplicate - Double-Categorized'),white)    
    sheetname.write('I6', find_keyNumber(df ,'Awaiting Question Response'),blue)
    chart2 = workbook.add_chart({'type': 'column'})
    chart2.set_title({'name': '%s  - Content Loading' %(clientName) })
    chart2.set_size({'width':720, 'height' :576})
    
    

# Configure the first series.
    chart2.add_series({
    'name':       '=Summary!$B$1',
    'categories': '=Summary!$A$2:$A$10',
    'values':     '=Summary!$B$2:$B$10',
    'fill' :{'color': '#ff9900'}
    })

# Configure second series.
    chart2.add_series({
    'name':       '=Summary!$C$1',
    'categories': '=Summary!$A$2:$A$7',
    'values':     '=Summary!$C$2:$C$7',
    'fill':{'color': '#0070c0'}
    })
    
    sheetname.insert_chart('A14', chart2)
    
    