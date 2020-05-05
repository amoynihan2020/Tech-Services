# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 13:28:03 2020

@author: amoynihan
"""

import pandas as pd 
import xlsxwriter as excel
import re
import excelFormatter 

#this creates a dataframe from the the original broken link report which gives us raw data
def getDF (report):
    masterDF=pd.DataFrame(columns=None)
    xls = pd.ExcelFile(report)
    exDF = pd.read_excel(xls, 'External Links')
    scDF = pd.read_excel(xls, 'SilverCloud Links')
    intraDF = pd.read_excel(xls, 'Internal Links')
    finalDF = pd.DataFrame(columns=None)
    finalDF = finalDF.append(exDF)
    finalDF = finalDF.append(scDF)
    finalDF = finalDF.append(intraDF)
    count = 0
    for indexs, row in finalDF.iterrows():
        if '@' in row['Link']:
            count+=1
            finalDF.drop(finalDF.index[indexs],inplace = True)
            finalDF.reset_index(drop=True, inplace = True)
            masterDF = finalDF
    return masterDF

#print(masterDF.columns)
def findSource (pID, csvName):
   
    csv = pd.read_csv(csvName)
    
    for index, rows in csv.iterrows():
       
        if rows['id'] == pID:
          
            return(rows['source'])
        
    
def formatDF(frame, csv_of_Source)  : 

    cols = ['Source', 'Parent Procedure ID', 'Procedure Title', 'Category', 'System ID','Content Type', 'Title', 'Published', 'BrokenLink', 'Linked Text', 'Request', 'Customer Response']
    thisDF = pd.DataFrame(columns = cols)
    for index, rows in frame.iterrows():
        if rows['Content Type'] == 'step':
            thisDF  = thisDF.append({
               'Source' : findSource(rows['Procedure ID'], csv_of_Source),
               'Parent Procedure ID' : frame.loc[index]['Procedure ID'],
               'Procedure Title' : frame.loc[index]['Parent Title'],
               'Category' : frame.loc[index]['Category'],
               'System ID' : frame.loc[index]['System ID'],
               'Content Type' : frame.loc[index]['Content Type'],
               'Title' : frame.loc[index]['Content Title'],
               'Published' : frame.loc[index]['published'],
               'BrokenLink' : frame.loc[index]['Link'],
               'Linked Text' : frame.loc[index]['Linked Text']
            
            
            
            },ignore_index = True)
 
        else:
            thisDF  = thisDF.append({
               'Source' : findSource(rows['System ID'], csv_of_Source),
               'Parent Procedure ID' : frame.loc[index]['Procedure ID'],
               'Procedure Title' : frame.loc[index]['Parent Title'],
               'Category' : frame.loc[index]['Category'],
               'System ID' : frame.loc[index]['System ID'],
               'Content Type' : frame.loc[index]['Content Type'],
               'Title' : frame.loc[index]['Content Title'],
               'Published' : frame.loc[index]['published'],
               'BrokenLink' : frame.loc[index]['Link'],
               'Linked Text' : frame.loc[index]['Linked Text']
            
            
            
            },ignore_index = True)
    return thisDF
    
def doExcel(theDF, name_of_final):
    uniqueLinks = theDF.BrokenLink.unique()



    workbook = excel.Workbook(name_of_final +'.xlsx', {'strings_to_urls':False})
    bld = workbook.add_worksheet('Broken Links-Detail')
    unique = workbook.add_worksheet('Unique URL Validation')
    summary = workbook.add_worksheet('Summary')
    excelFormatter.formatBLD(theDF, bld, workbook)
    excelFormatter.formatUnique(unique, workbook, uniqueLinks)
    excelFormatter.formatSum(summary,workbook)
    workbook.close()
    
def mainApp(rawReport, sourceReport, finalName):
    print('Starting Report')
    raw = getDF(rawReport)
    formattedDF = formatDF(raw,sourceReport)

    doExcel(formattedDF, finalName)
    print('All Done')
if __name__ == "__main__":
    raw = input('What Is the Name of the raw Report?')
    source = input('What is the name of the source report?')
    final = input('What do you want the name of the final report to be?')
    mainApp(raw, source, final)
    

    
    
    
    
    



