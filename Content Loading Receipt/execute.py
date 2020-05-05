# -*- coding: utf-8 -*-
"""
Created on Tue Mar 24 16:33:51 2020

@author: amoynihan
"""
import pandas as pd 
import xlsxwriter as excel

import excel as xl

#nameOfReport = 'new.xlsx'
#exDF = pd.read_excel(nameOfReport, 'Detail')
#
### Dropping all empty columns
#exDF = exDF[exDF['Phase'].notna()]
#
#
#
#workbook = excel.Workbook('Adam.xlsx', {'strings_to_urls':False, 'nan_inf_to_errors':True})
##detail = workbook.add_worksheet('Detail')
#summary = workbook.add_worksheet('Summary')
#
#
##xl.formatExcel(exDF, detail, workbook)
#xl.formatSum(exDF, summary, workbook, 'Adam\'s')
##print(xl.find_complete(exDF, 'Documents'))
##print(exDF.columns)

if __name__ == '__main__':
    nameOfLoad = input('What is the Name of the content Loading receipt?')
    exDF = pd.read_excel(nameOfLoad, 'Detail')
    exDF = exDF[exDF['Phase'].notna()]
    nameOfSummary = input('What do you want the name of the Summary to be?')
    workbook = excel.Workbook(nameOfSummary + '.xlsx', {'strings_to_urls':False, 'nan_inf_to_errors':True})
    summary = workbook.add_worksheet('Summary')
    nameOfClient = input ('What is the name of the client?' )
    xl.formatSum(exDF, summary, workbook, nameOfClient)
    workbook.close()

