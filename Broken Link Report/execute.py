# -*- coding: utf-8 -*-
"""
Created on Tue Mar  3 16:55:20 2020

@author: amoynihan
"""

import BrokeEmployee as broke
import v3

if __name__ =="__main__":
    json = input('What is the name of the All Content JSON?')
    rawData = input('What do you want the name of the internal report to be?')
    finalReport = input('What do you want the name of the customer(external) report to be ?')
    sourceReport = input('What is the name of the Source report CSV?')
    
    broke.main_app(json, rawData)
    v3.mainApp(rawData + '.xlsx', sourceReport, finalReport )
    
    