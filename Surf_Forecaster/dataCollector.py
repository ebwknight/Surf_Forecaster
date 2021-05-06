#, attrs = {'class': 'js-forecast-table-content forecast-table__table forecast-table__table--content'}
"""
Created on Sun Apr 11 18:23:11 2021

@author: evanb
"""

import pandas as pd
import numpy as np
from bs4 import BeautifulSoup as soup
from urllib.request import Request, urlopen
import re


def surf_forecast_req():
    
    southernMaine = [
        "Old Orchard Beach", 
        "Pine Point",  
        "Scarborough Beach",
        "Higgins Beach",
        "Popham Read",
        "Doc Browns"]
    
    spotData = {}
    
    for spot in southernMaine:
        try:
            print("Getting data on " + spot)
            # Set up scraper
            surfBreak = re.sub(r' ', '-', spot)
            url = ("https://www.surf-forecast.com/breaks/" + surfBreak + "/forecasts/latest")
            req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            webpage = urlopen(req).read()
            html = soup(webpage, "html.parser")
            # Find fundamentals table
            forecastData = pd.read_html(str(html), attrs = {'class': 'js-forecast-table-content forecast-table__table forecast-table__table--content'})[0] 
            
            #grab data for today
            date = forecastData[0][0]
            #list to hold data
            todaysData = []
            col = 0
            row = 0
            
            while col < 13:
                today = []
                today.append(forecastData[col][row])
                while row < 11:
                    row += 1
                    today.append(forecastData[col][row])
                row = 0
                col += 1
                todaysData.append(today)
            
            try:
                surfDataFrame = pd.DataFrame(todaysData, columns = ['Date', 'Time', 'temp', 'temp', 'Wave Height + Direction', 'Period', 'temp', 'Energy', 'Wind Speed', 'Wind State', 'High Tide', 'Low Tide'])
                #cleaning data
                del surfDataFrame['temp']
                spotData[spot] = surfDataFrame
                
            except Exception as e:
                print(e)
                print("Make sure Excel is closed")
                
        except Exception as e:
            print("No report found for " + spot)
            
    writer = pd.ExcelWriter('forecast.xlsx', engine='xlsxwriter')
    workbook  = writer.book
    # Green fill with dark green text.
    format1 = workbook.add_format({'bg_color':   '#C6EFCE',
                               'font_color': '#006100'})
    # Light yellow fill with dark yellow text.
    format2 = workbook.add_format({'bg_color':   '#FFEB9C',
                               'font_color': '#9C6500'})
    #red
    format3 = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
        
    for spot, data in spotData.items():
        
        data.to_excel(writer, sheet_name=spot)
        worksheet = writer.sheets[spot]
        worksheet.conditional_format('H1:H14', {'type': 'text', 
                                             'criteria': 'containing',
                                             'value': 'glass',
                                             'format': format1})
        worksheet.conditional_format('H1:H14', {'type': 'text', 
                                             'criteria': 'containing',
                                             'value': 'cross',
                                             'format': format2})
        worksheet.conditional_format('H1:H14', {'type': 'text', 
                                             'criteria': 'containing',
                                             'value': 'on',
                                             'format': format3})
        worksheet.conditional_format('E1:E14', {'type': '3_color_scale',
                                        'min_value': 1,
                                        'mid_value': 10,
                                        'max_value': 20,
                                        'min_color': '#FFC7CE',
                                        'mid_color': '#FFEB9C',
                                        'max_color': '#C6EFCE'})
        
    writer.save()
    
    

#def magic_req():
    
    


surf_forecast_req()