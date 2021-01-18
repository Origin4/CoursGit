#Libraries Import
try:
    import numpy as np #numpy numerical camputing library 
    import pandas as pd #pandas data science library
    import requests # HTTP requests 
    import xlsxwriter as xlsxw # module for creating Excel xlsx files
    import math # math module
    import logging # logs modules
    from secret import IEX_CLOUD_API_TOKEN # import the IEX_CLOUD_API_TOKEN from python file secret.py
    print("Libraries import.............. OK")
except :
    print("Libraries import failure") #exception import failure

#List of stocks import
try:
    stocks = pd.read_csv('stocks.csv')
 
except:
    print(' Creating Formats for xlsx file failure')

try:
    column_formats = { 
                        'A': ['Ticker', string_format],
                        'B': ['Price', dollar_format],
                        'C': ['Market Capitalization', dollar_format],
                        'D': ['Number of Shares to Buy', integer_format]
                        }

    for column in column_formats.keys():
        writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
        writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)
    print('append xlsx file .....OK')
except:
    print('append xlsx file failure')

writer.save()


