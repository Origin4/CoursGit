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
    print("List of stocks import........... OK")
except:
    print("List of stocks import failure")

#symbol = 'AAPL' #symbom of stocks name to get data

try:
    my_columns = ['Ticker','Price', 'Market Capitalization', 'Number Of Shares to buy']
    final_dataframe = pd.DataFrame(columns = my_columns)
    print("DataFrame initialization ...... OK")
except :
    print("DataFrame initialization failure")

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n] 

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

for symbol_string in symbol_strings:    
    try:
        batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}' # api url where we will send data
        data = requests.get(batch_api_url).json() # send HTTP request and get data in JSON form
        print('GET stocks data of ' + symbol_string + ' from IEX CLOUD API......... OK')
    except :
        print("HTTP request to IEX CLOUD API failure")

    try:
        for symbol in symbol_string.split(','):
            final_dataframe = final_dataframe.append(
                                                    pd.Series([symbol,
                                                                data[symbol]['quote']['latestPrice'],
                                                                data[symbol]['quote']['marketCap'],
                                                                'N/A'],
                                                                index = my_columns),
                                                                ignore_index = True)
            print("Adding "+ symbol+" Stocks Data to a Pandas DataFrame .......OK")
    except :
            print('Adding '+ symbol+' Stocks Data to a Pandas DataFrame failure')

print(final_dataframe)

print(len(final_dataframe.index))


portfolio_size = input("Enter the value of your portfolio:")

try:
    val = float(portfolio_size)
    print('portfolio size .....OK')
except ValueError:
    print("That's not a number! \n Try again:")
    portfolio_size = input("Enter the value of your portfolio:")

try:
    position_size = float(portfolio_size) / len(final_dataframe.index)
    for i in range(0, len(final_dataframe['Ticker'])-1):
        final_dataframe.loc[i, 'Number Of Shares to Buy'] = math.floor(position_size / final_dataframe['Price'][i])
    print('Number Of Shares to Buy ........... OK')
except : 
    print('Number Of Shares to Buy failure')

try:
    writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
    final_dataframe.to_excel(writer, sheet_name='Recommended Trades', index = False)
    print('Initializing XlsxWriter Object..........OK')

except:
    print('Initializing XlsxWriter Object failure')

try:
    background_color = '#0a0a23'
    font_color = '#ffffff'

    string_format = writer.book.add_format(
            {
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )

    dollar_format = writer.book.add_format(
            {
                'num_format':'$0.00',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )

    integer_format = writer.book.add_format(
            {
                'num_format':'0',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
    print(' Creating Formats for xlsx file .....OK')
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


