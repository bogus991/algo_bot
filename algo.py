#!/usr/bin/env python
# coding: utf-8

# In[13]:


#tixisi1594@azduan.com - 15938430


import time
import openpyxl
from API import XTB
import numpy as np

from IPython.display import clear_output





interval = 'M15'
close_prices_prev_one = {'M1':1, 'M5':2, 'M15':2}
close_prices_prev_two = {'M1':2, 'M5':3, 'M15':3}
close_candle = 13 #14 for M1, 13 for M5
can_len = 14







#symbol_div = {'US500':10, 'US30':1}
#list_of_symbols= ['US500', 'US30']
#symbol_volume = {'US500':0.01, 'US30':0.01}

symbol_div = {'US500':10, 'US30':1, 'DE40':10, 'FRA40':10, 'NED25':100}
list_of_symbols= ['US500', 'US30', 'DE40', 'FRA40', 'NED25']
symbol_volume = {'US500':0.01, 'US30':0.01, 'DE40':0.01, 'FRA40':0.01, 'NED25':0.01}

#symbol_div = {'ETHEREUM':1000, 'BITCOIN':100}
#list_of_symbols= ['ETHEREUM', 'BITCOIN']
#symbol_volume = {'ETHEREUM':0.01, 'BITCOIN':0.01}


open_position = 0
open_long = 0
open_short = 0
check_long = 1
check_short = 1
pnl = 0
close_check1 = 0
closed_time = 0
atr_closing_short = 0
atr_closing_long = 0
atr_sl_short = 0
atr_sl_long = 0
opened_symbol = 'x'
#beciv11348@ebuthor.com / 15795751
#sevex95592@huizk.com / 15809918
API = XTB("15938430", "Qweasd12")
i=1


def add_text_to_excel(file_path, sheet_name, cell_address, text_to_add):
    # Load the existing workbook
    workbook = openpyxl.load_workbook(file_path)

    # Select the desired sheet
    sheet = workbook[sheet_name]

    # Update the specified cell with the given text
    sheet[cell_address] = text_to_add

    # Save the changes
    workbook.save(file_path)

    print(f"Text '{text_to_add}' added to cell {cell_address} in {file_path}.")

excel_file_path = 'pnl.xlsx'
sheet_name = 'Sheet1'  # Change this to the actual sheet name
ex = 1
cell_address = f'A{ex}'

ii = 0
while ii < len(list_of_symbols):
   
    #time.sleep(1)
    last_candles = API.get_Candles(interval, list_of_symbols[ii], qty_candles=can_len)
    time.sleep(1)
    last_candles_b = API.get_Candles(interval, list_of_symbols[ii], qty_candles=can_len+13)
    time.sleep(1)

    

    sma_s = 1
    close_prices = []
    while sma_s <= can_len:
        close_prices.append((last_candles[sma_s]['open']+last_candles[sma_s]['close'])/symbol_div[list_of_symbols[ii]])
        sma_s=sma_s+1
    

    sma_h = 1
    high_prices = []
    while sma_h <= can_len:
        high_prices.append((last_candles[sma_h]['open']+last_candles[sma_h]['high'])/symbol_div[list_of_symbols[ii]])
        sma_h=sma_h+1

    sma_l = 1
    low_prices = []
    while sma_l <= can_len:
        low_prices.append((last_candles[sma_l]['open']+last_candles[sma_l]['low'])/symbol_div[list_of_symbols[ii]])
        sma_l=sma_l+1
    
    sma_b = 1
    close_prices_b = []
    while sma_b <= can_len+13:
        close_prices_b.append((last_candles_b[sma_b]['open']+last_candles_b[sma_b]['close'])/symbol_div[list_of_symbols[ii]])
        sma_b=sma_b+1
    

    

#     def calculate_bollinger_bands(prices, window_size=can_len, num_std_dev=2):
#         """
#         Calculate Bollinger Bands and Bollinger Bandwidth.

#         Parameters:
#         - prices: List of closing prices.
#         - window_size: Size of the moving average window.
#         - num_std_dev: Number of standard deviations for upper and lower bands.

#         Returns:
#         - middle_band: Single value for the middle band.
#         - upper_band: Single value for the upper band.
#         - lower_band: Single value for the lower band.
#         - bandwidth_range: Tuple representing the range of Bollinger Bandwidth.
#         """
#         middle_band = []
#         upper_band = []
#         lower_band = []
#         bandwidth_range = []

#         for i in range(len(prices)):
#             if i < window_size - 1:
#                 continue
#             else:
#                 window = prices[i - window_size + 1 : i + 1]
#                 middle = sum(window) / window_size
#                 std_dev = (sum((x - middle) ** 2 for x in window) / window_size) ** 0.5

#                 middle_band.append(middle)
#                 upper_band.append(middle + num_std_dev * std_dev)
#                 lower_band.append(middle - num_std_dev * std_dev)
#                 bandwidth = (upper_band[-1] - lower_band[-1]) / middle * 100
#                 bandwidth_range.append(bandwidth)

#         return middle_band[-1], upper_band[-1], lower_band[-1], (min(bandwidth_range), max(bandwidth_range))

#     middle_band, upper_band, lower_band, bandwidth_range = calculate_bollinger_bands(close_prices)




    def calculate_bollinger_bands(data, window_size=14, num_std_dev=2):
        """
        Calculate Bollinger Bands for a given data set.

        Parameters:
            data (numpy array or list): The input data set.
            window_size (int): The window size for the moving average.
            num_std_dev (int): The number of standard deviations for the bands.

        Returns:
            upper_bands (numpy array): The upper Bollinger Band values.
            lower_bands (numpy array): The lower Bollinger Band values.
        """
        # Convert data to numpy array
        data = np.array(data)

        # Initialize arrays to store upper and lower bands
        upper_bands = np.zeros(len(data) - window_size + 1)
        lower_bands = np.zeros(len(data) - window_size + 1)

        # Calculate rolling mean and standard deviation
        for i in range(len(data) - window_size + 1):
            window_data = data[i:i+window_size]
            rolling_mean = np.mean(window_data)
            rolling_std = np.std(window_data, ddof=1)
            upper_bands[i] = rolling_mean + num_std_dev * rolling_std
            lower_bands[i] = rolling_mean - num_std_dev * rolling_std

        return upper_bands, lower_bands

    upper_band, lower_band = calculate_bollinger_bands(close_prices_b)


    print("Upper Band:", upper_band[len(upper_band)-2])
    print("Lower Band:", lower_band[len(lower_band)-2])
    
    print("\nUpper Band prev:", upper_band[len(upper_band)-3])
    print("Lower Band prev:", lower_band[len(lower_band)-3])
    
    if open_short == 1:
        print('Targer price:', atr_closing_short)
        print('Stop loss:', atr_sl_short)
    elif open_long == 1:
        print('Targer price:', atr_closing_long)
        print('Stop loss:', atr_sl_long)
    
    print('\nLast close -1::', close_prices[len(close_prices)-close_prices_prev_one[interval]])
    print('\nLast close -2::', close_prices[len(close_prices)-close_prices_prev_two[interval]])
    
    def calculate_rsi(prices, period=can_len):
        """
        Calculate the Relative Strength Index (RSI) for a given price series.

        Parameters:
        - prices (list or array): List of closing prices.
        - period (int): Number of periods to use for RSI calculation.

        Returns:
        - rsi_values (list): List of RSI values for each corresponding price.
        """

        if len(prices) < period + 1:
            raise ValueError("Insufficient data for RSI calculation")

        # Calculate price changes
        price_changes = [prices[i] - prices[i - 1] for i in range(1, len(prices))]

        # Separate gains and losses
        gains = [change if change > 0 else 0 for change in price_changes]
        losses = [-change if change < 0 else 0 for change in price_changes]

        # Calculate average gains and losses
        avg_gain = sum(gains[:period]) / period
        avg_loss = sum(losses[:period]) / period

        rsi_values = [100 - (100 / (1 + avg_gain / avg_loss))] if avg_loss != 0 else [100]

        # Calculate RSI for the remaining periods
        for i in range(period, len(prices) - 1):
            current_gain = gains[i - 1] if gains[i - 1] > 0 else 0
            current_loss = losses[i - 1] if losses[i - 1] > 0 else 0

            avg_gain = (avg_gain * (period - 1) + current_gain) / period
            avg_loss = (avg_loss * (period - 1) + current_loss) / period

            rsi = 100 - (100 / (1 + avg_gain / avg_loss)) if avg_loss != 0 else 100
            rsi_values.append(rsi)

        return rsi_values


    rsi_values = calculate_rsi(close_prices, period=13)
    print('\nRSI:', rsi_values[0])
    
    
    def atr(high_prices, low_prices, close_prices, period=14):
        atr_values = []
        tr_values = []

        for i in range(len(high_prices)):
            if i == 0:
                tr_values.append(high_prices[i] - low_prices[i])
            else:
                tr = max(high_prices[i] - low_prices[i], abs(high_prices[i] - close_prices[i - 1]), abs(low_prices[i] - close_prices[i - 1]))
                tr_values.append(tr)

            if len(tr_values) >= period:
                atr_value = sum(tr_values[-period:]) / period
                atr_values.append(atr_value)
            else:
                atr_values.append(None)

        return atr_values
 
    # Calculate ATR with default period of 14
    atr_values = atr(high_prices, low_prices, close_prices)
    atr = atr_values[len(atr_values)-1]
    # Print the ATR values
    print("\nATR:", atr)
    time.sleep(1)
    price2 = API.get_TickPrices(level=0, symbols=list_of_symbols[ii], timestamp=0)
    price3 = price2['quotations'][0]['bid']
    
    print('\nCurrent price:', price3)
    print('\n', list_of_symbols[ii])
    
    
    market_time = str(API.get_time())
    market_time_split = market_time.split(sep=' ')
    market_minutes = int(market_time_split[1][3:5])
    print(market_minutes)

    if open_position == 1:
        profit_loss = API.get_Trades()
        pnl = (profit_loss[0]['profit'])
        print('\nPNL:', pnl)
        
    #long open
    if (
    open_position == 0 
    #and opened_symbol == list_of_symbols[ii]
    #and check_long == 0
    and close_prices[len(close_prices)-close_prices_prev_two[interval]] < lower_band[len(lower_band)-3]
    and close_prices[len(close_prices)-close_prices_prev_one[interval]] > lower_band[len(lower_band)-2]
    and rsi_values[0] < 50
    and closed_time <= market_minutes-5
    ):
        
        status, order_code = API.make_Trade(list_of_symbols[ii], 0, 0, volume=symbol_volume[list_of_symbols[ii]], comment='1111')
        #close_check1 = close_prices[len(close_prices)-3]
        #close_check1 = market_minutes + 7
        atr_closing_long = close_prices[len(close_prices)-1] + (atr*0.6)
        atr_sl_long = close_prices[len(close_prices)-1] - (atr*0.4)
        opened_symbol = list_of_symbols[ii]
        open_position = 1
        check_long = 1
        open_long = 1
        time.sleep(1)
        trade_info=API.get_Trades()
        order_number=trade_info[0]['order']
        symbol_div = {list_of_symbols[ii]: symbol_div.get(list_of_symbols[ii])}
        symbol_volume = {list_of_symbols[ii]: symbol_volume.get(list_of_symbols[ii])}
        list_of_symbols = [list_of_symbols[ii]]
        ii = 0
        closed_time = 0
        


    #long close
    if (open_position == 1 
    and open_long == 1 
    and opened_symbol == list_of_symbols[ii]
    #and close_check1 == close_prices[len(close_prices)-5]):
    #and pnl > 1 or pnl < -1):
    #and close_check1 == market_minutes
    and (price3 > atr_closing_long or price3 < atr_sl_long)
       ):
        
        time.sleep(1)
        status, order_code = API.make_Trade(list_of_symbols[ii], 0, 2, volume=symbol_volume[list_of_symbols[ii]], comment='1111', order=order_number)
        close_check1 = 0
        open_position = 0
        atr_closing_long = 0
        atr_sl_long = 0
        time.sleep(1)
        open_long = 0
        position_date = last_candles[close_candle]['datetime']
        pnl_excel = f'Long, target price, {position_date}, {[list_of_symbols[ii]]}, {pnl}' 
        add_text_to_excel(excel_file_path, sheet_name, cell_address, pnl_excel)
        ex = ex + 1
        cell_address = f'A{ex}'
        pnl = 0
        opened_symbol = 'x'
        symbol_div = {'US500':10, 'US30':1, 'DE40':10, 'FRA40':10, 'NED25':100}
        list_of_symbols= ['US500', 'US30', 'DE40', 'FRA40', 'NED25']
        symbol_volume = {'US500':0.01, 'US30':0.01, 'DE40':0.01, 'FRA40':0.01, 'NED25':0.01}
        ii = 0
        closed_time = market_minutes
 
    
    #short open
    if (
    open_position == 0
    #and opened_symbol == list_of_symbols[ii]
    #and check_short == 0
    and close_prices[len(close_prices)-close_prices_prev_two[interval]] > upper_band[len(upper_band)-3] 
    and close_prices[len(close_prices)-close_prices_prev_one[interval]] < upper_band[len(upper_band)-2] 
    and rsi_values[0] > 50
    and closed_time <= market_minutes-5
    
    ):
        
        status, order_code = API.make_Trade(list_of_symbols[ii], 1, 0, volume=symbol_volume[list_of_symbols[ii]], comment='1111')
        #close_check1 = close_prices[len(close_prices)-3]
        #close_check1 = market_minutes + 7
        atr_closing_short = close_prices[len(close_prices)-1] - (atr*0.6)
        atr_sl_short = close_prices[len(close_prices)-1] + (atr*0.4)
        opened_symbol = list_of_symbols[ii]
        open_position = 1
        check_short = 1
        open_short = 1
        time.sleep(1)
        trade_info=API.get_Trades()
        order_number=trade_info[0]['order']
        symbol_div = {list_of_symbols[ii]: symbol_div.get(list_of_symbols[ii])}
        symbol_volume = {list_of_symbols[ii]: symbol_volume.get(list_of_symbols[ii])}
        list_of_symbols = [list_of_symbols[ii]]
        ii = 0
        closed_time = 0
    
    
    
    #short close
    if (open_position == 1 
    and open_short == 1 
    and opened_symbol == list_of_symbols[ii]
    #and close_check1 == close_prices[len(close_prices)-5]):
    #and pnl > 1 or pnl < -1):
    #and close_check1 == market_minutes
    and (price3 < atr_closing_short or price3 > atr_sl_short)
       ):
        
        time.sleep(1)
        status, order_code = API.make_Trade(list_of_symbols[ii], 1, 2, volume=symbol_volume[list_of_symbols[ii]], comment='1111', order=order_number)
        close_check1 = 0
        open_position = 0
        atr_closing_short = 0
        atr_sl_short = 0
        time.sleep(1)
        open_short = 0
        position_date = last_candles[close_candle]['datetime']
        pnl_excel = f'Short, target price, {position_date}, {[list_of_symbols[ii]]}, {pnl}' 
        add_text_to_excel(excel_file_path, sheet_name, cell_address, pnl_excel)
        ex = ex + 1
        cell_address = f'A{ex}'
        pnl = 0
        opened_symbol = 'x'
        symbol_div = {'US500':10, 'US30':1, 'DE40':10, 'FRA40':10, 'NED25':100}
        list_of_symbols= ['US500', 'US30', 'DE40', 'FRA40', 'NED25']
        symbol_volume = {'US500':0.01, 'US30':0.01, 'DE40':0.01, 'FRA40':0.01, 'NED25':0.01}
        ii = 0
        closed_time = market_minutes

#     if (
#     price3 < lower_band
#     and open_position == 0
#     and check_long == 1):
#         check_long = 0
#         opened_symbol = list_of_symbols[ii]
#         time.sleep(1)

            
    
#     if (
#     price3 > upper_band
#     and open_position == 0
#     and check_short == 1):
#         check_short = 0
#         opened_symbol = list_of_symbols[ii]
#         time.sleep(1)

        
    ii=ii+1
    if ii == len(list_of_symbols):
        ii = 0    
        
    if market_minutes == 59:
        closed_time = 0
        
    
    clear_output(wait=True)
    


# In[125]:


close_prices_b


# In[3]:


close_prices[len(close_prices)-close_prices_prev_one[interval]] + (atr*2)

