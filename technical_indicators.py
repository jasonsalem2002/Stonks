# technical_indicators.py
import pandas as pd
import logging

def calculate_rsi(hist, window=14):
    delta = hist['Close'].diff()
    up = delta.clip(lower=0)
    down = -delta.clip(upper=0)
    ema_up = up.ewm(com=window-1, adjust=False).mean()
    ema_down = down.ewm(com=window-1, adjust=False).mean()
    rs = ema_up / ema_down
    rsi = 100 - (100 / (1 + rs))
    hist['RSI'] = rsi
    return hist

def calculate_macd(hist):
    ema_12 = hist['Close'].ewm(span=12, adjust=False).mean()
    ema_26 = hist['Close'].ewm(span=26, adjust=False).mean()
    macd = ema_12 - ema_26
    signal = macd.ewm(span=9, adjust=False).mean()
    hist['MACD'] = macd
    hist['MACD_Signal'] = signal
    hist['MACD_Hist'] = hist['MACD'] - hist['MACD_Signal']
    return hist

def calculate_bollinger_bands(hist, window=20, num_std=2):
    rolling_mean = hist['Close'].rolling(window=window).mean()
    rolling_std = hist['Close'].rolling(window=window).std()
    hist['Bollinger_High'] = rolling_mean + (rolling_std * num_std)
    hist['Bollinger_Low'] = rolling_mean - (rolling_std * num_std)
    return hist

def calculate_moving_averages(hist):
    hist['50_MA'] = hist['Close'].rolling(window=50).mean()
    hist['200_MA'] = hist['Close'].rolling(window=200).mean()
    return hist
