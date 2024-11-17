# data_fetcher.py
import yfinance as yf
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
import numpy as np
import logging

def fetch_stock_data(symbol, start_date, end_date):
    """
    Fetches historical stock data for a given symbol.
    """
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(start=start_date, end=end_date)
        if hist.empty:
            logging.warning(f"No historical data for {symbol}.")
        return hist
    except Exception as e:
        logging.error(f"Error fetching historical data for {symbol}: {e}")
        return None

def fetch_sentiment(symbol, num_news=20):
    """
    Fetches and analyzes sentiment from the latest news headlines for a given symbol.
    """
    try:
        url = f'https://finviz.com/quote.ashx?t={symbol}'
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        news_table = soup.find(id='news-table')
        if not news_table:
            logging.warning(f"No news table found for {symbol}.")
            return "N/A"
        sentiments = []
        for row in news_table.findAll('tr')[:num_news]:
            text = row.a.get_text()
            blob = TextBlob(text)
            sentiments.append(blob.sentiment.polarity)
        if sentiments:
            avg_sentiment = np.mean(sentiments)
            return round(avg_sentiment, 2)
        else:
            return "N/A"
    except Exception as e:
        logging.error(f"Error fetching sentiment for {symbol}: {e}")
        return "N/A"
