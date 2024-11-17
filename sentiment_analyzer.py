# sentiment_analyzer.py
from textblob import TextBlob
import numpy as np
import logging

def analyze_sentiment(news_headlines, num_news=20):
    sentiments = []
    try:
        for headline in news_headlines[:num_news]:
            blob = TextBlob(headline)
            sentiments.append(blob.sentiment.polarity)
        if sentiments:
            avg_sentiment = np.mean(sentiments)
            return round(avg_sentiment, 2)
        else:
            return "N/A"
    except Exception as e:
        logging.error(f"Error analyzing sentiment: {e}")
        return "N/A"
