# utils.py
import logging

def setup_logging():
    logging.basicConfig(
        filename='stock_analysis.log',
        filemode='a',
        format='%(asctime)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )
