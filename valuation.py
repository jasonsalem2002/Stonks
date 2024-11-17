# valuation.py
import logging

def dcf_valuation(ticker):
    try:
        fcfs = ticker.cashflow.loc['Free Cash Flow']
        if fcfs.empty:
            logging.warning("Free Cash Flow data is empty.")
            return "N/A"
        recent_fcf = fcfs.iloc[0]
        growth_rate = 0.05  # Assumed growth rate
        discount_rate = 0.10  # Assumed discount rate
        terminal_value = recent_fcf * (1 + growth_rate) / (discount_rate - growth_rate)
        intrinsic_value = terminal_value / (1 + discount_rate) ** 1  # Simplified for 1 year
        return round(intrinsic_value, 2)
    except Exception as e:
        logging.error(f"Error performing DCF valuation: {e}")
        return "N/A"
