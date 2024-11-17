# main.py
import os
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import threading
import yfinance as yf
from utils import setup_logging
from data_fetcher import fetch_stock_data, fetch_sentiment
from technical_indicators import calculate_rsi, calculate_macd, calculate_bollinger_bands, calculate_moving_averages
from sentiment_analyzer import analyze_sentiment
from risk_metrics import calculate_risk_metrics
from valuation import dcf_valuation
from report_generator import add_df_to_sheet, add_individual_symbol_sheets, insert_charts_into_workbook, cleanup_charts
from dashboard import create_dashboard

import logging
from openpyxl import Workbook

from datetime import datetime, timedelta
import pytz

# Constants
TIMEZONE = pytz.timezone('America/New_York')
TODAY = datetime.now(TIMEZONE)
ONE_YEAR_AGO = TODAY - timedelta(days=365)

def generate_candlestick_chart(symbol, hist):
    """
    Generates and saves a candlestick chart for a given symbol.
    """
    try:
        mc = mpf.make_marketcolors(up='g', down='r', inherit=True)
        s = mpf.make_mpf_style(marketcolors=mc)
        candle_chart = f"{symbol}_candlestick.png"
        mpf.plot(hist, type='candle', style=s, title=f"{symbol} Candlestick Chart", volume=True, savefig=candle_chart)
        return candle_chart
    except Exception as e:
        logging.error(f"Error generating candlestick chart for {symbol}: {e}")
        return None

def generate_comparative_chart(symbols, hist_data):
    """
    Generates and saves a comparative closing price chart for all symbols.
    """
    try:
        plt.figure(figsize=(12, 8))
        for symbol, hist in hist_data.items():
            if not hist.empty:
                plt.plot(hist.index, hist['Close'], label=symbol)
        plt.title("Comparative Closing Prices")
        plt.xlabel("Date")
        plt.ylabel("Price ($)")
        plt.legend()
        plt.grid(True)
        comparative_chart = "All_Stocks_Comparative.png"
        plt.savefig(comparative_chart)
        plt.close()
        return comparative_chart
    except Exception as e:
        logging.error(f"Error generating comparative chart: {e}")
        return None

def get_stock_symbols():
    """
    Prompts the user to input stock symbols separated by commas.
    """
    user_input = input("Enter stock symbols separated by commas (e.g., AAPL, MSFT, GOOGL): ")
    symbols = [symbol.strip().upper() for symbol in user_input.split(',') if symbol.strip()]
    if not symbols:
        print("No valid symbols entered. Exiting.")
        logging.warning("No valid symbols entered by the user.")
        exit()
    return symbols

def main():
    # Setup logging
    setup_logging()
    
    # Get user input for stock symbols
    symbols = get_stock_symbols()
    
    # Initialize lists to collect data
    all_data = []
    technical_data = []
    sentiment_data = []
    dividend_data = []
    risk_data = []
    valuation_data = []
    charts = []
    hist_data = {}
    
    # Initialize Excel writer with xlsxwriter
    excel_filename = "Enhanced_Stock_Analysis_with_Charts.xlsx"
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
    
    # Fetch and process data for each symbol
    for symbol in symbols:
        logging.info(f"Processing symbol: {symbol}")
        
        # Fetch historical data
        hist = fetch_stock_data(symbol, start_date=ONE_YEAR_AGO, end_date=TODAY)
        if hist is None or hist.empty:
            print(f"Warning: No historical data for {symbol}. Skipping.")
            logging.warning(f"No historical data for {symbol}. Skipping.")
            continue
        hist_data[symbol] = hist
        
        # Calculate moving averages first
        hist = calculate_moving_averages(hist)
        
        # Calculate technical indicators
        hist = calculate_rsi(hist)
        hist = calculate_macd(hist)
        hist = calculate_bollinger_bands(hist)
        
        # Calculate percentage changes
        try:
            current_price = hist['Close'].iloc[-1]
            price_3m = hist['Close'].iloc[-63] if len(hist) >= 63 else hist['Close'].iloc[0]
            price_6m = hist['Close'].iloc[-126] if len(hist) >= 126 else hist['Close'].iloc[0]
            price_1y = hist['Close'].iloc[0]
            pct_change_3m = round(((current_price - price_3m) / price_3m) * 100, 2)
            pct_change_6m = round(((current_price - price_6m) / price_6m) * 100, 2)
            pct_change_1y = round(((current_price - price_1y) / price_1y) * 100, 2)
        except Exception as e:
            print(f"Error calculating percentage changes for {symbol}: {e}")
            logging.error(f"Error calculating percentage changes for {symbol}: {e}")
            pct_change_3m = pct_change_6m = pct_change_1y = "N/A"
        
        # Fetch sentiment
        sentiment = fetch_sentiment(symbol)
        
        # Calculate risk metrics
        volatility, max_drawdown, sharpe_ratio = calculate_risk_metrics(hist)
        
        # Perform DCF valuation
        ticker = yf.Ticker(symbol)
        intrinsic_value = dcf_valuation(ticker)
        fair_value = intrinsic_value  # Simplified assumption
        
        # Get company info
        info = ticker.info
        company_name = info.get('longName', 'N/A')
        sector = info.get('sector', 'N/A')
        industry = info.get('industry', 'N/A')
        market_cap = info.get('marketCap', 'N/A')
        pe_ratio = info.get('trailingPE', 'N/A')
        peg_ratio = info.get('pegRatio', 'N/A')
        eps = info.get('trailingEps', 'N/A')
        beta = info.get('beta', 'N/A')
        revenue_growth = info.get('revenueGrowth', 'N/A')
        debt_to_equity = info.get('debtToEquity', 'N/A')
        free_cash_flow = info.get('freeCashflow', 'N/A')
        dividend_yield = round(info.get('dividendYield', 0) * 100, 2) if info.get('dividendYield') else "N/A"
        payout_ratio = info.get('payoutRatio', 'N/A')
        
        # Dividend History
        dividends = ticker.dividends.tail(5)
        if not dividends.empty:
            dividends = dividends.sort_index()
            dividend_dates = dividends.index.strftime('%Y-%m-%d').tolist()
            dividend_amounts = dividends.values.tolist()
            dividend_dict = {'Date': dividend_dates, 'Amount': dividend_amounts}
        else:
            dividend_dict = {'Date': ['N/A'], 'Amount': ['N/A']}
        
        # Dividend Trend (%)
        if len(dividends) >= 2:
            try:
                dividend_trend = [round(((dividend_amounts[i] - dividend_amounts[i-1]) / dividend_amounts[i-1]) * 100, 2) 
                                  for i in range(1, len(dividend_amounts))]
                dividend_trend = dividend_trend[::-1]  # Reverse to match chronological order
            except Exception as e:
                logging.error(f"Error calculating dividend trend for {symbol}: {e}")
                dividend_trend = ["N/A"]
        else:
            dividend_trend = ["N/A"]
        
        # Determine MA Crossover
        if pd.notna(hist['50_MA'].iloc[-1]) and pd.notna(hist['200_MA'].iloc[-1]):
            ma_crossover = "Golden Cross" if hist['50_MA'].iloc[-1] > hist['200_MA'].iloc[-1] else "Death Cross"
        else:
            ma_crossover = "N/A"
        
        # Add to main data list
        try:
            all_data.append({
                "Stock Symbol": symbol,
                "Company Name": company_name,
                "Sector": sector,
                "Industry": industry,
                "Market Cap": f"${market_cap:,}" if isinstance(market_cap, int) else "N/A",
                "P/E Ratio": pe_ratio if pe_ratio else "N/A",
                "PEG Ratio": peg_ratio if peg_ratio else "N/A",
                "EPS": eps if eps else "N/A",
                "Beta": beta if beta else "N/A",
                "Revenue Growth (%)": round(revenue_growth * 100, 2) if revenue_growth else "N/A",
                "Debt-to-Equity": debt_to_equity if debt_to_equity else "N/A",
                "Free Cash Flow": f"${free_cash_flow:,}" if isinstance(free_cash_flow, (int, float)) else "N/A",
                "Current Price": f"${current_price:.2f}" if isinstance(current_price, (int, float)) else "N/A",
                "Price 3 Months Ago": f"${price_3m:.2f}" if isinstance(price_3m, (int, float)) else "N/A",
                "Price 6 Months Ago": f"${price_6m:.2f}" if isinstance(price_6m, (int, float)) else "N/A",
                "Price 1 Year Ago": f"${price_1y:.2f}" if isinstance(price_1y, (int, float)) else "N/A",
                "% Change (3M)": pct_change_3m,
                "% Change (6M)": pct_change_6m,
                "% Change (1Y)": pct_change_1y,
                "Dividend Yield (%)": dividend_yield,
                "Payout Ratio": round(payout_ratio, 2) if isinstance(payout_ratio, (int, float)) else "N/A",
                "Sentiment Score": sentiment,
                "Intrinsic Value": f"${intrinsic_value:.2f}" if isinstance(intrinsic_value, (int, float)) else "N/A",
                "Fair Value": f"${fair_value:.2f}" if isinstance(fair_value, (int, float)) else "N/A",
                "Volatility": volatility,
                "Max Drawdown": max_drawdown,
                "Sharpe Ratio": sharpe_ratio,
                "RSI": round(hist['RSI'].iloc[-1], 2) if not pd.isna(hist['RSI'].iloc[-1]) else "N/A",
                "MACD": round(hist['MACD'].iloc[-1], 2) if not pd.isna(hist['MACD'].iloc[-1]) else "N/A",
                "MACD Signal": round(hist['MACD_Signal'].iloc[-1], 2) if not pd.isna(hist['MACD_Signal'].iloc[-1]) else "N/A",
                "MACD Histogram": round(hist['MACD_Hist'].iloc[-1], 2) if not pd.isna(hist['MACD_Hist'].iloc[-1]) else "N/A",
                "Bollinger High": round(hist['Bollinger_High'].iloc[-1], 2) if not pd.isna(hist['Bollinger_High'].iloc[-1]) else "N/A",
                "Bollinger Low": round(hist['Bollinger_Low'].iloc[-1], 2) if not pd.isna(hist['Bollinger_Low'].iloc[-1]) else "N/A",
                "MA Crossover": ma_crossover
            })
        except Exception as e:
            logging.error(f"Error adding data to main list for {symbol}: {e}")
        
        # Add to technical data
        technical_data.append({
            "Stock Symbol": symbol,
            "RSI": round(hist['RSI'].iloc[-1], 2) if not pd.isna(hist['RSI'].iloc[-1]) else "N/A",
            "MACD": round(hist['MACD'].iloc[-1], 2) if not pd.isna(hist['MACD'].iloc[-1]) else "N/A",
            "MACD Signal": round(hist['MACD_Signal'].iloc[-1], 2) if not pd.isna(hist['MACD_Signal'].iloc[-1]) else "N/A",
            "MACD Histogram": round(hist['MACD_Hist'].iloc[-1], 2) if not pd.isna(hist['MACD_Hist'].iloc[-1]) else "N/A",
            "Bollinger High": round(hist['Bollinger_High'].iloc[-1], 2) if not pd.isna(hist['Bollinger_High'].iloc[-1]) else "N/A",
            "Bollinger Low": round(hist['Bollinger_Low'].iloc[-1], 2) if not pd.isna(hist['Bollinger_Low'].iloc[-1]) else "N/A"
        })
        
        # Add to sentiment data
        sentiment_data.append({
            "Stock Symbol": symbol,
            "Sentiment Score": sentiment
        })
        
        # Add to dividend data
        for date, amount in zip(dividend_dict['Date'], dividend_dict['Amount']):
            dividend_data.append({
                "Stock Symbol": symbol,
                "Dividend Date": date,
                "Dividend Amount": amount
            })
        
        # Add to risk data
        risk_data.append({
            "Stock Symbol": symbol,
            "Volatility": volatility,
            "Max Drawdown": max_drawdown,
            "Sharpe Ratio": sharpe_ratio
        })
        
        # Add to valuation data
        valuation_data.append({
            "Stock Symbol": symbol,
            "Intrinsic Value": intrinsic_value,
            "Fair Value": fair_value
        })
        
        # Generate and save candlestick chart
        candle_chart = generate_candlestick_chart(symbol, hist)
        if candle_chart:
            charts.append(candle_chart)
        
        # Create Individual Sheets for Symbol
        # Inside the for-loop after processing historical data for each symbol

        # Calculate percentage changes without dropping NaN yet
        # Inside the for-loop after fetching and processing historical data for each symbol

        # Inside the for-loop after fetching and processing historical data for each symbol

        # Calculate percentage changes without dropping NaN yet
        # Inside the for-loop after fetching and processing historical data for each symbol

        # Calculate percentage changes
        daily_change = hist['Close'].pct_change() * 100
        weekly_change = hist['Close'].pct_change(periods=5) * 100
        monthly_change = hist['Close'].pct_change(periods=21) * 100

        # Create the DataFrame with aligned indices
        changes_df = pd.DataFrame({
            'Date': hist.index,
            'Daily Change (%)': daily_change,
            'Weekly Change (%)': weekly_change,
            'Monthly Change (%)': monthly_change
        })

        # Remove timezone information to make datetime objects timezone-naive
        if changes_df['Date'].dtype.kind == 'M':  # 'M' stands for datetime
            changes_df['Date'] = changes_df['Date'].dt.tz_localize(None)

        # **Important:** Drop rows with NaNs in percentage change columns **before** converting 'Date' to string
        changes_df.dropna(subset=['Daily Change (%)', 'Weekly Change (%)', 'Monthly Change (%)'], inplace=True)

        # Convert 'Date' to string format for Excel compatibility
        changes_df['Date'] = changes_df['Date'].dt.strftime('%Y-%m-%d')

        # **Add Logging and Print Statements to Verify Data**
        logging.info(f"{symbol} changes_df has {len(changes_df)} rows after dropna.")
        print(f"{symbol} changes_df has {len(changes_df)} rows after dropna.")

        # Optionally, log a sample of the DataFrame
        logging.info(f"Sample changes_df for {symbol}:\n{changes_df.head()}")
        print(f"Sample changes_df for {symbol}:\n{changes_df.head()}")

        # Now add the individual symbol sheets
        add_individual_symbol_sheets(writer, symbol, changes_df)


    
    # After processing all symbols, generate DataFrames
    df_overview = pd.DataFrame(all_data)
    df_technical = pd.DataFrame(technical_data)
    df_sentiment = pd.DataFrame(sentiment_data)
    df_dividends = pd.DataFrame(dividend_data)
    df_risk = pd.DataFrame(risk_data)
    df_valuation = pd.DataFrame(valuation_data)
    
    # Add DataFrames to Excel workbook
    add_df_to_sheet(writer, df_overview, "Stock Overview", conditional_format_cols=(12, 14))  # Columns M to O for % Changes
    add_df_to_sheet(writer, df_technical, "Technical Indicators")
    add_df_to_sheet(writer, df_sentiment, "Sentiment Analysis")
    add_df_to_sheet(writer, df_dividends, "Dividend History")
    add_df_to_sheet(writer, df_risk, "Risk Analysis")
    add_df_to_sheet(writer, df_valuation, "Valuation Metrics")
    
    # Generate and save comparative chart
    comparative_chart = generate_comparative_chart(symbols, hist_data)
    if comparative_chart:
        charts.append(comparative_chart)
    
    # Insert Charts into the workbook
    insert_charts_into_workbook(writer, charts, df_overview)
    
    # Save the Excel workbook
    try:
        writer.close()
        print(f"Excel report generated: '{excel_filename}'")
        logging.info(f"Excel report generated: '{excel_filename}'")
    except Exception as e:
        print(f"Error saving Excel workbook: {e}")
        logging.error(f"Error saving Excel workbook: {e}")

    
    # Clean up chart images
    cleanup_charts(charts)
    
    # Start the dashboard in a separate thread
    dashboard_thread = threading.Thread(target=create_dashboard, args=(symbols, hist_data))
    dashboard_thread.start()
    
    print("Stock analysis completed. Results saved to 'Enhanced_Stock_Analysis_with_Charts.xlsx'.")
    logging.info("Stock analysis completed successfully.")

if __name__ == "__main__":
    main()
