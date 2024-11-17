# risk_metrics.py
import numpy as np
import logging

def calculate_risk_metrics(hist):
    try:
        returns = hist['Close'].pct_change().dropna()
        volatility = returns.std() * np.sqrt(252)  # Annualized volatility
        cumulative_returns = (1 + returns).cumprod()
        peak = cumulative_returns.cummax()
        drawdown = (cumulative_returns - peak) / peak
        max_drawdown = drawdown.min()
        sharpe_ratio = (returns.mean() / returns.std()) * np.sqrt(252) if returns.std() != 0 else "N/A"
        return round(volatility, 4), round(max_drawdown, 4), round(sharpe_ratio, 4) if sharpe_ratio != "N/A" else "N/A"
    except Exception as e:
        logging.error(f"Error calculating risk metrics: {e}")
        return "N/A", "N/A", "N/A"
