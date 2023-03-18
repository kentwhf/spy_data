import yfinance as yf
import pandas as pd
import json
import requests
from datetime import datetime, timedelta, date


def load_data_api(symbol, start_date, end_date):
    """
    Directly use yfinance lib
    """
    df = yf.download(symbol, start=start_date, end=end_date, interval="1d")
    df.index = df.index.tz_localize(None) # Remove timezone information from the index
    return df[['Open', 'High', 'Low', 'Close', 'Volume']]


def scrape_data(symbol, start_date, end_date):
    """
    scrape data from requests
    """
    
    # format conversion
    start_datetime = datetime.combine(start_date, datetime.min.time())
    end_datetime = datetime.combine(end_date, datetime.max.time())

    start_timestamp = int(start_datetime.timestamp())
    end_timestamp = int(end_datetime.timestamp())
    
    # set headers to avoid blocks from API
    url = f"https://query1.finance.yahoo.com/v7/finance/chart/{symbol}?period1={start_timestamp}&period2={end_timestamp}&interval=1d&indicators=quote&includeTimestamps=true"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
        }
    
    response = requests.get(url, headers=headers)
    data = json.loads(response.text)
    
    timestamps = data["chart"]["result"][0]["timestamp"]
    opens = data["chart"]["result"][0]["indicators"]["quote"][0]["open"]
    highs = data["chart"]["result"][0]["indicators"]["quote"][0]["high"]
    lows = data["chart"]["result"][0]["indicators"]["quote"][0]["low"]
    closes = data["chart"]["result"][0]["indicators"]["quote"][0]["close"]
    volumes = data["chart"]["result"][0]["indicators"]["quote"][0]["volume"]

    historical_prices = []
    for i in range(len(timestamps)):
        date = datetime.fromtimestamp(timestamps[i])
        if date >= start_date:
            historical_prices.append({
                "date": date,
                "open": opens[i],
                "high": highs[i],
                "low": lows[i],
                "close": closes[i],
                "volume": volumes[i]
            })

    return pd.DataFrame(historical_prices)


if __name__ == "__main__":
    ticker = "SPY"    
    start_date = datetime.now() - timedelta(days=365)
    end_date = date.today()
    
    df = load_data_api(ticker, start_date, end_date)
    # df = scrape_data(ticker, start_date, end_date)
    
    # Export the data to an Excel file
    df.to_excel('spy_data.xlsx')