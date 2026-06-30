import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
import os
import pandas as pd
import math   # ★ これが必須

def sanitize(v):
    if v is None:
        return "N/A"
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return "N/A"
    return v

# Google Sheets 認証
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

creds_json = os.environ["GCP_CREDENTIALS"]
with open("creds.json", "w") as f:
    f.write(creds_json)

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

sheet_id = os.environ["SHEET_ID"]
sheet = client.open_by_key(sheet_id).worksheet("Portfolio")

data = sheet.get_all_records()
df = pd.DataFrame(data)

prices = []
prev_closes = []
changes = []
dividends = []
yields = []

for ticker in df["Ticker"]:
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period="5d")

        if len(hist) >= 2:
            today_close = hist["Close"].iloc[-1]
            prev_close = hist["Close"].iloc[-2]
            change = today_close - prev_close
        else:
            today_close = None
            prev_close = None
            change = None

        info = stock.info
        dividend = info.get("dividendRate")
        dividend_yield = info.get("dividendYield")

        if dividend_yield is not None:
            dividend_yield = round(dividend_yield * 100, 2)

    except Exception:
        today_close = None
        prev_close = None
        change = None
        dividend = None
        dividend_yield = None

    prices.append(today_close)
    prev_closes.append(prev_close)
    changes.append(change)
    dividends.append(dividend)
    yields.append(dividend_yield)

write_values = []

for i in range(len(df)):
    y = sanitize(yields[i])
    y = "N/A" if y == "N/A" else f"{y}%"

    row = [
        sanitize(prices[i]),
        sanitize(prev_closes[i]),
        sanitize(changes[i]),
        sanitize(dividends[i]),
        y
    ]
    write_values.append(row)

sheet.update(
    range_name=f"E2:I{len(df)+1}",
    values=write_values
)
