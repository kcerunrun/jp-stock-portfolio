import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
import os
import pandas as pd

# Google Sheets 認証
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

creds_json = os.environ["GCP_CREDENTIALS"]
with open("creds.json", "w") as f:
    f.write(creds_json)

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

# シート取得
sheet_id = os.environ["SHEET_ID"]
sheet = client.open_by_key(sheet_id).worksheet("Portfolio")

# データ読み込み
data = sheet.get_all_records()
df = pd.DataFrame(data)

# 株価更新
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

# シートに書き込み
# まとめて書き込むデータを作成
write_values = []

for i in range(len(df)):
    row = [
        prices[i] if prices[i] is not None else "N/A",
        prev_closes[i] if prev_closes[i] is not None else "N/A",
        changes[i] if changes[i] is not None else "N/A",
        dividends[i] if dividends[i] is not None else "N/A",
        f"{yields[i]}%" if yields[i] is not None else "N/A"
    ]
    write_values.append(row)

# E2:I(最終行) に一括書き込み
sheet.update(
    range_name=f"E2:I{len(df)+1}",
    values=write_values
)
