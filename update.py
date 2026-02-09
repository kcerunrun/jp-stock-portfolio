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
print(df.columns)


# 株価更新
prices = []
for ticker in df["Ticker"]:
    stock = yf.Ticker(ticker)
    price = stock.history(period="1d")["Close"].iloc[-1]
    prices.append(price)

# シートに書き込み
for i, price in enumerate(prices, start=2):
    if price is not None:
        sheet.update(f"E{i}", [[float(price)]])
    else:
        sheet.update(f"E{i}", [["N/A"]])
