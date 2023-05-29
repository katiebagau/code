import requests
import pandas as pd
from openpyxl import Workbook
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

url = "https://iss.moex.com/iss/history/engines/stock/markets/shares/boards/TQBR/securities/GAZP.json"
params = {
    "iss.meta": "off",
    "iss.only": "history",
    "history.columns": "TRADEDATE,CLOSE",
    "limit": "252",
}  # определение парамтров запроса к API

response = requests.get(url, params=params)
data = response.json()

df = pd.DataFrame(data["history"]["data"], columns=data["history"]["columns"])  # пробразование данных в Датафрэйм
df["TRADEDATE"] = pd.to_datetime(df["TRADEDATE"], format="%Y-%m-%d")  # преобразование TRADEDATE в формат даты

df.to_excel("GAZP_Price.xlsx", index=False)  # сохранение в Excel

df["Dt"] = (df["CLOSE"] - df["CLOSE"].shift(1)) / df["CLOSE"].shift(1)  # подсчет приростов стоимости акций за 1 день

# расчет Var
alpha = 0.01  # ур знач для 99% дов инт
T = 1
VaR_long = -np.percentile(df["Dt"], alpha)
VaR_short = -np.percentile(df["Dt"], 100 - alpha)

wb = Workbook()  # файл Excel для VAR
ws = wb.active
ws.append(["VAR_long", "VAR_short"])  # заголовки
ws.append([VaR_long, VaR_short])
wb.save("VAR.xlsx")

plt.figure(figsize=(10, 6))  # график изменения цены за 30 дней

if df["CLOSE"].iloc[-1] > df["CLOSE"].iloc[0]:  # настройка цвета в зависимости от направления тренда
    color = "green"
else:
    color = "red"

plt.plot(df.index, df["CLOSE"], color=color)
plt.xlabel("Date")
plt.ylabel("Price")
plt.title("Stock Price Trend")
plt.grid(True)

with PdfPages("Trend.pdf") as pdf:  # сохранение графика в PDF
    pdf.savefig()

plt.close()
