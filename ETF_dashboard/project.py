import xlwings as xw
import pandas as pd
import numpy as np
import yfinance as yf
import datetime as dt
import matplotlib.pyplot as plt
import seaborn as sns
plt.style.use("seaborn")

def main():
    wb = xw.Book.caller()
    data = wb.sheets[0]
    data.range("B3").columns.autofit()

# get the name of ETF
@xw.func
@xw.arg("company", dict, expand= "table")
def show_com(symbol, company):
    return "<" + company[symbol] + ">"

# get ETF description
@xw.func
def summary(symbol):
    com_sum = yf.Ticker(symbol).info
    if "longBusinessSummary" in com_sum.keys():
        return com_sum["longBusinessSummary"]
    else:
        return "No Business Summary for this ETF"


# plot ETF price time data
@xw.func
@xw.arg("symbol")
def plot_etf(symbol):   
    df = data["K3"].options(pd.DataFrame, expand="table").value
    chart = plt.figure(figsize=(12,8))
    df[symbol].plot(fontsize=15)
    plt.title(symbol, fontsize = 20)
    plt.xlabel("Date", fontsize = 15)
    plt.ylabel("ETF Price", fontsize = 15)
    data.pictures.add(chart, update=True, name = "chart",
                       left = sheet.range("D4").left,
                       top = sheet.range("D4").top,
                       scale = 0.6)
    return chart

# Get the ETF's full name
@xw.func
@xw.arg("symbols", expand="down")
@xw.ret(expand="table", index=False)
def full_name(symbols):
    name_list = []
    for symbol in symbols:
        name = yf.Ticker(symbol).info["shortName"]
        name_list.append(name)
    name_list = pd.Series(name_list)
    return name_list


# get ETF price date
@xw.func
@xw.arg("symbols", expand = "down")
@xw.ret(expand="table", index=True, header=True)
def data(symbols, startdate, enddate):
    table_lists=[]
    for symbol in symbols:
        etf_price = yf.Ticker(symbol).history(start = startdate, end = enddate).reset_index()
        index = etf_price.iloc[:, 0]
        table_lists.append(etf_price.Close)
        
        df = pd.concat(table_lists, keys=symbols, axis=1)
    df["date"] = index.astype(str)
    df.set_index("date", inplace=True)
    return df 



@xw.func
@xw.arg("price", pd.DataFrame, expand = "table", index = False, header = True)
@xw.ret(expand="table", index=True, header= True)
def show_stats(symbol, price):
    
    return price[symbol].describe()

