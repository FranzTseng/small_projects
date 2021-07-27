
# import libraries
import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
plt.style.use("seaborn")


# plot single ETF price
def plotETF():
    # connect the workbook
    wb = xw.Book.caller()
    data = wb.sheets[0]
    
    symbol = data.range("H2").value
    df = data.range("K3").options(pd.DataFrame, expand="table").value
    chart = plt.figure(figsize=(12,8))
    df[symbol].plot(fontsize=15)
    plt.title(symbol, fontsize = 20)
    plt.xlabel("Date", fontsize = 15)
    plt.ylabel("ETF Price", fontsize = 15)
    plt.xticks(rotation=45)
    data.pictures.add(chart, update=True, name = "chart",
                      left = data.range("D6").left,
                      top = data.range("D6").top,
                      scale = 0.43)

# plot correlation matrix
def plotCorr():
    wb = xw.Book.caller()
    data = wb.sheets[0]

    df = data.range("K3").options(pd.DataFrame, expand="table").value
    df_corr = df.corr()
    chart1 = plt.figure(figsize = (15,12))
    sns.heatmap(df_corr, vmin=0, vmax=1, annot=True,  
                fmt='.1f', linewidth=.4, annot_kws={"fontsize":20})
    data.pictures.add(chart1, update=True, name="chart1",
                      left = data.range("D20").left,
                      top = data.range("D20").top,
                      scale = 0.43)


