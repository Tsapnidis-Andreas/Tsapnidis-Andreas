import pandas as pd
import numpy as np
from pandas_datareader import data as wb
import matplotlib.pyplot as plt
import yfinance as yf
import datetime as dt
from datetime import datetime, date
import statsmodels.api as sm
from scipy import stats
import openpyxl
import xlsxwriter
from tkinter import *
from scipy.optimize import minimize


def disappear(a):
    a.place(x=0,y=0,width=0,height=0)

def mode_1():
    for i in objects:
        disappear(i)
    label1.place(x=20, y=80, width=200, height=20)
    label2.place(x=240, y=80, width=200, height=20)
    entry1.place(x=20, y=140, width=200, height=20)
    entry2.place(x=20, y=200, width=200, height=20)
    entry3.place(x=20, y=260, width=200, height=20)
    entry4.place(x=240, y=140, width=200, height=20)
    entry5.place(x=240, y=200, width=200, height=20)
    entry6.place(x=240, y=260, width=200, height=20)
    button1.config(command=OK_1,text='RUN')
    button1.place(x=300,y=300,width=60,height=20)
    global mode
    mode=1
def mode_2():
    for i in objects:
        disappear(i)
    label1.place(x=20, y=80, width=200, height=20)
    entry1.place(x=20, y=140, width=200, height=20)
    entry2.place(x=20, y=200, width=200, height=20)
    entry3.place(x=20, y=260, width=200, height=20)
    button1.config(command=OK_2,text='RUN')
    button1.place(x=300,y=300,width=60,height=20)
    global mode
    mode=2

def OK_1():
    t1=entry1.get()
    t2=entry2.get()
    t3=entry3.get()
    w1=float(entry4.get())
    w2=float(entry5.get())
    w3=float(entry6.get())
    for i in objects:
        disappear(i)
    stocks=[t1,t2,t3]
    weights=[w1,w2,w3]
    analyze(stocks,weights)

def OK_2():
    t1=entry1.get()
    t2=entry2.get()
    t3=entry3.get()
    for i in objects:
        disappear(i)
    stocks=[t1,t2,t3]
    optimize(stocks)

def analyze(stocks,weights):
    weights=np.array(weights)
    global inflation_rate
    stock_data = pd.DataFrame()
    average_monthly_returns=[]
    for i in stocks:
        stock_data[i],average_monthly_return= obtain_returns(i)
        average_monthly_returns.append(average_monthly_return)
    average_monthly_returns=np.array(average_monthly_returns)
    print(stock_data)
    stock_data.columns = stocks
    portfolio_df = calculate_portfolio_returns(weights, stock_data)
    portfolio_df.columns = ['portfolio']
    portfolio_average_monthly_return=average_monthly_returns.dot(weights)
    portfolio_stddev,correlation = calculate_portfolio_risk(weights, stock_data)
    portfolio_df_per_cent=pd.DataFrame()
    portfolio_df_per_cent['portfolio']=portfolio_df['portfolio']*100

    cov_matrix = stock_data.cov()
    sharpe = -calculate_sharpe_ratio(weights,average_monthly_returns,cov_matrix)
    sp_monthly_returns,sp_average_monthly_return=obtain_returns('^GSPC')
    sp_data=pd.DataFrame()
    sp_data_per_cent=pd.DataFrame()
    sp_data['monthly returns']=sp_monthly_returns
    beta,r_square=linear_regression(sp_data['monthly returns'].tail(-1),portfolio_df['portfolio'].tail(-1))
    results = pd.DataFrame()
    results['metric'] = ['average monthly return % (nominal)','average monthly return % (inflation adjusted)','standard deviation %', 'sharpe ratio', 'beta', 'R^2']
    results['value'] = [portfolio_average_monthly_return*100,portfolio_average_monthly_return*100-inflation_rate, portfolio_stddev, sharpe, beta,r_square]

    global mode
    print(results)
    if mode==1:
        saving_1(results,correlation)
    else:
        global optimal_weights
        optimal_weights = pd.DataFrame(optimal_weights)
        optimal_weights.index = [stocks]
        optimal_weights.columns = ['weight %']
        optimal_weights['weight %'] = optimal_weights * 100
        saving_2(results,correlation,optimal_weights,stocks)
    updating()


def optimize(stocks):
    global optimal_weights
    stock_data = pd.DataFrame()
    avg_returns = []
    for i in stocks:
        stock_data[i],average_monthly_return=obtain_returns(i)
        avg_returns.append(average_monthly_return)
    avg_returns = np.array(avg_returns)
    cov_matrix = stock_data.cov()

    constraints = {'type': 'eq', 'fun': lambda weights: np.sum(weights) - 1}
    bounds = [(0, 1) for i in range(len(stocks))]

    initial_weights = np.array([1 / len(stocks)] * len(stocks))

    optimized_results = minimize(calculate_sharpe_ratio, initial_weights, args=(avg_returns, cov_matrix), method='SLSQP',
                                 constraints=constraints, bounds=bounds)
    optimal_weights = optimized_results.x
    sharpe=calculate_sharpe_ratio(optimal_weights,avg_returns,cov_matrix)
    analyze(stocks,optimal_weights)


def obtain_returns(stock):
    global n
    global date_index
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.strptime(end_date, '%Y-%m-%d') - dt.timedelta(days=5*365)).strftime('%Y-%m-%d')
    data = yf.download(stock, start=start_date, end=end_date,interval='1mo')
    data=pd.DataFrame(data)
    date_index=data.index
    data['returns']= (data['Adj Close']-data['Adj Close'].shift(1))/data['Adj Close'].shift(1)
    data['returns']=data['returns'].dropna()
    average_monthly_return=data['returns'].mean()
    data.index=range(0,len(data))
    return(data['returns'],average_monthly_return)


def calculate_portfolio_returns(weights,stock_data):
    portfolio_df=pd.DataFrame(stock_data.dot(weights))
    return(portfolio_df)

def calculate_portfolio_risk(weights,stock_data):
    cov_matrix = stock_data.cov()
    corr_matrix = stock_data.corr()
    weights=np.array(weights)
    portfolio_variance = np.dot(weights.T,(np.dot(cov_matrix, weights)))
    portfolio_stddev = np.sqrt(portfolio_variance)*100
    return(portfolio_stddev,corr_matrix)

def calculate_sharpe_ratio(weights,avg_returns,cov_matrix):
    global monthly_risk_free_rate_per_cent
    portfolio_return = np.dot(avg_returns, weights)*100
    portfolio_variance = (np.dot(weights.T, np.dot(cov_matrix, weights)))
    portfolio_stddev = np.sqrt(portfolio_variance) * 100
    sharpe = (portfolio_return - monthly_risk_free_rate_per_cent) / portfolio_stddev
    return (-sharpe)

def linear_regression(x,y):
    global daily_risk_free_rate
    y=y-monthly_risk_free_rate
    x=x-monthly_risk_free_rate
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    r = r_value ** 2

    cov = pd.DataFrame(np.cov(x,y) )
    market_var = x.var()
    cov_with_market = cov.iloc[0, 1]

    beta = cov_with_market / market_var

    return(beta,r)

def saving_1(results,correlation):
    global path
    global date_index
    dfs = {'correlation matrix': correlation, 'analysis': results,}
    writer = pd.ExcelWriter(path + 'Portfolio_Analysis.xlsx', engine='xlsxwriter')
    for sheet_name in dfs.keys():
        if sheet_name == 'analysis':
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=True)
        for i, col in enumerate(dfs[sheet_name].columns):
            worksheet = writer.sheets[sheet_name]
            width = max(dfs[sheet_name][col].apply(lambda x: len(str(x))).max(), len(dfs[sheet_name][col]))
            worksheet.set_column(i, i, width)
    writer.close()

def saving_2(results,correlation,optimal_weights,stocks):
    global path
    portfolio_df=pd.DataFrame()
    optimal_weights=pd.DataFrame(optimal_weights)
    optimal_weights.index=stocks
    dfs={'correlation matrix':correlation, 'analysis':results,'optimal weights':optimal_weights}
    writer=pd.ExcelWriter(path+'Portfolio_Analysis.xlsx',engine='xlsxwriter')
    for sheet_name in dfs.keys():
        if sheet_name == 'analysis':
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=True)
        for i, col in enumerate(dfs[sheet_name].columns):
            worksheet = writer.sheets[sheet_name]
            width = max(dfs[sheet_name][col].apply(lambda x: len(str(x))).max(), len(dfs[sheet_name][col]))
            worksheet.set_column(i, i, width)
    writer.close()

def updating():
    label1.place(x=150, y=80, width=200, height=20)
    label1.config(text='Done')

window=Tk()
window.geometry('500x500')
window.title('Portfolio Analysis and Optimization')

label1=Label(window,text='Ticker')
label2=Label(window,text='weight(eg: msft 0.2)')
entry1=Entry(window)
entry2=Entry(window)
entry3=Entry(window)
entry4=Entry(window)
entry5=Entry(window)
entry6=Entry(window)
button1=Button(window,text='Portfolio Analysis',command=mode_1)
button2=Button(window,text='Portfolio Optimization',command=mode_2)

objects=[label1,label2,entry1,entry2,entry3,entry4,entry5,entry6,button1,button2]

button1.place(x=20, y=80, width=200, height=20)
button2.place(x=240, y=80, width=200, height=20)

#INITIALIZE!!!!!!!!!!!!!!!!

#stocks=['msft','tsla','aapl']
#weights=[0.5,0.3,0.2]
global path
path="C:/Users/tsapn/OneDrive/Υπολογιστής/halof wealth management/"

global inflation_rate
inflation_rate=((1.02)**(1/12)-1)*100

risk_free_rate_per_cent=2
monthly_risk_free_rate = ((1 + risk_free_rate_per_cent/100) ** (1 / 12) - 1)
monthly_risk_free_rate_per_cent=monthly_risk_free_rate*100
#monthly_risk_free_rate_per_cent='{:f}'.format(monthly_risk_free_rate_per_cent)
global mode
mode=1


window.mainloop()