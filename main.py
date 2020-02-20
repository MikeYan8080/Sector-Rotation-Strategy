import pandas as pd
import numpy as np
from scipy import stats
from pyecharts.charts import Bar, Page
from pyecharts import options as opts

# import the initial ETF return data
data0 = pd.read_excel(r'C:\Users\严书航\Desktop\PM Project\ETF returns volume.xlsx')

# convert the date to Year & Month
data0.rename(columns={'Names Date': 'YearMonth'}, inplace=True)
data0.YearMonth = data0.YearMonth.map(lambda x: x//100)

# transpose the data to (time. etf)table show the monthly return
data1 = data0.pivot(index='YearMonth', columns='Ticker Symbol', values='Returns')

# drop the ETF does not have historical data from Jan 2009
data2 = data1.dropna(axis=1, how='any')

# Calculate the yearly return for each ETF
data3 = data2.reset_index()
data3.YearMonth = data3.YearMonth.map(lambda x: x//100)
data3 = data3.groupby('YearMonth').apply(lambda x: np.prod(x+1)-1)
data3.drop(columns='YearMonth', inplace=True)

# Output data to excel
# with pd.ExcelWriter(r'C:\Users\严书航\Desktop\PM Project\ETFs_Return.xlsx') as writer:
#   data2.to_excel(writer, sheet_name='Monthly Return')
#   data3.to_excel(writer, sheet_name='Yearly Return')

# Change the order of the data3
etfdata = pd.read_excel(r'C:\Users\严书航\Desktop\PM Project\PM ETF Chosen.xlsx', sheet_name='ETF selected')
etfdata.set_index(etfdata.Ticker, inplace=True)
data3 = data3[list(etfdata.Ticker)]


# transpose the data3 and drop the benchmark('SPY')
data4 = data3.drop(columns='SPY', inplace=False)
data4 = data4.T

# transpose the data to (time. etf)table show the monthly Volume
data11 = data0.pivot(index='YearMonth', columns='Ticker Symbol', values='Volume')

# drop the ETF does not have historical data from Jan 2009
data22 = data11.dropna(axis=1, how='any')

# Calculate the yearly Volume for each ETF
data33 = data22.reset_index()
data33.YearMonth = data33.YearMonth.map(lambda x: x//100)
data33 = data33.groupby('YearMonth').apply(lambda x: np.sum(x))
data33.drop(columns='YearMonth', inplace=True)
data33 = data33[list(etfdata.Ticker)]

# Output data to excel
# with pd.ExcelWriter(r'C:\Users\严书航\Desktop\PM Project\ETFs_Return.xlsx') as writer:
#   data22.to_excel(writer, sheet_name='Monthly Volume')
#   data33.to_excel(writer, sheet_name='Yearly Volume' )

# transpose the data3 and drop the benchmark('SPY')
data44 = data33.drop(columns='SPY', inplace=False)
data44 = data44.T


# sector rotation stratege, numETF means number of ETF under watch, num means number of ETF bought yearly
# ascend means the stratege choosing max past or min past, Fee means considerring the ETF's expense
def rotation(data, numetf=10, num=2, ascend=False, fee=False):
    data5 = data.drop(columns=2019, inplace=False)
    data5 = data5[0: numetf]
    etf_choose = []
    for i in data5:
        data6 = data5[i].sort_values(ascending=ascend)
        for x in range(num):
            nyr = data4.loc[data6.index[x], i+1]
            if fee:
                nyr -= etfdata.loc[data6.index[x], 'Expense Ratio']
            item1 = [i+1, data6.index[x], data6[x], nyr]

            etf_choose.append(item1)

    data7 = pd.DataFrame(etf_choose, columns=['Year', 'Ticker', 'Yearly Return', 'Next Yearly Return'])
    # print(data7)

    # Calculate the mean return
    data8 = data7.groupby(by='Year').mean()
    # Portfolio's total Return from 2010 to 2019
    # portTotalR = data8.apply(lambda y: np.prod(y+1))[1]
    # portstd = data8.apply(lambda y: np.std(y, ddof=1))[1]
    return data8['Next Yearly Return']


# SPY's  Returns from 2010 to 2019
spyR = data3.drop(2009).SPY
# SPY's  Returns discount the expense from 2010 to 2019
spyRD = spyR - etfdata.loc['SPY', 'Expense Ratio']


# risk free rates come form 30 day T-bill form 2010 to 2019
risk_free_rate = [0.001216, 0.000428, 0.000565, 0.000277, 0.000164, 0.000092, 0.001886, 0.007914, 0.017066, 0.02147]


# Function to calculate : Average Return, Sharp Ratio, Alpha, M-Measure
def measure(returns, RFR=risk_free_rate, BM=spyR):
    AR = np.mean(returns)
    SR = np.mean(returns-RFR)/np.std(returns-RFR, ddof=1)
    Alpha = stats.linregress(x=(BM-risk_free_rate), y=(returns-risk_free_rate)).intercept
    MS = SR*np.std(BM-risk_free_rate, ddof=1) + np.mean(risk_free_rate)
    return [AR, SR, Alpha, MS]


# Function to show the results comparing to Bench Mark
def result(data, ascends, names, fees=True):
    table = pd.DataFrame(measure(spyR),
                         index=['Average Return', 'Sharp Ratio', 'Alpha', 'M-Measure'],
                         columns=['SPY-Bench Mark'])
    for x in range(3):
        name = str(x+1) + names
        table[name] = measure(rotation(data, num=x+1, ascend=ascends, fee=fees))
    return table


# rotation based on return, picked the last year best 1,2,3
a = result(data4, False, 'Best Past Returns')
# rotation based on return, picked the last year worst 1,2,3
b = result(data4, True, 'Worst Past Returns')
# rotation based on volume, picked the last year best 1,2,3
c = result(data44, False, 'Best Past Volume')
# rotation based on volume, picked the last year best 1,2,3
d = result(data44, True, 'Worst Past Volume')


page = Page(layout=Page.SimplePageLayout, page_title='Projection about Portfolio strategy')
for i, x in enumerate([a, b, c, d]):
    bar = Bar(init_opts=opts.InitOpts(width='600px', height='350px'))
    bar.add_xaxis(list(x.index))
    # x.applymap(lambda t: format(t, '.2%'))
    for j in x:
        bar.add_yaxis(j, ['{:.3f}'.format(t) for t in list(x[j])])
    bar.set_global_opts(title_opts=opts.TitleOpts(title="results", pos_top='top'),
                        legend_opts=opts.LegendOpts(pos_top='10%'))
    bar.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
    page.add(bar)
    with pd.ExcelWriter(r'C:\Users\严书航\Desktop\Results.xlsx', mode='a') as writer:
        x.to_excel(writer, sheet_name=str(i))

page.render(r'Results.html')
