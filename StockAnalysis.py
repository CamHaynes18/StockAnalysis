# Import Packages
# import subprocess
import pandas as pd
import yfinance as yf
import cpi
from datetime import datetime


# with open('requirements.txt', 'w') as file_:
#     subprocess.Popen(['pip', 'freeze', '-l'], stdout=file_).communicate()
# print('requirements.txt saved')


print('Started')


tickers = yf.Tickers('VBTLX VGIT VTIAX VXUS VTEB VOHIX VTSAX VTI VOO VIG SCHD VYM VGSLX')
i = 0
divInfo = pd.DataFrame()


# div = tickers.tickers['SCHD'].dividends
#
# div = div.to_frame().reset_index()
# div['Year'] = div['Date'].dt.year
# div['Date'] = pd.to_datetime(div['Date']).dt.date
# div['Real Dividends'] = ''
# div['Price'] = ''
# div = div[div['Year'] != datetime.now().year]
#
# for x in range(0, div.shape[0]):
#     div.iloc[x, 3] = cpi.inflate(div.iloc[x, 1], div.iloc[x, 2])
#     div.iloc[x, 4] = yf.download('SCHD', div.iloc[x, 0], div.iloc[x, 0] +
#                                  pd.Timedelta(days=1))['Adj Close'].values[0]
#
# div['Dividend Yield'] = div['Dividends'] / div['Price']
#
# divYearly = pd.concat([div.groupby('Year')['Dividends'].sum(), div.groupby('Year')['Real Dividends'].sum(),
#                        div.groupby('Year')['Dividend Yield'].sum()], axis=1)
# divYearly = divYearly.iloc[1:]
#
# divGrowth = divYearly['Dividends'].pct_change().iloc[1:]
# realDivGrowth = divYearly['Real Dividends'].pct_change().iloc[1:]
# divInfo = pd.DataFrame({'Symbol': ['SCHD'], 'Dividend Yield': [divYearly['Dividend Yield'].iloc[-1]],
#                                 'DGR': [divGrowth.mean()], 'RDGR': [realDivGrowth.mean()]})
#
# filePath = r'C:\Users\CamHa\Downloads\Dividend Growth Rate.xlsx'
#
# with pd.ExcelWriter(filePath) as writer:
#     div.to_excel(writer, sheet_name='Div', index=False)
#     divYearly.to_excel(writer, sheet_name='Div Yearly', index=False)
#     divInfo.to_excel(writer, sheet_name='Dividend Info', index=False)


for symbol in tickers.symbols:
    print(symbol)
    div = tickers.tickers[symbol].dividends

    div = div.to_frame().reset_index()
    div['Year'] = div['Date'].dt.year
    div['Date'] = pd.to_datetime(div['Date']).dt.date
    div['Real Dividends'] = ''
    div['Price'] = ''
    div = div[div['Year'] != datetime.now().year]

    for x in range(0, div.shape[0]):
        div.iloc[x, 3] = cpi.inflate(div.iloc[x, 1], div.iloc[x, 2])
        print(div.iloc[x, 0])
        print(div.iloc[x, 0] + pd.Timedelta(days=1))
        try:
            div.iloc[x, 4] = yf.download(symbol, div.iloc[x, 0], div.iloc[x, 0] +
                                         pd.Timedelta(days=1))['Adj Close'].values[0]
        except IndexError:
            div.iloc[x, 4] = yf.download(symbol, div.iloc[x, 0] - pd.Timedelta(days=1),
                                         div.iloc[x, 0])['Adj Close'].values[0]

    div['Dividend Yield'] = div['Dividends'] / div['Price']

    divYearly = pd.concat([div.groupby('Year')['Dividends'].sum(), div.groupby('Year')['Real Dividends'].sum(),
                           div.groupby('Year')['Dividend Yield'].sum()], axis=1)
    divYearly = divYearly.iloc[1:]

    divGrowth = divYearly['Dividends'].pct_change().iloc[1:]
    realDivGrowth = divYearly['Real Dividends'].pct_change().iloc[1:]

    if i == 0:
        divInfo = pd.DataFrame({'Symbol': [symbol], 'Dividend Yield': [divYearly['Dividend Yield'].iloc[-1]],
                                'DGR': [divGrowth.mean()], 'RDGR': [realDivGrowth.mean()]})
    else:
        divInfo = pd.concat([divInfo, pd.DataFrame({'Symbol': [symbol],
                                                    'Dividend Yield': [divYearly['Dividend Yield'].iloc[-1]],
                                                    'DGR': [divGrowth.mean()],
                                                    'RDGR': [realDivGrowth.mean()]})])

    i += 1


filePath = r'C:\Users\CamHa\Downloads\Dividend Growth Rate.xlsx'

with pd.ExcelWriter(filePath) as writer:
    div.to_excel(writer, sheet_name='Div', index=False)
    divYearly.to_excel(writer, sheet_name='Div Yearly', index=False)
    divInfo.to_excel(writer, sheet_name='Dividend Info', index=False)


print('Complete')
