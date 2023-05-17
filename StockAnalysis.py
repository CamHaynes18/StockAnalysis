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


mode = 2

if mode == 1:
    symbol = 'VOO'
    tickers = yf.Tickers(symbol)
    div = tickers.tickers[symbol].dividends

    div = div.to_frame().reset_index()
    div['Year'] = div['Date'].dt.year
    div['Date'] = pd.to_datetime(div['Date']).dt.date
    div['Real Dividends'] = ''
    div['Price'] = ''
    div['Adjusted Price'] = ''
    div['Real Price'] = ''
    div['Real Adjusted Price'] = ''
    div = div[div['Year'] != datetime.now().year]

    for x in range(0, div.shape[0]):
        div.iloc[x, 3] = cpi.inflate(div.iloc[x, 1], div.iloc[x, 2])
        try:
            div.iloc[x, 4] = yf.download(symbol, div.iloc[x, 0], div.iloc[x, 0] +
                                         pd.Timedelta(days=1), progress=False)['Close'].values[0]
            div.iloc[x, 5] = yf.download(symbol, div.iloc[x, 0], div.iloc[x, 0] +
                                         pd.Timedelta(days=1), progress=False)['Adj Close'].values[0]
        except IndexError:
            div.iloc[x, 4] = yf.download(symbol, div.iloc[x, 0] -
                                         pd.Timedelta(days=1), div.iloc[x, 0], progress=False)['Close'].values[0]
            div.iloc[x, 5] = yf.download(symbol, div.iloc[x, 0] -
                                         pd.Timedelta(days=1), div.iloc[x, 0], progress=False)['Adj Close'].values[0]
        div.iloc[x, 6] = cpi.inflate(div.iloc[x, 4], div.iloc[x, 2])
        div.iloc[x, 7] = cpi.inflate(div.iloc[x, 5], div.iloc[x, 2])

    div['Dividend Yield'] = div['Dividends'] / div['Price']

    divYearly = pd.concat([div.groupby('Year')['Dividends'].sum(), div.groupby('Year')['Real Dividends'].sum(),
                           div.groupby('Year')['Dividend Yield'].sum()], axis=1)
    divYearly = divYearly.iloc[1:]

    divGrowth = divYearly['Dividends'].pct_change().iloc[1:]
    realDivGrowth = divYearly['Real Dividends'].pct_change().iloc[1:]
    changeInValuation = div['Price'].pct_change().iloc[1:]
    adjustedChangeInValuation = div['Adjusted Price'].pct_change().iloc[1:]
    realChangeInValuation = div['Real Price'].pct_change().iloc[1:]
    realAdjustedChangeInValuation = div['Real Adjusted Price'].pct_change().iloc[1:]

    divInfo = pd.DataFrame({'Symbol': [symbol], 'Dividend Yield': [divYearly['Dividend Yield'].iloc[-1]],
                            'DGR': [divGrowth.mean()], 'RDGR': [realDivGrowth.mean()],
                            'CV': [changeInValuation.mean()],
                            'ACV': [adjustedChangeInValuation.mean()],
                            'RCV': [realChangeInValuation.mean()],
                            'RACV': [realAdjustedChangeInValuation.mean()]})

    filePath = r'C:\Users\CamHa\Downloads\Dividend Growth Rate - Single.xlsx'

    with pd.ExcelWriter(filePath) as writer:
        div.to_excel(writer, sheet_name='Div', index=False)
        divYearly.to_excel(writer, sheet_name='Div Yearly')
        divInfo.to_excel(writer, sheet_name='Dividend Info', index=False)

else:
    tickers = yf.Tickers(
        'VBTLX VGIT VTIAX VXUS VTEB VOHIX VTSAX VTI VFIAX VOO VIG SCHD VYM VVIAX VTV VIMAX VO VMVAX VOE '
        'VSMAX VB VSIAX VBR VGSLX')
    i = 0
    divInfo = pd.DataFrame()

    for symbol in tickers.symbols:
        print(symbol)
        div = tickers.tickers[symbol].dividends

        div = div.to_frame().reset_index()
        div['Year'] = div['Date'].dt.year
        div['Date'] = pd.to_datetime(div['Date']).dt.date
        div['Real Dividends'] = ''
        div['Price'] = ''
        div['Adjusted Price'] = ''
        div['Real Price'] = ''
        div['Real Adjusted Price'] = ''
        div = div[div['Year'] != datetime.now().year]

        for x in range(0, div.shape[0]):
            div.iloc[x, 3] = cpi.inflate(div.iloc[x, 1], div.iloc[x, 2])
            try:
                div.iloc[x, 4] = yf.download(symbol, div.iloc[x, 0], div.iloc[x, 0] +
                                             pd.Timedelta(days=1), progress=False)['Close'].values[0]
                div.iloc[x, 5] = yf.download(symbol, div.iloc[x, 0], div.iloc[x, 0] +
                                             pd.Timedelta(days=1), progress=False)['Adj Close'].values[0]
            except IndexError:
                div.iloc[x, 4] = yf.download(symbol, div.iloc[x, 0] -
                                             pd.Timedelta(days=1), div.iloc[x, 0], progress=False)['Close'].values[0]
                div.iloc[x, 5] = yf.download(symbol, div.iloc[x, 0] -
                                             pd.Timedelta(days=1), div.iloc[x, 0], progress=False)['Adj Close'].values[
                    0]
            div.iloc[x, 6] = cpi.inflate(div.iloc[x, 4], div.iloc[x, 2])
            div.iloc[x, 7] = cpi.inflate(div.iloc[x, 5], div.iloc[x, 2])

        div['Dividend Yield'] = div['Dividends'] / div['Price']

        divYearly = pd.concat([div.groupby('Year')['Dividends'].sum(), div.groupby('Year')['Real Dividends'].sum(),
                               div.groupby('Year')['Dividend Yield'].sum()], axis=1)
        divYearly = divYearly.iloc[1:]

        divGrowth = divYearly['Dividends'].pct_change().iloc[1:]
        realDivGrowth = divYearly['Real Dividends'].pct_change().iloc[1:]
        changeInValuation = div['Price'].pct_change().iloc[1:]
        adjustedChangeInValuation = div['Adjusted Price'].pct_change().iloc[1:]
        realChangeInValuation = div['Real Price'].pct_change().iloc[1:]
        realAdjustedChangeInValuation = div['Real Adjusted Price'].pct_change().iloc[1:]

        if i == 0:
            divInfo = pd.DataFrame({'Symbol': [symbol], 'Dividend Yield': [divYearly['Dividend Yield'].iloc[-1]],
                                    'DGR': [divGrowth.mean()], 'RDGR': [realDivGrowth.mean()],
                                    'CV': [changeInValuation.mean()],
                                    'ACV': [adjustedChangeInValuation.mean()],
                                    'RCV': [realChangeInValuation.mean()],
                                    'RACV': [realAdjustedChangeInValuation.mean()]})
        else:
            divInfo = pd.concat([divInfo, pd.DataFrame({'Symbol': [symbol],
                                                        'Dividend Yield': [divYearly['Dividend Yield'].iloc[-1]],
                                                        'DGR': [divGrowth.mean()],
                                                        'RDGR': [realDivGrowth.mean()],
                                                        'CV': [changeInValuation.mean()],
                                                        'ACV': [adjustedChangeInValuation.mean()],
                                                        'RCV': [realChangeInValuation.mean()],
                                                        'RACV': [realAdjustedChangeInValuation.mean()]})])

        i += 1

        filePath = r'C:\Users\CamHa\Downloads\Dividend Growth Rate - Group.xlsx'

        with pd.ExcelWriter(filePath) as writer:
            divInfo.to_excel(writer, sheet_name='Dividend Info', index=False)


print('Complete')
