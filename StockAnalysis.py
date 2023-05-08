# Import Packages
# import subprocess
import pandas as pd
import yfinance as yf


# with open('requirements.txt', 'w') as file_:
#     subprocess.Popen(['pip', 'freeze', '-l'], stdout=file_).communicate()
# print('requirements.txt saved')


tickers = yf.Tickers('VTSAX VTI VOO VTIAX VXUS SCHD VIG VYM DGRO VBTLX BND VTEB VTIP')
i = 0
divGrowthRate = pd.DataFrame()

for symbol in tickers.symbols:
    div = tickers.tickers[symbol].dividends

    div = div.to_frame().reset_index()
    div['Year'] = div['Date'].dt.year

    divYield = div.groupby('Year')['Dividends'].sum()

    divYield = divYield.iloc[1:]
    divYield = divYield.iloc[:-1]

    divGrowth = divYield.pct_change().iloc[1:]

    if i == 0:
        divGrowthRate = pd.DataFrame({'Symbol': [symbol], 'DGR': [divGrowth.mean()]})
    else:
        divGrowthRate = pd.concat([divGrowthRate, pd.DataFrame({'Symbol': [symbol], 'DGR': [divGrowth.mean()]})])

    i += 1


filePath = r'C:\Users\CamHa\Downloads\Dividend Growth Rate.xlsx'

with pd.ExcelWriter(filePath) as writer:
    divGrowthRate.to_excel(writer, sheet_name='Div Growth Rate', index=False)


print('Complete')
