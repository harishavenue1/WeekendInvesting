import yfinance as yf
import pandas as pd
from datetime import datetime

cnx_200 = "ABB.NS,ACC.NS,APLAPOLLO.NS,AUBANK.NS,ADANIENT.NS,ADANIGREEN.NS,ADANIPORTS.NS,ADANIPOWER.NS,ATGL.NS,AWL.NS,ABCAPITAL.NS,ABFRL.NS,ALKEM.NS,AMBUJACEM.NS,APOLLOHOSP.NS,APOLLOTYRE.NS,ASHOKLEY.NS,ASIANPAINT.NS,ASTRAL.NS,AUROPHARMA.NS,DMART.NS,AXISBANK.NS,BAJAJ-AUTO.NS,BAJFINANCE.NS,BAJAJFINSV.NS,BAJAJHLDNG.NS,BALKRISIND.NS,BANDHANBNK.NS,BANKBARODA.NS,BANKINDIA.NS,BATAINDIA.NS,BERGEPAINT.NS,BDL.NS,BEL.NS,BHARATFORG.NS,BHEL.NS,BPCL.NS,BHARTIARTL.NS,BIOCON.NS,BOSCHLTD.NS,BRITANNIA.NS,CGPOWER.NS,CANBK.NS,CHOLAFIN.NS,CIPLA.NS,COALINDIA.NS,COFORGE.NS,COLPAL.NS,CONCOR.NS,COROMANDEL.NS,CROMPTON.NS,CUMMINSIND.NS,DLF.NS,DABUR.NS,DALBHARAT.NS,DEEPAKNTR.NS,DELHIVERY.NS,DEVYANI.NS,DIVISLAB.NS,DIXON.NS,LALPATHLAB.NS,DRREDDY.NS,EICHERMOT.NS,ESCORTS.NS,NYKAA.NS,FEDERALBNK.NS,FACT.NS,FORTIS.NS,GAIL.NS,GLAND.NS,GODREJCP.NS,GODREJPROP.NS,GRASIM.NS,FLUOROCHEM.NS,GUJGASLTD.NS,HCLTECH.NS,HDFCAMC.NS,HDFCBANK.NS,HDFCLIFE.NS,HAVELLS.NS,HEROMOTOCO.NS,HINDALCO.NS,HAL.NS,HINDPETRO.NS,HINDUNILVR.NS,ICICIBANK.NS,ICICIGI.NS,ICICIPRULI.NS,IDFCFIRSTB.NS,ITC.NS,INDIANB.NS,INDHOTEL.NS,IOC.NS,IRCTC.NS,IRFC.NS,IGL.NS,INDUSTOWER.NS,INDUSINDBK.NS,NAUKRI.NS,INFY.NS,INDIGO.NS,IPCALAB.NS,JSWENERGY.NS,JSWSTEEL.NS,JINDALSTEL.NS,JUBLFOOD.NS,KPITTECH.NS,KOTAKBANK.NS,L&TFH.NS,LTTS.NS,LICHSGFIN.NS,LT.NS,LAURUSLABS.NS,LICI.NS,LUPIN.NS,MRF.NS,LODHA.NS,M&MFIN.NS,M&M.NS,MARICO.NS,MARUTI.NS,MFSL.NS,MAXHEALTH.NS,MAZDOCK.NS,MSUMI.NS,MPHASIS.NS,MUTHOOTFIN.NS,NHPC.NS,NMDC.NS,NTPC.NS,NAVINFLUOR.NS,NESTLEIND.NS,OBEROIRLTY.NS,ONGC.NS,OIL.NS,PAYTM.NS,POLICYBZR.NS,PIIND.NS,PAGEIND.NS,PERSISTENT.NS,PETRONET.NS,PIDILITIND.NS,PEL.NS,POLYCAB.NS,POONAWALLA.NS,PFC.NS,POWERGRID.NS,PRESTIGE.NS,PGHH.NS,PNB.NS,RECLTD.NS,RVNL.NS,RELIANCE.NS,SBICARD.NS,SBILIFE.NS,SRF.NS,MOTHERSON.NS,SHREECEM.NS,SHRIRAMFIN.NS,SIEMENS.NS,SONACOMS.NS,SBIN.NS,SAIL.NS,SUNPHARMA.NS,SUNTV.NS,SYNGENE.NS,TVSMOTOR.NS,TATACHEM.NS,TATACOMM.NS,TCS.NS,TATACONSUM.NS,TATAELXSI.NS,TATAMTRDVR.NS,TATAMOTORS.NS,TATAPOWER.NS,TATASTEEL.NS,TECHM.NS,RAMCOCEM.NS,TITAN.NS,TORNTPHARM.NS,TORNTPOWER.NS,TRENT.NS,TIINDIA.NS,UPL.NS,ULTRACEMCO.NS,UNIONBANK.NS,UBL.NS,MCDOWELL-N.NS,VBL.NS,VEDL.NS,IDEA.NS,VOLTAS.NS,WIPRO.NS,YESBANK.NS,ZEEL.NS,ZOMATO.NS,ZYDUSLIFE.NS"
cnx_200_list = cnx_200.split(",")

# fetch YOY prices and price changes in percentage for all CNX200 stocks
map = {}
for stockName in cnx_200_list:
  try:
    closing_price_1 = round(yf.download(stockName, start='2023-01-17', end='2023-01-18').Close.iloc[0],2)
  except:
    closing_price_1 = round(yf.download(stockName, start='2023-01-15', end='2023-01-16').Close.iloc[0],2)

  try:
    closing_price_2 = round(yf.download(stockName, start='2024-01-17', end='2024-01-18').Close.iloc[0],2)
  except:
    closing_price_2 = round(yf.download(stockName, start='2024-01-15', end='2024-01-16').Close.iloc[0],2)

  YOY = round(((closing_price_2 - closing_price_1) * 100)/closing_price_1,2)
  map[stockName] = [closing_price_1, closing_price_2, YOY]


df = pd.DataFrame(map, index=["pricePrevYear","priceCurYear","YOY_pct_change"])

# transpose
df = df.transpose()

# add rank column based on yoy change
df['rank'] = df['YOY_pct_change'].rank(ascending=False)

# sort for rank
df_ranked = df.sort_values('rank')

# change index name to stocks
df_ranked['stocks']=df_ranked.index
df_ranked = df_ranked.reset_index(drop=True)

# save to csv/excel
writer = pd.ExcelWriter("./excel_"+str(datetime.now())+".xlsx", engine = 'xlsxwriter')
#df_ranked.to_csv('rankedDf.csv')
df_ranked.to_excel(writer, sheet_name = 'ranked', index=False)

# fetch only top 40 names to list from dataObject
top_20_ranked = df_ranked['stocks'].head(20).tolist()

# fetch mom price changes in percentage for top 20 ranked stocks from previous month to current month
map_mom = {}
for stockName in top_20_ranked:
  try:
    closing_price_1 = round(yf.download(stockName, start='2023-12-17', end='2023-12-18').Close.iloc[0],2)
  except:
    closing_price_1 = round(yf.download(stockName, start='2023-12-15', end='2023-12-16').Close.iloc[0],2)

  try:
    closing_price_2 = round(yf.download(stockName, start='2024-01-17', end='2024-01-18').Close.iloc[0],2)
  except:
    closing_price_2 = round(yf.download(stockName, start='2024-01-15', end='2024-01-16').Close.iloc[0],2)

  mom = round(((closing_price_2 - closing_price_1) * 100)/closing_price_1,2)
  map_mom[stockName] = [closing_price_1, closing_price_2, mom]

df_mom = pd.DataFrame(map_mom, index=["pricePrevYear","priceCurYear","mom_pct_change"])

# transpose
df_mom = df_mom.transpose()

# add rank column based on yoy change
df_mom['rank'] = df_mom['mom_pct_change'].rank(ascending=False)

# sort for rank
df_mom_ranked = df_mom.sort_values('rank')

# change index name to stocks
df_mom_ranked['stocks']=df_mom_ranked.index
df_mom_ranked = df_mom_ranked.reset_index(drop=True)

# save to csv
#df_mom_ranked.to_csv('ranked_df_mom.csv')
df_mom_ranked.to_excel(writer, sheet_name = 'mom', index=False)
writer.close()
