import os, logging, sys, json, csv, time
sys.path.insert(0, '../')

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Color, PatternFill, Font, Border
from pymongo import MongoClient
from multiprocessing.dummy import Pool as ThreadPool

import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np

from regression_high import process_regression_high
from regression_low import process_regression_low
from technical import ta_lib_data_df
from util.util import historical_data
from talib.abstract import *

directory = '../../output' + '/regression/' + time.strftime("%d%m%y-%H%M%S")
logname = '../../output' + '/regression/mllog' + time.strftime("%d%m%y-%H%M%S")
logging.basicConfig(filename=logname, filemode='a', stream=sys.stdout, level=logging.INFO)
log = logging.getLogger(__name__)

connection = MongoClient('localhost', 27017)
db = connection.Nsedata


def regression_ta_data(scrip):
    data = db.history.find_one({'dataset_code':scrip})
    if(data is None or (np.array(data['data'])).size < 1000):
        print('Missing or very less Data for ', scrip)
        return
        
    hsdate, hsopen, hshigh, hslow, hslast, hsclose, hsquantity, hsturnover = historical_data(data)   
    df = pd.DataFrame({
        'date': hsdate,
        'open': hsopen,
        'high': hshigh,
        'low': hslow,
        'close': hsclose,
        'volume': hsquantity,
        'turnover':hsturnover
    })
    df = df[['date','open','high','low','close','volume','turnover']]
    print(scrip)
    df=df.rename(columns = {'total trade quantity':'volume'})
    df=df.rename(columns = {'turnover (lacs)': 'turnover'})
    df['volume_pre'] = df['volume'].shift(+1)
    df['open_pre'] = df['open'].shift(+1)
    df['high_pre'] = df['high'].shift(+1)
    df['low_pre'] = df['low'].shift(+1)
    df['close_pre'] = df['close'].shift(+1)
    df['VOL_change'] = (((df['volume'] - df['volume_pre'])/df['volume_pre'])*100)
    df['PCT_change'] = (((df['close'] - df['close_pre'])/df['close_pre'])*100)
    df['PCT_day_change'] = (((df['close'] - df['open'])/df['open'])*100)
    df['HL_change'] = (((df['high'] - df['low'])/df['low'])*100).astype(int)
    df['CL_change'] = (((df['close'] - df['low'])/df['low'])*100).astype(int)
    df['CH_change'] = (((df['close'] - df['high'])/df['high'])*100).astype(int)
    df['OL_change'] = (((df['low'] - df['open'])/df['open'])*100).astype(float)
    df['HO_change'] = (((df['high'] - df['open'])/df['open'])*100).astype(float)
    df['bar_high'] = np.where(df['close'] > df['open'], df['close'], df['open'])
    df['bar_low'] = np.where(df['close'] > df['open'], df['open'], df['close'])
    df['bar_high_pre'] = np.where(df['close_pre'] > df['open_pre'], df['close_pre'], df['open_pre'])
    df['bar_low_pre'] = np.where(df['close_pre'] > df['open_pre'], df['open_pre'], df['close_pre'])
    df['uptrend'] = np.where((df['bar_high'] >  df['bar_high_pre']) & (df['high'] > df['high_pre']), 1, 0)
    df['downtrend'] = np.where((df['bar_low'] <  df['bar_low_pre']) & (df['low'] < df['low_pre']), -1, 0)
    
    df.dropna(inplace=True)
    df['EMA9'] = EMA(df,9)
    df['EMA21'] = EMA(df,21)
    
    size = int(int(np.floor(df.shape[0]))/3)
    for x in range(size):
        buy, sell, trend, yearHighChange, yearLowChange = ta_lib_data_df(scrip, df, False) 
        process_regression_high(scrip, df, buy, sell, trend, yearHighChange, yearLowChange, directory)
        process_regression_low(scrip, df, buy, sell, trend, yearHighChange, yearLowChange, directory)
        df = df[:-1]

def calculateParallel(threads=1):
    pool = ThreadPool(threads)
    with open('../../data-import/nselist/test.csv') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        scrips = []
        for row in readCSV:
            scrips.append(row[0].replace('&','').replace('-','_'))   
        scrips.sort()
        pool.map(regression_ta_data, scrips)   
                     
if __name__ == "__main__":
    calculateParallel(1)
    connection.close()