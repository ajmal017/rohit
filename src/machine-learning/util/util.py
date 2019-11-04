import csv
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Color, PatternFill, Font, Border
from pymongo import MongoClient
from multiprocessing.dummy import Pool as ThreadPool

import quandl, math, time
from nsepy import get_history
from nsepy.derivatives import get_expiry_date
from datetime import date
import datetime   
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np

connection = MongoClient('localhost',27017)
db = connection.Nsedata

soft=False
NIFTY_HIGH = False
NIFTY_LOW = False
BUY_VERY_LESS_DATA=True
SELL_VERY_LESS_DATA=True
MARKET_IN_UPTREND=False
MARKET_IN_DOWNTREND=False
TEST = False

buyMLP = 0.1
buyMLP_MIN = 0
buyKN = 0.1
buyKN_MIN = 0
sellMLP = -0.1
sellMLP_MIN = 0
sellKN = -0.1
sellKN_MIN = 0

def patterns_to_dict(filename):  
    tempDict = {}
    count = 0
    with open(filename) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            try:
                if (count != 0):
                    dictValue = {}
                    dictValue['avg'] = row[1]
                    dictValue['count'] = row[2]
                    dictValue['countgt'] = row[3]
                    dictValue['countlt'] = row[4]
                    tempDict[row[0]] = dictValue
                count = count + 1
            except:
                pass
    return tempDict

filter345sell = patterns_to_dict('../../data-import/nselist/filter-345-sell.csv')
filtersell = patterns_to_dict('../../data-import/nselist/filter-sell.csv')
filterpctchangesell = patterns_to_dict('../../data-import/nselist/filter-pct-change-sell.csv')
filterallsell = patterns_to_dict('../../data-import/nselist/filter-all-sell.csv')
filtertechsell = patterns_to_dict('../../data-import/nselist/filter-tech-buy.csv')
filtertechallsell = patterns_to_dict('../../data-import/nselist/filter-tech-all-buy.csv')
filtertechallpctchangesell = patterns_to_dict('../../data-import/nselist/filter-tech-all-pct-change-buy.csv')
filter345buy = patterns_to_dict('../../data-import/nselist/filter-345-buy.csv')
filterbuy = patterns_to_dict('../../data-import/nselist/filter-buy.csv')
filterpctchangebuy = patterns_to_dict('../../data-import/nselist/filter-pct-change-buy.csv')
filterallbuy = patterns_to_dict('../../data-import/nselist/filter-all-buy.csv')
filtertechbuy = patterns_to_dict('../../data-import/nselist/filter-tech-buy.csv')
filtertechallbuy = patterns_to_dict('../../data-import/nselist/filter-tech-all-buy.csv')
filtertechallpctchangebuy = patterns_to_dict('../../data-import/nselist/filter-tech-all-pct-change-buy.csv')

    
def add_in_csv(regression_data, regressionResult, ws=None, filter=None, filter1=None, filter2=None, filter3=None, filter4=None, filter5=None, filter6=None):
    if(TEST != True):
        if(is_algo_buy(regression_data) and (filter is not None)):
            if '[MLBuy]' not in regression_data['filter']:
                regression_data['filter'] ='[MLBuy]' + ':'
        if(is_algo_sell(regression_data) and (filter is not None)):
            if '[MLSell]' not in regression_data['filter']:
                regression_data['filter'] ='[MLSell]' + ':'
        if ((filter is not None) and (filter not in regression_data['filter'])):
            list = regression_data['filter'].partition(']:')
            if(list[1] == ']:'):
                regression_data['filter'] = list[0] + list[1] + filter + ','  + list[2]
            else:
                regression_data['filter'] = filter + ','  + regression_data['filter']
            if ('P@[' in str(regression_data['sellIndia'])) and (('buy' or 'Buy') in regression_data['filter']):
                if '***SELLPATTERN***' not in regression_data['filter']:
                   regression_data['filter'] = regression_data['filter'] + '***SELLPATTERN***' 
            if ('P@[' in str(regression_data['buyIndia'])) and (('sell' or 'Sell') in regression_data['filter']):
                if '***BUYPATTERN***' not in regression_data['filter']:
                   regression_data['filter'] = regression_data['filter'] + '***BUYPATTERN***'            
                
        if ((filter1 is not None) and (filter1 not in regression_data['filter1'])):
            regression_data['filter1'] = filter1 + ',' + regression_data['filter1']
        if ((filter2 is not None) and (filter2 not in regression_data['filter2'])):
            regression_data['filter2'] = filter2 + ',' + regression_data['filter2']
        if ((filter3 is not None) and (filter3 not in regression_data['filter3'])):
            regression_data['filter3'] = filter3 + ',' + regression_data['filter3']
        if ((filter4 is not None) and (filter4 not in regression_data['filter4'])):
            regression_data['filter4'] = filter4 + ',' + regression_data['filter4']  
        if ((filter5 is not None) and (filter5 not in regression_data['filter5'])):
            regression_data['filter5'] = filter5 + ',' + regression_data['filter5'] 
        tempRegressionResult = regressionResult.copy() 
        tempRegressionResult.append(regression_data['filter1'])
        tempRegressionResult.append(regression_data['filter2'])
        tempRegressionResult.append(regression_data['filter3'])
        tempRegressionResult.append(regression_data['filter4'])
        tempRegressionResult.append(regression_data['filter5'])
        tempRegressionResult.append(regression_data['filter'])
        tempRegressionResult.append(regression_data['filter_345_avg'])
        tempRegressionResult.append(regression_data['filter_345_count'])
        tempRegressionResult.append(regression_data['filter_345_pct'])
        tempRegressionResult.append(regression_data['filter_avg'])
        tempRegressionResult.append(regression_data['filter_count'])
        tempRegressionResult.append(regression_data['filter_pct'])
        tempRegressionResult.append(regression_data['filter_pct_change_avg'])
        tempRegressionResult.append(regression_data['filter_pct_change_count'])
        tempRegressionResult.append(regression_data['filter_pct_change_pct'])
        tempRegressionResult.append(regression_data['filter_all_avg'])
        tempRegressionResult.append(regression_data['filter_all_count'])
        tempRegressionResult.append(regression_data['filter_all_pct'])
        tempRegressionResult.append(regression_data['filter_tech_avg'])
        tempRegressionResult.append(regression_data['filter_tech_count'])
        tempRegressionResult.append(regression_data['filter_tech_pct'])
        tempRegressionResult.append(regression_data['filter_tech_all_avg'])
        tempRegressionResult.append(regression_data['filter_tech_all_count'])
        tempRegressionResult.append(regression_data['filter_tech_all_pct'])
        tempRegressionResult.append(regression_data['filter_tech_all_pct_change_avg'])
        tempRegressionResult.append(regression_data['filter_tech_all_pct_change_count'])
        tempRegressionResult.append(regression_data['filter_tech_all_pct_change_pct'])
        ws.append(tempRegressionResult) if (ws is not None) else False
        if(db.resultScripFutures.find_one({'scrip':regression_data['scrip']}) is None):
            db.resultScripFutures.insert_one({
                "scrip": regression_data['scrip'],
                "date": regression_data['date']
                })
    else:
        if ((filter is not None) and (filter not in regression_data['filterTest'])):
            list = regression_data['filterTest'].partition(']:')
            if(list[1] == ']:'):
                regression_data['filterTest'] = list[0] + list[1] + filter + ','  + list[2]
            else:
                regression_data['filterTest'] = filter + ','  + regression_data['filterTest']
            if ('P@[' in str(regression_data['sellIndia'])) and (('buy' or 'Buy') in regression_data['filterTest']):
                if '***SELLPATTERN***' not in regression_data['filterTest']:
                   regression_data['filterTest'] = regression_data['filterTest'] + '***SELLPATTERN***'
            if ('P@[' in str(regression_data['buyIndia'])) and (('sell' or 'Sell') in regression_data['filterTest']):
                if '***BUYPATTERN***' not in regression_data['filterTest']:
                   regression_data['filterTest'] = regression_data['filterTest'] + '***BUYPATTERN***' 
        if ((filter1 is not None) and (filter1 not in regression_data['filter1'])):
            regression_data['filter1'] = filter1 + ',' + regression_data['filter1']
        if ((filter2 is not None) and (filter2 not in regression_data['filter2'])):
            regression_data['filter2'] = filter2 + ',' + regression_data['filter2']
        if ((filter3 is not None) and (filter3 not in regression_data['filter3'])):
            regression_data['filter3'] = filter3 + ',' + regression_data['filter3']
        if ((filter4 is not None) and (filter4 not in regression_data['filter4'])):
            regression_data['filter4'] = filter4 + ',' + regression_data['filter4']  
        if ((filter5 is not None) and (filter5 not in regression_data['filter5'])):
            regression_data['filter5'] = filter5 + ',' + regression_data['filter5'] 
        if ((filter6 is not None) and (filter6 not in regression_data['filter6'])):
            regression_data['filter6'] = filter6 + ',' + regression_data['filter6']
        
def add_in_csv_hist_pattern(regression_data, regressionResult, ws, filter, avg, count):
    if(TEST != True):
        if ((filter is not None) and (filter not in regression_data['filter'])):
            regression_data['filter'] = regression_data['filter'] + filter + ','
        tempRegressionResult = regressionResult.copy() 
        tempRegressionResult.append(regression_data['filter'])
        tempRegressionResult.append(avg)
        tempRegressionResult.append(count)
        ws.append(tempRegressionResult) if (ws is not None) else False
    else:
        if ((filter is not None) and (filter not in regression_data['filterTest'])):
            regression_data['filterTest'] = regression_data['filterTest'] + filter + ','
        tempRegressionResult = regressionResult.copy() 
        tempRegressionResult.append(regression_data['filterTest'])
        tempRegressionResult.append(avg)
        tempRegressionResult.append(count)
        ws.append(tempRegressionResult) if (ws is not None) else False

def is_any_reg_algo_gt1(regression_data):
    if((regression_data['mlpValue_reg'] >= 1) 
        or (regression_data['kNeighboursValue_reg'] >= 1)
        or (regression_data['mlpValue_reg_other'] >= 1) 
        or (regression_data['kNeighboursValue_reg_other'] >= 1)
        ):
        return True
    else:
        return False
    
def is_any_reg_algo_lt_minus1(regression_data):
    if((regression_data['mlpValue_reg'] <= -1) 
        or (regression_data['kNeighboursValue_reg'] <= -1)
        or (regression_data['mlpValue_reg_other'] <= -1) 
        or (regression_data['kNeighboursValue_reg_other'] <= -1)
        ):
        return True
    else:
        return False

def is_algo_buy(regression_data, resticted=False):
    if TEST:
        return True
#     if not_from_rule == False:
#         return True
    REG_MIN = -1.5
    CLA_MIN = -1
    if resticted:
        REG_MIN = -0.5
        CLA_MIN = 0
    if((regression_data['mlpValue_reg'] >= REG_MIN) and (regression_data['kNeighboursValue_reg'] >= REG_MIN)
        and (regression_data['mlpValue_reg_other'] >= REG_MIN) and (regression_data['kNeighboursValue_reg_other'] >= REG_MIN)
        and (regression_data['mlpValue_cla'] >= CLA_MIN) and (regression_data['kNeighboursValue_cla'] >= CLA_MIN)
        and (regression_data['mlpValue_cla_other'] >= CLA_MIN) and (regression_data['kNeighboursValue_cla_other'] >= CLA_MIN)
        ):
        if((regression_data['mlpValue_reg'] > buyMLP and regression_data['kNeighboursValue_reg'] >= buyKN_MIN) 
            or (regression_data['mlpValue_reg'] >= buyMLP_MIN and regression_data['kNeighboursValue_reg'] > buyKN)
            or (regression_data['mlpValue_reg'] > 0.5 and regression_data['kNeighboursValue_reg'] > 0.5)
            or ((regression_data['mlpValue_reg'] + regression_data['kNeighboursValue_reg'] + regression_data['mlpValue_reg_other'] + regression_data['kNeighboursValue_reg_other']) > 4)
            ):
            if((regression_data['mlpValue_cla'] < 0) or (regression_data['kNeighboursValue_cla'] < 0)
                or (regression_data['mlpValue_cla_other'] < 0) or (regression_data['kNeighboursValue_cla_other'] < 0)):
                if((regression_data['mlpValue_reg'] < 0) or (regression_data['kNeighboursValue_reg'] < 0)
                    or (regression_data['mlpValue_reg_other'] < 0) or (regression_data['kNeighboursValue_reg_other'] < 0)):
                    return False
            if resticted:
                if((regression_data['mlpValue_reg_other'] >= 0 or regression_data['kNeighboursValue_reg_other'] >= 0)):
                    return True
            else:
                return True
    return False   
    
def is_algo_sell(regression_data, resticted=False):
    if TEST:
        return True
#     if not_from_rule == False:
#         return True
    REG_MAX = 1.5
    CLA_MAX = 1
    if resticted:
        REG_MAX = 0.5
        CLA_MAX = 0
    if((regression_data['mlpValue_reg'] <= REG_MAX) and (regression_data['kNeighboursValue_reg'] <= REG_MAX)
        and (regression_data['mlpValue_reg_other'] <= REG_MAX) and (regression_data['kNeighboursValue_reg_other'] <= REG_MAX)
        and (regression_data['mlpValue_cla'] <= CLA_MAX) and (regression_data['kNeighboursValue_cla'] <= CLA_MAX)
        and (regression_data['mlpValue_cla_other'] <= CLA_MAX) and (regression_data['kNeighboursValue_cla_other'] <= CLA_MAX)
        ):
        if((regression_data['mlpValue_reg'] < sellMLP and regression_data['kNeighboursValue_reg'] <= sellKN_MIN)
            or (regression_data['mlpValue_reg'] <= sellMLP_MIN and regression_data['kNeighboursValue_reg'] < sellKN)
            or (regression_data['mlpValue_reg'] < -0.5 and regression_data['kNeighboursValue_reg'] < -0.5)
            or ((regression_data['mlpValue_reg'] + regression_data['kNeighboursValue_reg'] + regression_data['mlpValue_reg_other'] + regression_data['kNeighboursValue_reg_other']) < -4)
            ):
            if((regression_data['mlpValue_cla'] > 0) or (regression_data['kNeighboursValue_cla'] > 0)
                or (regression_data['mlpValue_cla_other'] > 0) or (regression_data['kNeighboursValue_cla_other'] > 0)):
                if((regression_data['mlpValue_reg'] > 0) or (regression_data['kNeighboursValue_reg'] > 0)
                    or (regression_data['mlpValue_reg_other'] > 0) or (regression_data['kNeighboursValue_reg_other'] > 0)):
                    return False
            if resticted:
                if((-5 < regression_data['PCT_day_change'] < 0) and (regression_data['mlpValue_reg_other'] <= 0 or regression_data['kNeighboursValue_reg_other'] <= 0)):
                    return True
            else:
                return True
    return False

def is_algo_buy_classifier(regression_data, resticted=False):
    REG_MIN = -1.5
    CLA_MIN = -1
    if resticted:
        REG_MIN = -0.5
        CLA_MIN = 0
    if((regression_data['mlpValue_reg'] >= REG_MIN) and (regression_data['kNeighboursValue_reg'] >= REG_MIN)
        and (regression_data['mlpValue_reg_other'] >= REG_MIN) and (regression_data['kNeighboursValue_reg_other'] >= REG_MIN)
        and (regression_data['mlpValue_cla'] >= CLA_MIN) and (regression_data['kNeighboursValue_cla'] >= CLA_MIN)
        and (regression_data['mlpValue_cla_other'] >= CLA_MIN) and (regression_data['kNeighboursValue_cla_other'] >= CLA_MIN)
        ):
        if((regression_data['mlpValue_cla'] >= buyMLP and regression_data['kNeighboursValue_cla'] >= buyKN_MIN) 
            or (regression_data['mlpValue_cla'] >= buyMLP_MIN and regression_data['kNeighboursValue_cla'] >= buyKN)
            or (regression_data['mlpValue_cla'] >= 0.5 and regression_data['kNeighboursValue_cla'] >= 0.5)
            or ((regression_data['mlpValue_cla'] + regression_data['kNeighboursValue_cla'] + regression_data['mlpValue_cla_other'] + regression_data['kNeighboursValue_cla_other']) > 6)
            ):
            if((0 < regression_data['PCT_day_change'] < 5) and (regression_data['mlpValue_cla_other'] >= 0 or regression_data['kNeighboursValue_cla_other'] >= 0)):
                return True
            elif(regression_data['PCT_day_change'] <= 0 or regression_data['PCT_day_change'] >=5):
                return True
    return False   
    
def is_algo_sell_classifier(regression_data, resticted=False):
    REG_MAX = 1.5
    CLA_MAX = 1
    if resticted:
        REG_MAX = 0.5
        CLA_MAX = 0
    if((regression_data['mlpValue_reg'] <= REG_MAX) and (regression_data['kNeighboursValue_reg'] <= REG_MAX)
        and (regression_data['mlpValue_reg_other'] <= REG_MAX) and (regression_data['kNeighboursValue_reg_other'] <= REG_MAX)
        and (regression_data['mlpValue_cla'] <= CLA_MAX) and (regression_data['kNeighboursValue_cla'] <= CLA_MAX)
        and (regression_data['mlpValue_cla_other'] <= CLA_MAX) and (regression_data['kNeighboursValue_cla_other'] <= CLA_MAX)
        ):
        if((regression_data['mlpValue_cla'] <= sellMLP and regression_data['kNeighboursValue_cla'] <= sellKN_MIN)
            or (regression_data['mlpValue_cla'] <= sellMLP_MIN and regression_data['kNeighboursValue_cla'] <= sellKN)
            or (regression_data['mlpValue_cla'] <= -0.5 and regression_data['kNeighboursValue_cla'] <= -0.5)
            or ((regression_data['mlpValue_cla'] + regression_data['kNeighboursValue_cla'] + regression_data['mlpValue_cla_other'] + regression_data['kNeighboursValue_cla_other']) < -6)
            ):
            if((-5 < regression_data['PCT_day_change'] < 0) and (regression_data['mlpValue_cla_other'] <= 0 or regression_data['kNeighboursValue_cla_other'] <= 0)):
                return True
            elif(regression_data['PCT_day_change'] >= 0 or regression_data['PCT_day_change'] <=-5):
                return True
    return False

def get_reg_or_cla(regression_data, reg=True):
    if(TEST != True):
        if reg:
            mlpValue = regression_data['mlpValue_reg']
            kNeighboursValue = regression_data['kNeighboursValue_reg']
            return mlpValue, kNeighboursValue
        else:
            mlpValue = regression_data['mlpValue_cla']
            kNeighboursValue = regression_data['kNeighboursValue_cla']
            return mlpValue, kNeighboursValue
    else:
        return 0, 0
    
def get_reg_or_cla_other(regression_data, reg=True):
    if(TEST != True):
        if reg:
            mlpValue_other = regression_data['mlpValue_reg_other']
            kNeighboursValue_other = regression_data['kNeighboursValue_reg_other']
            return mlpValue_other, kNeighboursValue_other
        else:
            mlpValue_other = regression_data['mlpValue_cla_other']
            kNeighboursValue_other = regression_data['kNeighboursValue_cla_other']
            return mlpValue_other, kNeighboursValue_other
    else:
        return 0, 0

def getScore(vol_change, pct_change):
    try:
        return float(vol_change)/float(pct_change) 
    except ZeroDivisionError:
        return 0

def all_day_pct_change_negative(regression_data):
    if(regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] <= 0
        and regression_data['forecast_day_PCT3_change'] <= 0
        and regression_data['forecast_day_PCT4_change'] <= 0
        and regression_data['forecast_day_PCT5_change'] <= 0
        and regression_data['forecast_day_PCT7_change'] <= 0
        and regression_data['forecast_day_PCT10_change'] <= 0):
        return True;
    else:
        return False;
    
def all_day_pct_change_positive(regression_data):
    if(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > -0.5
        and regression_data['forecast_day_PCT3_change'] > -0.5
        and regression_data['forecast_day_PCT4_change'] > -0.5
        and regression_data['forecast_day_PCT5_change'] > -0.5
        and regression_data['forecast_day_PCT7_change'] > -0.5
        and regression_data['forecast_day_PCT10_change'] > -0.5):
        return True;
    else:
        return False;

def all_day_pct_change_negative_except_today(regression_data):
    if(regression_data['forecast_day_PCT2_change'] <= 0
        and regression_data['forecast_day_PCT3_change'] <= 0
        and regression_data['forecast_day_PCT4_change'] <= 0
        and regression_data['forecast_day_PCT5_change'] <= 0
        and regression_data['forecast_day_PCT7_change'] <= 0
        and regression_data['forecast_day_PCT10_change'] <= 0):
        return True;
    else:
        return False
    
def all_day_pct_change_positive_except_today(regression_data):
    if(regression_data['forecast_day_PCT2_change'] >= 0
        and regression_data['forecast_day_PCT3_change'] >= 0
        and regression_data['forecast_day_PCT4_change'] >= 0
        and regression_data['forecast_day_PCT5_change'] >= 0
        and regression_data['forecast_day_PCT7_change'] >= 0
        and regression_data['forecast_day_PCT10_change'] >= 0):
        return True;
    else:
        return False; 

def high_counter(regression_data):
    count = 0
    if(regression_data['high'] > regression_data['high_pre1']):
        count = count + 1;
    if(regression_data['high_pre1'] > regression_data['high_pre2']):
        count = count + 1;
    if(regression_data['high_pre2'] > regression_data['high_pre3']):
        count = count + 1;
    if(regression_data['high_pre3'] > regression_data['high_pre4']):
        count = count + 1;
    return count

def low_counter(regression_data):
    count = 0
    if(regression_data['low'] < regression_data['low_pre1']):
        count = count + 1;
    if(regression_data['low_pre1'] < regression_data['low_pre2']):
        count = count + 1;
    if(regression_data['low_pre2'] < regression_data['low_pre3']):
        count = count + 1;
    if(regression_data['low_pre3'] < regression_data['low_pre4']):
        count = count + 1;
    return count

def pct_day_change_counter(regression_data):
    countGt = 0
    countLt = 0
    if(regression_data['PCT_day_change'] > 0):
        countGt = countGt + 1
    if(regression_data['PCT_day_change'] < 0):
        countLt = countLt + 1
        
    if(regression_data['PCT_day_change_pre1'] > 0):
        countGt = countGt + 1
    if(regression_data['PCT_day_change_pre1'] < 0):
        countLt = countLt + 1
        
    if(regression_data['PCT_day_change_pre2'] > 0):
        countGt = countGt + 1
    if(regression_data['PCT_day_change_pre2'] < 0):
        countLt = countLt + 1
        
    if(regression_data['PCT_day_change_pre3'] > 0):
        countGt = countGt + 1
    if(regression_data['PCT_day_change_pre3'] < 0):
        countLt = countLt + 1
    
    return countGt, countLt

def pct_day_change_trend(regression_data):
    countGt = 0
    countLt = 0
    if(regression_data['PCT_day_change'] > 0):
        countGt = countGt + 1
    else:
        countLt = countLt - 1
        
    if(regression_data['PCT_day_change_pre1'] > 0):
        countGt = countGt + 1
    else:
        countLt = countLt - 1
        
    if(regression_data['PCT_day_change_pre2'] > 0):
        countGt = countGt + 1
    else:
        countLt = countLt - 1
        
    if(regression_data['PCT_day_change_pre3'] > 0):
        countGt = countGt + 1
    else:
        countLt = countLt - 1
        
    if(regression_data['PCT_day_change_pre4'] > 0):
        countGt = countGt + 1
    else:
        countLt = countLt - 1
        
    if (high_counter(regression_data) >= 2 and countGt >= 4):
        return countGt
    elif(low_counter(regression_data) >= 2 and countLt <= -4):
        return countLt
    else:
        return 0
    
    return count
    
def pct_change_negative_trend(regression_data):
    if (regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0
        and ((regression_data['low'] < regression_data['low_pre1'] and regression_data['low_pre1'] < regression_data['low_pre2'])
            or (regression_data['bar_low'] < regression_data['bar_low_pre1'] and regression_data['bar_low_pre1'] < regression_data['bar_low_pre2'])
            )
        and low_counter(regression_data) >= 3
        ):
        pct_change_list = [regression_data['forecast_day_PCT_change'],
                           regression_data['forecast_day_PCT2_change'],
                           regression_data['forecast_day_PCT3_change'],
                           regression_data['forecast_day_PCT4_change'],
                           regression_data['forecast_day_PCT5_change'],
                           regression_data['forecast_day_PCT7_change'],
                           regression_data['forecast_day_PCT10_change']
                           ]
        trend_change = False
        trend = True
        for pct_change in pct_change_list:
            if pct_change < 0 and trend_change == False:
                trend = True
            elif pct_change > 0 and trend_change == True: 
                trend = True    
            elif pct_change > 0 and trend_change == False:
                trend_change = True
                trend = True
            else:
                trend = False
        if trend:
            if(regression_data['bar_low'] < regression_data['bar_low_pre1']
                and regression_data['bar_low_pre1'] < regression_data['bar_low_pre2']
                ):
                return '(downTrend)'
            else:
                return '(downTrend-Risky)'
                
    elif (#regression_data['forecast_day_PCT_change'] < 0
        #regression_data['forecast_day_PCT2_change'] < 0
        regression_data['forecast_day_PCT3_change'] < 0
        and low_counter(regression_data) >= 3
        ):
        pct_change_list = [#regression_data['forecast_day_PCT_change'],
                           #regression_data['forecast_day_PCT2_change'],
                           regression_data['forecast_day_PCT3_change'],
                           regression_data['forecast_day_PCT4_change'],
                           regression_data['forecast_day_PCT5_change'],
                           regression_data['forecast_day_PCT7_change'],
                           regression_data['forecast_day_PCT10_change']
                           ]
        trend_change = False
        trend = True
        for pct_change in pct_change_list:
            if pct_change < 0 and trend_change == False:
                trend = True
            elif pct_change > 0 and trend_change == True: 
                trend = True    
            elif pct_change > 0 and trend_change == False:
                trend_change = True
                trend = True
            else:
                trend = False
        if trend:
            if(regression_data['bar_low'] < regression_data['bar_low_pre1']
                #or regression_data['bar_low_pre1'] < regression_data['bar_low_pre2']
                ):
                return '(downTrend-min3Day)'
            else:
                return '(downTrend-min3Day-Risky)'
    elif (regression_data['forecast_day_PCT_change'] > 0  
        and regression_data['forecast_day_PCT_change'] < regression_data['forecast_day_PCT2_change'] < regression_data['forecast_day_PCT3_change']
        and high_counter(regression_data) >= 3
        and ((regression_data['forecast_day_PCT3_change'] > regression_data['forecast_day_PCT4_change'] > regression_data['forecast_day_PCT5_change'] > regression_data['forecast_day_PCT7_change'] > regression_data['forecast_day_PCT10_change'])
             or (regression_data['forecast_day_PCT4_change'] > regression_data['forecast_day_PCT5_change'] > regression_data['forecast_day_PCT7_change'] > regression_data['forecast_day_PCT10_change'])
            )
        ):
        if(abs(regression_data['monthHighChange']) > abs(regression_data['monthLowChange'])):
            if(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumDownTrend)'  
            elif(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumDownTrend-crossed10Days)'
        elif(abs(regression_data['monthHighChange']) < abs(regression_data['monthLowChange'])): 
            if(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumDownTrend-monthLowGTmonthHigh)'  
            elif(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumDownTrend-monthLowGTmonthHigh-crossed10Days)'                       
    return 'NA'           
    
def pct_change_positive_trend(regression_data):
    if (regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and ((regression_data['high'] > regression_data['high_pre1'] and regression_data['high_pre1'] > regression_data['high_pre2'])
             or (regression_data['bar_high'] > regression_data['bar_high_pre1'] and regression_data['bar_high_pre1'] > regression_data['bar_high_pre2'])
             )
        and high_counter(regression_data) >= 3
        ):
        pct_change_list = [regression_data['forecast_day_PCT_change'],
                           regression_data['forecast_day_PCT2_change'],
                           regression_data['forecast_day_PCT3_change'],
                           regression_data['forecast_day_PCT4_change'],
                           regression_data['forecast_day_PCT5_change'],
                           regression_data['forecast_day_PCT7_change'],
                           regression_data['forecast_day_PCT10_change']
                           ]
        trend_change = False
        trend = True
        for pct_change in pct_change_list:
            if pct_change > 0 and trend_change == False:
                trend = True
            elif pct_change < 0 and trend_change == True: 
                trend = True    
            elif pct_change < 0 and trend_change == False:
                trend_change = True
                trend = True
            else:
                trend = False
        if trend:
            if(regression_data['bar_high'] > regression_data['bar_high_pre1']
                and regression_data['bar_high_pre1'] > regression_data['bar_high_pre2']
                ):
                return '(upTrend)' 
            else:
                return '(upTrend-Risky)' 
    elif (
        #regression_data['forecast_day_PCT_change'] > 0
        #regression_data['forecast_day_PCT2_change'] > 0
        regression_data['forecast_day_PCT3_change'] > 0
        and high_counter(regression_data) >= 3
        ):
        pct_change_list = [#regression_data['forecast_day_PCT_change'],
                           #regression_data['forecast_day_PCT2_change'],
                           regression_data['forecast_day_PCT3_change'],
                           regression_data['forecast_day_PCT4_change'],
                           regression_data['forecast_day_PCT5_change'],
                           regression_data['forecast_day_PCT7_change'],
                           regression_data['forecast_day_PCT10_change']
                           ]
        trend_change = False
        trend = True
        for pct_change in pct_change_list:
            if pct_change > 0 and trend_change == False:
                trend = True
            elif pct_change < 0 and trend_change == True: 
                trend = True    
            elif pct_change < 0 and trend_change == False:
                trend_change = True
                trend = True
            else:
                trend = False
        if trend:
            if(regression_data['bar_high'] > regression_data['bar_high_pre1']
                #or regression_data['bar_high_pre1'] > regression_data['bar_high_pre2']
                ):
                return '(upTrend-min3Day)' 
            else:
                return '(upTrend-min3Day-Risky)'
    elif (regression_data['forecast_day_PCT_change'] < 0  
        and regression_data['forecast_day_PCT_change'] > regression_data['forecast_day_PCT2_change'] > regression_data['forecast_day_PCT3_change'] > 0
        and low_counter(regression_data) >= 3
        and ((regression_data['forecast_day_PCT3_change'] < regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change'] < regression_data['forecast_day_PCT7_change'] < regression_data['forecast_day_PCT10_change'])
             or (regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change'] < regression_data['forecast_day_PCT7_change'] < regression_data['forecast_day_PCT10_change'])
             )
        ):
        if(abs(regression_data['monthHighChange']) < abs(regression_data['monthLowChange'])):
            if(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumUpTrend)'
            elif(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumUpTrend-crossed10Days)'
        elif(abs(regression_data['monthHighChange']) > abs(regression_data['monthLowChange'])):
            if(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumUpTrend-monthHighGTmonthLow)'
            elif(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumUpTrend-monthHighGTmonthLow-crossed10Days)'
                
    return 'NA' 

def pct_change_negative_trend_medium(regression_data):
    if (regression_data['forecast_day_PCT_change'] > 0  
        and regression_data['forecast_day_PCT_change'] < regression_data['forecast_day_PCT2_change'] < regression_data['forecast_day_PCT3_change']
        and high_counter(regression_data) >= 3
        and ((regression_data['forecast_day_PCT3_change'] > regression_data['forecast_day_PCT4_change'] > regression_data['forecast_day_PCT5_change'] > regression_data['forecast_day_PCT7_change'] > regression_data['forecast_day_PCT10_change'])
             or (regression_data['forecast_day_PCT4_change'] > regression_data['forecast_day_PCT5_change'] > regression_data['forecast_day_PCT7_change'] > regression_data['forecast_day_PCT10_change'])
            )
        ):
        if(abs(regression_data['monthHighChange']) > abs(regression_data['monthLowChange'])):
            if(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumDownTrend)'  
            elif(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumDownTrend-crossed10Days)'
        elif(abs(regression_data['monthHighChange']) < abs(regression_data['monthLowChange'])): 
            if(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumDownTrend-monthLowGTmonthHigh)'  
            elif(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumDownTrend-monthLowGTmonthHigh-crossed10Days)'           
    return 'NA'           
    
def pct_change_positive_trend_medium(regression_data):
    if (regression_data['forecast_day_PCT_change'] < 0  
        and regression_data['forecast_day_PCT_change'] > regression_data['forecast_day_PCT2_change'] > regression_data['forecast_day_PCT3_change'] > 0
        and low_counter(regression_data) >= 3
        and ((regression_data['forecast_day_PCT3_change'] < regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change'] < regression_data['forecast_day_PCT7_change'] < regression_data['forecast_day_PCT10_change'])
             or (regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change'] < regression_data['forecast_day_PCT7_change'] < regression_data['forecast_day_PCT10_change'])
             )
        ):
        if(abs(regression_data['monthHighChange']) < abs(regression_data['monthLowChange'])):
            if(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumUpTrend)'
            elif(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumUpTrend-crossed10Days)'
        elif(abs(regression_data['monthHighChange']) > abs(regression_data['monthLowChange'])):
            if(regression_data['forecast_day_PCT10_change'] > 0):
                return '(mediumUpTrend-monthHighGTmonthLow)'
            elif(regression_data['forecast_day_PCT10_change'] < 0):
                return '(mediumUpTrend-monthHighGTmonthLow-crossed10Days)'
    return 'NA'          

def pct_change_negative_trend_short(regression_data):
    if(regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0
        and ((regression_data['low'] < regression_data['low_pre1'] and regression_data['low_pre1'] < regression_data['low_pre2'])
            or (regression_data['bar_low'] < regression_data['bar_low_pre1'] and regression_data['bar_low_pre1'] < regression_data['bar_low_pre2'])
            )
        and low_counter(regression_data) >= 3
        ):
        if(regression_data['bar_low'] < regression_data['bar_low_pre1']
            and regression_data['bar_low_pre1'] < regression_data['bar_low_pre2']
            ):
            return '(shortDownTrend)'
        else:
            return '(shortDownTrend-Risky)'
    elif(#regression_data['forecast_day_PCT_change'] < 0
        #and regression_data['forecast_day_PCT2_change'] < 0
        regression_data['forecast_day_PCT3_change'] < 0
        and low_counter(regression_data) >= 3
        ):
        if(regression_data['bar_low'] < regression_data['bar_low_pre1']
            #or regression_data['bar_low_pre1'] < regression_data['bar_low_pre2']
            ):
            return '(shortDownTrend-min3Day)' 
        else:
            return '(shortDownTrend-min3Day-Risky)' 
                      
    return 'NA'           
    
def pct_change_positive_trend_short(regression_data):
    if(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and ((regression_data['high'] > regression_data['high_pre1'] and regression_data['high_pre1'] > regression_data['high_pre2'])
             or (regression_data['bar_high'] > regression_data['bar_high_pre1'] and regression_data['bar_high_pre1'] > regression_data['bar_high_pre2'])
             )
        and high_counter(regression_data) >= 3
        ):
        if(regression_data['bar_high'] > regression_data['bar_high_pre1']
            and regression_data['bar_high_pre1'] > regression_data['bar_high_pre2']
            ):
            return '(shortUpTrend)'
        else:
            return '(shortUpTrend-Risky)'
    elif(#regression_data['forecast_day_PCT_change'] > 0
        #and regression_data['forecast_day_PCT2_change'] > 0
        regression_data['forecast_day_PCT3_change'] > 0
        and high_counter(regression_data) >= 3
        ):
        if(regression_data['bar_high'] > regression_data['bar_high_pre1']
            #or regression_data['bar_high_pre1'] > regression_data['bar_high_pre2']
            ):
            return '(shortUpTrend-min3Day)'
        else:
            return '(shortUpTrend-min3Day-Risky)'
    return 'NA'       

def trend_calculator(regression_data):
    trend = pct_change_positive_trend(regression_data) + '$' + pct_change_negative_trend(regression_data)
    shortTrend = pct_change_positive_trend_short(regression_data) + '$' +  pct_change_negative_trend_short(regression_data)
    #mediumTrend = pct_change_positive_trend_medium(regression_data) + '$' +  pct_change_negative_trend_medium(regression_data)
           
    return trend + ':' + shortTrend 

def preDayPctChangeUp_orVolHigh(regression_data):
    if(regression_data['PCT_day_change_pre1'] > 0 
       or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
    ):
        return True
    else:
        return False
    
def preDayPctChangeDown_orVolHigh(regression_data):
    if(regression_data['PCT_day_change_pre1'] < 0 
       or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
    ):
        return True
    else:
        return False    

def all_between_five_and_up_score(regression_data):
    if(-5 < regression_data['forecast_day_PCT_change'] < 5
        and -5 < regression_data['forecast_day_PCT2_change'] < 5
        and -5 < regression_data['forecast_day_PCT3_change'] < 5
        and -5 < regression_data['forecast_day_PCT4_change'] < 5
        and -5 < regression_data['forecast_day_PCT5_change'] < 5
        and -5 < regression_data['forecast_day_PCT7_change'] < 5
        and -6 < regression_data['forecast_day_PCT10_change'] < 6
        and str(regression_data['score']) == '10'
        and regression_data['trend'] != 'down'
        ):
        return True
    else:
        return False

def all_between_five_and_down_score(regression_data):
    if(-5 < regression_data['forecast_day_PCT_change'] < 5
        and -5 < regression_data['forecast_day_PCT2_change'] < 5
        and -5 < regression_data['forecast_day_PCT3_change'] < 5
        and -5 < regression_data['forecast_day_PCT4_change'] < 5
        and -5 < regression_data['forecast_day_PCT5_change'] < 5
        and -5 < regression_data['forecast_day_PCT7_change'] < 5
        and -6 < regression_data['forecast_day_PCT10_change'] < 6
        and str(regression_data['score']) == '0-1'
        and regression_data['trend'] != 'up'
        ):
        return True
    else:
        return False        

def all_between_zero_and_five_up_score(regression_data):
    if(0 < regression_data['forecast_day_PCT_change'] < 5
        and 0 < regression_data['forecast_day_PCT2_change'] < 5
        and 0 < regression_data['forecast_day_PCT3_change'] < 5
        and 0 < regression_data['forecast_day_PCT4_change'] < 5
        and 0 < regression_data['forecast_day_PCT5_change'] < 5
        and 0 < regression_data['forecast_day_PCT7_change'] < 5
        and 0 < regression_data['forecast_day_PCT10_change'] < 6
        and str(regression_data['score']) == '10'
        and regression_data['trend'] != 'down'
        ):
        return True
    else:
        return False

def all_between_zero_and_five_down_score(regression_data):
    if(-5 < regression_data['forecast_day_PCT_change'] < 0
        and -5 < regression_data['forecast_day_PCT2_change'] < 0
        and -5 < regression_data['forecast_day_PCT3_change'] < 0
        and -5 < regression_data['forecast_day_PCT4_change'] < 0
        and -5 < regression_data['forecast_day_PCT5_change'] < 0
        and -5 < regression_data['forecast_day_PCT7_change'] < 0
        and -6 < regression_data['forecast_day_PCT10_change'] < 0
        and str(regression_data['score']) == '0-1'
        and regression_data['trend'] != 'up'
        ):
        return True
    else:
        return False           

def abs_yearHigh_more_than_yearLow(regression_data):
    if(abs(regression_data['yearHighChange']) > abs(regression_data['yearLowChange'])):
        return True;
    else:
        return False;

def abs_yearHigh_less_than_yearLow(regression_data):
    if(abs(regression_data['yearHighChange']) < abs(regression_data['yearLowChange'])):
        return True;
    else:
        return False;

def abs_month3High_more_than_month3Low(regression_data):
    if(abs(regression_data['month3HighChange']) > abs(regression_data['month3LowChange'])):
        return True;
    else:
        return False;

def abs_month3High_less_than_month3Low(regression_data):
    if(abs(regression_data['month3HighChange']) < abs(regression_data['month3LowChange'])):
        return True;
    else:
        return False;
    
def last_5_day_all_up_except_today(regression_data):
    if(regression_data['PCT_day_change_pre1'] > 0
        and regression_data['PCT_day_change_pre2'] > 0
        and regression_data['PCT_day_change_pre3'] > 0
        and regression_data['PCT_day_change_pre4'] > 0): 
        return True
    else:
        return False 
    
def last_5_day_all_down_except_today(regression_data):
    if(regression_data['PCT_day_change_pre1'] < 0
        and regression_data['PCT_day_change_pre2'] < 0
        and regression_data['PCT_day_change_pre3'] < 0
        and regression_data['PCT_day_change_pre4'] < 0): 
        return True
    else:
        return False      

def last_4_day_all_up(regression_data):
    if(regression_data['PCT_day_change'] > 0
        and regression_data['PCT_day_change_pre1'] > 0
        and regression_data['PCT_day_change_pre2'] > 0
        and regression_data['PCT_day_change_pre3'] > 0
        #and regression_data['PCT_day_change_pre4'] > 0
        ): 
        return True
    else:
        return False 
    
def last_4_day_all_down(regression_data):
    if(regression_data['PCT_day_change'] < 0
        and regression_data['PCT_day_change_pre1'] < 0
        and regression_data['PCT_day_change_pre2'] < 0
        and regression_data['PCT_day_change_pre3'] < 0
        #and regression_data['PCT_day_change_pre4'] < 0
        ): 
        return True
    else:
        return False      

def last_7_day_all_up(regression_data):
    if(regression_data['PCT_day_change_pre1'] > 0
        and regression_data['PCT_day_change_pre2'] > 0
        and regression_data['PCT_day_change_pre3'] > 0
        and regression_data['PCT_day_change_pre4'] > 0
        and regression_data['PCT_day_change_pre5'] > 0
        and regression_data['PCT_day_change_pre6'] > 0
        and regression_data['PCT_day_change_pre7'] > 0
        ): 
        return True
    else:
        return False 
    
def last_7_day_all_down(regression_data):
    if(regression_data['PCT_day_change_pre1'] < 0
        and regression_data['PCT_day_change_pre2'] < 0
        and regression_data['PCT_day_change_pre3'] < 0
        and regression_data['PCT_day_change_pre4'] < 0
        and regression_data['PCT_day_change_pre5'] < 0
        and regression_data['PCT_day_change_pre6'] < 0
        and regression_data['PCT_day_change_pre7'] < 0
        ): 
        return True
    else:
        return False      

def historical_data_OI(data):
    ardate = np.array([str(x) for x in (np.array(data['data'])[:,3][::-1]).tolist()])
    aroipctchange = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,6][::-1]).tolist()])
    arcontractpctchange = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,7][::-1]).tolist()])
    return ardate, aroipctchange, arcontractpctchange

def high_tail_pct(regression_data):
    if(regression_data['high'] - regression_data['bar_high'] == 0):
        return 0
    else:
        return (((regression_data['high'] - regression_data['bar_high'])/regression_data['bar_high'])*100)
    
def low_tail_pct(regression_data):
    if((regression_data['bar_low'] - regression_data['low']) == 0):
        return 0
    else:
        return (((regression_data['bar_low'] - regression_data['low'])/regression_data['bar_low'])*100)

def high_tail_pct_pre1(regression_data):
    if(regression_data['high_pre1'] - regression_data['bar_high_pre1'] == 0):
        return 0
    else:
        return (((regression_data['high_pre1'] - regression_data['bar_high_pre1'])/regression_data['bar_high_pre1'])*100)
    
def low_tail_pct_pre1(regression_data):
    if((regression_data['bar_low_pre1'] - regression_data['low_pre1']) == 0):
        return 0
    else:
        return (((regression_data['bar_low_pre1'] - regression_data['low_pre1'])/regression_data['bar_low_pre1'])*100)

def high_tail_pct_pre2(regression_data):
    if(regression_data['high_pre2'] - regression_data['bar_high_pre2'] == 0):
        return 0
    else:
        return (((regression_data['high_pre2'] - regression_data['bar_high_pre2'])/regression_data['bar_high_pre2'])*100)
    
def low_tail_pct_pre2(regression_data):
    if((regression_data['bar_low_pre2'] - regression_data['low_pre2']) == 0):
        return 0
    else:
        return (((regression_data['bar_low_pre2'] - regression_data['low_pre2'])/regression_data['bar_low_pre2'])*100)

def is_recent_consolidation(regression_data):
#     if(abs(regression_data['forecast_day_PCT_change']) < 1
#         and abs(regression_data['forecast_day_PCT2_change']) < 1
#         and (abs(regression_data['forecast_day_PCT3_change']) < 1 or abs(regression_data['forecast_day_PCT4_change']) < 1)
#         and abs(regression_data['forecast_day_PCT3_change']) < 2
#         and abs(regression_data['forecast_day_PCT4_change']) < 2
#         #and abs(regression_data['PCT_day_change']) < 1
#         #and abs(regression_data['PCT_day_change_pre1']) < 1
#         ): 
#         return True
    if(abs(regression_data['forecast_day_PCT_change']) < 1.5
        and abs(regression_data['forecast_day_PCT2_change']) < 1.5
        and abs(regression_data['forecast_day_PCT3_change']) < 1.5
        and abs(regression_data['forecast_day_PCT4_change']) < 1.5
        and abs(regression_data['forecast_day_PCT5_change']) < 1.5
        and abs(regression_data['forecast_day_PCT7_change']) < 1.5
        and abs(regression_data['forecast_day_PCT10_change']) < 1.5
        ): 
        return True
    else:
        return False 

def is_short_consolidation_breakout(regression_data):
    if(is_recent_consolidation(regression_data) and abs(regression_data['PCT_day_change']) > 2.5 and abs(regression_data['PCT_change']) > 2.5):
        return True
    else:
        return False

def breakout_or_no_consolidation(regression_data):
    if(is_recent_consolidation(regression_data) == False 
       or is_short_consolidation_breakout(regression_data) == True
       ):
        return True
    else:
        return False

def is_ema6_sliding_up(regression_data):
    if((regression_data['EMA6'] > regression_data['EMA6_1daysBack'])
        and (regression_data['EMA6'] > regression_data['EMA6_2daysBack'])):
        return True
    else:
        return False
  
def is_ema14_sliding_up(regression_data):
    if((regression_data['EMA14'] > regression_data['EMA14_1daysBack'])
        and (regression_data['EMA14'] > regression_data['EMA14_2daysBack'])):
        return True
    else:
        return False 
    
def is_ema6_sliding_down(regression_data):
    if((regression_data['EMA6'] < regression_data['EMA6_1daysBack'])
        and (regression_data['EMA6'] < regression_data['EMA6_2daysBack'])):
        return True
    else:
        return False
  
def is_ema14_sliding_down(regression_data):
    if((regression_data['EMA14'] < regression_data['EMA14_1daysBack'])
        and (regression_data['EMA14'] < regression_data['EMA14_2daysBack'])):
        return True
    else:
        return False    

def high_tail_pct_filter(regression_data, regressionResult):
    if((-0.5 < regression_data['PCT_day_change'] < 0.5 and -0.5 < regression_data['PCT_change'] < 0.5)
       or (0 < regression_data['PCT_day_change'] < 1)
       ):
        if(low_tail_pct(regression_data) > 3):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GLowTail-3pc')
        elif(low_tail_pct(regression_data) > 2):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GLowTail-2pc')
        elif(low_tail_pct(regression_data) > 1):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GLowTail-1pc')
    
    if((-0.5 < regression_data['PCT_day_change'] < 0.5 and -0.5 < regression_data['PCT_change'] < 0.5)
       or (-1 < regression_data['PCT_day_change'] < 0)
        ):
        if(high_tail_pct(regression_data) > 3):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GHighTail-3pc')
        elif(high_tail_pct(regression_data) > 2):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GHighTail-2pc')
        elif(high_tail_pct(regression_data) > 1):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GHighTail-1pc')
            
    if(0.75 < regression_data['PCT_day_change'] < 4):
        if(high_tail_pct(regression_data) > 3):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GHighTail-3pc')
        elif(high_tail_pct(regression_data) > 2):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GHighTail-2pc')
        elif(high_tail_pct(regression_data) > 1.5):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GHighTail-1.5pc')
    if(-4 < regression_data['PCT_day_change'] < -0.75):
        if(low_tail_pct(regression_data) > 3):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GLowTail-3pc')
        elif(low_tail_pct(regression_data) > 2):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GLowTail-2pc')
        elif(low_tail_pct(regression_data) > 1.5):
            add_in_csv(regression_data, regressionResult, None, None, None, 'GLowTail-1.5pc')

def tail_pct_filter(regression_data, regressionResult):
#     if(regression_data['PCT_day_change'] > 0):
#         if(high_tail_pct(regression_data) < 0.1):
#             add_in_csv(regression_data, regressionResult, None, None, None, 'HighTailLessThan-0.1pc')
#         if(high_tail_pct(regression_data) < 0.3):
#             add_in_csv(regression_data, regressionResult, None, None, None, 'HighTailLessThan-0.3pc')
#             if(low_tail_pct(regression_data) < 0.1):
#                 add_in_csv(regression_data, regressionResult, None, None, None, 'LowTailLessThan-0.1pc')
#             if(low_tail_pct(regression_data) < 0.3):
#                 add_in_csv(regression_data, regressionResult, None, None, None, 'LowTailLessThan-0.3pc')
#     if(regression_data['PCT_day_change'] < 0):
#         if(low_tail_pct(regression_data) < 0.1):
#             add_in_csv(regression_data, regressionResult, None, None, None, 'LowTailLessThan-0.1pc')
#         if(low_tail_pct(regression_data) < 0.3):
#             add_in_csv(regression_data, regressionResult, None, None, None, 'LowTailLessThan-0.3pc')
#             if(high_tail_pct(regression_data) < 0.1):
#                 add_in_csv(regression_data, regressionResult, None, None, None, 'HighTailLessThan-0.1pc')
#             if(high_tail_pct(regression_data) < 0.3):
#                 add_in_csv(regression_data, regressionResult, None, None, None, 'HighTailLessThan-0.3pc')
#    high_tail_pct_filter(regression_data, regressionResult)
    return False

def tail_reversal_filter(regression_data, regressionResult):
#     if(3 > low_tail_pct(regression_data) > 2
#             and high_tail_pct(regression_data) < 1
#             and low_tail_pct_pre1(regression_data) < 1.5
#             #and regression_data['SMA9'] < 0
#             #and 'MayBuyCheckChart' in regression_data['filter1'] 
#             and -3 < regression_data['PCT_day_change_pre1'] < -1
#             and -3 < regression_data['PCT_day_change'] < -1
#             and (regression_data['PCT_day_change_pre2'] > 0
#                 or regression_data['PCT_day_change_pre3'] > 0
#                 )
#             ):
#             add_in_csv(regression_data, regressionResult, None, '(Check-chart-market2to3down)MayBuy-LastDayDown', None)
#     elif(3 > low_tail_pct(regression_data) > 1.8
#             and high_tail_pct(regression_data) < 1
#             and low_tail_pct_pre1(regression_data) < 1.5
#             #and regression_data['SMA9'] < 0
#             #and 'MayBuyCheckChart' in regression_data['filter1'] 
#             and -3 < regression_data['PCT_day_change_pre1'] < -1
#             and -3 < regression_data['PCT_day_change'] < -1
#             and regression_data['PCT_day_change_pre2'] < 0
#             and regression_data['PCT_day_change_pre3'] < 0
#             ):
#             add_in_csv(regression_data, regressionResult, None, 'MaySell-downTrend', None)
#     if(3 > high_tail_pct(regression_data) > 2
#             and low_tail_pct(regression_data) < 1
#             and high_tail_pct_pre1(regression_data) < 1.5
#             #and regression_data['SMA9'] > 0
#             #and 'MaySellCheckChart' in regression_data['filter1'] 
#             and 3 > regression_data['PCT_day_change_pre1'] > 1
#             and 3 > regression_data['PCT_day_change'] > 1
#             and (regression_data['PCT_day_change_pre2'] < 0
#                 or regression_data['PCT_day_change_pre3'] < 0
#                 )
#             ):
#             add_in_csv(regression_data, regressionResult, None, '(Check-chart-market2to3up)MaySell-LastDayUp', None)
#     elif(3 > high_tail_pct(regression_data) > 1.8
#             and low_tail_pct(regression_data) < 1
#             and high_tail_pct_pre1(regression_data) < 1.5
#             #and regression_data['SMA9'] > 0
#             #and 'MaySellCheckChart' in regression_data['filter1'] 
#             and 3 > regression_data['PCT_day_change_pre1'] > 1
#             and 3 > regression_data['PCT_day_change'] > 1
#             and regression_data['PCT_day_change_pre2'] > 0
#             and regression_data['PCT_day_change_pre3'] > 0
#             ):
#             add_in_csv(regression_data, regressionResult, None, 'MayBuy-upTrend', None)
    
    if(-5 < regression_data['PCT_day_change'] < -1.5
        and 4 > low_tail_pct(regression_data) >= 1.49 and high_tail_pct(regression_data) <= 0.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuy-CheckChart)-1', None)
    elif(-5 < regression_data['PCT_day_change'] < -1
        and 4 > low_tail_pct(regression_data) >= 1.49 and high_tail_pct(regression_data) <= 0.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuy-CheckChart)-1-Risky', None)
    elif(-6 < regression_data['PCT_day_change'] < -1.5
        and 3 > low_tail_pct(regression_data) >= 0.99 and high_tail_pct(regression_data) <= 0.6):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuy-CheckChart)-2', None)
    elif(-6 < regression_data['PCT_day_change'] < -1
        and 3 > low_tail_pct(regression_data) >= 0.99 and high_tail_pct(regression_data) <= 0.6):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuy-CheckChart)-2-Risky', None)
            
    if(1.5 < regression_data['PCT_day_change'] < 5
        and 4 > high_tail_pct(regression_data) >= 1.49 and low_tail_pct(regression_data) <= 0.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySell-CheckChart)-1', None) 
    elif(1 < regression_data['PCT_day_change'] < 5
        and 4 > high_tail_pct(regression_data) >= 1.49 and low_tail_pct(regression_data) <= 0.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySell-CheckChart)-1-Risky', None)   
    elif(1.5 < regression_data['PCT_day_change'] < 6
        and 3 > high_tail_pct(regression_data) >= 0.99 and low_tail_pct(regression_data) <= 0.6):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySell-CheckChart)-2', None)
    elif(1 < regression_data['PCT_day_change'] < 6
        and 3 > high_tail_pct(regression_data) >= 0.99 and low_tail_pct(regression_data) <= 0.6):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySell-CheckChart)-2-Risky', None)
    
    if(-4 < regression_data['PCT_day_change'] < -1.5 and -4 < regression_data['PCT_change'] < -1.5
        ):
        if(3.5 > low_tail_pct(regression_data) > 1.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuyCheckChart-Strong(check-chart-MarketNotUp2-3))', None)
        elif(low_tail_pct(regression_data) > 1.8 and regression_data['PCT_day_change'] < -2):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuyCheckChart-Strong(check-chart-MarketNotUp2-3))', None)
    elif(-4 < regression_data['PCT_day_change'] < 1 and -4 < regression_data['PCT_change'] < 1
        ):
        if(3.5 > low_tail_pct(regression_data) > 1.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuyCheckChart(check-chart-MarketNotUp2-3))', None)
        elif(low_tail_pct(regression_data) > 1.8 and regression_data['PCT_day_change'] < -2):
            add_in_csv(regression_data, regressionResult, None, None, '(MayBuyCheckChart(check-chart-MarketNotUp2-3))', None)
    
    if(1.5 < regression_data['PCT_day_change'] < 4 and 1.5 < regression_data['PCT_change'] < 4
        ):
        if(3.5 > high_tail_pct(regression_data) > 1.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySellCheckChart-Strong(check-chart-MarketNotDown2-3))', None)
        if(high_tail_pct(regression_data) > 1.8 and regression_data['PCT_day_change'] > 2):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySellCheckChart-Strong(check-chart-MarketNotDown2-3))', None)
    elif(-1 < regression_data['PCT_day_change'] < 4 and -1 < regression_data['PCT_change'] < 4
        ):
        if(3.5 > high_tail_pct(regression_data) > 1.8):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySellCheckChart(check-chart-MarketNotDown2-3))', None)
        elif(high_tail_pct(regression_data) > 1.8 and regression_data['PCT_day_change'] > 2):
            add_in_csv(regression_data, regressionResult, None, None, '(MaySellCheckChart(check-chart-MarketNotDown2-3))', None)
        
    if(('MayBuy-CheckChart' in regression_data['filter1']) 
        or ('MayBuyCheckChart' in regression_data['filter1'])
        ):
        if(
            (low_tail_pct(regression_data) > 1.8)
            and low_tail_pct(regression_data) < abs(regression_data['PCT_change'])
            ):
            add_in_csv(regression_data, regressionResult, None, None, 'Strong', None)
    if(('MaySell-CheckChart' in regression_data['filter1']) 
        or ('MaySellCheckChart' in regression_data['filter1'])
        ):
        if(
            (high_tail_pct(regression_data) > 1.8)
            and high_tail_pct(regression_data) < abs(regression_data['PCT_change'])
            ):
            add_in_csv(regression_data, regressionResult, None, None, 'Strong', None)

def pct_change_filter(regression_data, regressionResult, save):
    filterName = ''
    if(regression_data['PCT_change'] < 0):
        if(regression_data['PCT_change'] < -10):
            filterName = 'PCTChangeLT-10'
        elif(regression_data['PCT_change'] < -8):
            filterName = 'PCTChangeLT-8'
        elif(regression_data['PCT_change'] < -6):
            filterName = 'PCTChangeLT-6'
        elif(regression_data['PCT_change'] < -5):
            filterName = 'PCTChangeLT-5'
        elif(regression_data['PCT_change'] < -4):
            filterName = 'PCTChangeLT-4'
        elif(regression_data['PCT_change'] < -3):
            filterName = 'PCTChangeLT-3'
        elif(regression_data['PCT_change'] < -2):
            filterName = 'PCTChangeLT-2'
        elif(regression_data['PCT_change'] < -1):
            filterName = 'PCTChangeLT-1'
        elif(regression_data['PCT_change'] < -0.5):
            filterName = 'PCTChangeLT-0.5'
        else:
            filterName = 'PCTDayChangeLT0'
    elif(regression_data['PCT_change'] > 0):
        if(regression_data['PCT_change'] > 10):
            filterName = 'PCTChangeGT10'
        elif(regression_data['PCT_change'] > 8):
            filterName = 'PCTChangeGT8'
        elif(regression_data['PCT_change'] > 6):
            filterName = 'PCTChangeGT6'
        elif(regression_data['PCT_change'] > 5):
            filterName = 'PCTChangeGT5'
        elif(regression_data['PCT_change'] > 4):
            filterName = 'PCTChangeGT4'
        elif(regression_data['PCT_change'] > 3):
            filterName = 'PCTChangeGT3'
        elif(regression_data['PCT_change'] > 2):
            filterName = 'PCTChangeGT2'
        elif(regression_data['PCT_change'] > 1):
            filterName = 'PCTChangeGT1'
        elif(regression_data['PCT_change'] > 0.5):
            filterName = 'PCTChangeGT0.5'
        else:
            filterName = 'PCTChangeGT0'
    elif(regression_data['PCT_change'] == 0):
        filterName = 'PCTChangeEQ0'
        
    if(save):
        add_in_csv(regression_data, regressionResult, None, filterName)
    else:
        return filterName

def tail_change_filter(regression_data, regressionResult, save):
    filterTail = ''
    
    if(high_tail_pct(regression_data) > 1.5):
        filterTail = filterTail + 'HTGT1.5'
        if(high_tail_pct(regression_data) > abs(regression_data['PCT_day_change'])):
            filterTail = filterTail + ',HTGTPCTDayChange'
#     elif(high_tail_pct(regression_data) > 2.5):
#         filterTail = filterTail + 'HTGT2.5'
#     elif(high_tail_pct(regression_data) > 1.5):
#         filterTail = filterTail + 'HTGT1.5'
#     elif(high_tail_pct(regression_data) < 1):
#         filterTail = filterTail + 'HTLT1'
        
    
    if(low_tail_pct(regression_data) > 1.5):
        filterTail = filterTail + 'LTGT1.5'
        if(low_tail_pct(regression_data) > abs(regression_data['PCT_day_change'])):
            filterTail = filterTail + ',LTGTPCTDayChange'
#     elif(low_tail_pct(regression_data) > 2.5):
#         filterTail = filterTail + 'LTGT2.5'
#     elif(low_tail_pct(regression_data) > 1.5):
#         filterTail = filterTail + 'LTGT1.5'
#     elif(low_tail_pct(regression_data) < 1):
#         filterTail = filterTail + 'LTLT1'
        
    if(save):
        add_in_csv(regression_data, regressionResult, None, filterTail)
    else:
        return filterTail

def tail2_change_filter(regression_data, regressionResult, save):
    filterTail = ''
    
    if(high_tail_pct(regression_data) > 3.5):
        filterTail = filterTail + 'HTGT3.5,'
    elif(high_tail_pct(regression_data) > 2.5):
        filterTail = filterTail + 'HTGT2.5,'
    elif(high_tail_pct(regression_data) > 1.5):
        filterTail = filterTail + 'HTGT1.5,'
    if(high_tail_pct(regression_data) > abs(regression_data['PCT_day_change']) 
        and high_tail_pct(regression_data) > 1.5
        ):
        filterTail = filterTail + 'HTGTPCTDayChange'

    if(low_tail_pct(regression_data) > 3.5):
        filterTail = filterTail + 'LTGT3.5,'
    elif(low_tail_pct(regression_data) > 2.5):
        filterTail = filterTail + 'LTGT2.5,'
    elif(low_tail_pct(regression_data) > 1.5):
        filterTail = filterTail + 'LTGT1.5,'
    if(low_tail_pct(regression_data) > abs(regression_data['PCT_day_change']) 
        and low_tail_pct(regression_data) > 1.5
        ):
        filterTail = filterTail + 'LTGTPCTDayChange'

    if(save):
        add_in_csv(regression_data, regressionResult, None, filterTail)
    else:
        return filterTail
        
def filterMA(regression_data, regressionResult):
    ws = None
    if(2 < regression_data['PCT_day_change'] < 4
       and 0 < regression_data['PCT_change'] < 4.5
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['PCT_day_change_pre1'] < 0
       #and regression_data['PCT_day_change_pre2'] < 0.75
       and high_tail_pct(regression_data) < 1
       and low_tail_pct(regression_data) < 1
       #and low_tail_pct_pre1(regression_data) < 1.5
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, "UP")
    elif(2 < regression_data['PCT_day_change'] < 4
       and 0 < regression_data['PCT_change'] < 4.5
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['forecast_day_PCT2_change'] > 0
       and regression_data['PCT_day_change_pre1'] < 0.75
       and regression_data['PCT_day_change_pre2'] < 0.5
       and 0.75 > regression_data['PCT_day_change_pre3'] > 0
       and high_tail_pct(regression_data) < 1
       and low_tail_pct(regression_data) < 1
       and low_tail_pct_pre1(regression_data) < 1
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, "UP-1")
    elif(2.5 < regression_data['PCT_day_change'] < 4
       and 0 < regression_data['PCT_change'] < 4.5
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['forecast_day_PCT2_change'] > 0
       and regression_data['PCT_day_change_pre1'] < 0.75
       and regression_data['PCT_day_change_pre2'] < 0.5
       and regression_data['PCT_day_change_pre3'] < 0
       and high_tail_pct(regression_data) < 1
       and low_tail_pct(regression_data) < 1
       #and low_tail_pct_pre1(regression_data) < 1
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, "UP-2")
    elif(-4 < regression_data['PCT_day_change'] < -2
       and -4.5 < regression_data['PCT_change'] < 0
       and regression_data['forecast_day_PCT_change'] < 0
       and regression_data['PCT_day_change_pre1'] > 0
       #and regression_data['PCT_day_change_pre2'] > -0.75
       and low_tail_pct(regression_data) < 1
       and high_tail_pct(regression_data) < 1
       #and high_tail_pct_pre1(regression_data) < 1.5
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, "DOWN")
    elif(-4 < regression_data['PCT_day_change'] < -2
       and -4.5 < regression_data['PCT_change'] < 0
       and regression_data['forecast_day_PCT_change'] < 0
       and regression_data['forecast_day_PCT2_change'] < 0
       and regression_data['PCT_day_change_pre1'] > -0.75
       and regression_data['PCT_day_change_pre2'] > -0.5
       and -0.75 < regression_data['PCT_day_change_pre3'] < 0
       and low_tail_pct(regression_data) < 1
       and high_tail_pct(regression_data) < 1
       and high_tail_pct_pre1(regression_data) < 1
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, "DOWN-1")
    elif(-4 < regression_data['PCT_day_change'] < -2
       and -4.5 < regression_data['PCT_change'] < 0
       and regression_data['forecast_day_PCT_change'] < 0
       and regression_data['forecast_day_PCT2_change'] < 0
       and regression_data['PCT_day_change_pre1'] > -0.75
       and regression_data['PCT_day_change_pre2'] > -0.5
       and regression_data['PCT_day_change_pre3'] > 0
       and low_tail_pct(regression_data) < 1
       and high_tail_pct(regression_data) < 1
       #and high_tail_pct_pre1(regression_data) < 1
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, "DOWN-2")
    
    if(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['PCT_day_change'] > 2.5
        and -1.5 < regression_data['PCT_day_change_pre1'] < 0
        and -1.5 < regression_data['PCT_day_change_pre2'] < 0
        and -1.5 < regression_data['PCT_day_change_pre3'] < 0 
        and (regression_data['yearHighChange'] < -10 or regression_data['PCT_day_change_pre3'] < 0)
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakout-0')
    elif(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['PCT_day_change'] > 2.5
        and -1.5 < regression_data['PCT_day_change_pre1'] < 0
        and -1.5 < regression_data['PCT_day_change_pre2'] < 0
        and -1.5 < regression_data['PCT_day_change_pre3'] < 1.5 
        and (regression_data['yearHighChange'] < -10 or regression_data['PCT_day_change_pre3'] < 0)
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakout-1')
    elif(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['PCT_day_change'] > 2.5
        and -1.5 < regression_data['PCT_day_change_pre1'] < 0
        and -1.5 < regression_data['PCT_day_change_pre2'] < 1.5
        and -1.5 < regression_data['PCT_day_change_pre3'] < 1.5 
        and regression_data['yearHighChange'] < -10
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakout-Risky')
    elif(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['PCT_day_change'] > 2.5
        and regression_data['PCT_day_change_pre1'] > 2.5
        and -1.5 < regression_data['PCT_day_change_pre2'] < 0
        and -1.5 < regression_data['PCT_day_change_pre3'] < 0
        and -1.5 < regression_data['PCT_day_change_pre4'] < 0
        and regression_data['yearHighChange'] < -10
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakoutLastDay-Risky')
    
    
    if(regression_data['forecast_day_PCT_change'] < 0
        and regression_data['PCT_day_change'] < -2.5
        and 0 < regression_data['PCT_day_change_pre1'] < 1.5
        and 0 < regression_data['PCT_day_change_pre2'] < 1.5
        and 0 < regression_data['PCT_day_change_pre3'] < 1.5 
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakdown-0')
    elif(regression_data['forecast_day_PCT_change'] < 0
        and regression_data['PCT_day_change'] < -2.5
        and 0 < regression_data['PCT_day_change_pre1'] < 1.5
        and 0 < regression_data['PCT_day_change_pre2'] < 1.5
        and -1.5 < regression_data['PCT_day_change_pre3'] < 1.5 
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakdown-1')
    elif(regression_data['forecast_day_PCT_change'] < 0
        and regression_data['PCT_day_change'] < -2.5
        and 0 < regression_data['PCT_day_change_pre1'] < 1.5
        and -1.5 < regression_data['PCT_day_change_pre2'] < 1.5
        and -1.5 < regression_data['PCT_day_change_pre3'] < 1.5 
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakdown-Risky')
    elif(regression_data['forecast_day_PCT_change'] < 0
        and regression_data['PCT_day_change'] < -2.5
        and regression_data['PCT_day_change_pre1'] < -2.5
        and 0 < regression_data['PCT_day_change_pre2'] < 1.5
        and 0 < regression_data['PCT_day_change_pre3'] < 1.5
        and 0 < regression_data['PCT_day_change_pre4'] < 1.5 
        ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, None, '(check-chart)consolidationBreakdownLastDay-Risky')
    
    
    ema_diff = regression_data['ema6-14']
    ema_diff_pre1 = regression_data['ema6-14_pre1']
    ema_diff_pre2 = regression_data['ema6-14_pre2']
    
    if(ema_diff > 0
       and ema_diff_pre1 < 0
       and ema_diff_pre2 < 0
       and ema_diff_pre2 < ema_diff_pre1
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '(Confirmed)EMA6>EMA14')
        
    if(ema_diff < 0
       and ema_diff_pre1 > 0
       and ema_diff_pre2 > 0
       and ema_diff_pre2 > ema_diff_pre1
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '(Confirmed)EMA6<EMA14')
        
        
    if(((regression_data['EMA6'] < regression_data['EMA6_1daysBack'] < regression_data['EMA6_2daysBack'])
       or (regression_data['EMA14'] < regression_data['EMA14_1daysBack'] < regression_data['EMA14_2daysBack']))
       and ema_diff < ema_diff_pre1 < ema_diff_pre2
       and 0 < ema_diff < 1 and ema_diff_pre2 > 2*ema_diff
       and (regression_data['PCT_day_change'] > 0
             or regression_data['PCT_day_change_pre1'] > 0
             or regression_data['PCT_day_change_pre2'] > 0
             or regression_data['PCT_day_change_pre3'] > 0
            )
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '(May)EMA6-MT-May-LT-EMA14')
        
    if(((regression_data['EMA6'] > regression_data['EMA6_1daysBack'] > regression_data['EMA6_2daysBack'])
       or (regression_data['EMA14'] > regression_data['EMA14_1daysBack'] > regression_data['EMA14_2daysBack']))
       and ema_diff_pre2 < ema_diff_pre1 < ema_diff
       and -1 < ema_diff < 0 and ema_diff_pre2 < 2*ema_diff
       and (regression_data['PCT_day_change'] < 0
             or regression_data['PCT_day_change_pre1'] < 0
             or regression_data['PCT_day_change_pre2'] < 0
             or regression_data['PCT_day_change_pre3'] < 0
            )
       ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '(May)EMA6-LT-May-MT-EMA14')
    
    if(is_ema6_sliding_up(regression_data)):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@EMA6Up@e@')
    elif(is_ema6_sliding_down(regression_data)):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@EMA6Down@e@')
    if(is_ema14_sliding_up(regression_data)):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@EMA14Up@e@')
    elif(is_ema14_sliding_down(regression_data)):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@EMA14Down@e@')
    
    if(regression_data['SMA25']) < 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA25LT0@e@')
    elif(regression_data['SMA25']) > 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA25GT0@e@')
    if(regression_data['SMA50']) < 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA50LT0@e@')
    elif(regression_data['SMA50']) > 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA50GT0@e@')
    if(regression_data['SMA100']) < 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA100LT0@e@')
    elif(regression_data['SMA100']) > 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA100GT0@e@')
    if(regression_data['SMA200']) < 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA200LT0@e@')
    elif(regression_data['SMA200']) > 0:
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '@s@SMA200GT0@e@')
            
    if((regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50'] < regression_data['SMA25']
        or(
            regression_data['SMA200'] < 0
            and regression_data['SMA100'] < 0
            and regression_data['SMA50'] < 0
            and regression_data['SMA50'] < regression_data['SMA25']
        )
        )
       and regression_data['SMA100'] < -5
       and (regression_data['SMA100'] < -10 or regression_data['SMA200'] < -10)
       and (regression_data['SMA9'] > -1 and regression_data['SMA4'] > -1)
       and (
            (regression_data['SMA9'] > 1 or regression_data['SMA4'] > 1)
            or (regression_data['SMA9'] > 0 and regression_data['SMA4'] > 0)
            )
       #and (regression_data['SMA9'] > 0 and regression_data['SMA4'] > 0)
        ):
        if(is_ema14_sliding_up(regression_data)):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:RisingMA-Confirm')
        elif((regression_data['SMA9'] > 1 or regression_data['SMA4'] > 1)
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:RisingMA')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:RisingMA-Risky')
        
        if(-5 < regression_data['SMA25'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##LONGTERM:NearSMA25')
        elif(0 < regression_data['SMA25'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##LONGTERM:CrossoverSMA25')
        elif(-5 < regression_data['SMA50'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##MEDIUMTERM:NearSMA50')
        elif(0 < regression_data['SMA50'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##MEDIUMTERM:CrossoverSMA50')
        elif(-5 < regression_data['SMA100'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##SHORTTERM:NearSMA100')
        elif(0 < regression_data['SMA100'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##SHORTTERM:CrossoverSMA100')  
        
    elif(regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50'] < regression_data['SMA25'] < 0
         and (regression_data['SMA9'] > 0 or regression_data['SMA4'] > 0)
    ):
        if(regression_data['SMA9'] > 0 and regression_data['SMA4'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, 'SMA9&SMA4GT0')
        elif(regression_data['SMA9'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, 'SMA9GT0')
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:NegativeMA')
    elif(regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50'] < regression_data['SMA25'] < 0
         and regression_data['SMA9'] < 0
         and regression_data['SMA4'] < 0
    ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:AllNegativeMA')
    elif(regression_data['SMA200'] > 0
         and regression_data['SMA100'] > 0 
         and regression_data['SMA50'] > 0
         and (regression_data['SMA9'] < 1 and regression_data['SMA4'] < 1)
         and (regression_data['SMA9'] < -1 or regression_data['SMA4'] < -1
              or (regression_data['SMA9'] < 0 or regression_data['SMA4'] < 0))
         ):
        if(is_ema14_sliding_down(regression_data)):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:DowningMA-Confirm')
        elif((regression_data['SMA9'] < -1 or regression_data['SMA4'] < -1)
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:DowningMA')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:DowningMA-Risky')
        
        if(-5 < regression_data['SMA25'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##LONGTERM:NearSMA25')
        elif(0 < regression_data['SMA25'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##LONGTERM:CrossoverSMA25')
        elif(-5 < regression_data['SMA50'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##MEDIUMTERM:NearSMA50')
        elif(0 < regression_data['SMA50'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##MEDIUMTERM:CrossoverSMA50')
        elif(-5 < regression_data['SMA100'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##SHORTTERM:NearSMA100')
        elif(0 < regression_data['SMA100'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##SHORTTERM:CrossoverSMA100')
    elif(regression_data['SMA200'] > regression_data['SMA100'] > regression_data['SMA50'] > regression_data['SMA25'] > 0
        and (regression_data['SMA9'] < 0 or regression_data['SMA4'] < 0)
        ):
        if(regression_data['SMA9'] < 0 and regression_data['SMA4'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, 'SMA9&SMA4LT0')
        elif(regression_data['SMA9'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, 'SMA9LT0')
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:PositiveMA')
    elif(regression_data['SMA200'] > regression_data['SMA100'] > regression_data['SMA50'] > regression_data['SMA25'] > 0
         and regression_data['SMA9'] > 0
         and regression_data['SMA4'] > 0
    ):
        add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '$$(Study)$$:AllPositiveMA')
    else:
        if(-5 < regression_data['SMA25'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##LONGTERM:NearSMA25')
        elif(0 < regression_data['SMA25'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##LONGTERM:CrossoverSMA25')
        elif(-5 < regression_data['SMA50'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##MEDIUMTERM:NearSMA50')
        elif(0 < regression_data['SMA50'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##MEDIUMTERM:CrossoverSMA50')
        elif(-5 < regression_data['SMA100'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##SHORTTERM:NearSMA100')
        elif(0 < regression_data['SMA100'] < 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, None, '##SHORTTERM:CrossoverSMA100')
            
def base_line(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if((0 < regression_data['year2HighChange'] < 10) 
        and (regression_data['year2LowChange'] > 40)
        ):
        if(regression_data['year2High'] == regression_data['month6High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['year2HighChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2GT7')
        elif(regression_data['year2HighChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2GT5')
        elif(regression_data['year2HighChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2GT2')
        elif(regression_data['year2HighChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2GT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakHighYear2')
        
    elif((0 < regression_data['yearHighChange'] < 10) 
        and (regression_data['yearLowChange'] > 30)
        ):
        if(regression_data['yearHigh'] == regression_data['month6High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['yearHighChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearGT7')
        elif(regression_data['yearHighChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearGT5')
        elif(regression_data['yearHighChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearGT2')
        elif(regression_data['yearHighChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearGT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakHighYear')
            
    elif((0 < regression_data['month6HighChange'] < 10)
        and (regression_data['month6LowChange'] > 20)
        ):
        if(regression_data['month6High'] == regression_data['month3High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['month6HighChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6GT7')
        elif(regression_data['month6HighChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6GT5')
        elif(regression_data['month6HighChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6GT2')
        elif(regression_data['month6HighChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6GT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakHighMonth6')
                
    elif((0 < regression_data['month3HighChange'] < 10) 
        and (regression_data['month3LowChange'] > 15)
        ):
        if(regression_data['month3High'] == regression_data['monthHigh']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['month3HighChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3GT7')
        elif(regression_data['month3HighChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3GT5')
        elif(regression_data['month3HighChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3GT2')
        elif(regression_data['month3HighChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3GT0')
           
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakHighMonth3')
        
    elif((0 < regression_data['monthHighChange'] < 10) 
        and (regression_data['monthLowChange'] > 10)
        ):
        if(regression_data['monthHigh'] == regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['monthHighChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthGT7')
        elif(regression_data['monthHighChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthGT5')
        elif(regression_data['monthHighChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthGT2')
        elif(regression_data['monthHighChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthGT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakHighMonth')    
    
    if((-10 < regression_data['year2HighChange'] < 0) 
        and (regression_data['year2LowChange'] > 40)
        ):
        if(regression_data['year2High'] == regression_data['month6High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
            
        if(regression_data['year2HighChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2LT-7')
        elif(regression_data['year2HighChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2LT-5')
        elif(regression_data['year2HighChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2LT-2')
        elif(regression_data['year2HighChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYear2LT0')
            
        if(regression_data['weekHigh'] >= regression_data['year2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalHighYear2')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearHighYear2')   
    elif((-10 < regression_data['yearHighChange'] < 0) 
        and (regression_data['yearLowChange'] > 30)
        ):
        if(regression_data['yearHigh'] == regression_data['month6High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['yearHighChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearLT-7')
        elif(regression_data['yearHighChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearLT-5')
        elif(regression_data['yearHighChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearLT-2')
        elif(regression_data['yearHighChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighYearLT0')
        
        if(regression_data['weekHigh'] >= regression_data['yearHigh']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalHighYear')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearHighYear')
        
    elif((-10 < regression_data['month6HighChange'] < 0) 
        and (regression_data['month6LowChange'] > 20)
        ):
        if(regression_data['month6High'] == regression_data['month3High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['month6HighChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6LT-7')
        elif(regression_data['month6HighChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6LT-5')
        elif(regression_data['month6HighChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6LT-2')
        elif(regression_data['month6HighChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth6LT0')
           
        if(regression_data['weekHigh'] >= regression_data['month6High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalHighMonth6')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearHighMonth6')
        
    elif((-10 < regression_data['month3HighChange'] < 0) 
        and (regression_data['month3LowChange'] > 15)
        ):
        if(regression_data['month3High'] == regression_data['monthHigh']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['month3HighChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3LT-7')
        elif(regression_data['month3HighChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3LT-5')
        elif(regression_data['month3HighChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3LT-2')
        elif(regression_data['month3HighChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonth3LT0')
        
        if(regression_data['weekHigh'] >= regression_data['month3High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalHighMonth3')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearHighMonth3')
        
    elif((-10 < regression_data['monthHighChange'] < 0) 
        and (regression_data['monthLowChange'] > 10)
        ):
        if(regression_data['monthHigh'] == regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekHighGTweek2High)')
        
        if(regression_data['monthHighChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthLT-7')
        elif(regression_data['monthHighChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthLT-5')
        elif(regression_data['monthHighChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthLT-2')
        elif(regression_data['monthHighChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'HighMonthLT0')
        
        if(regression_data['weekHigh'] >= regression_data['month3High']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalHighMonth')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearHighMonth')
        
        
    if((-10 < regression_data['year2LowChange'] < 0 ) 
        and (regression_data['year2HighChange'] < -40)
        ):
        if(regression_data['year2Low'] == regression_data['month6Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
            
        if(regression_data['year2LowChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2LT-7')
        elif(regression_data['year2LowChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2LT-5')
        elif(regression_data['year2LowChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2LT-2')
        elif(regression_data['year2LowChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2LT0')    
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakLowYear2')
        
    elif((-10 < regression_data['yearLowChange'] < 0) 
        and (regression_data['year2LowChange'] > 20)
        and (regression_data['yearHighChange'] < -30)
        ):
        if(regression_data['yearLow'] == regression_data['month6Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
            
        if(regression_data['yearLowChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearLT-7')
        elif(regression_data['yearLowChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearLT-5')
        elif(regression_data['yearLowChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearLT-2')
        elif(regression_data['yearLowChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearLT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakLowYear')
        
    elif((-10 < regression_data['month6LowChange'] < 0)
        and (regression_data['yearLowChange'] > 20)
        and (regression_data['month6HighChange'] < -20)
        ):
        if(regression_data['month6Low'] == regression_data['month3Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
            
        if(regression_data['month6LowChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6LT-7')
        elif(regression_data['month6LowChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6LT-5')
        elif(regression_data['month6LowChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6LT-2')
        elif(regression_data['month6LowChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6LT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakLowMonth6')
        
    elif((-10 < regression_data['month3LowChange'] < 0)
        and (regression_data['month6LowChange'] > 10)
        and (regression_data['yearLowChange'] > 20)
        and (regression_data['month3HighChange'] < -15)
        ):
        if(regression_data['month3Low'] == regression_data['monthLow']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
            
        if(regression_data['month3LowChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3LT-7')
        elif(regression_data['month3LowChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3LT-5')
        elif(regression_data['month3LowChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3LT-2')
        elif(regression_data['month3LowChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3LT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakLowMonth3')
        
    elif((-10 < regression_data['monthLowChange'] < 0)
        and (regression_data['month3LowChange'] > 10)
        and (regression_data['yearLowChange'] > 20)
        and (regression_data['monthHighChange'] < -10)
        ):
        if(regression_data['monthLow'] == regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
            
        if(regression_data['monthLowChange'] < -7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthLT-7')
        elif(regression_data['monthLowChange'] < -5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthLT-5')
        elif(regression_data['monthLowChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthLT-2')
        elif(regression_data['monthLowChange'] < 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthLT0')
            
        add_in_csv(regression_data, regressionResult, ws, None, None, None, 'BreakLowMonth')
        
    
    if((0 < regression_data['year2LowChange'] < 10) 
        and (regression_data['year2HighChange'] < -40)
        ):
        if(regression_data['year2Low'] == regression_data['month6Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
        
        if(regression_data['year2LowChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2GT7')
        elif(regression_data['year2LowChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2GT5')
        elif(regression_data['year2LowChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2GT2')
        elif(regression_data['year2LowChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYear2GT0')
           
        if(regression_data['weekLow'] < regression_data['year2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalLowYear2')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearLowYear2')
        
    elif((0 < regression_data['yearLowChange'] < 10)
        and (regression_data['year2LowChange'] > 7.5)
        and (regression_data['yearHighChange'] < -30)
        ):
        if(regression_data['yearLow'] == regression_data['month6Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
        
        if(regression_data['yearLowChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearGT7')
        elif(regression_data['yearLowChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearGT5')
        elif(regression_data['yearLowChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearGT2')
        elif(regression_data['yearLowChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowYearGT0')
        
        if(regression_data['weekLow'] < regression_data['yearLow']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalLowYear')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearLowYear')
        
    elif((0 < regression_data['month6LowChange'] < 10)
        and (regression_data['yearLowChange'] > 7.5)
        and (regression_data['month6HighChange'] < -20)
        ):
        if(regression_data['month6Low'] == regression_data['month3Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
            
        if(regression_data['month6LowChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6GT7')
        elif(regression_data['month6LowChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6GT5')
        elif(regression_data['month6LowChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6GT2')
        elif(regression_data['month6LowChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth6GT0')
            
        if(regression_data['weekLow'] < regression_data['month6Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalLowMonth6')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearLowMonth6')
        
    elif((0 < regression_data['month3LowChange'] < 10)
        and (regression_data['month6LowChange'] > 7.5)
        and (regression_data['month3HighChange'] < -15)
        ):
        if(regression_data['month3Low'] == regression_data['monthLow']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
        
        if(regression_data['month3LowChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3GT7')
        elif(regression_data['month3LowChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3GT5')
        elif(regression_data['month3LowChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3GT2')
        elif(regression_data['month3LowChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonth3GT0')
        
        if(regression_data['weekLow'] < regression_data['month3Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalLowMonth3')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearLowMonth3')
        

    elif((0 < regression_data['monthLowChange'] < 10)
        and (regression_data['month3LowChange'] > 7.5)
        and (regression_data['monthHighChange'] < -10)
        ):
        if(regression_data['monthLow'] == regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(Recent)')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, '(weekLowLTweek2Low)')
        
        if(regression_data['monthLowChange'] > 7):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthGT7')
        elif(regression_data['monthLowChange'] > 5):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthGT5')
        elif(regression_data['monthLowChange'] > 2):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthGT2')
        elif(regression_data['monthLowChange'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'LowMonthGT0')
            
        if(regression_data['weekLow'] < regression_data['month3Low']):
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'ReversalLowMonth')
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None, None, 'NearLowMonth')
                        
#     elif(regression_data['year2LowChange'] < 0):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**LowYear2')
#     elif(regression_data['yearLowChange'] < 0):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**LowYear')
#     elif(regression_data['month6LowChange'] < 0):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**LowMonth6')
#     elif(regression_data['month3LowChange'] < 0):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**LowMonth3')
#     elif(regression_data['year2HighChange'] > 5):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**HighYear2')
#     elif(regression_data['yearHighChange'] > 5):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**HighYear')
#     elif(regression_data['month6HighChange'] > 5):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**HighMonth6')
#     elif(regression_data['month3HighChange'] > 5):
#         add_in_csv(regression_data, regressionResult, ws, None, None, None, '**HighMonth3')
    
    if(-1 < regression_data['monthHighChange'] < 3
        and regression_data['month3HighChange'] < -5
        and regression_data['PCT_day_change'] > 1
        and (regression_data['PCT_day_change_pre2'] < 0.5
             or regression_data['PCT_day_change_pre3'] < 0.5
            )
        ):
        if(abs(regression_data['year2HighChange']) > abs(regression_data['year2LowChange'])
            and regression_data['year2LowChange'] < 25  
            and regression_data['yearLowChange'] < 15
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, '(check-chart):buyMonthHighBreak-atNearYearLow')
        elif(abs(regression_data['year2HighChange']) < abs(regression_data['year2LowChange'])
            and regression_data['high'] > regression_data['high_pre1']
            and regression_data['bar_high'] > regression_data['bar_high_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, '(check-chart):buyMonthHighBreak-checkATRSellBreak')
        elif(regression_data['high'] > regression_data['high_pre1']
            and regression_data['bar_high'] > regression_data['bar_high_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, '(check-chart):buyMonthHighBreak-checkATRSellBreak-Risky') 
    
    if(-3 < regression_data['monthLowChange'] < 1
        and regression_data['month3LowChange'] > 5
        and regression_data['PCT_day_change'] < -1
        and (regression_data['PCT_day_change_pre2'] > -0.5
             or regression_data['PCT_day_change_pre3'] > -0.5
            )
        ):
        if(abs(regression_data['year2HighChange']) < abs(regression_data['year2LowChange'])
            and regression_data['year2HighChange'] > -25  
            and regression_data['yearHighChange'] > -15  
            and (regression_data['PCT_day_change_pre2'] > 0
                 or regression_data['PCT_day_change_pre1'] > 0
                )                                      
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, '(check-chart):sellMonthLowBreak-atNearYearHigh')
        elif(abs(regression_data['year2HighChange']) > abs(regression_data['year2LowChange'])
            and regression_data['low'] < regression_data['low_pre1']
            and regression_data['bar_low'] < regression_data['bar_low_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, '(check-chart):sellMonthLowBreak-checkATRBuyBreak')
        elif(regression_data['low'] < regression_data['low_pre1']
            and regression_data['bar_low'] < regression_data['bar_low_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None, '(check-chart):sellMonthLowBreak-checkATRBuyBreak-Risky')                    

def historical_data(data):
    ardate = np.array([str(x) for x in (np.array(data['data'])[:,0][::-1]).tolist()])
    aropen = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,1][::-1]).tolist()])
    arhigh = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,2][::-1]).tolist()])
    arlow  = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,3][::-1]).tolist()])
    arlast = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,4][::-1]).tolist()])
    arclose= np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,5][::-1]).tolist()])
    arquantity = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,6][::-1]).tolist()])
    arturnover = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,7][::-1]).tolist()])
    return ardate, aropen, arhigh, arlow, arlast, arclose, arquantity, arturnover

def no_doji_or_spinning_buy_india(regression_data):
    if ('SPINNINGTOP' not in str(regression_data['buyIndia']) and 'DOJI' not in str(regression_data['buyIndia'])):
        return True;
    else:
        return False
    
def no_doji_or_spinning_sell_india(regression_data): 
    if ('SPINNINGTOP' not in str(regression_data['sellIndia']) and 'DOJI' not in str(regression_data['sellIndia'])):
        return True;
    else:
        return False   

def get_open_interest_data(regression_data, db_collection):
    data = db_collection.find_one({'scrip':regression_data['scrip']})
    if(data is None or (np.array(data['data'])).size < 1):
        return '0', '0'
    ardate, aroipctchange, arcontractpctchange = historical_data_OI(data)   
    if(ardate[-1] == regression_data['date']):
        return str(round(aroipctchange[-1], 2)), str(round(arcontractpctchange[-1], 2))
    return '0', '0'
 
def scrip_patterns_to_dict(filename):  
    tempDict = {}
    count = 0
    with open(filename) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            try:
                if (count != 0):
                    dictValue = {}
                    dictValue['avg'] = row[1]
                    dictValue['count'] = row[2]
                    tempDict[row[0]] = dictValue
                count = count + 1
            except:
                pass
    return tempDict  

def get_regressionResult(regression_data, scrip, db, mlp_r_o, kneighbours_r_o, mlp_c_o, kneighbours_c_o):
    regression_data['mlpValue_reg_other'] = float(mlp_r_o)
    regression_data['kNeighboursValue_reg_other'] = float(kneighbours_r_o)
    regression_data['mlpValue_cla_other'] = float(mlp_c_o)
    regression_data['kNeighboursValue_cla_other'] = float(kneighbours_c_o)
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['filter_345_avg'] = 0
    regression_data['filter_345_count'] = 0
    regression_data['filter_345_pct'] = 0
    regression_data['filter_avg'] = 0
    regression_data['filter_count'] = 0
    regression_data['filter_pct'] = 0
    regression_data['filter_pct_change_avg']  = 0
    regression_data['filter_pct_change_count'] = 0
    regression_data['filter_pct_change_pct'] = 0
    regression_data['filter_all_avg']  = 0
    regression_data['filter_all_count'] = 0
    regression_data['filter_all_pct'] = 0
    regression_data['filter_tech_avg'] = 0
    regression_data['filter_tech_count'] = 0
    regression_data['filter_tech_pct'] = 0
    regression_data['filter_tech_all_avg'] = 0
    regression_data['filter_tech_all_count'] = 0
    regression_data['filter_tech_all_pct'] = 0
    regression_data['filter_tech_all_pct_change_avg'] = 0
    regression_data['filter_tech_all_pct_change_count'] = 0
    regression_data['filter_tech_all_pct_change_pct'] = 0
    regression_data['series_trend'] = trend_calculator(regression_data)
    regression_data['oi'] = float(-10000)
    regression_data['contract'] = float(-10000)
    if(db is not None):
        oi, contract = get_open_interest_data(regression_data, db.historyOpenInterest)
        regression_data['oi'] = float(oi)
        regression_data['contract'] = float(contract)
    
    regression_data['oi_next'] = float(-10000)
    regression_data['contract_next'] = float(-10000)
    if(db is not None):
        oi, contract = get_open_interest_data(regression_data, db.historyOpenInterestNext)
        regression_data['oi_next'] = float(oi)
        regression_data['contract_next'] = float(contract)
    resultDeclared = ""
    resultDate = ""
    resultSentiment = ""
    resultComment = ""
    if(db is not None):
        result_data = db.scrip_result.find_one({'scrip':scrip})
        if(result_data is not None):
            resultDate = result_data['result_date'].strip()
            resultSentiment = result_data['result_sentiment']
            resultComment = result_data['comment']
            start_date = (datetime.datetime.now() - datetime.timedelta(hours=0))
            start_date = datetime.datetime(start_date.year, start_date.month, start_date.day, start_date.hour)
            result_time = datetime.datetime.strptime(resultDate, "%Y-%m-%d")
            if result_time < start_date: 
                resultDeclared = resultDate 
                resultDate = ""
    regressionResult = [ ]
    regressionResult.append(regression_data['buyIndia'])
    regressionResult.append(regression_data['sellIndia'])
    if(TEST != True):
        regressionResult.append(regression_data['buyIndia_avg'])
        regressionResult.append(regression_data['buyIndia_count'])
        regressionResult.append(regression_data['sellIndia_avg'])
        regressionResult.append(regression_data['sellIndia_count'])
    regressionResult.append(regression_data['scrip'])
    regressionResult.append(regression_data['series_trend'])
    regressionResult.append(regression_data['SMA4_2daysBack'])
    regressionResult.append(regression_data['SMA9_2daysBack'])
    regressionResult.append(regression_data['SMA4'])
    regressionResult.append(regression_data['SMA9'])
    regressionResult.append(regression_data['SMA25'])
    regressionResult.append(regression_data['SMA50'])
    regressionResult.append(regression_data['SMA100'])
    regressionResult.append(regression_data['SMA200'])
    regressionResult.append(resultDate)
    regressionResult.append(resultDeclared)
    regressionResult.append(resultSentiment)
    regressionResult.append(resultComment)
    regressionResult.append(regression_data['forecast_day_VOL_change'])
    regressionResult.append(regression_data['oi'])
    regressionResult.append(regression_data['contract'])
    regressionResult.append(regression_data['oi_next'])
    regressionResult.append(regression_data['contract_next'])
    regressionResult.append(regression_data['forecast_day_PCT_change'])
    regressionResult.append(regression_data['forecast_day_PCT2_change'])
    regressionResult.append(regression_data['forecast_day_PCT3_change'])
    regressionResult.append(regression_data['forecast_day_PCT4_change'])
    regressionResult.append(regression_data['forecast_day_PCT5_change'])
    regressionResult.append(regression_data['forecast_day_PCT7_change'])
    regressionResult.append(regression_data['forecast_day_PCT10_change'])
    regressionResult.append(regression_data['mlpValue_reg'])
    regressionResult.append(regression_data['kNeighboursValue_reg'])
    regressionResult.append(regression_data['mlpValue_cla'])
    regressionResult.append(regression_data['kNeighboursValue_cla'])
    regressionResult.append(regression_data['mlpValue_reg_other'])
    regressionResult.append(regression_data['kNeighboursValue_reg_other'])
    regressionResult.append(regression_data['mlpValue_cla_other'])
    regressionResult.append(regression_data['kNeighboursValue_cla_other'])
    regressionResult.append(regression_data['year2HighChange'])
    regressionResult.append(regression_data['year2LowChange'])
    regressionResult.append(regression_data['yearHighChange'])
    regressionResult.append(regression_data['yearLowChange'])
    regressionResult.append(regression_data['month6HighChange'])
    regressionResult.append(regression_data['month6LowChange'])
    regressionResult.append(regression_data['month3HighChange'])
    regressionResult.append(regression_data['month3LowChange'])
    regressionResult.append(regression_data['monthHighChange'])
    regressionResult.append(regression_data['monthLowChange'])
    regressionResult.append(regression_data['week2HighChange'])
    regressionResult.append(regression_data['week2LowChange'])
    regressionResult.append(regression_data['weekHighChange'])
    regressionResult.append(regression_data['weekLowChange'])
    regressionResult.append(regression_data['trend'])
    regressionResult.append(regression_data['score'])
    regressionResult.append(str(round(high_tail_pct(regression_data), 1))) 
    regressionResult.append(str(round(low_tail_pct(regression_data), 1)))
    regressionResult.append(regression_data['close'])
    regressionResult.append(regression_data['PCT_day_change'])
    regressionResult.append(regression_data['PCT_change'])
    regressionResult.append(regression_data['scrip'])
    return regressionResult

def all_withoutml(regression_data, regressionResult, ws):
    filter_accuracy_finder_regression(regression_data, regressionResult, None)
    add_in_csv(regression_data, regressionResult, ws, '')

def ten_days_more_than_fifteen(regression_data):
    if(25 > regression_data['forecast_day_PCT10_change'] > 5
       and (regression_data['forecast_day_PCT5_change'] > 15
         or regression_data['forecast_day_PCT7_change'] > 15
         or regression_data['forecast_day_PCT10_change'] > 15)
    ):
        return True
    else:
        return False

def ten_days_more_than_ten(regression_data):
    if(25 > regression_data['forecast_day_PCT10_change'] > 5
       and (regression_data['forecast_day_PCT5_change'] > 10
         or regression_data['forecast_day_PCT7_change'] > 10
         or regression_data['forecast_day_PCT10_change'] > 10)
    ):
        return True
    else:
        return False

def ten_days_less_than_ten(regression_data):
    if(regression_data['forecast_day_PCT5_change'] < 10
       and regression_data['forecast_day_PCT7_change'] < 10
       and regression_data['forecast_day_PCT10_change'] < 10
    ):
        return True
    else:
        return False

def ten_days_less_than_five(regression_data):
    if(regression_data['forecast_day_PCT_change'] < 5
       and regression_data['forecast_day_PCT2_change'] < 5
       and regression_data['forecast_day_PCT3_change'] < 5
       and regression_data['forecast_day_PCT4_change'] < 5
       and regression_data['forecast_day_PCT5_change'] < 5
       and regression_data['forecast_day_PCT7_change'] < 5
       and regression_data['forecast_day_PCT10_change'] < 5
    ):
        return True
    else:
        return False
 
def ten_days_less_than_minus_fifteen(regression_data):
    if(-25 < regression_data['forecast_day_PCT10_change'] < -5
       and (regression_data['forecast_day_PCT5_change'] < -15
         or regression_data['forecast_day_PCT7_change'] < -15
         or regression_data['forecast_day_PCT10_change'] < -15)
    ):
        return True
    else:
        return False 
       
def ten_days_less_than_minus_ten(regression_data):
    if(-25 < regression_data['forecast_day_PCT10_change'] < -5
       and (regression_data['forecast_day_PCT5_change'] < -10
         or regression_data['forecast_day_PCT7_change'] < -10
         or regression_data['forecast_day_PCT10_change'] < -10)
    ):
        return True
    else:
        return False

def ten_days_more_than_minus_ten(regression_data):
    if(regression_data['forecast_day_PCT5_change'] > -10 
       and regression_data['forecast_day_PCT7_change'] > -10 
       and regression_data['forecast_day_PCT10_change'] > -10
    ):
        return True
    else:
        return False
  
def ten_days_more_than_minus_five(regression_data):
    if(regression_data['forecast_day_PCT_change'] > -5
       and regression_data['forecast_day_PCT2_change'] > -5 
       and regression_data['forecast_day_PCT3_change'] > -5 
       and regression_data['forecast_day_PCT4_change'] > -5
       and regression_data['forecast_day_PCT5_change'] > -5 
       and regression_data['forecast_day_PCT7_change'] > -5 
       and regression_data['forecast_day_PCT10_change'] > -5
    ):
        return True
    else:
        return False
       
def ten_days_more_than_seven(regression_data):
    if(25 > regression_data['forecast_day_PCT10_change'] > 5
       and (regression_data['forecast_day_PCT5_change'] > 7
         or regression_data['forecast_day_PCT7_change'] > 7
         or regression_data['forecast_day_PCT10_change'] > 7)
    ):
        return True
    else:
        return False
        
def ten_days_less_than_minus_seven(regression_data):
    if(-25 < regression_data['forecast_day_PCT10_change'] < -5
       and (regression_data['forecast_day_PCT5_change'] < -7
         or regression_data['forecast_day_PCT7_change'] < -7
         or regression_data['forecast_day_PCT10_change'] < -7)
    ):
        return True
    else:
        return False    

def ten_days_more_than_five(regression_data):
    if(25 > regression_data['forecast_day_PCT10_change'] > 5
       and regression_data['forecast_day_PCT7_change'] > 5
       and regression_data['forecast_day_PCT5_change'] > 1
       and regression_data['forecast_day_PCT4_change'] > 1
       and regression_data['forecast_day_PCT3_change'] > 1
    ):
        return True
    else:
        return False
        
def ten_days_less_than_minus_five(regression_data):
    if(-25 < regression_data['forecast_day_PCT10_change'] < -5
       and regression_data['forecast_day_PCT7_change'] < -5
       and regression_data['forecast_day_PCT5_change'] < -1
       and regression_data['forecast_day_PCT4_change'] < -1
       and regression_data['forecast_day_PCT3_change'] < -1
    ):
        return True
    else:
        return False    

def buy_pattern_without_mlalgo(regression_data, regressionResult):
    if(regression_data['PCT_day_change'] < 3.5
        and regression_data['year2LowChange'] > 5):
        if(regression_data['buyIndia_avg'] > 0.9 and regression_data['buyIndia_count'] > 1
            and (regression_data['SMA9'] > 0 or regression_data['SMA4'] > 1.5)
            ):
            if(regression_data['SMA9'] > 0):
                add_in_csv(regression_data, regressionResult, None, '(SMA9GT0)')
            elif(regression_data['SMA4'] > 1.5):
                add_in_csv(regression_data, regressionResult, None, '(SMA4GT(1.5))')
            if(regression_data['yearLowChange'] > 10 and regression_data['buyIndia_count'] > 1):
                if(regression_data['buyIndia_avg'] > 1.5):
                    add_in_csv(regression_data, regressionResult, None, 'buyPattern-GT1.5-SMAGT0')
                elif(regression_data['buyIndia_avg'] > 0.9
                    ):
                    add_in_csv(regression_data, regressionResult, None, 'buyPattern-Risky-GT0.9-SMAGT0')
                #add_in_csv(regression_data, regressionResult, None, 'buyPatterns-1-Risky')

    if(regression_data['PCT_day_change'] < 3.5
        and regression_data['year2LowChange'] > 5
        and regression_data['year2HighChange'] < -5):
        if(regression_data['buyIndia_avg'] < -1.8):
            add_in_csv(regression_data, regressionResult, None, 'maySellFromBuyPattern-LT-1.8')
        if(regression_data['buyIndia_avg'] < -0.9 and regression_data['buyIndia_count'] > 1
            and (regression_data['SMA9'] < 0 or regression_data['SMA4'] < -1.5)
            ):
            if(regression_data['SMA9'] < 0):
                add_in_csv(regression_data, regressionResult, None, '(SMA9LT0)')
            elif(regression_data['SMA4'] < -1.5):
                add_in_csv(regression_data, regressionResult, None, '(SMA4LT(-1.5))')
            if(regression_data['yearLowChange'] > 10 and regression_data['buyIndia_count'] > 1):
                if(regression_data['buyIndia_avg'] < -1.5):
                    add_in_csv(regression_data, regressionResult, None, 'sellPattern-LT-1.5-SMALT0')
                elif(regression_data['buyIndia_avg'] < -0.9
                    ):
                    add_in_csv(regression_data, regressionResult, None, 'sellPattern-Risky-LT-0.9-SMALT0')
            #add_in_csv(regression_data, regressionResult, None, 'sellPatterns-1-Risky')
#     buyPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/all-buy-filter-by-PCT-Change.csv')
#     if regression_data['buyIndia'] != '' and regression_data['buyIndia'] in buyPatternsDict:
#         if (abs(float(buyPatternsDict[regression_data['buyIndia']]['avg'])) >= .1 and float(buyPatternsDict[regression_data['buyIndia']]['count']) >= 2):
#             if(-0.5 < regression_data['PCT_day_change'] < 3 and float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 1):
#                 avg = buyPatternsDict[regression_data['buyIndia']]['avg']
#                 count = buyPatternsDict[regression_data['buyIndia']]['count']
#                 #add_in_csv_hist_pattern(regression_data, regressionResult, ws2, 'wml_buy', avg, count)
#             elif(-3 < regression_data['PCT_day_change'] < 0.5 and float(buyPatternsDict[regression_data['buyIndia']]['avg']) < -1):
#                 avg = buyPatternsDict[regression_data['buyIndia']]['avg']
#                 count = buyPatternsDict[regression_data['buyIndia']]['count']
#                 #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'wml_buy', avg, count)

def buy_pattern_from_history(regression_data, ws):
    buyPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/patterns-buy.csv')
    buyIndiaAvg = 0
    regression_data['buyIndia_avg'] = 0
    regression_data['buyIndia_count'] = 0
    regression_data['sellIndia_avg'] = 0
    regression_data['sellIndia_count'] = 0
    regression_data['filter_avg'] = 0
    regression_data['filter_count'] = 0
    flag = False
    if regression_data['buyIndia'] != '' and regression_data['buyIndia'] in buyPatternsDict:
        regression_data['buyIndia_avg'] = float(buyPatternsDict[regression_data['buyIndia']]['avg'])
        regression_data['buyIndia_count'] = float(buyPatternsDict[regression_data['buyIndia']]['count'])
        if (abs(float(buyPatternsDict[regression_data['buyIndia']]['avg'])) >= .1):
            buyIndiaAvg = float(buyPatternsDict[regression_data['buyIndia']]['avg'])
            if(int(buyPatternsDict[regression_data['buyIndia']]['count']) >= 2):
                if(is_algo_buy(regression_data)
                    and 'P@[' not in str(regression_data['sellIndia'])
                    #and regression_data['trend'] != 'up'
                    and -1 < regression_data['PCT_day_change'] < 3):
                    avg = buyPatternsDict[regression_data['buyIndia']]['avg']
                    count = buyPatternsDict[regression_data['buyIndia']]['count']
                    if(float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 0.8 and int(buyPatternsDict[regression_data['buyIndia']]['count']) >= 5):
                       flag = True
                       #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'buyPattern2', avg, count)
                    elif(float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 0.5 
                       or (float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 0.3 and (ten_days_less_than_minus_ten(regression_data) or regression_data['yearHighChange'] < -40))):
                        if(regression_data['forecast_day_PCT10_change'] < 0 and regression_data['forecast_day_PCT_change'] >= 0):
                            flag = True
                            #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'buyPattern2', avg, count)
                        elif(regression_data['forecast_day_PCT10_change'] > 0):    
                            flag = True
                            #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'buyPattern2', avg, count)     
    return buyIndiaAvg, flag

def buy_all_rule(regression_data, regressionResult, buyIndiaAvg, ws):
    if(is_algo_buy(regression_data)
        and buyIndiaAvg >= -.70
        and ((last_7_day_all_up(regression_data) == False) or (regression_data['forecast_day_PCT10_change'] < 10))
        and (MARKET_IN_UPTREND or (last_4_day_all_up(regression_data) == False)) #Uncomment0 If very less data
        and breakout_or_no_consolidation(regression_data) == True
        and '(Confirmed)EMA6>EMA14' not in regression_data['filter4']
        and regression_data['close'] > 50
        and high_tail_pct(regression_data) < 1
        and low_tail_pct(regression_data) > 1.2
        and regression_data['low'] > regression_data['low_pre1']
        ):
        if(regression_data['PCT_day_change'] < -0.75 and regression_data['PCT_change'] < 0
            and regression_data['month3HighChange'] < -15  
            ):
            add_in_csv(regression_data, regressionResult, None, 'ML:buy-0')
#     elif(is_algo_sell(regression_data)):
#         if(2 < regression_data['PCT_day_change'] < 3.5
#             and  2 < regression_data['PCT_change'] < 4
#             and high_tail_pct(regression_data) > 1.5
#             and low_tail_pct(regression_data) < 1
#             ):
#             add_in_csv(regression_data, regressionResult, None, 'ML:sell-1')
    
    if(is_algo_buy(regression_data, True)
        and ((regression_data['PCT_day_change'] > 1 
                and high_tail_pct(regression_data) < 2 
                and ('P@[' not in str(regression_data['sellIndia']))
                )
             or (low_tail_pct(regression_data) > 1.5 
                and (high_tail_pct(regression_data) < low_tail_pct(regression_data))
                )
             or (regression_data['PCT_day_change'] < -2
                and regression_data['forecast_day_PCT_change'] > 0.3
                )
            )
        and (BUY_VERY_LESS_DATA or regression_data['PCT_change'] > -1)
        and (BUY_VERY_LESS_DATA or ((regression_data['PCT_day_change_pre1'] < 0) or (regression_data['forecast_day_VOL_change'] > 0))) #Uncomment1 If very less data
        and (BUY_VERY_LESS_DATA or (regression_data['high']-regression_data['bar_high']) < (regression_data['bar_high']-regression_data['bar_low']))
        and buyIndiaAvg >= -.70
        and ((last_7_day_all_up(regression_data) == False) or (regression_data['forecast_day_PCT10_change'] < 10))
        and (MARKET_IN_UPTREND or (last_4_day_all_up(regression_data) == False)) #Uncomment0 If very less data
        and breakout_or_no_consolidation(regression_data) == True
        and '(Confirmed)EMA6>EMA14' not in regression_data['filter4']
        and regression_data['close'] > 50
        ):
        add_in_csv(regression_data, regressionResult, ws, None)
        return True
    return False

def buy_all_rule_classifier(regression_data, regressionResult, buyIndiaAvg, ws):
    if(is_algo_buy_classifier(regression_data, True)
        and ((regression_data['PCT_day_change'] > 1 
                and high_tail_pct(regression_data) < 2 
                and ('P@[' not in str(regression_data['sellIndia']))
                )
             or (low_tail_pct(regression_data) > 1.5 
                and (high_tail_pct(regression_data) < low_tail_pct(regression_data))
                )
             or (regression_data['PCT_day_change'] < -2
                and regression_data['forecast_day_PCT_change'] > 0.3
                )
            )
        and (BUY_VERY_LESS_DATA or regression_data['PCT_change'] > -1)
        and (BUY_VERY_LESS_DATA or ((regression_data['PCT_day_change_pre1'] < 0) or (regression_data['forecast_day_VOL_change'] > 0))) #Uncomment1 If very less data
        and (BUY_VERY_LESS_DATA or (regression_data['high']-regression_data['bar_high']) < (regression_data['bar_high']-regression_data['bar_low']))
        and buyIndiaAvg >= -.70
        and ((last_7_day_all_up(regression_data) == False) or (regression_data['forecast_day_PCT10_change'] < 10))
        and (MARKET_IN_UPTREND or (last_4_day_all_up(regression_data) == False)) #Uncomment0 If very less data
        and breakout_or_no_consolidation(regression_data) == True
        and '(Confirmed)EMA6>EMA14' not in regression_data['filter4']
        and regression_data['close'] > 50
        ):
        add_in_csv(regression_data, regressionResult, ws, None)
        return True
    return False

def buy_all_common(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, True)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, True)
    mlpValue_cla, kNeighboursValue_cla = get_reg_or_cla(regression_data, False)
    mlpValue_other_cla, kNeighboursValue_other_cla = get_reg_or_cla_other(regression_data, False)
                
    if(is_algo_buy(regression_data)
        and (2 < regression_data['PCT_day_change'] < 4) and (1 < regression_data['PCT_change'] < 4)
        and ((regression_data['PCT_day_change_pre1'] < 0 and regression_data['forecast_day_PCT_change'] > 0)
             )
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and regression_data['forecast_day_PCT4_change'] > 1
        #and regression_data['forecast_day_PCT10_change'] < 15
        and regression_data['month3HighChange'] < -10
        and regression_data['month3LowChange'] > 10
        ):    
        if(regression_data['SMA9'] > 1):
            add_in_csv(regression_data, regressionResult, ws, None, '##Common:buyNotM3HighLow-0(SMA9GT1)')
        elif(regression_data['SMA25'] > 0):
            add_in_csv(regression_data, regressionResult, ws, None, '##(Test):buyNotM3HighLow-0(SMA25GT0)')  
        else:
            add_in_csv(regression_data, regressionResult, ws, None)
             
    return False

def buy_all_common_High_Low(regression_data, regressionResult, reg, ws):
    if( ("upTrend" in regression_data['series_trend'])
        and (-1 < regression_data['PCT_day_change'] < 0) and (-1 < regression_data['PCT_change'] < 0)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:LastDayDownInUptrend')
        
    if((-2 < regression_data['PCT_day_change'] < 0) and (-2.5 < regression_data['PCT_change'] < 0)
        and (1 <= low_tail_pct(regression_data) < 3)
        ):
        if(regression_data['bar_high'] < regression_data['bar_high_pre1']
            and regression_data['bar_high_pre1'] > regression_data['bar_high_pre2']
            and regression_data['high'] < regression_data['high_pre1']
            and regression_data['high_pre1'] > regression_data['high_pre2'] > regression_data['high_pre3']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuyContinueLowTail-downtrendStart(CheckChart-Risky)')
        elif(regression_data['low'] < regression_data['low_pre1']
            and regression_data['low_pre1'] < regression_data['low_pre2']
            and regression_data['PCT_day_change_pre1'] < 0
            and regression_data['PCT_day_change_pre2'] < 0
            and regression_data['PCT_day_change_pre3'] < 0
            and (regression_data['forecast_day_PCT3_change'] < 0)
            and (regression_data['forecast_day_PCT5_change'] < -5)
            and (regression_data['forecast_day_PCT7_change'] < -10 or regression_data['forecast_day_PCT10_change'] < -10)
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuyStartLowTail-downtrend(CheckChart-Risky)')
        elif((high_tail_pct(regression_data) < 1 and 2.5 > low_tail_pct(regression_data) >= 1.5)
            and regression_data['PCT_day_change'] < 0
            and regression_data['PCT_change'] < 0
            ):
            if(regression_data['PCT_day_change_pre1'] > 0):
                add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuy-CheckChart-LastDayUp(|/mayFail(|before10AM))')
            else: 
                add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuy-CheckChart(|/mayFail(|before10AM))')
        
    if((0.3 < regression_data['PCT_day_change'] < 2) and (0.3 < regression_data['PCT_change'] < 2)
        and (1 <= high_tail_pct(regression_data) < 2)
        ):
        if(regression_data['high'] > regression_data['high_pre1']
            and regression_data['high_pre1'] > regression_data['high_pre2']
            and 0 < regression_data['PCT_day_change_pre1'] < 5
            and -1 < regression_data['PCT_day_change_pre2'] < 5
            and (regression_data['forecast_day_PCT3_change'] < 5)
            and (regression_data['forecast_day_PCT5_change'] < 5)
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuyContinueHighTail-uptrend(CheckChart-risky)')
        elif(regression_data['high'] > regression_data['high_pre1']
            and regression_data['high_pre1'] > regression_data['high_pre2']
            and -1 < regression_data['PCT_day_change_pre2'] < 0
            and (regression_data['forecast_day_PCT3_change'] < 5)
            and (regression_data['forecast_day_PCT5_change'] < 5)
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuyContinueHighTail-uptrend-0(CheckChart-risky)')
        elif(((regression_data['high'] > regression_data['high_pre1']
            and regression_data['forecast_day_PCT2_change'] > 0)
              or (regression_data['PCT_day_change_pre1'] < -5)
            )
            and regression_data['forecast_day_PCT_change'] > 0  
            and regression_data['forecast_day_PCT3_change'] < 0
            and regression_data['forecast_day_PCT5_change'] < 0 
            and regression_data['forecast_day_PCT7_change'] < 0
            and regression_data['forecast_day_PCT10_change'] < 0
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MayBuyContinueHighTail-uptrend-1(CheckChart-risky)')
         
        
    if((-0.75 < regression_data['PCT_day_change'] < 0) and (-0.75 < regression_data['PCT_change'] < 0)
        and (regression_data['SMA4_2daysBack'] > 0 or regression_data['SMA9_2daysBack'] > 0)
        and regression_data['SMA4'] < 0
        and regression_data['PCT_day_change_pre1'] < -0.5
        and regression_data['PCT_day_change_pre2'] < -0.5
        and (regression_data['PCT_day_change_pre1'] < -1.5 or regression_data['PCT_day_change_pre2'] < -1.5)
        ):   
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:buySMA4Reversal')
    if('RisingMA-Risky' not in regression_data['filter4'] 
        and 'RisingMA' in regression_data['filter4'] 
        and (regression_data['PCT_day_change'] < -0.2)
        and (-1 < regression_data['PCT_day_change'] < 0.5) and (-1.5 < regression_data['PCT_change'] < 0.75)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:buyRisingMA')
    elif(regression_data['SMA4'] > 0 and regression_data['SMA4_2daysBack'] > 0
        and regression_data['SMA4'] > regression_data['SMA4_2daysBack'] 
        and regression_data['SMA4'] > regression_data['SMA9'] > regression_data['SMA25']
        and (-1 < regression_data['PCT_day_change'] < 0) and (-1.5 < regression_data['PCT_change'] < 0.5)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:buyRisingMAShortTerm')
    elif(regression_data['SMA4'] > regression_data['SMA4_2daysBack']  
        and regression_data['SMA9'] < -2 < regression_data['SMA4'] < 0
        and regression_data['SMA25'] < -2 < regression_data['SMA4'] < 0 
        and (-3 < regression_data['PCT_day_change'] < -0.5) and (-3 < regression_data['PCT_change'] < -0.5)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:buyRisingMAShortTerm-Risky')
#     if('ReversalLow' in regression_data['filter3'] 
#         ):
#         add_in_csv(regression_data, regressionResult, ws, 'CommonHL:UPTREND'+regression_data['filter3'])
        
    if(last_4_day_all_up(regression_data) == False
        and -10 < regression_data['year2HighChange'] < -1 and regression_data['weekHighChange'] < -2
        and regression_data['yearLowChange'] > 30
        and ((-1 < regression_data['PCT_change'] < 1.75
            and -1 < regression_data['PCT_day_change'] < 0
            and (high_tail_pct(regression_data) > 0.8
                or low_tail_pct(regression_data) > 0.8
                or abs(regression_data['forecast_day_PCT_change']) > 0.5
                )
            )
            or(-2 < regression_data['PCT_change'] < 1.75
                and -2 < regression_data['PCT_day_change'] < -1
                and regression_data['PCT_day_change_pre1'] < -0.5
                )
            )
        ):
        if(((regression_data['low'] > regression_data['low_pre1'])
            or regression_data['week2LowChange'] < 0.5)):
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:DOWNTREND:buyYear2High')
        else:
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:DOWNTREND:buyYear2High-afterMorningBase')
            
    if((2 < regression_data['PCT_day_change'] < 3.25) and (1.5 < regression_data['PCT_change'] < 3.25)
        and regression_data['bar_high'] > regression_data['bar_high_pre1']
        ):
        if(regression_data['high_pre1'] < regression_data['high_pre2'] < regression_data['high_pre3']):
            add_in_csv(regression_data, regressionResult, ws, None)
        elif(regression_data['year2HighChange'] < -40
            and regression_data['yearHighChange'] < -30
            and (regression_data['PCT_day_change_pre1'] < 0 or regression_data['PCT_day_change_pre2'] < -2)
            and ((regression_data['PCT_day_change_pre1'] > 0 or regression_data['PCT_day_change_pre2'] > -2)
                 or regression_data['high_pre1'] > regression_data['high_pre2'] > regression_data['high_pre3']
                )
            and regression_data['SMA25'] < 0
            and regression_data['SMA50'] < 0
            and regression_data['SMA100'] < 0
            #and regression_data['high_pre1'] > regression_data['high_pre5']
            ):   
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:UPTREND:buyYear2LowLT-50(triggerAfter-9:15)')    
       
def buy_other_indicator(regression_data, regressionResult, reg, ws):
    tail_pct_filter(regression_data, regressionResult)
    base_line(regression_data, regressionResult, reg, ws)
    filterMA(regression_data, regressionResult)
    tail_reversal_filter(regression_data, regressionResult)
    if(regression_data['close'] > 50):
        buy_year_high(regression_data, regressionResult, reg, ws)
        buy_year_low(regression_data, regressionResult, reg, ws, ws)
        buy_down_trend(regression_data, regressionResult, reg, ws)
        buy_final(regression_data, regressionResult, reg, ws, ws)
        buy_pattern(regression_data, regressionResult, reg, ws, ws)
        buy_morning_star_buy(regression_data, regressionResult, reg, ws)
        buy_evening_star_sell(regression_data, regressionResult, reg, ws)
        buy_day_low(regression_data, regressionResult, reg, ws)
        buy_trend_reversal(regression_data, regressionResult, reg, ws)
        buy_trend_break(regression_data, regressionResult, reg, ws)
        buy_consolidation_breakout(regression_data, regressionResult, reg, ws)
        buy_final_candidate(regression_data, regressionResult, reg, ws)
        buy_oi(regression_data, regressionResult, reg, ws)
        buy_up_trend(regression_data, regressionResult, reg, ws)
        buy_market_uptrend(regression_data, regressionResult, reg, ws)
        buy_check_chart(regression_data, regressionResult, reg, ws)
        buy_month3_high_continue(regression_data, regressionResult, reg, ws)
        buy_heavy_uptrend_reversal(regression_data, regressionResult, reg, ws)
        buy_supertrend(regression_data, regressionResult, reg, ws)
        buy_risingMA(regression_data, regressionResult, reg, ws)
        buy_study_risingMA(regression_data, regressionResult, reg, ws)
        buy_random_filters(regression_data, regressionResult, reg, ws)
        buy_tail_reversal_filter(regression_data, regressionResult, reg, ws)
        
        sell_up_trend(regression_data, regressionResult, reg, ws)
        sell_down_trend(regression_data, regressionResult, reg, ws)
        sell_final(regression_data, regressionResult, reg, ws, ws)
        sell_pattern(regression_data, regressionResult, reg, ws, ws)
        sell_base_line_buy(regression_data, regressionResult, reg, ws)
        sell_morning_star_buy(regression_data, regressionResult, reg, ws)
        sell_evening_star_sell(regression_data, regressionResult, reg, ws)
        sell_day_high(regression_data, regressionResult, reg, ws)
        sell_trend_reversal(regression_data, regressionResult, reg, ws)
        sell_trend_break(regression_data, regressionResult, reg, ws)
        sell_consolidation_breakdown(regression_data, regressionResult, reg, ws)
        sell_final_candidate(regression_data, regressionResult, reg, ws)
        sell_oi(regression_data, regressionResult, reg, ws)
        sell_downingMA(regression_data, regressionResult, reg, ws)
        sell_study_downingMA(regression_data, regressionResult, reg, ws)
        sell_market_downtrend(regression_data, regressionResult, reg, ws)
        sell_supertrend(regression_data, regressionResult, reg, ws)
        sell_heavy_downtrend(regression_data, regressionResult, reg, ws)
        sell_check_chart(regression_data, regressionResult, reg, ws)
        sell_random_filter(regression_data, regressionResult, reg, ws)
        sell_tail_reversal_filter(regression_data, regressionResult, reg, ws)
        return True
    if(buy_skip_close_lt_50(regression_data, regressionResult, reg, ws)):
        return True
    return False

def buy_indicator_after_filter_accuracy(regression_data, regressionResult, reg, ws):
    if(is_sell_from_filter_all_filter(regression_data) != True):
        buy_af_high_indicators(regression_data, regressionResult, reg, ws)
        buy_af_up_continued(regression_data, regressionResult, reg, ws)
        buy_af_high_volatility(regression_data, regressionResult, reg, ws)
        buy_af_oi_negative(regression_data, regressionResult, reg, ws)
        buy_af_vol_contract(regression_data, regressionResult, reg, ws)
        buy_af_vol_contract_contrarian(regression_data, regressionResult, reg, ws)
        buy_af_others(regression_data, regressionResult, reg, ws)
    buy_af_low_tail(regression_data, regressionResult, reg, ws)
    #sell_af_high_tail(regression_data, regressionResult, reg, ws)
        
def buy_tail_reversal_filter(regression_data, regressionResult, reg, ws):
    if('MayBuy-CheckChart' in regression_data['filter1']):
        if(regression_data['PCT_day_change_pre1'] > 2
            and regression_data['PCT_day_change_pre2'] < 0
            and is_ema14_sliding_up(regression_data)
            and (last_5_day_all_up_except_today(regression_data) != True)
            and regression_data['bar_low'] >  regression_data['bar_low_pre1']
            and regression_data['low'] >  regression_data['low_pre1']
            ):
            #add_in_csv(regression_data, regressionResult, ws, 'MayBuy-CheckChart(upTrend-lastDayDown)')
            add_in_csv(regression_data, regressionResult, ws, None)
    if(('MayBuy-CheckChart' in regression_data['filter1']) or ('MayBuyCheckChart' in regression_data['filter1'])):
#         if(-3 < regression_data['PCT_day_change'] < -1
#             and -3 < regression_data['PCT_change'] < 0
#             and regression_data['forecast_day_PCT_change'] < 0
#             and regression_data['month3HighChange'] > -5
#             and regression_data['month6HighChange'] > -5
#             and regression_data['yearHighChange'] > -5
#             and regression_data['year2HighChange'] > -5
#             ):
#             add_in_csv(regression_data, regressionResult, ws, 'sell(MayBuy-CheckChart)(downTrendstart-yearHigh)')
        if(-1 < regression_data['PCT_day_change'] < 0
            and -1 < regression_data['PCT_change'] < 0
            and regression_data['PCT_day_change_pre1'] < 0
            and regression_data['PCT_day_change_pre2'] < regression_data['PCT_day_change']
            and (regression_data['PCT_day_change_pre2'] < -1.5 or regression_data['PCT_day_change_pre3'] < -1)
            and regression_data['month3HighChange'] > -10
            and abs(regression_data['year2HighChange']) < abs(regression_data['year2LowChange'])
            #and regression_data['PCT_day_change_pre3'] < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MayBuy-CheckChart(downTrend-mayReverseAtMonth3High)-0')
        elif(-2 < regression_data['PCT_day_change'] < -1
            and -2 < regression_data['PCT_change'] < -1
            and regression_data['PCT_day_change_pre1'] < -1
            and regression_data['PCT_day_change_pre2'] < 0
            and regression_data['PCT_day_change_pre3'] < 0
            and regression_data['monthLow'] >= regression_data['low']
            and (regression_data['forecast_day_PCT7_change'] < -10 or regression_data['forecast_day_PCT10_change'] < -10)
            #and regression_data['PCT_day_change_pre3'] < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MayBuy-CheckChart(downTrend-mayReverseLast4DaysDown)')
        elif(-2 < regression_data['PCT_day_change'] < -1
            and -2 < regression_data['PCT_change'] < -1
            and regression_data['PCT_day_change_pre1'] < -1
            and regression_data['PCT_day_change_pre2'] < 0
            and regression_data['PCT_day_change_pre3'] < 0
            and regression_data['monthLow'] >= regression_data['low']
            #and regression_data['PCT_day_change_pre3'] < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MayBuy-CheckChart(downTrend-mayReverseLast4DaysDown)-Risky')
        elif(-3.5 < regression_data['PCT_day_change'] < -1
            and -3.5 < regression_data['PCT_change'] < -1
            and regression_data['PCT_day_change_pre1'] < -1
            and regression_data['monthLowChange'] < 1
            and regression_data['month3LowChange'] > 1
            and regression_data['yearLowChange'] > 10
            and low_tail_pct(regression_data) < 2.5
            #and regression_data['PCT_day_change_pre3'] < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MayBuy-CheckChart(monthLow-minimumLast2DayDown)')
    
    if((-0.5 < regression_data['PCT_change'] < 0.5)
        and (-1 < regression_data['PCT_day_change'] < 0)
        and high_tail_pct(regression_data) < 0.5
        and (0.6 <= low_tail_pct(regression_data) < 1.5)
        and (low_tail_pct(regression_data) - high_tail_pct(regression_data)) >= 0.6
        #and (regression_data['score'] == '10')
        and (regression_data['PCT_day_change_pre1'] > 2)
        ):
        if(abs_month3High_more_than_month3Low(regression_data)
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MayBuyCheckChart-(|`|_|)')
    elif(is_algo_buy(regression_data)
        and (-2.5 < regression_data['PCT_day_change'] <= -1)
        and (-4 < regression_data['PCT_change'] <= -1)
        ):
        if(('MayBuyCheckChart' in regression_data['filter1'])
            and (-15 < regression_data['forecast_day_PCT5_change'] < -1)
            and (-15 < regression_data['forecast_day_PCT7_change'] < -1)
            and (-15 < regression_data['forecast_day_PCT10_change'] < -1)
            and high_tail_pct(regression_data) < 0.8
            ):
            if(1.8 < low_tail_pct(regression_data) < 2.5):
                add_in_csv(regression_data, regressionResult, ws, '$$MayBuyCheckChart-(IndexNotUpInSecondHalf)')
                
    if(is_algo_buy(regression_data)):        
        if(('MayBuyCheckChart' in regression_data['filter1'])
            and ('Reversal' not in regression_data['filter3'])
            and high_tail_pct(regression_data) < 0.5
            ):
            if((-2 <= regression_data['PCT_day_change'] < -1) and (-2 <= regression_data['PCT_change'] < 0)
                and 3 > low_tail_pct(regression_data) > 1.8
                and regression_data['PCT_day_change_pre1'] < 0
                and regression_data['PCT_day_change_pre2'] < 0
                and regression_data['PCT_day_change_pre3'] < 0
                ):
                add_in_csv(regression_data, regressionResult, ws, "$$ML:MayBuyCheckChart-last3DayDown")
            elif((-2 <= regression_data['PCT_day_change'] < -1) and (-2 <= regression_data['PCT_change'] < 0)
                and 3 > low_tail_pct(regression_data) > 1.8
                ):
                add_in_csv(regression_data, regressionResult, ws, "$$ML:(check-chart-2-3MidCapCross)MayBuyCheckChart-Risky") 
    elif(('MayBuyCheckChart' in regression_data['filter1'])
        and ('Reversal' not in regression_data['filter3'])
        and high_tail_pct(regression_data) < 0.5
        ):
        if((-5 <= regression_data['PCT_day_change'] < -3) and (-6 <= regression_data['PCT_change'] < -2)
            and 5 > low_tail_pct(regression_data) > 2.8
            ):
            add_in_csv(regression_data, regressionResult, ws, "$$ML:(Test)MayBuyCheckChart-PCTDayChangeLT(-3)BigHighTail")
            
    if(regression_data['year2LowChange'] > 3
        and regression_data['low'] == regression_data['weekLow']
        and low_tail_pct(regression_data) > 1.5
        and (regression_data['PCT_day_change'] < 2 or is_algo_buy(regression_data))
        ):
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, "weekLowLTweek2")
        if(regression_data['monthLowChange'] < 3
            and regression_data['monthLow'] != regression_data['weekLow']
            and regression_data['monthLow'] == regression_data['month2Low']
            #and (is_algo_buy(regression_data)
                #or ((regression_data['bar_low'] - regression_data['month2BarLow'])/regression_data['month2BarLow'])*100 < 1.5
                #or regression_data['PCT_day_change'] > 0
                #)
            ):
            if(regression_data['monthLowChange'] < -5):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-5")
            elif(regression_data['monthLowChange'] < -4):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-4")
            elif(regression_data['monthLowChange'] < -3):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-3")
            elif(regression_data['monthLowChange'] < -2):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-2")
            elif(regression_data['monthLowChange'] < -1):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-1")
            elif(regression_data['monthLowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT0")
            elif(regression_data['monthLowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT1")     
            elif(regression_data['monthLowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT2")
            elif(regression_data['monthLowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT3")     
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-monthLowReversal")
        elif(regression_data['monthLowChange'] < 3
            and regression_data['monthLow'] != regression_data['weekLow']
            ):
            if(regression_data['monthLowChange'] < -5):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-5")
            elif(regression_data['monthLowChange'] < -4):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-4")
            elif(regression_data['monthLowChange'] < -3):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-3")
            elif(regression_data['monthLowChange'] < -2):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-2")
            elif(regression_data['monthLowChange'] < -1):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT-1")
            elif(regression_data['monthLowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT0")
            elif(regression_data['monthLowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT1")     
            elif(regression_data['monthLowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT2")
            elif(regression_data['monthLowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT3") 
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-monthLowReversal-risky")
        elif((regression_data['monthLowChange'] < 5 and regression_data['SMA9'] > 0)
            and regression_data['monthLow'] != regression_data['weekLow']
            and regression_data['monthLow'] == regression_data['month2Low']
            ):
            if(regression_data['monthLowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT0")
            elif(regression_data['monthLowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT1")     
            elif(regression_data['monthLowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT2")
            elif(regression_data['monthLowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT3")
            elif(regression_data['monthLowChange'] < 4):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT4")
            elif(regression_data['monthLowChange'] < 5):
                add_in_csv(regression_data, regressionResult, ws, "monthLowChangeLT5") 
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-monthLowReversal-SMA9GT0")
        elif(regression_data['month2LowChange'] < 3
            and regression_data['month2Low'] != regression_data['weekLow']
            and regression_data['month2Low'] == regression_data['month3Low']
            #and (is_algo_buy(regression_data)
                #or ((regression_data['bar_low'] - regression_data['month3BarLow'])/regression_data['month3BarLow'])*100 < 1.5
                #or regression_data['PCT_day_change'] > 0
                #)
            ):
            if(regression_data['month2LowChange'] < -5):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-5")
            elif(regression_data['month2LowChange'] < -4):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-4")
            elif(regression_data['month2LowChange'] < -3):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-3")
            elif(regression_data['month2LowChange'] < -2):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-2")
            elif(regression_data['month2LowChange'] < -1):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-1")
            elif(regression_data['month2LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT0")
            elif(regression_data['month2LowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT1")     
            elif(regression_data['month2LowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT2")
            elif(regression_data['month2LowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT3") 
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month2LowReversal")
        elif(regression_data['month2LowChange'] < 3
            and regression_data['month2Low'] != regression_data['weekLow']
            ):
            if(regression_data['month2LowChange'] < -5):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-5")
            elif(regression_data['month2LowChange'] < -4):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-4")
            elif(regression_data['month2LowChange'] < -3):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-3")
            elif(regression_data['month2LowChange'] < -2):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-2")
            elif(regression_data['month2LowChange'] < -1):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT-1")
            elif(regression_data['month2LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT0")
            elif(regression_data['month2LowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT1")     
            elif(regression_data['month2LowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT2")
            elif(regression_data['month2LowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT3") 
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month2LowReversal-risky")
        elif((regression_data['month2LowChange'] < 5 and regression_data['SMA9'] > 0)
            and regression_data['month2Low'] != regression_data['weekLow']
            and regression_data['month2Low'] == regression_data['month3Low']
            ):
            if(regression_data['month2LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT0")
            elif(regression_data['month2LowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT1")     
            elif(regression_data['month2LowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT2")
            elif(regression_data['month2LowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT3")
            elif(regression_data['month2LowChange'] < 4):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT4")
            elif(regression_data['month2LowChange'] < 5):
                add_in_csv(regression_data, regressionResult, ws, "month2LowChangeLT5")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month2LowReversal-SMA9GT0")
        elif(regression_data['month3LowChange'] < 3
            and regression_data['month3Low'] != regression_data['weekLow']
            #and regression_data['month3Low'] != regression_data['low_month3']
            and regression_data['month6Low'] == regression_data['yearLow']
            ):
            if(regression_data['month3LowChange'] < -5):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT-5")
            elif(regression_data['month3LowChange'] < -4):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT-4")
            elif(regression_data['month3LowChange'] < -3):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT-3")
            elif(regression_data['month3LowChange'] < -2):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT-2")
            elif(regression_data['month3LowChange'] < -1):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT-1")
            elif(regression_data['month3LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT0")
            elif(regression_data['month3LowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT1")     
            elif(regression_data['month3LowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT2")
            elif(regression_data['month3LowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "monthLow3ChangeLT3") 
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month3LowReversal")
        elif(regression_data['month3LowChange'] < 0
            and regression_data['month3Low'] != regression_data['weekLow']
            and regression_data['month3Low'] == regression_data['month6Low']
            ):
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month3LowBreakReversal")
        elif(regression_data['month6LowChange'] < 3
            and regression_data['month6Low'] != regression_data['weekLow']
            #and regression_data['month3Low'] != regression_data['low_month6']
            and regression_data['yearLow'] == regression_data['year2Low']
            ):
            if(regression_data['month6LowChange'] < -5):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT-5")
            elif(regression_data['month6LowChange'] < -4):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT-4")
            elif(regression_data['month6LowChange'] < -3):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT-3")
            elif(regression_data['month6LowChange'] < -2):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT-2")
            elif(regression_data['month6LowChange'] < -1):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT-1")
            elif(regression_data['month6LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT0")
            elif(regression_data['month6LowChange'] < 1):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT1")     
            elif(regression_data['month6LowChange'] < 2):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT2")
            elif(regression_data['month6LowChange'] < 3):
                add_in_csv(regression_data, regressionResult, ws, "monthLow6ChangeLT3") 
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month6LowReversal")
        elif(regression_data['month6LowChange'] < 0
            and regression_data['month6Low'] != regression_data['weekLow']
            and regression_data['month6Low'] == regression_data['yearLow']
            ):
            add_in_csv(regression_data, regressionResult, ws, "(Test)MayBuyCheckChart-month6LowBreakReversal")
                          
def buy_year_high(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(float(regression_data['forecast_day_VOL_change']) > 70
       and regression_data['PCT_day_change_pre1'] > -0.5
       and no_doji_or_spinning_buy_india(regression_data)
       ):
        if(-15 < regression_data['yearHighChange'] < -5 and regression_data['yearLowChange'] > 30
            and -0.5 < regression_data['PCT_day_change'] < 2 and regression_data['forecast_day_PCT2_change'] <= 2
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyYearHigh-0')
            return True
    elif(float(regression_data['forecast_day_VOL_change']) > 35
       and regression_data['PCT_day_change_pre1'] > -0.5
       and no_doji_or_spinning_buy_india(regression_data)
       ):
        if(-15 < regression_data['yearHighChange'] < -5 and regression_data['yearLowChange'] > 30
            and -0.5 < regression_data['PCT_day_change'] < 2 and regression_data['forecast_day_PCT2_change'] <= 2
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyYearHigh-1')
            return True
    elif(float(regression_data['forecast_day_VOL_change']) > 50
       and regression_data['PCT_day_change_pre1'] > -0.5
       and no_doji_or_spinning_buy_india(regression_data)
       ):
        if(-5 <= regression_data['yearHighChange'] < -1 and regression_data['yearLowChange'] > 15 
            and -0.5 < regression_data['PCT_day_change'] < 2 and regression_data['forecast_day_PCT2_change'] <= 2
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyYearHigh-2')
            return True
    return False

def buy_year_low(regression_data, regressionResult, reg, ws, ws1):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if((-1.5 < regression_data['PCT_day_change'] < 0) and (-2 < regression_data['PCT_change'] < 0)
        and regression_data['year2HighChange'] < -50
        and regression_data['yearHighChange'] < -30
        and regression_data['month3LowChange'] > 10
        and regression_data['monthLowChange'] > 10
        and 0 < regression_data['SMA4'] < 10
        and 2 < regression_data['SMA9'] < 10
        ):   
        add_in_csv(regression_data, regressionResult, ws, 'buyYear2LowMovingSMA')
        return True
    elif(1 < regression_data['yearLowChange'] < 5 and regression_data['yearHighChange'] < -30 
        and 2 < regression_data['PCT_day_change'] < 6 and 2 < regression_data['PCT_day_change'] < 6
        and regression_data['forecast_day_PCT2_change'] > 0 and regression_data['forecast_day_PCT_change'] > 0
        and float(regression_data['forecast_day_VOL_change']) > 35
        ):
        add_in_csv(regression_data, regressionResult, ws, 'buyYearLow')
        return True
    elif(5 < regression_data['yearLowChange'] < 15 and regression_data['yearHighChange'] < -75 
        and 2 < regression_data['PCT_day_change'] < 5 and 2 < regression_data['PCT_day_change'] < 5
        and 5 > regression_data['forecast_day_PCT2_change'] > -0.5 and regression_data['forecast_day_PCT_change'] > 0
        and float(regression_data['forecast_day_VOL_change']) > 35
        ):
        add_in_csv(regression_data, regressionResult, ws1, 'buyYearLow1')
        return True
    return False

def buy_down_trend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if((('NearLowYear2' in regression_data['filter3'])
        or ('BreakLowYear2' in regression_data['filter3'])
        or (regression_data['year2HighChange'] < -50 and regression_data['month3LowChange'] < 10)
        ) 
        and (regression_data['year2HighChange'] < -40)
        and high_tail_pct(regression_data) < 1
        and low_tail_pct(regression_data) > 1.1
        and regression_data['PCT_day_change'] < -1
        and regression_data['PCT_day_change_pre1'] < -2
        and regression_data['PCT_day_change_pre2'] < -2
        ):
        if('NearLowYear2' in regression_data['filter3']):
            add_in_csv(regression_data, regressionResult, ws, 'NearLowYear2')
        elif('BreakLowYear2' in regression_data['filter3']):
            add_in_csv(regression_data, regressionResult, ws, 'BreakLowYear2')
        add_in_csv(regression_data, regressionResult, ws, 'buyYear2Low')
        return True
    elif(regression_data['forecast_day_PCT10_change'] < -10
        and regression_data['year2HighChange'] < -60
        and regression_data['month3LowChange'] < 10
        and -1.5 < regression_data['forecast_day_PCT_change']
        and 3 < regression_data['PCT_day_change'] < 7
        and 2 < regression_data['PCT_change'] < 8
        and (regression_data['PCT_day_change_pre1'] < -4 or regression_data['PCT_day_change_pre2'] < -4)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'sellYear2LowContinue')
        return True
    if(all_day_pct_change_negative_except_today(regression_data) 
        and regression_data['forecast_day_PCT_change'] > 0 
        and 0 < regression_data['PCT_day_change'] < 4 and 0 < regression_data['PCT_change']
        and regression_data['yearHighChange'] < -10 
        and high_tail_pct(regression_data) < 0.5
       ):
        if(last_5_day_all_down_except_today(regression_data)
            and ten_days_less_than_minus_ten(regression_data)
            and regression_data['forecast_day_PCT7_change'] < -5
            and regression_data['forecast_day_PCT10_change'] < -10
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyDownTrend-0(Risky)')
            return True
        elif(last_5_day_all_down_except_today(regression_data)
            and ten_days_less_than_minus_seven(regression_data)
            and regression_data['forecast_day_PCT7_change'] < -4
            and regression_data['forecast_day_PCT10_change'] < -4
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyDownTrend-1')
            return True
        elif(regression_data['forecast_day_PCT7_change'] < -4
            and regression_data['forecast_day_PCT10_change'] < -4
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyDownTrend-2(Risky)')
            return True
    
    if(ten_days_less_than_minus_ten(regression_data)
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0
        and regression_data['forecast_day_PCT5_change'] < 0
        and regression_data['forecast_day_PCT7_change'] < 0
        and regression_data['forecast_day_PCT7_change'] < 0
        and 2 < regression_data['PCT_day_change'] < 5
        and 2 < regression_data['PCT_change'] < 5
        and regression_data['PCT_day_change_pre2'] < 0
        and regression_data['PCT_day_change_pre3'] < 0
        ):
        add_in_csv(regression_data, regressionResult, ws, 'buyTenDaysLessThanMinusTen')
        return True
        
    return False

def buy_final(regression_data, regressionResult, reg, ws, ws1):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['yearHighChange'] < -10 
        and regression_data['yearLowChange'] > 5
        and regression_data['month3LowChange'] > 3
        and 4 > regression_data['PCT_day_change'] > 1 and 4 > regression_data['PCT_change'] > 1
        and abs(regression_data['PCT_day_CH']) < 0.3
        and regression_data['forecast_day_VOL_change'] > 0
        and high_tail_pct(regression_data) < 1
        and str(regression_data['score']) != '0-1'
        ):   
        if(str(regression_data['sellIndia']) == '' and -90 < regression_data['yearHighChange'] < -10
            and(ten_days_less_than_minus_ten(regression_data)
                 or (last_5_day_all_down_except_today(regression_data)
                    and ten_days_less_than_minus_seven(regression_data)
                    )
            ) 
            and regression_data['forecast_day_PCT7_change'] < -5 
            and regression_data['forecast_day_PCT5_change'] < 0.5 
            and regression_data['forecast_day_PCT4_change'] < 0.5 
            and 5 > regression_data['forecast_day_PCT2_change'] > -0.5 
            and regression_data['forecast_day_PCT_change'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyFinal')
            return True
        elif(str(regression_data['sellIndia']) == '' and -90 < regression_data['yearHighChange'] < -10
            and(ten_days_less_than_minus_ten(regression_data)
                 or (last_5_day_all_down_except_today(regression_data)
                    and ten_days_less_than_minus_seven(regression_data)
                    )
            ) 
            and regression_data['forecast_day_PCT7_change'] <= -1 
            and regression_data['forecast_day_PCT5_change'] <= 1  
            and regression_data['forecast_day_PCT2_change'] > 0 
            and regression_data['forecast_day_PCT_change'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyFinal1')
            return True
    return False

def buy_af_high_indicators(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_cla, kNeighboursValue_cla = get_reg_or_cla(regression_data, False)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if(regression_data['PCT_day_change'] < 5 and abs(regression_data['PCT_day_change']) > 1.5 or (low_tail_pct(regression_data) > 1.5 and high_tail_pct(regression_data) < 1)):
        if(is_algo_buy(regression_data, True)
            and ((mlpValue_other >= 1 or kNeighboursValue_other >= 1) or regression_data['PCT_day_change'] < 0)
            and (abs(regression_data['filter_345_avg']) == 0
                 or (float(regression_data['filter_345_pct']) > -1.5 
                     and abs(regression_data['filter_345_avg']) > -50)
                )
            and (abs(regression_data['filter_all_avg']) == 0
                 or (float(regression_data['filter_all_pct']) > -1.5 
                     and abs(regression_data['filter_all_avg']) > -50)
                )
            and (abs(regression_data['filter_pct_change_avg']) == 0
                 or (float(regression_data['filter_pct_change_pct']) > -1.5 
                     and abs(regression_data['filter_pct_change_avg']) > -50)
                )
            and (abs(regression_data['filter_avg']) == 0
                 or (float(regression_data['filter_pct']) > -1.5 
                     and abs(regression_data['filter_avg']) > -50)
                )
            ):
            if(2 < high_tail_pct(regression_data) < 4
                and (regression_data['PCT_day_change'] + high_tail_pct(regression_data)) > 6.5):
                return False
            if(mlpValue >= 2.0 and kNeighboursValue >= 2.0
               and regression_data['month3HighChange'] < -3
               and regression_data['month3LowChange'] > 3
               and regression_data['yearHighChange'] < -10
               and regression_data['yearLowChange'] > 10
               #and regression_data['forecast_day_PCT10_change'] < 15
               and (high_tail_pct(regression_data) < 1.5 and (high_tail_pct(regression_data) < low_tail_pct(regression_data)))
               ):
                if(3 > regression_data['PCT_day_change'] > 1.5 and 3 > regression_data['PCT_change'] > -0.5
                   and regression_data['forecast_day_PCT_change'] > 0
                   and high_tail_pct(regression_data) < 1
                   ):
                    add_in_csv(regression_data, regressionResult, ws, 'buyHighIndicators(shouldNotOpenInMinus)')
                    return True
                if(1 > regression_data['PCT_day_change'] > 0 and 2.5 > regression_data['PCT_change'] > -0.5):
                    #add_in_csv(regression_data, regressionResult, ws, '(longDownTrend)buyHighIndicators')
                    add_in_csv(regression_data, regressionResult, ws, None)
                    return True
            
            if(is_algo_buy(regression_data)
                and (regression_data['weekLowChange'] < 0) 
                and (4 < regression_data['PCT_change'] < 10)
                and (4 < regression_data['PCT_day_change'] < 10)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**(HighBothBuyAll)buyDowntrendReversal') 
            
            if(is_algo_buy(regression_data)
                and (mlpValue_other >= 0 and kNeighboursValue_other >= 0)
                and ((4.0 <= mlpValue < 10.0 and 2.0 <= kNeighboursValue < 10.0) or (2.0 <= mlpValue < 10.0 and 4.0 <= kNeighboursValue < 10.0))
                ):
                if(mlpValue_cla > 0 or kNeighboursValue_cla > 0):
                    add_in_csv(regression_data, regressionResult, ws, '**buyHighIndicators-0')
            elif(is_algo_buy(regression_data)
                and (mlpValue_other >= 0 and kNeighboursValue_other >= 0)
                #and (mlpValue_cla > 0 or kNeighboursValue_cla > 0)
                and (2.0 <= mlpValue < 10.0 and 1.5 <= kNeighboursValue < 10.0)
                and ((mlpValue + kNeighboursValue) > 5)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**buyHighIndicators-Risky')
            elif(is_algo_buy(regression_data)
                and (mlpValue_other >= 0 and kNeighboursValue_other >= 0)
                and (1 < mlpValue_cla <= 5)
                and (0 <= mlpValue < 10.0 and 0 <= kNeighboursValue < 10.0)
                and (1 <= mlpValue or 1 <= kNeighboursValue)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**(uptrend)buyHighMLPClaIndicators')
            elif((5 <= mlpValue < 15 and 2 < mlpValue_other)
                and (-1 <= kNeighboursValue < 15)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**(Test)buyHighMLPClaIndicators-Risky')
        elif((5 <= mlpValue < 15 and 2 < mlpValue_other)
            and (-1 <= kNeighboursValue < 15)
            ):
            add_in_csv(regression_data, regressionResult, ws, '**(Test)buyHighMLPClaIndicators-Risky')
    elif(regression_data['PCT_day_change'] < 5
        and (5 <= mlpValue < 15 and 2 < mlpValue_other)
        and (-1 <= kNeighboursValue < 15)
        ):
        add_in_csv(regression_data, regressionResult, ws, '**(Test)buyHighMLPClaIndicators-Risky')
    return False
    
def buy_af_low_tail(regression_data, regressionResult, reg, ws):
    if(high_tail_pct(regression_data) <= 1 and 1.3 <= low_tail_pct(regression_data) <= 2
        and low_tail_pct(regression_data) > (high_tail_pct(regression_data) + 0.5)
        ):
        if(-3 < regression_data['PCT_day_change'] < -0.5 and -3 < regression_data['PCT_change'] < -0.5
           and (regression_data['PCT_day_change'] < -1 or regression_data['PCT_change'] < -1)
           ):
           add_in_csv(regression_data, regressionResult, ws, '%%mayBuyTail')
        elif(-3 < regression_data['PCT_day_change'] < 1 and -3 < regression_data['PCT_change'] < 1 and 1.4 <= low_tail_pct(regression_data)):
           add_in_csv(regression_data, regressionResult, ws, '%%mayBuyTail-Risky')
    elif(high_tail_pct(regression_data) <= 1 and 2 <= low_tail_pct(regression_data) <= 4):
        if(-4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1):
           add_in_csv(regression_data, regressionResult, ws, '%%mayBuyTail')
    
def buy_af_up_continued(regression_data, regressionResult, reg, ws):
    if(high_tail_pct(regression_data) < 1.1 and low_tail_pct(regression_data) < 2):
        if(1.9 < regression_data['PCT_day_change'] < 3 and 1 < regression_data['PCT_change'] < 3):
            if(regression_data['PCT_day_change_pre1'] < 0.75 and regression_data['PCT_change_pre1'] < 1):
                add_in_csv(regression_data, regressionResult, ws, '%%mayBuyUpContinueLT3')
            else:
                add_in_csv(regression_data, regressionResult, ws, '%%mayBuyUpContinueLT3-Risky')
        elif(3.25 < regression_data['PCT_day_change'] < 4 and 3 < regression_data['PCT_change'] < 4):
            if(regression_data['PCT_day_change_pre1'] < 0.75 and regression_data['PCT_change_pre1'] < 1):
                add_in_csv(regression_data, regressionResult, ws, '%%mayBuyUpContinueGT3')
            else:
                add_in_csv(regression_data, regressionResult, ws, '%%mayBuyUpContinueGT3-Risky')

def buy_af_high_volatility(regression_data, regressionResult, reg, ws):
    if(regression_data['PCT_day_change'] < -8
        and regression_data['PCT_change'] < -8
        and regression_data['forecast_day_PCT2_change'] < -20
        ):
        add_in_csv(regression_data, regressionResult, ws, None, '%%mayBuyHighVolatileAfter10:20DownL2T-20')
        return True
    elif('%%maySellTail' in regression_data['filter']
        and (('upTrend' in regression_data['series_trend'])
             or pct_day_change_trend(regression_data) >= 3
             )
        ):
        countGt, countLt = pct_day_change_counter(regression_data)
        if(countGt > countLt):
            add_in_csv(regression_data, regressionResult, ws, None, '%%BuyUpTrend-highTail')
        elif(countGt == countLt
            and regression_data['PCT_day_change'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws, None, '%%SellUpTrend-highTail-alternateUpDown')
        return True
    else:
        return False

def buy_af_oi_negative(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['redtrend'] == -1
        and regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < -0.5
        and regression_data['forecast_day_PCT3_change'] < -0.5
        and regression_data['forecast_day_PCT4_change'] < -0.5
        and regression_data['forecast_day_PCT5_change'] < -0.5
        and regression_data['forecast_day_PCT7_change'] < -0.5
        and ten_days_less_than_minus_ten(regression_data)
        and float(regression_data['forecast_day_VOL_change']) < -40
        and regression_data['PCT_day_change_pre1'] < regression_data['PCT_day_change']
        and float(regression_data['contract']) < 0
        and float(regression_data['oi']) < 5
        and (regression_data['yearHighChange'] < -15 or regression_data['yearLowChange'] < 15)
        ):
        if(is_algo_buy(regression_data)):
            if(((-1 < regression_data['PCT_day_change'] < 0 and -1 < regression_data['PCT_change'] < -0.5)
                or (-1 < regression_data['PCT_day_change'] < -0.5 and -1 < regression_data['PCT_change'] < 0))
                ):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyNegativeOI-0-checkBase(1%down)')
                return True
            if(-2 < regression_data['PCT_day_change'] < 0 and -2 < regression_data['PCT_change'] < 0
                ):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyNegativeOI-1-checkBase(1%down)')
                return True
        return False
    return False

def buy_af_vol_contract(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if((ten_days_more_than_ten(regression_data) == False
        or ten_days_more_than_fifteen(regression_data) == True
        or (regression_data['forecast_day_PCT5_change'] < 5 and regression_data['forecast_day_PCT7_change'] < 5)) 
        #and ((mlpValue >= 0.3 and kNeighboursValue >= 0.3) or is_algo_buy(regression_data))
        and (float(regression_data['contract']) != 0 or float(regression_data['oi']) != 0)
        and float(regression_data['contract']) > 10
        and ((regression_data['PCT_day_change'] > 1 or regression_data['PCT_change'] > 1)
            or
            (regression_data['PCT_day_change'] > 0.75 and regression_data['PCT_change'] > 0.75
            and regression_data['forecast_day_PCT_change'] < 2
            and regression_data['forecast_day_PCT2_change'] < 2
            and regression_data['forecast_day_PCT3_change'] < 2
            and regression_data['forecast_day_PCT4_change'] < 2
            and regression_data['forecast_day_PCT4_change'] < 2
            and regression_data['forecast_day_PCT5_change'] < 2
            and regression_data['forecast_day_PCT7_change'] < 2
            and regression_data['forecast_day_PCT10_change'] < 2
            )
        )
        and regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and 5 > regression_data['forecast_day_PCT4_change'] > 0
        and regression_data['forecast_day_PCT5_change'] < 5
        and regression_data['forecast_day_PCT7_change'] < 5
        and regression_data['forecast_day_PCT10_change'] < 5
        and (regression_data['PCT_day_change'] < 0
            or regression_data['PCT_day_change_pre1'] < 0
            or regression_data['PCT_day_change_pre2'] < 0
            or regression_data['PCT_day_change_pre3'] < 0
            or regression_data['PCT_day_change_pre4'] < 0
           )
        and regression_data['yearHighChange'] < -2
        and preDayPctChangeUp_orVolHigh(regression_data)
        and regression_data['open'] > 50
        and last_4_day_all_up(regression_data) == False
        and high_tail_pct(regression_data) < 1.5
        and regression_data['month3HighChange'] < -7.5
        and (regression_data['forecast_day_VOL_change'] > 150
            or (regression_data['PCT_day_change_pre2'] < 0
                and (((regression_data['volume'] - regression_data['volume_pre2'])*100)/regression_data['volume_pre2']) > 100
                and (((regression_data['volume'] - regression_data['volume_pre3'])*100)/regression_data['volume_pre3']) > 100
               )
           )
       ):
        if((regression_data['forecast_day_VOL_change'] > 70 and 0.5 < regression_data['PCT_day_change'] < 2 and 1 < regression_data['PCT_change'] < 2)
            and float(regression_data['contract']) > 10
            ):
            if((mlpValue >= 0.3 and kNeighboursValue >= 0.3) or is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyOI-0')
                return True
            return False
        elif((regression_data['forecast_day_VOL_change'] > 35 and 0.5 < regression_data['PCT_day_change'] < 2 and 1 < regression_data['PCT_change'] < 2)
            and float(regression_data['contract']) > 20
            #and regression_data['PCT_day_change_pre1'] > -0.5
            ):
            if((mlpValue >= 0.3 and kNeighboursValue >= 0.3) or is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyOI-1')
                return True
            return False
        elif((regression_data['forecast_day_VOL_change'] > 150 and 1 < regression_data['PCT_day_change'] < 3 and 1 < regression_data['PCT_change'] < 3)
            #and regression_data['PCT_day_change_pre1'] > -0.5
            ):
            if((mlpValue >= 0.3 and kNeighboursValue >= 0.3) or is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyOI-2-checkBase-Risky')
                return True
            return False
        elif(((regression_data['forecast_day_VOL_change'] > 400 and 1 < regression_data['PCT_day_change'] < 3.5 and 1 < regression_data['PCT_change'] < 3.5)
            or (regression_data['forecast_day_VOL_change'] > 500 and 1 < regression_data['PCT_day_change'] < 4.5 and 1 < regression_data['PCT_change'] < 4.5)
            )
            #and regression_data['PCT_day_change_pre1'] > -0.5
            ):
            if((mlpValue >= 0.3 and kNeighboursValue >= 0.3) or is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyOI-3-checkBase-Risky')
                return True
            return False    
        elif((regression_data['forecast_day_VOL_change'] > 500 and 1 < regression_data['PCT_day_change'] < 5 and 1 < regression_data['PCT_change'] < 5)
            and float(regression_data['contract']) > 50 
            and (regression_data['forecast_day_PCT10_change'] < -8 or regression_data['forecast_day_PCT7_change'] < -8)
            ):
            if((mlpValue >= 0.3 and kNeighboursValue >= 0.3) or is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyOI-4-checkBase-Risky')
                return True
            return False    
    return False

def buy_af_vol_contract_contrarian(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(
        (float(regression_data['contract']) != 0 or float(regression_data['oi']) != 0)
        and (ten_days_more_than_ten(regression_data) == True and ten_days_more_than_fifteen(regression_data) == False 
             and (regression_data['forecast_day_PCT5_change'] > 5 and  regression_data['forecast_day_PCT7_change'] > 5))
        and float(regression_data['contract']) > 10
        and (regression_data['PCT_day_change'] > 1 or regression_data['PCT_change'] > 1)
        and regression_data['forecast_day_PCT_change'] > 0.5
        and regression_data['forecast_day_PCT2_change'] > 0.5
        and regression_data['forecast_day_PCT3_change'] > 0
        and regression_data['forecast_day_PCT4_change'] > 0
        and preDayPctChangeUp_orVolHigh(regression_data)
        and regression_data['open'] > 50
        and last_7_day_all_up(regression_data) == False
        and high_tail_pct(regression_data) > 1.5
        and low_tail_pct(regression_data) < 1
        ):
        if((regression_data['forecast_day_VOL_change'] > 70 and 0.75 < regression_data['PCT_day_change'] < 2 and 0.5 < regression_data['PCT_change'] < 2)
            and float(regression_data['contract']) > 10
            ):
            add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):sellReversalOI-0(openAroundLastCloseAnd5MinuteChart)')
            return True
        elif((regression_data['forecast_day_VOL_change'] > 35 and 0.75 < regression_data['PCT_day_change'] < 2 and 0.5 < regression_data['PCT_change'] < 2)
            and float(regression_data['contract']) > 20
            ):
            add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):sellReversalOI-1(openAroundLastCloseAnd5MinuteChart)')
            return True
#         elif((regression_data['forecast_day_VOL_change'] > 150 and 0.75 < regression_data['PCT_day_change'] < 3 and 0.5 < regression_data['PCT_change'] < 3)
#             and regression_data['PCT_day_change_pre1'] > -0.5
#             ):
#             add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):sellReversalOI-2(openAroundLastCloseAnd5MinuteChart)')
#             return True
#         elif(((regression_data['forecast_day_VOL_change'] > 400 and 0.75 < regression_data['PCT_day_change'] < 3.5 and 0.5 < regression_data['PCT_change'] < 3.5)
#             or (regression_data['forecast_day_VOL_change'] > 500 and 0.75 < regression_data['PCT_day_change'] < 4.5 and 0.5 < regression_data['PCT_change'] < 4.5)
#             )
#             and regression_data['forecast_day_PCT10_change'] > 10
#             ):
#             if(('P@[') in regression_data['buyIndia']):
#                 add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):sellReversalOI-3-checkBase-(openAroundLastCloseAnd5MinuteChart)')
#                 return True
#             elif(preDayPctChangeUp_orVolHigh(regression_data)):
#                 add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):buyContinueOI-3-checkBase-(openAroundLastCloseAnd5MinuteChart)')
#                 return True
#         elif((regression_data['forecast_day_VOL_change'] > 500 and 0.75 < regression_data['PCT_day_change'] < 5 and 0.5 < regression_data['PCT_change'] < 5)
#             and float(regression_data['contract']) > 50 
#             and regression_data['forecast_day_PCT10_change'] > 10
#             ):
#             add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):sellReversalOI-4-checkBase-(openAroundLastCloseAnd5MinuteChart)')
#             return True
#     if((('P@[' not in str(regression_data['buyIndia'])) and ('P@[' not in str(regression_data['sellIndia'])))
#         and (-2 < regression_data['PCT_day_change'] < 0 and -2 < regression_data['PCT_change'] < 0)
#         and kNeighboursValue >= 1
#         and mlpValue >= 0
#         and is_recent_consolidation(regression_data) == False
#         and str(regression_data['score']) == '10'
#         and (high_tail_pct(regression_data) < 1 and (high_tail_pct(regression_data) < low_tail_pct(regression_data)))
#         ):
#             add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):buyReversalKNeighbours(downTrend)(downLastDay)')
#             return True    
    return False

def buy_af_others(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, True)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, True)
    mlpValue_cla, kNeighboursValue_cla = get_reg_or_cla(regression_data, False)
    mlpValue_other_cla, kNeighboursValue_other_cla = get_reg_or_cla_other(regression_data, False)
                       
    if(len(regression_data['filter']) > 9
        and 2 < regression_data['PCT_day_change'] < 4
        and 2 < regression_data['PCT_change'] < 4
        and high_tail_pct(regression_data) < 1.3
        and is_buy_from_filter_all_filter_relaxed(regression_data)
        ):
        if(regression_data['PCT_day_change_pre1'] < -1
            and regression_data['PCT_day_change_pre2'] > 1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Buy-Relaxed-01')
        elif(-1 < regression_data['PCT_day_change_pre1'] < 1
            and regression_data['PCT_day_change_pre2'] < -1
            and regression_data['PCT_day_change_pre3'] < -1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Buy-Relaxed-02')
        elif(regression_data['PCT_day_change_pre1'] > 1
            and regression_data['PCT_day_change_pre2'] > 1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None)
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None)
    
    if('NA$NA:(shortUpTrend)$NA' in regression_data['series_trend']):
        add_in_csv(regression_data, regressionResult, ws, None, 'shortUpTrend')
        
            
    return False
              
def buy_pattern(regression_data, regressionResult, reg, ws, ws1):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    score = ''
    buy_pattern_without_mlalgo(regression_data, regressionResult)
    if(regression_data['yearLowChange'] > 10):
        if(is_algo_buy(regression_data)
            and regression_data['buyIndia_avg'] > 1.5 
            and high_tail_pct(regression_data) < 1.5
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyPatternsML')
            return True   
    
def buy_morning_star_buy(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(-8 < regression_data['forecast_day_PCT_change'] < 0
        and ((regression_data['year2LowChange'] > 15) or (regression_data['year2LowChange'] < 2))
        and high_tail_pct(regression_data) < 1
        and low_tail_pct(regression_data) > 2.5
        and regression_data['forecast_day_PCT10_change'] < (regression_data['forecast_day_PCT_change'] - 5)
        ):
        if(-6 < regression_data['PCT_day_change'] < -2 and -6 < regression_data['PCT_change'] < -2
            and (regression_data['close'] - regression_data['low']) > ((regression_data['open'] - regression_data['close']))
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-HighLowerTail')
            return True
        elif(-6 < regression_data['PCT_day_change'] < 0 and -6 < regression_data['PCT_change'] < 0
            and (regression_data['bar_low'] - regression_data['low']) >= ((regression_data['bar_high'] - regression_data['bar_low']) * 3)
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-HighLowerTail-(checkChart)')
        elif(-6 < regression_data['PCT_day_change'] < 2 and -6 < regression_data['PCT_change'] < 2
            and regression_data['PCT_day_change_pre1'] < 0
            and regression_data['PCT_day_change_pre2'] < 0
            and regression_data['PCT_day_change_pre3'] < 0
            and (regression_data['bar_low'] - regression_data['low']) >= ((regression_data['bar_high'] - regression_data['bar_low']) * 2)
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-(checkChart)')
            return True
    elif(-10 < regression_data['forecast_day_PCT_change'] < -2
        and regression_data['PCT_day_change_pre1'] < 0
        and (ten_days_less_than_minus_seven(regression_data))
        and high_tail_pct(regression_data) < 0.5
        and 2 < low_tail_pct(regression_data) < 3.5
        ):
        if(-3 < regression_data['PCT_day_change'] < 0 and -3 < regression_data['PCT_change'] < 0
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
            and regression_data['forecast_day_PCT10_change'] < -10
            and regression_data['SMA25'] < -8.5
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-dayDown')
            return True
        if(0.3 < regression_data['PCT_day_change'] < 1
            and (regression_data['open'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
            and (regression_data['open'] - regression_data['low']) >= ((regression_data['high'] - regression_data['close']) * 3)
            and regression_data['year2HighChange'] > -50
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-dayUp')
            return True
#         if(high_tail_pct(regression_data) < 1.5
#            and low_tail_pct(regression_data) > 2
#            ):
#             if(-1.5 < regression_data['PCT_day_change'] < 0 and -1.5 < regression_data['PCT_change'] < 0
#                 and (regression_data['close'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
#                 and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##TEST:buyMorningStar-3(Check2-3MidCapCross)')
#                 return True
#             if(0 < regression_data['PCT_day_change'] < 1
#                 and (regression_data['open'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
#                 and (regression_data['open'] - regression_data['low']) >= ((regression_data['high'] - regression_data['close']) * 3)
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##TEST:buyMorningStar-4(Check2-3MidCapCross)')
#                 return True
    return False

def buy_evening_star_sell(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(8 > regression_data['forecast_day_PCT_change'] > 5
        and regression_data['yearLowChange'] > 15
        and low_tail_pct(regression_data) < 0.5
        and high_tail_pct(regression_data) > 3
        ):
        if(5 > regression_data['PCT_day_change'] > 2 and 5 > regression_data['PCT_change'] > 1
            and (regression_data['high']-regression_data['close']) > ((regression_data['close']-regression_data['open']))
            and (regression_data['high']-regression_data['close']) >= ((regression_data['open']-regression_data['low'])*3)
            and ((regression_data['yearHighChange'] < -15) or (regression_data['yearHighChange'] > -2))
            and regression_data['forecast_day_PCT5_change'] < 10
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyEveningStar-0')
            return True
        elif(4 > regression_data['PCT_day_change'] > 2 and 4 > regression_data['PCT_change'] > 2
            and -15 < regression_data['yearHighChange'] 
            and regression_data['PCT_day_change_pre3'] > -1
            and regression_data['PCT_day_change_pre1'] > -1
            and regression_data['PCT_change_pre1'] > -1
            and regression_data['PCT_change_pre2'] > -1
            ):
            add_in_csv(regression_data, regressionResult, ws, 'sellEveningStar-0')
            return True
        elif(5 > regression_data['PCT_day_change'] > 2 and 5 > regression_data['PCT_change'] > 2
            and (regression_data['high']-regression_data['close']) > ((regression_data['close']-regression_data['open']))
            and (regression_data['high']-regression_data['close']) >= ((regression_data['open']-regression_data['low'])*3)
            and -15 < regression_data['yearHighChange']
            ):
            if(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:EveningStarBuy-1(Check2-3MidCapCross)')
                return True
            return False
    elif(
        regression_data['PCT_day_change_pre1'] > 0
        and (ten_days_more_than_seven(regression_data))
        ):
        if(low_tail_pct(regression_data) < 1.5
            and high_tail_pct(regression_data) > 1.5
            ):
            if(1.5 > regression_data['PCT_day_change'] > 0.5 and 1.5 > regression_data['PCT_change'] > 0
                and (regression_data['high']-regression_data['close']) >= ((regression_data['close']-regression_data['open']) * 3)
                and (regression_data['high']-regression_data['close']) >= ((regression_data['open']-regression_data['low'])*3)
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellEveningStar-2(Check2-3MidCapCross)')
                    return True
                return False
        elif((low_tail_pct(regression_data) < 0.5 or (regression_data['forecast_day_PCT_change'] > 6 and low_tail_pct(regression_data) < 1))
            and 2 < high_tail_pct(regression_data) < 3.5
            and low_tail_pct(regression_data) < 0.5
            ):
            if(2 > regression_data['PCT_day_change'] > 0.5 and 2 > regression_data['PCT_change'] > 0
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellEveningStar-Risky-3(Check2-3MidCapCross)')
                    return True
                return False
    return False

def buy_day_low(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['low'] < regression_data['low_pre1']
       and regression_data['bar_low'] < regression_data['bar_low_pre1']
       ):
        if((regression_data['PCT_day_change'] < -5 or regression_data['PCT_change'] < -5)
           and float(regression_data['forecast_day_VOL_change']) < -30
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['forecast_day_PCT_change'] < -5
           and ten_days_less_than_minus_ten(regression_data)
           ):
            if(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:dayLowBuy')
                return True
            return False
        elif((regression_data['PCT_day_change'] < -5 or regression_data['PCT_change'] < -5)
           and float(regression_data['forecast_day_VOL_change']) < -30
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['forecast_day_PCT_change'] < -4.5
           and regression_data['forecast_day_PCT10_change'] > 5
           ):
            add_in_csv(regression_data, regressionResult, ws, 'buyDayLowVolLow-0')
            return True
        elif((regression_data['PCT_day_change'] < -5 or regression_data['PCT_change'] < -5)
           and float(regression_data['forecast_day_VOL_change']) < -20
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['forecast_day_PCT_change'] < -5
           #and regression_data['forecast_day_PCT10_change'] < -5
           ):
            add_in_csv(regression_data, regressionResult, ws, 'sellDayLowVolLow-0')
            return True
        elif((regression_data['PCT_day_change'] < -5 or regression_data['PCT_change'] < -5)
           and float(regression_data['forecast_day_VOL_change']) < 0
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['forecast_day_PCT10_change'] < -5
           ):
            if(regression_data['month6LowChange'] < -5
                and regression_data['forecast_day_PCT10_change'] < -10
                ):
                add_in_csv(regression_data, regressionResult, ws, 'buyDayLowVolLow-01')
                return True
            elif(regression_data['month6LowChange'] > 10
                and regression_data['yearLowChange'] > 20
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellDayLowVolLow-01')
                return True
        elif((regression_data['PCT_day_change'] < -5 and regression_data['PCT_change'] < -4)
           and float(regression_data['forecast_day_VOL_change']) < -30
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and -20 < regression_data['SMA50'] < 10
           and regression_data['SMA9'] > -7
           #and regression_data['PCT_day_change_pre2'] < -1.5
           ):
            add_in_csv(regression_data, regressionResult, ws, 'buyDayLowVolLow-02-checkMorningTrend(.5SL)')
            return True
        elif((regression_data['PCT_day_change'] < -5 and regression_data['PCT_change'] < -4)
           and float(regression_data['forecast_day_VOL_change']) < 0
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and -5 < regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['PCT_day_change_pre2'] > 1.5
           and regression_data['year2LowChange'] > 10
           ):
            add_in_csv(regression_data, regressionResult, ws, 'sellDayLowVolLow-02-checkMorningTrend(.5SL)')
            return True
        elif((regression_data['PCT_day_change'] < -4 and regression_data['PCT_change'] < -2)
           and float(regression_data['forecast_day_VOL_change']) < 0
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['PCT_day_change_pre2'] < -1.5
           and -20 < regression_data['SMA50']
           and regression_data['SMA9'] > -7
           ):
            add_in_csv(regression_data, regressionResult, ws, 'sellDayLowVolLow-03-checkMorningTrend(.5SL)')
            return True
        elif((regression_data['PCT_day_change'] < -4 and regression_data['PCT_change'] < -2)
           and float(regression_data['forecast_day_VOL_change']) < 0
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['PCT_day_change_pre2'] < -1.5
           ):
            if(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyDayLowVolLow-03-checkMorningTrend(.5SL)-after10:30')
                return True
            return False
        elif((regression_data['PCT_day_change'] < -4 and regression_data['PCT_change'] < -2)
           and float(regression_data['forecast_day_VOL_change']) < 0
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['PCT_day_change_pre2'] < -1.5
           ):
            if(is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:sellDayLowVolLow-03-checkMorningTrend(.5SL)')
                return True
            return False
        elif((regression_data['PCT_day_change'] < -4 and regression_data['PCT_change'] < -2)
           and float(regression_data['forecast_day_VOL_change']) < 0
           and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
           and regression_data['PCT_day_change_pre1'] < -1.5
           and regression_data['PCT_day_change_pre2'] > 1.5
           ):
            add_in_csv(regression_data, regressionResult, ws, None)
            return True
#     if((regression_data['PCT_day_change'] < -2 and regression_data['PCT_change'] < -2)
#        and float(regression_data['forecast_day_VOL_change']) < -30
#        and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
#        and regression_data['forecast_day_PCT7_change'] > 7
#        and regression_data['forecast_day_PCT10_change'] > 10
#        ):
#         add_in_csv(regression_data, regressionResult, ws, '##buyDayLowVolLow-2')
#         return True
    return False

def buy_trend_reversal(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if((ten_days_less_than_minus_ten(regression_data)
        or (last_5_day_all_down_except_today(regression_data)
            and ten_days_less_than_minus_seven(regression_data)
            )
        )
        and regression_data['yearLowChange'] > 15 and regression_data['yearHighChange'] < -15
        ):
        if(regression_data['forecast_day_PCT_change'] > 0 and regression_data['PCT_day_change'] > 0
            and (regression_data['PCT_day_change_pre1'] > 1)
            and regression_data['PCT_day_change'] > regression_data['PCT_day_change_pre1'] > 0
            and regression_data['forecast_day_VOL_change'] <= 20
            ):
            if(1 < regression_data['PCT_day_change'] < 5 and 1 < regression_data['PCT_change'] < 5
                and regression_data['forecast_day_VOL_change'] <= -30
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalContinue-00')
                return True
            elif(3 < regression_data['PCT_day_change'] < 5 and 3 < regression_data['PCT_change'] < 5
                ):
                if(is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalContinue-00')
                    return True
                return False
            elif(regression_data['forecast_day_PCT_change'] > 0
                and regression_data['forecast_day_VOL_change'] <= 0
                ):
                add_in_csv(regression_data, regressionResult, ws, 'buyFinalContinue-00')
                return True    
        return False
    
#     if((regression_data['PCT_day_change'] < 2 and low_tail_pct(regression_data) > 4)
#         #and regression_data['PCT_change'] <= 0
#         and -75 < regression_data['year2HighChange'] < -25
#         and low_tail_pct(regression_data) < 1
#         and ('MayBuyCheckChart' in regression_data['filter1']) 
#         and ('Reversal' not in regression_data['filter3'])
#         ):
#         add_in_csv(regression_data, regressionResult, ws, "Test:MayBuyCheckChart-HighTail(Check2-3MidCapCross)")
#         return True
#     
#     if((-0.5 < regression_data['PCT_day_change'] < 0.5)
#         and (-0.5 < regression_data['PCT_change'] < 0.5)
#         and high_tail_pct(regression_data) < 0.3
#         and (0.6 < low_tail_pct(regression_data) < 0.9)
#         and regression_data['PCT_day_change_pre1'] > 0
#         ):
#         add_in_csv(regression_data, regressionResult, ws, "Test:buyHighTailLessThan0.3-3(checkHillUp)")
        
    return False            

def buy_trend_break(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
#     if(regression_data['SMA200'] == 1
#        and regression_data['PCT_day_change'] > 2 and regression_data['PCT_change'] > 2
#        ):
#         add_in_csv(regression_data, regressionResult, ws, '##TestBreakOutBuyConsolidate-0')
#         return True
    flag = False
    if(NIFTY_LOW and low_tail_pct(regression_data) > 1.5
       and -1 < regression_data['PCT_change'] < 1 and -1 < regression_data['PCT_day_change'] < 1
       and regression_data['forecast_day_PCT10_change'] < 10
       ):
        add_in_csv(regression_data, regressionResult, ws, None, '##TEST:buyCheckLastHourUp')
        flag = True
        
#         if(2 < regression_data['month3LowChange'] < 10):
#             if(ten_days_less_than_minus_ten(regression_data)
#                and regression_data['forecast_day_PCT10_change'] < -15
#                and 5 > regression_data['PCT_day_change'] > 1 and 5 > regression_data['PCT_day_change'] > 1
#                #and regression_data['forecast_day_VOL_change'] > 50
#                and abs_yearHigh_more_than_yearLow(regression_data) == True
#                ):
#                 add_in_csv(regression_data, regressionResult, ws, '##BreakOutBuyCandidate(notUpLastDay)-1')
#                 return True
        
#     if(0.5 < regression_data['forecast_day_PCT10_change'] < 2
#         and regression_data['forecast_day_PCT_change'] > 1
#         and 2 < regression_data['PCT_day_change'] < 5.5
#         and 2 < regression_data['PCT_change'] < 5.5
#         and regression_data['PCT_day_change_pre1'] > -0.5
#         and 0.5 < high_tail_pct(regression_data) < 2
#         and regression_data['buyIndia_avg'] > -0.75
#         and regression_data['sellIndia_avg'] > -0.75
#         and regression_data['month3LowChange'] > 4
#         and ('P@[' not in regression_data['sellIndia'])
#         and abs_month3High_more_than_month3Low(regression_data)
#         and (abs(regression_data['month6HighChange']) > abs(regression_data['month6LowChange']))
#         ):
#         add_in_csv(regression_data, regressionResult, ws, 'buy-2week-breakoutup')
#     elif(0.5 < regression_data['forecast_day_PCT10_change'] < 3
#         #and regression_data['forecast_day_PCT_change'] > 1
#         and 2 < regression_data['PCT_day_change']
#         and 2 < regression_data['PCT_change']
#         and regression_data['PCT_day_change_pre1'] > -0.5
#         and 0.5 < high_tail_pct(regression_data) < 2
#         #and regression_data['buyIndia_avg'] > -0.75
#         #and regression_data['sellIndia_avg'] > -0.75
#         #and regression_data['month3LowChange'] > 4
#         #and ('P@[' not in regression_data['sellIndia'])
#         and abs_month3High_more_than_month3Low(regression_data)
#         and (abs(regression_data['month6HighChange']) > abs(regression_data['month6LowChange']))
#         ):
#         add_in_csv(regression_data, regressionResult, ws, 'sell-2week-breakoutup')
    if(0.5 < regression_data['forecast_day_PCT10_change'] < 3
        and regression_data['forecast_day_PCT_change'] > 1
        and 2 < regression_data['PCT_day_change'] < 5.5
        and 2 < regression_data['PCT_change'] < 5.5
        and high_tail_pct(regression_data) < 2.5
        and regression_data['buyIndia_avg'] > -0.75
        and regression_data['sellIndia_avg'] > -0.75
        and regression_data['month3LowChange'] > 4
        and ('P@[' not in regression_data['sellIndia'])
        and abs_month3High_more_than_month3Low(regression_data)
        ):
        if(is_algo_buy(regression_data)):
            add_in_csv(regression_data, regressionResult, ws, '(check-chart):ML:buy-2week-breakoutup')
            return True
        elif(is_algo_sell(regression_data)):
            add_in_csv(regression_data, regressionResult, ws, '(check-chart):ML:sell-2week-breakoutup')
            return True
    elif(0.5 < regression_data['forecast_day_PCT10_change'] < 3
        and regression_data['forecast_day_PCT_change'] > 1
        and 2 < regression_data['PCT_day_change'] < 5.5
        and 2 < regression_data['PCT_change'] < 5.5
        and high_tail_pct(regression_data) < 1
        #and regression_data['buyIndia_avg'] > -0.75
        #and regression_data['sellIndia_avg'] > -0.75
        and regression_data['month3LowChange'] > 4
        and regression_data['month3HighChange'] < -4
        ):
        add_in_csv(regression_data, regressionResult, ws, '(check-chart)buy-2week-breakoutup(|-|-| or --|)')
        return True
            
    
          
    #     if(regression_data['yearLowChange'] < 5):
#        if(regression_data['forecast_day_PCT_change'] > 3 and regression_data['PCT_day_change'] > 3
#            and abs(regression_data['PCT_day_change']) > abs(regression_data['PCT_day_change_pre1'])
#            #and regression_data['open'] == regression_data['low']
#            and regression_data['forecast_day_VOL_change'] >= -20
#            and high_tail_pct(regression_data) < 0.5
#            ):
#                add_in_csv(regression_data, regressionResult, ws, '##finalBreakOutBuy-1test-atYearLow')
#                flag = False
#     if(5 < regression_data['yearLowChange'] < 10 and abs_yearHigh_more_than_yearLow(regression_data)
#        and regression_data['forecast_day_PCT10_change'] < 10
#        and last_7_day_all_up(regression_data) == False
#        and high_tail_pct(regression_data) < 0.7
#        ):
#        if(3 > regression_data['forecast_day_PCT_change'] > 2 and 3 > regression_data['PCT_day_change'] > 2
#            and regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change_pre2'] > 0 and regression_data['PCT_day_change_pre3'] > 0
#            and (regression_data['open'] == regression_data['low'] or regression_data['forecast_day_VOL_change'] >= 0)
#            and regression_data['forecast_day_VOL_change'] >= -50
#            ):
#                add_in_csv(regression_data, regressionResult, ws, '##finalBreakOutBuyContinue-00-test-atYearLow')
#                flag = False
#        if(2 > regression_data['forecast_day_PCT_change'] > 0 and 2 > regression_data['PCT_day_change'] > 0
#            and regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change_pre2'] > 0 and regression_data['PCT_day_change_pre3'] > 0
#            and (regression_data['forecast_day_PCT_change'] > 0.75 or regression_data['PCT_day_change_pre1'] > 0.75 or regression_data['PCT_day_change_pre2'] > 0.75 or regression_data['PCT_day_change_pre3'] > 0.75)
#            and (regression_data['open'] == regression_data['low'] or regression_data['forecast_day_VOL_change'] >= 0)
#            and regression_data['forecast_day_VOL_change'] >= -50
#            ):
#                add_in_csv(regression_data, regressionResult, ws, '##finalBreakOutBuyContinue-0-test-atYearLow')
#                flag = False    
#     if(5 < regression_data['yearLowChange'] < 12 and abs_yearHigh_more_than_yearLow(regression_data)
#        and regression_data['forecast_day_PCT10_change'] < 10
#        and last_7_day_all_up(regression_data) == False
#        and high_tail_pct(regression_data) < 0.7
#        ):
#        if(3 > regression_data['forecast_day_PCT_change'] > 2 and 3 > regression_data['PCT_day_change'] > 2
#            and regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change_pre2'] > 0 and regression_data['PCT_day_change_pre3'] > 0
#            #and regression_data['open'] == regression_data['low']
#            and regression_data['forecast_day_VOL_change'] >= 0
#            ):
#                add_in_csv(regression_data, regressionResult, ws, '##finalBreakOutBuyContinue-11-test-atYearLow')
#                flag = False
#        if(2 > regression_data['forecast_day_PCT_change'] > 0 and 2 > regression_data['PCT_day_change'] > 0
#            and regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change_pre2'] > 0 and regression_data['PCT_day_change_pre3'] > 0
#            and (regression_data['forecast_day_PCT_change'] > 0.75 or regression_data['PCT_day_change_pre1'] > 0.75 or regression_data['PCT_day_change_pre2'] > 0.75 or regression_data['PCT_day_change_pre3'] > 0.75)
#            #and regression_data['open'] == regression_data['low']
#            and regression_data['forecast_day_VOL_change'] >= 0
#            ):
#                add_in_csv(regression_data, regressionResult, ws, '##finalBreakOutBuyContinue-1-test-atYearLow')
#                flag = False
    return flag   

def buy_consolidation_breakout(regression_data, regressionResult, reg, ws):
    week2BarHighChange = ((regression_data['bar_high'] - regression_data['week2BarHigh'])/regression_data['bar_high'])*100
    if(0 < regression_data['PCT_day_change'] < 6
        and regression_data['PCT_change'] < 6
        and (regression_data['PCT_day_change_pre1'] < 0 
             or regression_data['PCT_day_change_pre2'] < 0
             or regression_data['PCT_day_change_pre3'] < 0
             )
        #and regression_data['high'] > regression_data['high_pre3']
        and regression_data['high'] > regression_data['high_pre2']
        and regression_data['high'] > regression_data['high_pre1']
        and 0.5 < regression_data['forecast_day_PCT4_change'] < 3
        and 0 < regression_data['forecast_day_PCT3_change'] < 4
        and 0 < regression_data['forecast_day_PCT2_change'] < 4
        and 0 < regression_data['forecast_day_PCT_change'] < 3
        and 0 <= regression_data['forecast_day_PCT10_change'] <= 3
        and regression_data['bar_high'] >= regression_data['week2BarHigh']
        ):
        if(regression_data['weekBarHigh'] > regression_data['bar_high_pre1'] 
            and regression_data['week2BarHigh'] > regression_data['bar_high_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'brokenToday')
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, 'weekHighGTweek2High')
        if(week2BarHighChange > 5):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarHighChangeGT5')
        elif(week2BarHighChange > 4):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarHighChangeGT4')
        elif(week2BarHighChange > 3):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarHighChangeGT3')
        elif(week2BarHighChange > 2):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarHighChangeGT2')
        elif(week2BarHighChange > 1):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarHighChangeGT1')
        elif(week2BarHighChange > 0):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarHighChangeGT0')
        if(regression_data['PCT_day_change'] > 1
            and regression_data['year2HighChange'] < -5
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakUp-2week')
        elif(regression_data['PCT_day_change'] < 1
            and regression_data['year2HighChange'] < -5
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakUp-2week-(Risky)-PCTDayChangeLT1')
        elif(regression_data['year2HighChange'] >= -5):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakUp-2week-(Risky)-year2High')
            
    elif(1.5 < regression_data['PCT_day_change'] < 6
        and regression_data['PCT_change'] < 6
        and (regression_data['PCT_day_change_pre1'] < 0 or regression_data['PCT_day_change_pre2'] < 0)
        #and regression_data['high'] > regression_data['high_pre3']
        and regression_data['high'] > regression_data['high_pre2']
        and regression_data['high'] > regression_data['high_pre1']
        and 0.5 < regression_data['forecast_day_PCT4_change'] < 3
        and 0 < regression_data['forecast_day_PCT3_change'] < 4
        and 0 < regression_data['forecast_day_PCT2_change'] < 4
        and 0 < regression_data['forecast_day_PCT_change'] < 3
        and ((abs(regression_data['monthHighChange']) < abs(regression_data['monthLowChange']))
            )
        and regression_data['SMA4'] > regression_data['SMA4_2daysBack']
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakUp-month3HighLTMonth3Low')
    elif(1.5 < regression_data['PCT_day_change'] < 6
        and regression_data['PCT_change'] < 6
        and (regression_data['PCT_day_change_pre1'] < 0 or regression_data['PCT_day_change_pre2'] < 0)
        #and regression_data['high'] > regression_data['high_pre3']
        and regression_data['high'] > regression_data['high_pre2']
        and regression_data['high'] > regression_data['high_pre1']
        and 0.5 < regression_data['forecast_day_PCT4_change'] < 3
        and 0 < regression_data['forecast_day_PCT3_change'] < 4
        and 0 < regression_data['forecast_day_PCT2_change'] < 4
        and 0 < regression_data['forecast_day_PCT_change'] < 3
        and ((0 < regression_data['forecast_day_PCT5_change'] < 4
                and 0 < regression_data['forecast_day_PCT7_change'] < 4
               )
            )
        and regression_data['SMA4'] > regression_data['SMA4_2daysBack']
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakUp-forecastDayPCT7changeGT0')
     
def buy_final_candidate(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['forecast_day_PCT4_change'] <= 0.5
        and regression_data['forecast_day_PCT5_change'] <= 0.5
        and regression_data['forecast_day_PCT7_change'] <= 0.5
        and (ten_days_less_than_minus_ten(regression_data)
             or (last_5_day_all_down_except_today(regression_data)
                 and ten_days_less_than_minus_seven(regression_data)
                 )
             )
        and regression_data['yearLowChange'] > 15 and regression_data['yearHighChange'] < -15
        and regression_data['month3LowChange'] > 3
        ):
        if(regression_data['forecast_day_PCT_change'] > 0
           and regression_data['bar_high'] > regression_data['bar_high_pre1']
           #and regression_data['forecast_day_VOL_change'] > 0
           ):
            if(2 < regression_data['PCT_day_change'] < 4 and 2 < regression_data['PCT_change'] < 4
                and regression_data['PCT_day_change_pre1'] < 0
                and regression_data['forecast_day_PCT10_change'] > -20
                and regression_data['month3HighChange'] < -15
                and regression_data['SMA4'] > 0
                and regression_data['SMA9'] > -3
                and high_tail_pct(regression_data) < 1.5
                ):
                add_in_csv(regression_data, regressionResult, ws, 'buyFinalCandidate-(lastDayDown)')
                return True
            elif(2 < regression_data['PCT_day_change'] < 4 and 2 < regression_data['PCT_change'] < 4
                and no_doji_or_spinning_buy_india(regression_data)
                and regression_data['PCT_day_change_pre1'] > 0
                and regression_data['forecast_day_PCT10_change'] > -20
                and regression_data['month3HighChange'] < -15
                and regression_data['forecast_day_VOL_change'] > -20
                and high_tail_pct(regression_data) < 1.5
                ):
                if(is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalCandidate-(lastDayUp)')
                    return True
                return False
            elif(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4
                and no_doji_or_spinning_buy_india(regression_data)
                and regression_data['PCT_day_change_pre1'] > 0
                and regression_data['forecast_day_VOL_change'] < -20
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalContinue-(lastDayUp)-(volLow)')
                return True
            elif(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4
                and regression_data['PCT_day_change_pre1'] < 0
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalContinue-(lastDayDown)-1')
                    return True
                return False
            elif(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4
                and regression_data['PCT_day_change_pre1'] > 0
                ):
                if('P@[' in regression_data['sellIndia']):
                    add_in_csv(regression_data, regressionResult, ws, 'buyFinalCandidate-(lastDayUp)-(sellPattern)')
                    return True
                elif('P@[' not in regression_data['buyIndia']):
                    add_in_csv(regression_data, regressionResult, ws, 'sellFinalContinue-(lastDayUp)-1')
                    return True
                return False
            elif(0.5 < regression_data['PCT_day_change'] < 2.5 and 0.5 < regression_data['PCT_change'] < 2.5
                and regression_data['PCT_day_change_pre1'] < 0
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalContinue-(lastDayDown)-2')
                    return True
                return False
            elif(0.5 < regression_data['PCT_day_change'] < 2.5 and 0.5 < regression_data['PCT_change'] < 2.5
                and regression_data['PCT_day_change_pre1'] > 0
                ):
                if('P@[' in regression_data['sellIndia']):
                    add_in_csv(regression_data, regressionResult, ws, 'buyFinalCandidate-(lastDayUp)-(sellPattern)')
                    return True
                elif(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalContinue-(lastDayUp)-2')
                    return True
                return False
        if((((regression_data['close'] - regression_data['open']) * 1.5 > regression_data['high'] - regression_data['low']) or (regression_data['forecast_day_PCT_change'] > 0 and regression_data['PCT_day_change'] > 1))
            and (regression_data['yearHighChange'] < -30 or regression_data['yearLowChange'] < 30)
            ):
            if(1 < regression_data['PCT_day_change'] < 2.5 and 1 < regression_data['PCT_change'] < 2.5 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] > -20
                    ):
                    if(is_algo_buy(regression_data)):
                        add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalCandidate-2')
                        return True
                    return False
                elif(regression_data['forecast_day_VOL_change'] < -30
                    ):
                    if(is_algo_sell(regression_data)):
                        add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalContinue-2')
                        return True
                    return False
            if(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] < -30
                    ):
                    if(is_algo_buy(regression_data)):
                        add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalCandidate-3')
                        return True
                    return False
                elif(regression_data['forecast_day_VOL_change'] > 0
                    ):
                    if(is_algo_sell(regression_data)):
                        add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalContinue-3')
                        return True
                    return False
    return False
           
def buy_oi(regression_data, regressionResult, reg, ws):
    return False
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(((regression_data['PCT_day_change'] > 1 or regression_data['PCT_change'] > 1)
            or
            (regression_data['PCT_day_change'] > 0.75 and regression_data['PCT_change'] > 0.75
            and regression_data['forecast_day_PCT_change'] < 2
            and regression_data['forecast_day_PCT2_change'] < 2
            and regression_data['forecast_day_PCT3_change'] < 2
            and regression_data['forecast_day_PCT4_change'] < 2
            and regression_data['forecast_day_PCT4_change'] < 2
            and regression_data['forecast_day_PCT5_change'] < 2
            and regression_data['forecast_day_PCT7_change'] < 2
            and regression_data['forecast_day_PCT10_change'] < 2
            )
        )
        and regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and 5 > regression_data['forecast_day_PCT4_change'] > 0
        and regression_data['forecast_day_PCT5_change'] < 5
        and regression_data['forecast_day_PCT7_change'] < 5
        and regression_data['forecast_day_PCT10_change'] < 5
        and (regression_data['PCT_day_change'] < 0
            or regression_data['PCT_day_change_pre1'] < 0
            or regression_data['PCT_day_change_pre2'] < 0
            or regression_data['PCT_day_change_pre3'] < 0
            or regression_data['PCT_day_change_pre4'] < 0
           )
        and (regression_data['forecast_day_VOL_change'] > 150
            or (regression_data['PCT_day_change_pre2'] < 0
                and (((regression_data['volume'] - regression_data['volume_pre2'])*100)/regression_data['volume_pre2']) > 100
                and (((regression_data['volume'] - regression_data['volume_pre3'])*100)/regression_data['volume_pre3']) > 100
               )
           )
        and float(regression_data['contract']) > 100
        and(regression_data['PCT_day_change_pre1'] > 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
            )
        and regression_data['yearHighChange'] < -5
        and regression_data['open'] > 50
        and (last_4_day_all_up(regression_data) == False)
        and (high_tail_pct(regression_data) < 1)
        and regression_data['month3HighChange'] < -7.5
        ):
        if(1 < regression_data['PCT_day_change'] < 2.5 and 1 < regression_data['PCT_change'] < 2.5       
        ):
            if(regression_data['forecast_day_PCT10_change'] < 0 or regression_data['forecast_day_PCT7_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-0')
                return True
            elif(regression_data['forecast_day_PCT10_change'] < 10 or (regression_data['forecast_day_PCT5_change'] < 5 and regression_data['forecast_day_PCT7_change'] < 5)):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-1')
                return True
            else:
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-2-Risky')
                return True
        if(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4
            and float(regression_data['forecast_day_VOL_change']) > 300 
            ):
            if(regression_data['forecast_day_PCT10_change'] < 0 or regression_data['forecast_day_PCT7_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-0-checkConsolidation')
                return True
            elif(regression_data['forecast_day_PCT10_change'] < 10 or (regression_data['forecast_day_PCT5_change'] < 5 and regression_data['forecast_day_PCT7_change'] < 5)):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-1-checkConsolidation')
                return True
            else:
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-2-checkConsolidation-Risky')
                return True    
    return False

def buy_up_trend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if((regression_data['yearLowChange'] > 20 and regression_data['month3LowChange'] > 15)
       and (regression_data['yearHighChange'] < -15 or regression_data['month3HighChange'] < -10)
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['forecast_day_PCT2_change'] > 0
       and regression_data['forecast_day_PCT3_change'] > 0
       and (regression_data['PCT_day_change_pre1'] < 0 or regression_data['PCT_day_change_pre2'] < 0)
       ):
        if(regression_data['yearHighChange'] < -10
           and 2 < regression_data['PCT_day_change'] < 3.5 and 2 < regression_data['forecast_day_PCT_change'] < 3
           and regression_data['series_trend'] != 'downTrend'
           and str(regression_data['score']) != '0-1' 
           and regression_data['forecast_day_PCT7_change'] < 0 and -6 < regression_data['forecast_day_PCT10_change'] < 0
           and high_tail_pct(regression_data) > .5
           and regression_data['SMA25'] < 4
           ):
            add_in_csv(regression_data, regressionResult, None, 'sellUpTrend-0')
            return True
        elif(regression_data['yearHighChange'] < -10
           and 2 < regression_data['PCT_day_change'] < 3.5 and 2 < regression_data['forecast_day_PCT_change'] < 3
           and regression_data['series_trend'] != 'downTrend'
           and str(regression_data['score']) != '0-1' 
           and regression_data['forecast_day_PCT7_change'] < 0 and -6 < regression_data['forecast_day_PCT10_change'] < 0
           and 1 > low_tail_pct(regression_data) > .6
           and high_tail_pct(regression_data) < .5
           ):
            add_in_csv(regression_data, regressionResult, None, 'buyUpTrend-0')
            return True
        elif(
           regression_data['yearHighChange'] < -10
           and 2 < regression_data['PCT_day_change'] < 3.5 and regression_data['PCT_change'] < 3
           and regression_data['PCT_day_change_pre1'] < 0
           and regression_data['series_trend'] != 'downTrend'
           and str(regression_data['score']) != '0-1'
           and 3.5 > regression_data['forecast_day_PCT_change'] > 1
           and regression_data['forecast_day_PCT7_change'] > 4
           and regression_data['forecast_day_PCT10_change'] < 7
           and ('ENGULFING' not in regression_data['buyIndia'])
           and (regression_data['forecast_day_PCT4_change'] < 3 or ('[' in regression_data['buyIndia']))
           ):
            add_in_csv(regression_data, regressionResult, None, 'sellUpTrend-1')
            return True
        elif(
           regression_data['yearHighChange'] < -10
           and 2 < regression_data['PCT_day_change'] < 3.5 and regression_data['PCT_change'] < 3
           and regression_data['PCT_day_change_pre1'] < 0
           and regression_data['series_trend'] != 'downTrend'
           and str(regression_data['score']) != '0-1'
           and regression_data['forecast_day_PCT5_change'] < 3
           and regression_data['forecast_day_PCT7_change'] > 1.5
           and regression_data['forecast_day_PCT10_change'] < 7
           and regression_data['weekHighChange'] > -10
           and (('[' in regression_data['buyIndia'] and 'BELTHOLD' not in regression_data['buyIndia'] and 'LONGLINE' not in regression_data['buyIndia'])
                or ('[' not in regression_data['buyIndia'] and regression_data['SMA4_2daysBack'] < 2)
                )
           ):
            add_in_csv(regression_data, regressionResult, None, 'sellUpTrend-2')
            return True
        elif(
           abs_yearHigh_less_than_yearLow(regression_data)
           and regression_data['yearHighChange'] < -10
           and 2 < regression_data['PCT_day_change'] < 3.5 and 1.5 < regression_data['PCT_change'] < 3
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_change']
           and regression_data['PCT_day_change_pre1'] < 0
           and regression_data['series_trend'] != 'downTrend'
           and str(regression_data['score']) != '0-1'
           and regression_data['forecast_day_PCT5_change'] < 3
           and regression_data['forecast_day_PCT7_change'] < 1.5
           and regression_data['forecast_day_PCT10_change'] < 7
           and regression_data['month3HighChange'] < -2.5
           and '[' not in regression_data['sellIndia']
           ):
            add_in_csv(regression_data, regressionResult, None, 'buyUpTrend-1')
            return True
    elif((regression_data['yearHighChange'] < -9 or regression_data['yearLowChange'] < 15)
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['forecast_day_PCT2_change'] > 0
       and regression_data['forecast_day_PCT3_change'] > 0    
       ):
        if(2 < regression_data['PCT_day_change'] < 4 and regression_data['PCT_change'] < 4
            and regression_data['forecast_day_PCT_change'] < regression_data['PCT_change']
            and regression_data['series_trend'] != 'downTrend'
            and str(regression_data['score']) != '0-1'
            and (regression_data['forecast_day_PCT7_change'] < 5 or regression_data['forecast_day_PCT10_change'] < 7)
            and 'P@[' in regression_data['sellIndia']
            ):
             if(is_algo_buy(regression_data)):
                 add_in_csv(regression_data, regressionResult, None, 'buyUpTrend-1-YearHigh')
                 return True
             return False
    return False

def buy_market_uptrend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    return False
#     if(('[' not in regression_data['sellIndia'])
#         and ((1.5 < regression_data['PCT_change'] < 6) 
#             and (1.5 < regression_data['PCT_day_change'] < 6)
#             and regression_data['close'] > regression_data['bar_high_pre1']
#             )
#         and regression_data['SMA4_2daysBack'] > -1
#         and regression_data['SMA4'] > -1
#         and regression_data['SMA9'] > -10
#         #and regression_data['trend'] == 'up'
#         ):
#         if(('ReversalLowYear2' in regression_data['filter3']) and (regression_data['year2HighChange'] < -40)):
#             add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyYear2LowReversal(Confirm)')
#         if(('ReversalLowYear' in regression_data['filter3']) and (regression_data['yearHighChange'] < -30)):
#             add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyYearLowReversal(Confirm)')
#         if(('ReversalLowMonth6' in regression_data['filter3']) and (regression_data['month6HighChange'] < -25)):
#             add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyMonth6LowReversal(Confirm)')
#         if(('ReversalLowMonth3' in regression_data['filter3']) and (regression_data['month3HighChange'] < -15)):
#             add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyMonth3LowReversal(Confirm)')
#         if(regression_data['month3HighChange'] < -20 and (regression_data['month6HighChange'] < -20 or regression_data['yearHighChange'] < -30)
#             ):
#             if(('NearLowYear2' in regression_data['filter3']) and (regression_data['year2HighChange'] < -40)):
#                 add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyYear2Low')
#             if(('NearLowYear' in regression_data['filter3']) and (regression_data['yearHighChange'] < -30)):
#                 add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyYearLow')
#             if(('NearLowMonth6' in regression_data['filter3']) and (regression_data['month6HighChange'] < -25)):
#                 add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyMonth6Low')
#             if(('NearLowMonth3' in regression_data['filter3']) and (regression_data['month3HighChange'] < -20)):
#                 add_in_csv(regression_data, regressionResult, ws, '##UPTREND:buyMonth3Low')

def buy_check_chart(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    #Check for last 5 from latest up should crossover
    if(('NearHighMonth3' in regression_data['filter3']) 
        or ('BreakHighMonth3' in regression_data['filter3'])
        or ('ReversalHighMonth3' in regression_data['filter3'])
        ):
        if(regression_data['month3LowChange'] > 15
            and ((-5 < regression_data['PCT_change'] < -1) and (-5 < regression_data['PCT_day_change'] < -2))
            and regression_data['close'] < regression_data['bar_low_pre1']
            and regression_data['close'] < regression_data['bar_low_pre2']
            ):
            if('NearHighMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearHighMonth3')
            elif('ReversalHighMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalHighMonth3')
            elif('BreakHighMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'BreakHighMonth3')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-buy):month3High-InMinus')
            
    elif(('NearHighMonth6' in regression_data['filter3']) 
        or ('ReversalHighMonth6' in regression_data['filter3'])
        ):
        if(regression_data['month6LowChange'] > 20
            and ((-5 < regression_data['PCT_change'] < -1) and (-5 < regression_data['PCT_day_change'] < -2))
            and regression_data['close'] < regression_data['bar_low_pre1']
            and regression_data['close'] < regression_data['bar_low_pre2']
            ):
            if('NearHighMonth6' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearHighMonth6')
            elif('ReversalHighMonth6' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalHighMonth6')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-buy):month6High-InMinus')
            
    elif(('NearHighYear' in regression_data['filter3']) 
        or ('ReversalHighYear' in regression_data['filter3'])
        ):
        if(regression_data['month6LowChange'] > 20
            and ((-5 < regression_data['PCT_change'] < -1) and (-5 < regression_data['PCT_day_change'] < -2))
            and regression_data['close'] < regression_data['bar_low_pre1']
            and regression_data['close'] < regression_data['bar_low_pre2']
            ):
            if('NearHighYear' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearHighYear')
            elif('ReversalHighYear' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalHighYear')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-buy):YearHigh-InMinus')
        
#         if((('NearHighYear2' in regression_data['filter3']) 
#             or ('ReversalHighYear2' in regression_data['filter3'])
#             or ('BreakHighYear2' in regression_data['filter3'])
#             )):    
#             if(regression_data['year2LowChange'] > 50
#                 and regression_data['year2HighChange'] > -1
#                 and ((regression_data['PCT_change'] > 1) and (regression_data['PCT_day_change'] > 1))
#                 and regression_data['close'] > regression_data['bar_high_pre1']
#                 ):
#                 if(regression_data['PCT_change'] > 4
#                     and ('BreakHighYear2' in regression_data['filter3'])
#                     ):
#                     add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-sell):year2High-InPlus')
#                 else:
#                     add_in_csv(regression_data, regressionResult, ws, '##(Check-chart):year2High-InPlus')
#                     
#         elif((('NearHighYear' in regression_data['filter3'])
#             or ('ReversalHighYear' in regression_data['filter3'])
#             or ('BreakHighYear' in regression_data['filter3'])
#             )):
#             if(regression_data['yearLowChange'] > 40
#                 and regression_data['yearHighChange'] > -1
#                 and ((regression_data['PCT_change'] > 1) and (regression_data['PCT_day_change'] > 1))
#                 and regression_data['close'] > regression_data['bar_high_pre1']
#                 ):
#                 if(regression_data['PCT_change'] > 4
#                     and ('BreakHighYear' in regression_data['filter3'])
#                     ):
#                     add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-sell):yearHigh-InPlus')
#                 else:
#                     add_in_csv(regression_data, regressionResult, ws, '##(Check-chart):yearHigh-InPlus')
    
def buy_month3_high_continue(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if(('NearHighMonth3' in regression_data['filter3']) 
        or ('BreakHighMonth3' in regression_data['filter3'])
        or ('ReversalHighMonth3' in regression_data['filter3'])
        ):
        if((regression_data['forecast_day_PCT4_change'] < 0)
            and (regression_data['forecast_day_PCT5_change'] < 0)
            and (regression_data['forecast_day_PCT7_change'] < 0)
            and ((regression_data['forecast_day_PCT_change'] > 0 and (2 < regression_data['PCT_change'] < 3) and (2 < regression_data['PCT_day_change'] < 3))
                 or ( ("downTrend" in regression_data['series_trend']) and (1.5 < regression_data['PCT_change'] < 3) and (1.5 < regression_data['PCT_day_change'] < 3)))
            and regression_data['high'] > regression_data['high_pre1']
            and regression_data['bar_high'] > regression_data['bar_high_pre1']
            #and regression_data['SMA4_2daysBack'] < 0
            ):
            if(regression_data['SMA4_2daysBack'] > 0.5 and regression_data['SMA9_2daysBack'] > 0.5
                and (regression_data['SMA4_2daysBack'] > 1 or regression_data['SMA9_2daysBack'] > 1)
                and ('[') not in regression_data['sellIndia']
                ):
                add_in_csv(regression_data, regressionResult, ws, '##(Check-chart)buyMonth3High-Continue')
                return True
            elif(regression_data['SMA4_2daysBack'] < -0.5 and regression_data['SMA9_2daysBack'] < -0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, '##(Check-chart)sellMonth3High-Continue')
                return True
            elif(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, '##(Check-chart)ML:buyMonth3High-Continue')
                return True
            elif(is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, '##(Check-chart)ML:sellMonth3High-Continue')
                return True
            return False
    return False

def buy_heavy_uptrend_reversal(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if( ("upTrend" in regression_data['series_trend'])
        and (0 < regression_data['forecast_day_PCT4_change'] < 15)
        and (0 < regression_data['forecast_day_PCT5_change'] < 15)
        and (0 < regression_data['forecast_day_PCT7_change'] < 15)
        and (regression_data['forecast_day_PCT10_change'] < 15)
        ):
#             if((regression_data['month3HighChange'] < -5)
#                 and ((0 < regression_data['PCT_change'] < 5) and (0 < regression_data['PCT_day_change'] < 5))
#                 ):
#                 if((4 < regression_data['PCT_change'] < 5) and (4 < regression_data['PCT_day_change'] < 5)
#                 ):
#                     add_in_csv(regression_data, regressionResult, ws, '##buyHeavyUpTrend-0-Continue')
#                 elif(ten_days_more_than_seven(regression_data)
#                      and (('NearHighMonth3' in regression_data['filter3']) 
#                           or ('NearHighMonth6' in regression_data['filter3'])
#                          )
#                 ):
#                     add_in_csv(regression_data, regressionResult, ws, '##buyHeavyUpTrend-1-Continue')
        if((regression_data['month3HighChange'] > -5)
            and ((0 < regression_data['PCT_change'] < 5) and (0 < regression_data['PCT_day_change'] < 5))
            and regression_data['PCT_day_change_pre1'] > 1
            ):
            if((2 < regression_data['PCT_change'] < 7) and (3 < regression_data['PCT_day_change'] < 7)
                and ten_days_more_than_seven(regression_data)
                and (('NearHighMonth3' in regression_data['filter3']) 
                      or ('BreakHighMonth3' in regression_data['filter3'])
                      or ('NearHighMonth6' in regression_data['filter3'])
                      or ('BreakHighMonth6' in regression_data['filter3'])
                    )
            ):
                if(regression_data['month6HighChange'] > -5
                    and (3 < regression_data['PCT_change'] or 3 < regression_data['PCT_day_change'])
                    and('BELTHOLD' and 'LONGLINE' not in regression_data['buyIndia'])
                    and regression_data['year2HighChange'] > -20
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellHeavyUpTrend-Reversal')
                else:
                    if(regression_data['forecast_day_VOL_change'] < 0):
                        add_in_csv(regression_data, regressionResult, ws, 'buyHeavyUpTrend-Reversal-(Risky)')
                    elif('P@' in regression_data['buyIndia']
                        and regression_data['year2HighChange'] > -20
                        ):
                        add_in_csv(regression_data, regressionResult, ws, 'sellHeavyUpTrend-Reversal-(Risky)')
            elif((2 < regression_data['PCT_change'] < 7) and (3 < regression_data['PCT_day_change'] < 7)
                #and ten_days_more_than_seven(regression_data)
                and (('NearHighMonth3' in regression_data['filter3']) 
                      or ('BreakHighMonth3' in regression_data['filter3'])
                      or ('NearHighMonth6' in regression_data['filter3'])
                      or ('BreakHighMonth6' in regression_data['filter3'])
                    )
                and regression_data['year2HighChange'] < -20
                and regression_data['yearHighChange'] > -50
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)sellHeavyUpTrend-Reversal-1')
                   
def buy_supertrend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    
    if(regression_data['close'] > 50
        ):
        if(0 < regression_data['forecast_day_PCT2_change'] < regression_data['forecast_day_PCT3_change'] < regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change']
            and regression_data['forecast_day_PCT5_change'] > 5
            and regression_data['forecast_day_PCT10_change'] < regression_data['forecast_day_PCT5_change']
            and regression_data['forecast_day_PCT10_change'] < regression_data['forecast_day_PCT7_change']
            and -0.75 < regression_data['PCT_day_change_pre1'] < 2.5
            and -0.75 < regression_data['PCT_day_change'] < 0.5
            and regression_data['yearHighChange'] < -5
            and regression_data['yearLowChange'] > 15
            and regression_data['month3LowChange'] > 10 
            #and regression_data['low'] > regression_data['low_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'set1')
            if(regression_data['PCT_day_change_pre1'] < 0
                and -0.75 < regression_data['PCT_day_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, '--')
            elif(regression_data['PCT_day_change_pre1'] > 0
                and 0 < regression_data['PCT_day_change'] < 0.75):
                add_in_csv(regression_data, regressionResult, ws, '++')
            elif(regression_data['PCT_day_change_pre1'] < 0
                and 0 < regression_data['PCT_day_change'] < 0.75):
                add_in_csv(regression_data, regressionResult, ws, '-+')
            elif(regression_data['PCT_day_change_pre1'] > 0
                and -0.75 < regression_data['PCT_day_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, '+-')
#             if(regression_data['high_month6'] >= regression_data['high']):
#                 add_in_csv(regression_data, regressionResult, ws, 'MoreThanM6High')
#             else:
#                 add_in_csv(regression_data, regressionResult, ws, 'LessThanM6High')
            if((regression_data['PCT_day_change_pre1'] < 0 and regression_data['PCT_day_change'] < 0)
                or (regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change'] > 0)):    
                if((regression_data['forecast_day_PCT_change'] < 0
                        and regression_data['PCT_day_change'] < 0
                        and regression_data['SMA4'] < 1.5
                        and high_tail_pct(regression_data) < 1.5
                        and regression_data['low'] > regression_data['low_pre1']
                    )
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'buySuperTrend-0')
                elif('SPINNINGTOP' in regression_data['sellIndia']
                    and regression_data['low'] > regression_data['low_pre1']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-0')
                elif(regression_data['PCT_day_change_pre2'] > 2
                    and regression_data['PCT_change_pre2'] > 2
                    and regression_data['low'] < regression_data['low_pre1']
                    ):
                    if(is_algo_buy(regression_data)):
                       add_in_csv(regression_data, regressionResult, ws, 'ML:buySuperTrend-2')
                       return True
                    return False
                elif(regression_data['PCT_day_change_pre2'] < 0.5
                    and regression_data['PCT_change_pre2'] < 0.5
                    and regression_data['low'] < regression_data['low_pre1']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-1')
                elif((('P@' or 'M@') in regression_data['sellIndia'])
                     or(regression_data['month3HighChange'] > 1
                        and 'Break' in regression_data['filter3']
                        and regression_data['forecast_day_PCT_change'] > 0.35
                        )
                     or(regression_data['month3HighChange'] < -25
                        and regression_data['month6HighChange'] < -30
                        )
                     or('P@[,SPINNINGTOP]' in regression_data['buyIndia']
                        and regression_data['month3HighChange'] < 0
                        and regression_data['yearHighChange'] < -20
                        )
                     or('[' not in regression_data['buyIndia']
                        and -20 < regression_data['month3HighChange'] < -10
                        and 5 < regression_data['forecast_day_PCT7_change'] < 12
                        )
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-2')
                elif(5 > regression_data['month3HighChange'] > 0.45
                    or (regression_data['PCT_day_change'] < 0
                       and regression_data['month3LowChange'] < 25
                       and regression_data['year2HighChange'] < -10
                       and regression_data['year2HighChange'] < -30
                       )
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'buySuperTrend-1')
        elif(0 < regression_data['forecast_day_PCT2_change'] < regression_data['forecast_day_PCT3_change'] < regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change']
            and regression_data['forecast_day_PCT5_change'] > 5
            and -0.75 < regression_data['PCT_day_change_pre1'] < 2.5
            and -0.75 < regression_data['PCT_day_change'] < 0.5
            and regression_data['yearHighChange'] < -5
            and regression_data['yearLowChange'] > 15
            and regression_data['month3LowChange'] > 10   
            #and regression_data['low'] > regression_data['low_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'set2')
            if(regression_data['PCT_day_change_pre1'] < 0
                and -0.75 < regression_data['PCT_day_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, '--')
            elif(regression_data['PCT_day_change_pre1'] > 0
                and 0 < regression_data['PCT_day_change'] < 0.75):
                add_in_csv(regression_data, regressionResult, ws, '++')
            elif(regression_data['PCT_day_change_pre1'] < 0
                and 0 < regression_data['PCT_day_change'] < 0.75):
                add_in_csv(regression_data, regressionResult, ws, '-+')
            elif(regression_data['PCT_day_change_pre1'] > 0
                and -0.75 < regression_data['PCT_day_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, '+-')
            if((regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change'] > 0)
                 or (regression_data['PCT_day_change_pre1'] > 0 and regression_data['PCT_day_change'] < 0)
                ):
                if('SPINNINGTOP' in regression_data['sellIndia']
                    and regression_data['year2LowChange'] > 50
                    and regression_data['low'] > regression_data['low_pre1']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'buySuperTrend-0')
                elif(regression_data['PCT_day_change_pre2'] > 2
                    and regression_data['PCT_change_pre2'] > 2
                    and regression_data['yearHighChange'] < -20
                    and regression_data['low'] < regression_data['low_pre1']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'buySuperTrend-1')
                elif(((('P@' or 'M@') in regression_data['sellIndia'])
                     or(regression_data['month3HighChange'] > 1
                        and 'Break' in regression_data['filter3']
                        and regression_data['forecast_day_PCT_change'] > 0.35
                        )
                     or(regression_data['month3HighChange'] < -25
                        and regression_data['month6HighChange'] < -30
                        )
                     )
                     and regression_data['PCT_day_change_pre1'] < 1
                     and ('P@' or 'M@') in regression_data['buyIndia']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-0')
                elif((regression_data['PCT_day_change'] < 0
                       and regression_data['PCT_change_pre1'] > 1
                       and regression_data['month3LowChange'] < 25
                       and regression_data['year2HighChange'] < -10
                       and regression_data['year2HighChange'] < -30
                       )
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-1')  
                elif(0 < regression_data['forecast_day_PCT2_change'] < regression_data['forecast_day_PCT3_change'] < regression_data['forecast_day_PCT4_change'] < regression_data['forecast_day_PCT5_change']
                    and regression_data['forecast_day_PCT5_change'] > 5
                    and regression_data['forecast_day_PCT10_change'] > regression_data['forecast_day_PCT5_change']
                    and regression_data['forecast_day_PCT10_change'] > regression_data['forecast_day_PCT7_change']
                    and -0.5 < regression_data['PCT_day_change_pre1'] < 1.5
                    and -0.5 < regression_data['PCT_day_change'] < 0.5
                    and regression_data['yearHighChange'] < -5
                    ):
                    if(regression_data['forecast_day_PCT_change'] < 0
                        and regression_data['month3HighChange'] < 0
                        and ('P@' or 'M@') in regression_data['buyIndia']
                        ):
                        add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-1')
                    elif(regression_data['PCT_day_change'] < 0
                        and regression_data['low'] < regression_data['low_pre1']
                        ):
                        add_in_csv(regression_data, regressionResult, ws, None)
                    elif(regression_data['PCT_day_change'] < 0
                        ):
                        add_in_csv(regression_data, regressionResult, ws, None)
                    elif(('[') not in regression_data['buyIndia']):
                        add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-1')
        return True
    return False

def buy_risingMA(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    
#     if(is_algo_buy(regression_data)
#         and regression_data['SMA4'] > 0
#         and regression_data['SMA9'] > 0
#         and regression_data['SMA25'] > 0
#         and regression_data['SMA100'] > 10
#         and regression_data['SMA200'] > 10
#         and regression_data['year2HighChange'] < -5
#         and regression_data['yearHighChange'] < -5
#         and (regression_data['PCT_day_change_pre1'] < 0.5
#              or -0.5 < regression_data['PCT_day_change'] < 0.5)
#         ):
#         if(regression_data['PCT_change'] < -4.5
#             and regression_data['PCT_day_change'] < -4.5):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(LT-4.5)')
#         elif(-5 < regression_data['PCT_change'] < -2
#             and -5 < regression_data['PCT_day_change'] < -2):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(LT-2)')
#         elif(-5 < regression_data['PCT_change'] < 0
#             and -5 < regression_data['PCT_day_change'] < 0):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(LT0)')
#         elif(-1 < regression_data['PCT_change'] < 1
#             and -1 < regression_data['PCT_day_change'] < 1):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(BT+1-1)')
#         elif(0 < regression_data['PCT_change'] < 2
#             and 0 < regression_data['PCT_day_change'] < 2):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(GT0)')
#         elif(0 < regression_data['PCT_change'] < 5
#             and 0 < regression_data['PCT_day_change'] < 5):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(GT+2)')
#         elif(regression_data['PCT_change'] > 4.5
#             and regression_data['PCT_day_change'] > 4.5):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend_(GT+4.5)')
#         else:
#             add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMAPositive_Uptrend')
    
    if(regression_data['year2HighChange'] < -50
        ):
        if(regression_data['yearHighChange'] < -30
            and 2 < regression_data['PCT_day_change'] < 4
            #and regression_data['forecast_day_PCT_change'] < 1
            and regression_data['forecast_day_PCT_change'] < regression_data['forecast_day_PCT4_change']
            and regression_data['forecast_day_PCT2_change'] < 1
            and regression_data['forecast_day_PCT3_change'] < 1
            and regression_data['forecast_day_PCT4_change'] < 0
            and regression_data['forecast_day_PCT5_change'] < 0
            and regression_data['forecast_day_PCT7_change'] < 0
            and regression_data['forecast_day_PCT10_change'] < 0
            and high_tail_pct(regression_data) < 1
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyYear2LowBreakingUp')
        elif(regression_data['yearHighChange'] < -20
            and -1.5 < regression_data['PCT_day_change'] < -0.3
            and -2 < regression_data['PCT_change'] < 1 
            and regression_data['PCT_day_change_pre1'] < 0
            and regression_data['EMA14_2daysBack'] < regression_data['EMA14_1daysBack'] < regression_data['EMA14']
            and (((regression_data['bar_low'] - regression_data['bar_low_pre1'])/regression_data['bar_low'])*100) > 0
            and regression_data['low'] > regression_data['low_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMayYear2LowBreakingUp')
            
            
    
    if(regression_data['close'] > 50
        ):        
        if(regression_data['SMA25'] > 0
            and 0 < regression_data['SMA50'] < 10
            and regression_data['SMA200'] < regression_data['SMA100'] < 0
            and -70 < regression_data['SMA200'] < -25
            #and regression_data['PCT_day_change'] < -1.5
            and regression_data['PCT_change'] < -1.5
            #and regression_data['series_trend'] != "downTrend"
            #and all_day_pct_change_negative_except_today(regression_data) != True
            and regression_data['month3HighChange'] < -5 
            ):
            if(regression_data['PCT_day_change_pre1'] > 1
                and regression_data['SMA4'] > -2
                and '[' not in regression_data['sellIndia']
                ):
                add_in_csv(regression_data, regressionResult, ws, None)
                return True
            elif(regression_data['PCT_day_change_pre1'] < -1
                and -5 < regression_data['PCT_day_change'] < -1
                and all_day_pct_change_negative_except_today(regression_data) != True
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellRisingMA')
                return True
            elif(regression_data['PCT_day_change_pre1'] > 1
                and regression_data['SMA50'] < 1
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellRisingMA')
                return True
            elif(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:BuyRisingMA-Risky-0')
                return True
            elif(is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:SellRisingMA-Risky-0')
                return True
            return False
        elif(regression_data['SMA4_2daysBack'] > 0
            and regression_data['SMA9'] > 0
            and 2 < regression_data['SMA50'] < 10
            and regression_data['SMA200'] < regression_data['SMA100'] < 0
            and -70 < regression_data['SMA200'] < -10
            #and regression_data['PCT_day_change'] < -1.5
            and regression_data['PCT_change'] < -1.5
            #and regression_data['series_trend'] != "downTrend"
            #and all_day_pct_change_negative_except_today(regression_data) != True
            ):
            if(-4 < regression_data['PCT_day_change'] < -1
                and regression_data['forecast_day_PCT4_change'] < 10
                and ('[' not in regression_data['sellIndia'])
                ):
                add_in_csv(regression_data, regressionResult, ws, 'buyRisingMA-1')
                return True
            elif(('P@' or 'M@') in regression_data['buyIndia']):
                add_in_csv(regression_data, regressionResult, ws, 'sellRisingMA-1')
                return True
            elif(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:BuyRisingMA-Risky-1')
                return True
            elif(is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:SellRisingMA-Risky-1')
                return True
            return False
        elif(regression_data['SMA4_2daysBack'] > 0
            and regression_data['SMA9'] > 1
            and regression_data['SMA25'] > 1
            and 1 < regression_data['SMA50'] < 3
            and regression_data['SMA200'] < regression_data['SMA100'] < -1
            and -70 < regression_data['SMA200'] < -15
            and regression_data['PCT_day_change'] < 0
            and regression_data['PCT_change'] < 0
            and regression_data['series_trend'] != "downTrend"
            and all_day_pct_change_negative_except_today(regression_data) != True
            ):
            if(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:BuyRisingMA-Risky-2')
                return True
            elif(is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:SellRisingMA-Risky-2')
                return True
            return False
        elif(1 < regression_data['SMA9'] < 5
            and 1 < regression_data['SMA25'] < 5
            and regression_data['SMA50'] < 0
            and regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50']
            and regression_data['SMA200'] < -10
            and 2 < regression_data['PCT_day_change']
            and 1.5 < regression_data['PCT_change']
            and regression_data['year2HighChange'] < -40
            and regression_data['series_trend'] != "downTrend"
            and all_day_pct_change_negative_except_today(regression_data) != True
            ):
            if (regression_data['month3LowChange'] > 20
                and ('P@[' or 'M@[') in regression_data['buyIndia']
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellRisingMAuptrend-0')
                return True
            elif(regression_data['month3LowChange'] < 10
                and ('P@[' or 'M@[') in regression_data['buyIndia']
                
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellRisingMAuptrend-1')
                    return True
                return False
            elif(high_tail_pct(regression_data) > 1.5
                and regression_data['month6HighChange'] > -80
                and ('P@[' or 'M@[') not in regression_data['buyIndia']
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellRisingMAuptrend-2')
                return True
            elif(2 < regression_data['PCT_day_change'] < 4
                and 1.5 < regression_data['PCT_change'] < 4
                and regression_data['forecast_day_PCT7_change'] > 0
                and regression_data['forecast_day_PCT10_change'] > 0
                and (regression_data['month6HighChange'] > -80 and regression_data['month3HighChange'] > -80)
                and ('P@[' or 'M@[') in regression_data['buyIndia']
                #and is_algo_buy(regression_data)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'buyRisingMAuptrend-(Risky)')
                return True
            elif(2 < regression_data['PCT_day_change'] < 4
                and 1.5 < regression_data['PCT_change'] < 4
                and regression_data['forecast_day_PCT7_change'] > 0
                and regression_data['forecast_day_PCT10_change'] > 0
                and (regression_data['month6HighChange'] < -80 or regression_data['month3HighChange'] < -80)
                ):
                if(is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyRisingMAuptrend-Risky')
                    return True
                elif(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellRisingMAuptrend-Risky')
                    return True
                return False
        elif(regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50'] < regression_data['SMA25'] < regression_data['SMA9']
            ):
            if((((regression_data['PCT_change'] + regression_data['PCT_change_pre1'])  < -4
                and (regression_data['PCT_change'] < -3.5 or regression_data['PCT_change_pre1'] < -3.5)
                )
                )
                and regression_data['PCT_change'] < 0
                and regression_data['year2HighChange'] < -40
                and -5 < regression_data['SMA25'] < 0
                and ('[') not in regression_data['sellIndia']
                and regression_data['series_trend'] != "downTrend"
                and all_day_pct_change_negative_except_today(regression_data) != True
                and is_algo_buy(regression_data)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyRisingMA-down-0')
                return True
            elif((((regression_data['PCT_change'] + regression_data['PCT_change_pre1'])  < -4
                and (regression_data['PCT_change'] < -3.5 or regression_data['PCT_change_pre1'] < -3.5)
                )
                )
                and regression_data['PCT_change'] < 0
                and regression_data['year2HighChange'] < -40
                and -5 < regression_data['SMA25'] < 0
                and is_algo_sell(regression_data)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'ML:sellRisingMA-down-(Risky)')
                return True
            elif((
                -4 < regression_data['PCT_change'] < -2
                )
                and regression_data['year2HighChange'] < -30
                and -2.5 < regression_data['SMA9'] < 2
                and -5 < regression_data['SMA25'] < 0
                ):
                if(is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyRisingMA-down-Risky')
                    return True
                elif(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellRisingMA-down-Risky')
                    return True
                return False
#         elif(regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50'] < regression_data['SMA25']
#             and regression_data['SMA25'] > 0
#             and regression_data['SMA50'] < 2
#             and regression_data['SMA100'] < 0
#             and regression_data['SMA200'] < 0
#             and regression_data['year2HighChange'] < -40
#             and regression_data['series_trend'] != "downTrend"
#             and all_day_pct_change_negative_except_today(regression_data) != True
#             ):
#             if(-2 < regression_data['PCT_change'] < -0.5
#                 and -2 < regression_data['PCT_change'] < -1
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '(Test):buyRisingMA-uptrend-SMA25gt0')
                #add_in_csv(regression_data, regressionResult, ws, None)
        
        if((('NearLowYear2' in regression_data['filter3']) 
            or ('NearLowYear' in regression_data['filter3'])
            or ('NearLowMonth6' in regression_data['filter3'])
            or ('ReversalLowYear2' in regression_data['filter3'])
            or ('ReversalLowYear' in regression_data['filter3'])
            or ('ReversalLowMonth6' in regression_data['filter3'])
            or ('BreakLowYear2' in regression_data['filter3'])
            or ('BreakLowYear' in regression_data['filter3'])
            or ('BreakLowMonth6' in regression_data['filter3'])
            )
            and regression_data['SMA200'] < regression_data['SMA100'] < regression_data['SMA50'] < regression_data['SMA25'] < regression_data['SMA9']
            and  ("upTrend" in regression_data['series_trend'])
            ):
    #             if(((-3 < regression_data['PCT_change'] < 0) and (-3 < regression_data['PCT_day_change'] < 0))
    #                 and regression_data['SMA4'] > 0.5
    #                 and regression_data['SMA25'] > -10
    #                 and ('[') not in regression_data['sellIndia']
    #                 ):
    #                 add_in_csv(regression_data, regressionResult, ws, '(Test)buyBottomReversal-0')
            if(((-3 < regression_data['PCT_change'] < 0) and (-3 < regression_data['PCT_day_change'] < 0))
                ):
                if(is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, None)
                    return True
                return False
    #             elif(((-3 < regression_data['PCT_change'] < -1.5) and (-3 < regression_data['PCT_day_change'] < -1.5))
    #                 ):
    #                 add_in_csv(regression_data, regressionResult, ws, '##UPTREND:(Test)buyBottomReversal-2')
     
            
    #             elif(regression_data['PCT_day_change_pre2'] < 0 and (regression_data['PCT_day_change'] > 1.5)):
    #                 add_in_csv(regression_data, regressionResult, ws, '(Test)buyBottomReversal-1')
        return True

def buy_study_risingMA(regression_data, regressionResult, reg, ws):
#     if(('UP-1' in regression_data['filter5'])
#        ):
#         add_in_csv(regression_data, regressionResult, ws,"buy-1-UP")
#         return True
#     elif(('UP-2' in regression_data['filter5'])
#        ):
#         add_in_csv(regression_data, regressionResult, ws, None)
#         return True
#     elif(('UP' in regression_data['filter5'])
#        ):
#         if((-2 < regression_data['SMA9'] < 0 or -2 < regression_data['SMA25'] < 0)
#             and regression_data['SMA100'] < -10
#             and regression_data['SMA200'] < -10
#             and regression_data['year2HighChange'] < -50
#             and regression_data['year2LowChange'] > 8
#             ):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)sell-UP')
#         elif((1 < regression_data['SMA9'] < 3 or 1 < regression_data['SMA25'] < 3)
#             and regression_data['SMA100'] < -10
#             and regression_data['SMA200'] < -10
#             and regression_data['year2LowChange'] > 8
#             ):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)sell-UP')
#         elif((1 < regression_data['SMA9'] < 3 or 1 < regression_data['SMA25'] < 3)
#             and regression_data['SMA100'] < -10
#             and regression_data['SMA200'] < -10
#             and regression_data['year2LowChange'] < 8
#             ):
#             add_in_csv(regression_data, regressionResult, ws, '(Test)buy-UP')
#         return True
#     elif(2.5 < regression_data['PCT_day_change'] < 7
#        and 3 < regression_data['PCT_change'] < 7
#        and regression_data['forecast_day_PCT_change'] > 0
#        and regression_data['forecast_day_PCT2_change'] > 0
#        and regression_data['PCT_day_change_pre1'] < 0.75
#        and regression_data['PCT_day_change_pre2'] < 0.5
#        and regression_data['PCT_day_change_pre3'] < 0
#        and high_tail_pct(regression_data) < 1
#        and low_tail_pct(regression_data) < 1
#        #and low_tail_pct_pre1(regression_data) < 1
#        and regression_data['year2HighChange'] < -5
#        and regression_data['year2LowChange'] > 15
#        and (('P@' not in regression_data['buyIndia']) or (regression_data['forecast_day_PCT10_change'] > 4))
#        ):
#         add_in_csv(regression_data, regressionResult, ws, "sell-UP")
#         return True
#     elif(('DOWN-2' in regression_data['filter5'])
#        ):
#         add_in_csv(regression_data, regressionResult, ws, None)
#     elif(('DOWN-1' in regression_data['filter5'])
#        ):
#         add_in_csv(regression_data, regressionResult, ws, None)
#     elif(('DOWN' in regression_data['filter5'])
#        ):
#         add_in_csv(regression_data, regressionResult, ws, None)
#         
#     
#     
    if(('EMA6-LT-May-MT-EMA14' in regression_data['filter4'])
        and 2.5 < regression_data['PCT_day_change'] < 4
        and 0 < regression_data['PCT_change'] < 4.5
        and regression_data['PCT_day_change_pre1'] > 0
        and (regression_data['PCT_day_change'] > 3 or regression_data['PCT_change'] > 3)
        and high_tail_pct(regression_data) < 1
        ):
        add_in_csv(regression_data, regressionResult, ws, '(check-chart)mayBuy-EMA6-LT-May-MT-EMA14')
    elif(('EMA6-MT-May-LT-EMA14' in regression_data['filter4'])
        and -4 < regression_data['PCT_day_change'] < -2.5
        and -4.5 < regression_data['PCT_change'] < 0
        and regression_data['PCT_day_change_pre1'] < 0
        and (regression_data['PCT_day_change'] < -3 or regression_data['PCT_change'] < -3)
        and low_tail_pct(regression_data) < 1
        ):
        add_in_csv(regression_data, regressionResult, ws, '(check-chart)maySell-EMA6-MT-May-LT-EMA14')
        
        
    if(('(Confirmed)EMA6>EMA14' in regression_data['filter4'])
        and 3 < regression_data['PCT_day_change'] 
        and 0 < regression_data['PCT_change']
        and regression_data['PCT_day_change_pre1'] < -1
        and (regression_data['PCT_day_change'] > 3 or regression_data['PCT_change'] > 3)
        ):
        if(regression_data['month3HighChange'] < 10):
            add_in_csv(regression_data, regressionResult, ws, '(check-chart)sellEMA6>EMA14')
        elif(regression_data['PCT_day_change'] > 3.5 or regression_data['PCT_change'] > 3.5):
            add_in_csv(regression_data, regressionResult, ws, '(check-chart)sellEMA6>EMA14-Risky')
    
    if('$$(Study)$$:RisingMA' in regression_data['filter4']
        ):
        if(-2.5 < regression_data['PCT_day_change'] < 0 and -2.5 < regression_data['PCT_change'] < 0
            and (regression_data['SMA50'] < 0 and regression_data['SMA25'] > 10 and regression_data['SMA4'] > 0)
            ):
            add_in_csv(regression_data, regressionResult, ws, "sellSMA25GT10")
        if(-4 < regression_data['PCT_day_change'] < -0.5
           and -4 < regression_data['PCT_change']
           and 5 > low_tail_pct(regression_data) > 2.5
           and high_tail_pct(regression_data) < 0.6
           and (low_tail_pct(regression_data) - high_tail_pct(regression_data)) > 1
           ):
            add_in_csv(regression_data, regressionResult, ws, None)
            return True
        elif(-4 < regression_data['PCT_day_change'] < -0.5
           and -4 < regression_data['PCT_change']
           and 5 > low_tail_pct(regression_data) > 2.5
           and high_tail_pct(regression_data) < 2
           and (low_tail_pct(regression_data) - high_tail_pct(regression_data)) > 0.8
           ):
            add_in_csv(regression_data, regressionResult, ws, "##RisingMA(check-chart-buy)-MayBuyCheckChart")
            return True
        elif(0.5 < regression_data['PCT_day_change'] < 4
           and regression_data['PCT_change'] < 4
           and low_tail_pct(regression_data) < 0.6
           and 5 > high_tail_pct(regression_data) > 2.5
           and (high_tail_pct(regression_data) - low_tail_pct(regression_data)) > 0.8
           ):
            add_in_csv(regression_data, regressionResult, ws, "##RisingMA(check-chart-sell)-MaySellCheckChart")
            return True
        elif(0.5 < regression_data['PCT_day_change'] < 4
           and regression_data['PCT_change'] < 4
           and low_tail_pct(regression_data) < 2
           and 5 > high_tail_pct(regression_data) > 2.5
           and (high_tail_pct(regression_data) - low_tail_pct(regression_data)) > 0.8
           ):
            add_in_csv(regression_data, regressionResult, ws, None)
            return True
        elif(0.5 < regression_data['PCT_day_change'] < 4
           and regression_data['PCT_day_change_pre1'] < 0
           and abs(regression_data['PCT_day_change_pre1']) > abs(regression_data['PCT_day_change'])
           and ((regression_data['forecast_day_PCT_change'] < 0)
                #or (regression_data['open'] > regression_data['close_pre'])
                )
           and high_tail_pct(regression_data) < 1
           ):
            add_in_csv(regression_data, regressionResult, ws, "##RisingMA(Test)(check-chart-buy)-1DayUp")
            return True
        elif(0.5 < regression_data['PCT_day_change'] < 4
           and 0.5 < regression_data['PCT_day_change_pre1'] < 4
           and regression_data['PCT_day_change_pre2'] < 0
           and regression_data['PCT_day_change_pre3'] < 0
           and abs(regression_data['PCT_day_change_pre2']) > abs(regression_data['PCT_day_change_pre1'])
           and regression_data['forecast_day_PCT_change'] > 0
           and high_tail_pct(regression_data) < 1
           ):
            add_in_csv(regression_data, regressionResult, ws, "##RisingMA(Test)(check-chart-buy)-2DayUp")
            return True
        elif(0 < regression_data['PCT_day_change'] < 1.5
           and regression_data['PCT_day_change_pre1'] < -2
           and (regression_data['PCT_day_change_pre2'] < 0 
                or (regression_data['PCT_day_change_pre2'] < 1 and high_tail_pct_pre2(regression_data) > 1.5)
                )
           and abs(regression_data['PCT_day_change_pre1']) > abs(regression_data['PCT_day_change'])
           and regression_data['forecast_day_PCT_change'] < 0
           and regression_data['forecast_day_PCT2_change'] < 0
           and regression_data['forecast_day_PCT3_change'] < 0
           and regression_data['forecast_day_PCT4_change'] < 0
           and regression_data['forecast_day_PCT5_change'] < 0
           ):
            add_in_csv(regression_data, regressionResult, ws, "##RisingMA(Test)(check-chart-sell)-1DayUp")
            return True
        
    if(regression_data['SMA9'] < 0 
        and regression_data['SMA4'] < 0
        and regression_data['forecast_day_PCT_change'] < -0.5
        and regression_data['forecast_day_PCT2_change'] < -0.5
        and regression_data['forecast_day_PCT5_change'] < -0.5
        and regression_data['forecast_day_PCT5_change'] > -5
        and regression_data['forecast_day_PCT7_change'] > -10
        and abs(regression_data['PCT_day_change_pre1']) > 2* (regression_data['PCT_day_change'])
        and ((0 < regression_data['PCT_day_change'] < 1.5
            and regression_data['PCT_day_change_pre1'] < -2
            and regression_data['PCT_day_change_pre2'] < 1.5
            and regression_data['PCT_day_change_pre3'] < 1.5
            and (
                regression_data['PCT_day_change_pre2'] < -1
                or regression_data['PCT_day_change_pre3'] < -1
                )
            )
        )
        and high_tail_pct(regression_data) > 1
        and (low_tail_pct(regression_data) < 1.5)
        ):  
        add_in_csv(regression_data, regressionResult, ws, "##(Test)(check-chart)-sellSMADownTrend") 
    elif((regression_data['SMA9'] < 0 or regression_data['SMA25'] < 0)
        and 0 < regression_data['weekHighChange'] < 0.5
        and regression_data['close'] < regression_data['high_week']
        and regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and (regression_data['forecast_day_PCT5_change'] < 0
            or regression_data['forecast_day_PCT7_change'] < 0
            ) 
        and regression_data['forecast_day_PCT10_change'] < 0
        and ((0 < regression_data['PCT_day_change'] > 1
            and regression_data['PCT_day_change_pre1'] > 1
            and (
                regression_data['PCT_day_change_pre2'] < 0
                or regression_data['PCT_day_change_pre3'] < 0
                )
            )
        )
        and high_tail_pct(regression_data) > 1 
        and (low_tail_pct(regression_data) < 1.5)
        ):  
        add_in_csv(regression_data, regressionResult, ws, "##(Test)(check-chart)-sellSMADownTrend-2DayUp-weekHighTouch")
           
        return True

def buy_random_filters(regression_data, regressionResult, reg, ws):
    if(1 < regression_data['PCT_day_change'] < 5
        and (regression_data['PCT_day_change']-0.5) < regression_data['PCT_change'] < (regression_data['PCT_day_change'] + 0.5)
        and (regression_data['forecast_day_PCT10_change'] < 1)
        and (regression_data['forecast_day_PCT7_change'] < 0
            or regression_data['forecast_day_PCT10_change'] < 0
            )
        and (regression_data['forecast_day_PCT_change'] > 0
            or regression_data['forecast_day_PCT2_change'] > 0
            or regression_data['forecast_day_PCT3_change'] > 0
            )
        and (((regression_data['low'] - regression_data['low_pre1'])/regression_data['low_pre1'])*100) > 0
        ):
        if(regression_data['week2HighChange'] > -0.5):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-near2WeekHigh-mayReveresalSell')
        elif(regression_data['PCT_day_change_pre1'] > 2
            and regression_data['PCT_day_change_pre2'] > 2
            and regression_data['forecast_day_PCT5_change'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-3dayUp-mayReversalSell')
        elif(regression_data['PCT_day_change'] > 2
            and regression_data['PCT_day_change_pre1'] > 2
            and regression_data['forecast_day_PCT5_change'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-2dayUp-mayReversalSell')
        elif(2 < regression_data['PCT_day_change'] < 3
            and regression_data['week2HighChange'] < -3
            ):
            if(is_algo_buy(regression_data)
                and 2 > regression_data['PCT_day_change_pre1'] > 0.5
                and regression_data['PCT_day_change_pre2'] > 0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-3dayUp')
            elif(is_algo_buy(regression_data)
                and 2 > regression_data['PCT_day_change_pre1'] > 0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-2dayUp')
            elif(regression_data['forecast_day_PCT7_change'] < 1
                and regression_data['PCT_day_change_pre1'] < 0 
                and regression_data['PCT_day_change_pre2'] < 0
                and regression_data['PCT_day_change_pre3'] < 0
                and regression_data['forecast_day_PCT_change'] > 0
                and regression_data['forecast_day_PCT2_change'] > 0
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-downTrendLastDayUp-mayReversalSell')
            elif(is_algo_buy(regression_data) 
                and regression_data['forecast_day_PCT7_change'] < 1
                and (regression_data['PCT_day_change_pre2'] < 0
                    or regression_data['bar_high'] > regression_data['bar_high_pre1']
                    )
                and abs_month3High_less_than_month3Low(regression_data)
                ):
                if(regression_data['bar_high'] > regression_data['bar_high_pre1']):
                    add_in_csv(regression_data, regressionResult, ws, 'bar_high')
                if(regression_data['forecast_day_PCT2_change'] > 1
                    and regression_data['forecast_day_PCT3_change'] > 1  
                    ):
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp')
                else:
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-inUpTrend')
            elif(is_algo_buy(regression_data) 
                and regression_data['forecast_day_PCT7_change'] < 1
                and (regression_data['PCT_day_change_pre2'] < 0
                    or regression_data['bar_high'] > regression_data['bar_high_pre1']
                    )
                ):
                if(regression_data['bar_high'] > regression_data['bar_high_pre1']):
                    add_in_csv(regression_data, regressionResult, ws, 'bar_high')
                if(regression_data['forecast_day_PCT2_change'] > 1
                    and regression_data['forecast_day_PCT3_change'] > 1  
                    ):
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-month3HighMTmonth3Low')
                else:
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-inUpTrend-month3HighMTmonth3Low')
            elif(regression_data['forecast_day_PCT7_change'] < 1
                and regression_data['PCT_day_change_pre1'] < 0 
                and regression_data['PCT_day_change_pre2'] > 0
                and regression_data['bar_high'] < regression_data['bar_high_pre1']
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-mayReversalSell')
            
    if(-2.5 < regression_data['PCT_day_change'] < -1
        and -3.5 < regression_data['PCT_change'] < -1
        and regression_data['PCT_day_change_pre1'] > 1.5
        and (regression_data['PCT_day_change_pre2'] > 1
            or regression_data['PCT_day_change_pre3'] > 1
            )
        #and (regression_data['forecast_day_PCT5_change'] < 2)
        and (regression_data['forecast_day_PCT7_change'] < 2)
        and (regression_data['forecast_day_PCT10_change'] < 2)
        and (regression_data['forecast_day_PCT5_change'] < 0
            or regression_data['forecast_day_PCT7_change'] < 0
            or regression_data['forecast_day_PCT10_change'] < 0
            )
        and (regression_data['forecast_day_PCT2_change'] > 0
            and regression_data['forecast_day_PCT3_change'] > 0
            )
        and (((regression_data['low'] - regression_data['low_pre1'])/regression_data['low_pre1'])*100) > 0
        and (((regression_data['low'] - regression_data['bar_low_pre1'])/regression_data['bar_low_pre1'])*100) > 0 
        ):
        if(regression_data['week2HighChange'] > -0.5):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-lastDayDown-near2WeekHigh-mayReveresalSell')
        elif(regression_data['week2HighChange'] < -2):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupUp-lastDayDown')  
        
    if((-0.5 < regression_data['forecast_day_PCT4_change'] < 0.5
        or -0.5 < regression_data['forecast_day_PCT5_change'] < 0.5
        )
        and 
        (regression_data['forecast_day_PCT4_change'] > 1
        or regression_data['forecast_day_PCT5_change'] > 1
        )
        and regression_data['forecast_day_PCT_change'] > 1
        and regression_data['forecast_day_PCT2_change'] > 2
        and regression_data['forecast_day_PCT3_change'] > 3
#         and regression_data['forecast_day_PCT7_change'] > 3
#         and regression_data['forecast_day_PCT10_change'] > 3
        and regression_data['PCT_day_change'] > 1
        and regression_data['PCT_day_change_pre1'] > 1
        and regression_data['PCT_day_change_pre2'] > 1
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)buyMay5DayCeilReversal')
        
    if(1.5 < regression_data['PCT_day_change'] < 6
        and regression_data['PCT_change'] < 6
        and (regression_data['PCT_day_change_pre2'] < 0 or regression_data['PCT_day_change_pre3'] < 0)
        and regression_data['high'] > regression_data['high_pre2']
        and regression_data['forecast_day_PCT3_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT_change'] > 0
        and regression_data['SMA4'] > regression_data['SMA4_2daysBack']
        and regression_data['forecast_day_PCT10_change'] < 0
        and ((regression_data['forecast_day_PCT4_change'] > 0
              and regression_data['forecast_day_PCT7_change'] < 0
              and (regression_data['forecast_day_PCT7_change'] < -5
                   or regression_data['forecast_day_PCT10_change'] < -5
                   )
            )
            or(regression_data['monthHighChange'] < -15
                and regression_data['monthLowChange'] < 10
                )
            ) 
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)check10DayLowReversal')
        
    if('RisingMA-Risky' not in regression_data['filter4'] 
        and 'RisingMA' in regression_data['filter4']
        and (regression_data['low'] > regression_data['low_pre1']) 
        and (-2 < regression_data['PCT_day_change'] < 0) and (-2 < regression_data['PCT_change'] < 0)
        and (regression_data['forecast_day_PCT5_change'] > 3
             or regression_data['forecast_day_PCT7_change'] > 3
             or regression_data['forecast_day_PCT10_change'] > 3
            )
        ):
        if(-0.5 < regression_data['week2HighChange'] < 0.5 
            and regression_data['week2High'] != regression_data['high_pre1']
            and regression_data['week2High'] != regression_data['high_pre2']
            #and (-1 < regression_data['PCT_day_change'] < 0) and (-1 < regression_data['PCT_change'] < 0)
            ):
                add_in_csv(regression_data, regressionResult, ws, None)
        elif(regression_data['week2HighChange'] < -1):
            add_in_csv(regression_data, regressionResult, ws, '(Test)buyRisingMA-(check5MinuteUpTrendAndBuySupertrend)')
        elif(is_algo_buy(regression_data)):
            add_in_csv(regression_data, regressionResult, ws, '(Test)buyRisingMA-(check5MinuteUpTrendAndBuySupertrend)-Already10DayDown')
    
    if((-0.75 < regression_data['PCT_day_change'] < 0) and (-0.75 < regression_data['PCT_change'] < 0)
        and (regression_data['SMA4_2daysBack'] > 0 or regression_data['SMA9_2daysBack'] > 0)
        and regression_data['SMA4'] < 0
        and regression_data['PCT_day_change_pre1'] < -0.5
        and regression_data['PCT_day_change_pre2'] < -0.5
        and (regression_data['PCT_day_change_pre1'] < -1.5 or regression_data['PCT_day_change_pre2'] < -1.5)
        ):   
        add_in_csv(regression_data, regressionResult, ws, '(Test)buySMA4Reversal')
        
    if((-2 < regression_data['PCT_day_change'] < 0) and (-2 < regression_data['PCT_change'] < 0)
        and regression_data['PCT_day_change_pre1'] < -7
        and low_tail_pct(regression_data) > 2
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)buyLastDayDownReversal')
        
    if((0.5 < regression_data['PCT_day_change'] < 2) and (-10 < regression_data['PCT_change'] < -4)
        and low_tail_pct(regression_data) > 1
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)buyLastDayHighDownReversal')
         
def buy_skip_close_lt_50(regression_data, regressionResult, reg, ws):
    return False
        
def buy_test_345(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    filterName = pct_change_filter(regression_data, regressionResult, False)
    regression_data['filterTest'] = filterName + ',' \
                                    + regression_data['series_trend'] + ',' \
                                    + regression_data['filter2'] + ',' \
                                    + regression_data['filter3'] + ',' \
                                    + regression_data['filter4'] + ',' \
                                    + regression_data['filter5']
    
    if regression_data['filterTest'] != '':
        return True  
    
    return False
  
def buy_test(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    if regression_data['filterTest'] != '':
        regression_data['filterTest'] = regression_data['series_trend'] + ',' \
                                        + regression_data['filterTest'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    
    if regression_data['filterTest'] != '':
        return True
    
    return False

def buy_test_pct_change(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    if regression_data['filterTest'] != '':
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        regression_data['filterTest'] = filterName + ',' \
                                        + filterNameTail + ',' \
                                        + regression_data['series_trend'] + ',' \
                                        + regression_data['filterTest'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    
    if regression_data['filterTest'] != '':
        return True
    
    return False  

def buy_test_all(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    regression_data['filterTest'] = filterName + ',' \
                                    + filterNameTail + ',' \
                                    + regression_data['series_trend'] + ',' \
                                    + regression_data['filter2'] + ',' \
                                    + regression_data['filter3'] + ',' \
                                    + regression_data['filter4'] + ',' \
                                    + regression_data['filter5'] + ',' \
                                    + regression_data['filterTest'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    if regression_data['filterTest'] != '':
        return True  
    return False

def buy_test_tech(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    if (regression_data['buyIndia'] != '' or regression_data['sellIndia'] != ''):
        regression_data['filterTest'] = 'B@{' + regression_data['buyIndia'] + '}' \
                                        + 'S@{' + regression_data['sellIndia'] + '}'
        return True
    
    return False

def buy_test_tech_all(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    if (regression_data['buyIndia'] != '' or regression_data['sellIndia'] != ''):
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        regression_data['filterTest'] = filterNameTail + ',' \
                                        + regression_data['series_trend'] + ',' \
                                        + 'B@{' + regression_data['buyIndia'] + '}' \
                                        + 'S@{' + regression_data['sellIndia'] + '}' \
                                        + regression_data['filter3'] + ',' \
                                        + regression_data['filter4'] + ',' 
        return True
    
    return False

def buy_test_tech_all_pct_change(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    if (regression_data['buyIndia'] != '' or regression_data['sellIndia'] != ''):
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        regression_data['filterTest'] = filterName + ',' \
                                        + filterNameTail + ',' \
                                        + regression_data['series_trend'] + ',' \
                                        + 'B@{' + regression_data['buyIndia'] + '}' \
                                        + 'S@{' + regression_data['sellIndia'] + '}' \
                                        + regression_data['filter3'] + ',' \
                                        + regression_data['filter4'] + ',' 
        return True
    
    return False

def buy_oi_candidate(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    tail_pct_filter(regression_data, regressionResult)
    tail_reversal_filter(regression_data, regressionResult)
    flag = False
    
    if(buy_trend_reversal(regression_data, regressionResult, reg, ws)
        and regression_data['close'] > 50
        ):
        flag = True
    if(breakout_or_no_consolidation(regression_data) == True):
        if buy_evening_star_sell(regression_data, regressionResult, reg, ws):
            flag = True
        if buy_morning_star_buy(regression_data, regressionResult, reg, ws):
            flag = True
        if buy_day_low(regression_data, regressionResult, reg, ws):
            flag = True
        if buy_trend_break(regression_data, regressionResult, reg, ws):
            flag = True
        if buy_final_candidate(regression_data, regressionResult, reg, ws):
            flag = True
    return flag

def buy_all_filter(regression_data, regressionResult, reg, ws):
    flag = False
    if buy_year_high(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_year_low(regression_data, regressionResult, reg, None, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_up_trend(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_down_trend(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_final(regression_data, regressionResult, reg, None, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_af_high_indicators(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_pattern(regression_data, regressionResult, reg, None, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if buy_oi(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    return flag

def buy_filter_345_accuracy(regression_data, regressionResult):
    filtersDict=filter345buy
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filter = filterName + ',' \
            + regression_data['series_trend'] + ',' \
            + regression_data['filter2'] + ',' \
            + regression_data['filter3'] + ',' \
            + regression_data['filter4'] + ',' \
            + regression_data['filter5']
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_345_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_345_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_345_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_345_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def buy_filter_accuracy(regression_data, regressionResult):
    filtersDict=filterbuy
    if regression_data['filter'] != '':
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        filter = regression_data['series_trend'] + ',' \
                + regression_data['filter'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
        if filter != '' and filter in filtersDict:
            regression_data['filter_avg'] = float(filtersDict[filter]['avg'])
            regression_data['filter_count'] = float(filtersDict[filter]['count'])
            if float(filtersDict[filter]['count']) >= 2:
                if float(filtersDict[filter]['avg']) >= 0:
                    regression_data['filter_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
                else:
                    regression_data['filter_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def buy_filter_pct_change_accuracy(regression_data, regressionResult):
    filtersDict=filterpctchangebuy
    if regression_data['filter'] != '':
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        filter = filterName + ',' \
                + filterNameTail + ',' \
                + regression_data['series_trend'] + ',' \
                + regression_data['filter'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
        if filter != '' and filter in filtersDict:
            regression_data['filter_pct_change_avg'] = float(filtersDict[filter]['avg'])
            regression_data['filter_pct_change_count'] = float(filtersDict[filter]['count'])
            if float(filtersDict[filter]['count']) >= 2:
                if float(filtersDict[filter]['avg']) >= 0:
                    regression_data['filter_pct_change_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
                else:
                    regression_data['filter_pct_change_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])
                
def buy_filter_345_all_accuracy(regression_data, regressionResult):
    filtersDict=filterallbuy
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    filter = filterName + ',' \
            + filterNameTail + ',' \
            + regression_data['series_trend'] + ',' \
            + regression_data['filter2'] + ',' \
            + regression_data['filter3'] + ',' \
            + regression_data['filter4'] + ',' \
            + regression_data['filter5'] + ',' \
            + regression_data['filter'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_all_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_all_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_all_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_all_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def buy_filter_tech_accuracy(regression_data, regressionResult):
    filtersDict=filtertechbuy
    filter = 'B@{' + regression_data['buyIndia'] + '}' \
            + 'S@{' + regression_data['sellIndia'] + '}' 
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_tech_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_tech_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_tech_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_tech_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def buy_filter_tech_all_accuracy(regression_data, regressionResult):
    filtersDict=filtertechallbuy
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    filter = filterNameTail + ',' \
            + regression_data['series_trend'] + ',' \
            + 'B@{' + regression_data['buyIndia'] + '}' \
            + 'S@{' + regression_data['sellIndia'] + '}' \
            + regression_data['filter3'] + ',' \
            + regression_data['filter4'] + ','
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_tech_all_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_tech_all_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_tech_all_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_tech_all_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def buy_filter_tech_all_pct_change_accuracy(regression_data, regressionResult):
    filtersDict=filtertechallpctchangebuy
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    filter = filterName + ',' \
            + filterNameTail + ',' \
            + regression_data['series_trend'] + ',' \
            + 'B@{' + regression_data['buyIndia'] + '}' \
            + 'S@{' + regression_data['sellIndia'] + '}' \
            + regression_data['filter3'] + ',' \
            + regression_data['filter4'] + ','
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_tech_all_pct_change_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_tech_all_pct_change_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_tech_all_pct_change_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_tech_all_pct_change_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def buy_filter_all_accuracy(regression_data, regressionResult):
    buy_filter_345_accuracy(regression_data, regressionResult)
    buy_filter_accuracy(regression_data, regressionResult)
    buy_filter_pct_change_accuracy(regression_data, regressionResult) 
    buy_filter_345_all_accuracy(regression_data, regressionResult)
    buy_filter_tech_accuracy(regression_data, regressionResult)
    buy_filter_tech_all_accuracy(regression_data, regressionResult)
    buy_filter_tech_all_pct_change_accuracy(regression_data, regressionResult)
                           
def sell_pattern_without_mlalgo(regression_data, regressionResult):
    if(regression_data['PCT_day_change'] > -3.5
        and regression_data['year2HighChange'] < -5):
        if(regression_data['sellIndia_avg'] < -0.9 and regression_data['sellIndia_count'] > 1
            and (regression_data['SMA9'] < 0 or regression_data['SMA4'] < -1.5)
            ):
            if(regression_data['SMA9'] < 0):
                add_in_csv(regression_data, regressionResult, None, '(SMA9LT0)')
            elif(regression_data['SMA4'] < -1.5):
                add_in_csv(regression_data, regressionResult, None, '(SMA4LT(-1.5))')
            if(regression_data['yearHighChange'] < -10 and regression_data['sellIndia_count'] > 1):
                if(regression_data['sellIndia_avg'] < -1.5):
                    add_in_csv(regression_data, regressionResult, None, 'sellPatterns-LT-1.5-SMALT0')
                elif(regression_data['sellIndia_avg'] < -0.9
                    ):
                    add_in_csv(regression_data, regressionResult, None, 'sellPatterns-Risky-LT-0.9-SMALT0')   
                    
    if(regression_data['PCT_day_change'] > -3.5
        and regression_data['year2HighChange'] < -5
        and regression_data['year2LowChange'] > 5):
        if(regression_data['sellIndia_avg'] > 1.8):
            add_in_csv(regression_data, regressionResult, None, 'mayBuyFromSellPattern-GT1.8')
        if(regression_data['sellIndia_avg'] > 0.9 and regression_data['sellIndia_count'] > 1
            and (regression_data['SMA9'] > 0 or regression_data['SMA4'] > 1.5)
            ):
            if(regression_data['SMA9'] > 0):
                add_in_csv(regression_data, regressionResult, None, '(SMA9GT0)')
            elif(regression_data['SMA4'] > 1.5):
                add_in_csv(regression_data, regressionResult, None, '(SMA4GT(1.5))')
            if(regression_data['yearLowChange'] < -10 and regression_data['sellIndia_count'] > 1):
                if(regression_data['sellIndia_avg'] > 1.5):
                    add_in_csv(regression_data, regressionResult, None, 'buyPatterns-GT1.5-SMAGT0')
                elif(regression_data['sellIndia_avg'] > 0.9
                    ):
                    add_in_csv(regression_data, regressionResult, None, 'buyPatterns-Risky-GT0.9-SMAGT0')                          
            #add_in_csv(regression_data, regressionResult, None, 'buyPatterns-1-Risky')   
#     sellPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/all-buy-filter-by-PCT-Change.csv')
#     if regression_data['sellIndia'] != '' and regression_data['sellIndia'] in sellPatternsDict:
#         if (abs(float(sellPatternsDict[regression_data['sellIndia']]['avg'])) >= .1 and float(sellPatternsDict[regression_data['sellIndia']]['count']) >= 2):
#             if(-3 < regression_data['PCT_day_change'] < 0.5 and float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -1):
#                 avg = sellPatternsDict[regression_data['sellIndia']]['avg']
#                 count = sellPatternsDict[regression_data['sellIndia']]['count']
#                 #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'wml_sell', avg, count)
#             if(-0.5 < regression_data['PCT_day_change'] < 3 and float(sellPatternsDict[regression_data['sellIndia']]['avg']) > 1): 
#                 avg = sellPatternsDict[regression_data['sellIndia']]['avg']
#                 count = sellPatternsDict[regression_data['sellIndia']]['count']
#                 #add_in_csv_hist_pattern(regression_data, regressionResult, ws2, 'wml_sell', avg, count)

def sell_pattern_from_history(regression_data, ws):
    sellPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/patterns-sell.csv')
    sellIndiaAvg = 0
    regression_data['buyIndia_avg'] = 0
    regression_data['buyIndia_count'] = 0
    regression_data['sellIndia_avg'] = 0
    regression_data['sellIndia_count'] = 0
    regression_data['filter_avg'] = 0
    regression_data['filter_count'] = 0
    flag = False
    if regression_data['sellIndia'] != '' and regression_data['sellIndia'] in sellPatternsDict: 
        regression_data['sellIndia_avg'] = float(sellPatternsDict[regression_data['sellIndia']]['avg'])
        regression_data['sellIndia_count'] = float(sellPatternsDict[regression_data['sellIndia']]['count'])
        if (abs(float(sellPatternsDict[regression_data['sellIndia']]['avg'])) >= .1):
            sellIndiaAvg = float(sellPatternsDict[regression_data['sellIndia']]['avg'])
            if(int(sellPatternsDict[regression_data['sellIndia']]['count']) >= 2):
                if(is_algo_sell(regression_data)
                    and 'P@[' not in str(regression_data['buyIndia'])
                    and regression_data['trend'] != 'down'
                    and -3 < regression_data['PCT_day_change'] < 0.5):
                    avg = sellPatternsDict[regression_data['sellIndia']]['avg']
                    count = sellPatternsDict[regression_data['sellIndia']]['count']
                    if(float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -1 and int(sellPatternsDict[regression_data['sellIndia']]['count']) >= 5):
                        flag = True
                        #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'sellPattern2', avg, count) 
                    if(float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -0.5 
                        or (float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -0.3 and (ten_days_more_than_ten(regression_data) or regression_data['yearLowChange'] > 40))):
                        if(regression_data['forecast_day_PCT10_change'] > 0 and regression_data['forecast_day_PCT_change'] <= 0):
                            flag = True
                            #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'sellPattern2', avg, count)
                        elif(regression_data['forecast_day_PCT10_change'] < 0):    
                            flag = True
                            #add_in_csv_hist_pattern(regression_data, regressionResult, ws, 'sellPattern2', avg, count)
    return sellIndiaAvg, flag

def sell_all_rule(regression_data, regressionResult, sellIndiaAvg, ws):
    if(is_algo_sell(regression_data)
        and sellIndiaAvg <= 0.70
        and ((last_7_day_all_down(regression_data) == False) or (regression_data['forecast_day_PCT10_change'] > -10))
        and (MARKET_IN_DOWNTREND or (last_4_day_all_down(regression_data) == False)) #Uncomment0 If very less data
        and breakout_or_no_consolidation(regression_data) == True
        and '(Confirmed)EMA6<EMA14' not in regression_data['filter4']
        and regression_data['close'] > 50
        and low_tail_pct(regression_data) < 1
        and high_tail_pct(regression_data) > 1.2
        and regression_data['low'] < regression_data['low_pre1']
        ):
        if(regression_data['PCT_day_change'] > 0.75 and regression_data['PCT_change'] > 0
            and regression_data['month3LowChange'] > 15  
            ):
            add_in_csv(regression_data, regressionResult, None, 'ML:sell-0')
#     elif(is_algo_buy(regression_data)):
#         if(-3.5 < regression_data['PCT_day_change'] < -2
#             and -4 < regression_data['PCT_change'] < -2
#             and low_tail_pct(regression_data) > 1.5
#             and high_tail_pct(regression_data) < 1
#             ):
#             add_in_csv(regression_data, regressionResult, None, 'ML:buy-1')
    
    if(is_algo_sell(regression_data, True)
        and ((regression_data['PCT_day_change'] < -1 
                and low_tail_pct(regression_data) < 2
                and ('P@[' not in str(regression_data['buyIndia']))
                ) 
             or (high_tail_pct(regression_data) > 1.5 
                and (low_tail_pct(regression_data) < high_tail_pct(regression_data))
                )
             or (regression_data['PCT_day_change'] > 2
                and regression_data['forecast_day_PCT_change'] < -0.3
                )
            )
        and (SELL_VERY_LESS_DATA or regression_data['PCT_change'] < 1)
        and (SELL_VERY_LESS_DATA or ((regression_data['PCT_day_change_pre1'] > 0) or (regression_data['forecast_day_VOL_change'] > 0))) #Uncomment1 If very less data
        and (SELL_VERY_LESS_DATA or (regression_data['bar_low']-regression_data['low']) < (regression_data['bar_high']-regression_data['bar_low']))
        and sellIndiaAvg <= 0.70
        and ((last_7_day_all_down(regression_data) == False) or (regression_data['forecast_day_PCT10_change'] > -10))
        and (MARKET_IN_DOWNTREND or (last_4_day_all_down(regression_data) == False)) #Uncomment0 If very less data
        and breakout_or_no_consolidation(regression_data) == True
        and '(Confirmed)EMA6<EMA14' not in regression_data['filter4']
        and regression_data['close'] > 50
        ):
        add_in_csv(regression_data, regressionResult, ws, None)
        return True
    return False

def sell_all_rule_classifier(regression_data, regressionResult, sellIndiaAvg, ws):
    if(is_algo_sell_classifier(regression_data, True)
        and ((regression_data['PCT_day_change'] < -1 
                and low_tail_pct(regression_data) < 2
                and ('P@[' not in str(regression_data['buyIndia']))
                ) 
             or (high_tail_pct(regression_data) > 1.5 
                and (low_tail_pct(regression_data) < high_tail_pct(regression_data))
                )
             or (regression_data['PCT_day_change'] > 2
                and regression_data['forecast_day_PCT_change'] < -0.3
                )
            )
        and (SELL_VERY_LESS_DATA or regression_data['PCT_change'] < 1)
        and (SELL_VERY_LESS_DATA or ((regression_data['PCT_day_change_pre1'] > 0) or (regression_data['forecast_day_VOL_change'] > 0))) #Uncomment1 If very less data
        and (SELL_VERY_LESS_DATA or (regression_data['bar_low']-regression_data['low']) < (regression_data['bar_high']-regression_data['bar_low']))
        and sellIndiaAvg <= 0.70
        and ((last_7_day_all_down(regression_data) == False) or (regression_data['forecast_day_PCT10_change'] > -10))
        and (MARKET_IN_DOWNTREND or (last_4_day_all_down(regression_data) == False)) #Uncomment0 If very less data
        and breakout_or_no_consolidation(regression_data) == True
        and '(Confirmed)EMA6<EMA14' not in regression_data['filter4']
        and regression_data['close'] > 50
        ):
        add_in_csv(regression_data, regressionResult, ws, None)
        return True
    return False

def sell_all_common(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, True)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, True)
    mlpValue_cla, kNeighboursValue_cla = get_reg_or_cla(regression_data, False)
    mlpValue_other_cla, kNeighboursValue_other_cla = get_reg_or_cla_other(regression_data, False)
    
    if(is_algo_sell(regression_data)
        and (-4 < regression_data['PCT_day_change'] < -2) and (-4 < regression_data['PCT_change'] < -2)
        and ((regression_data['PCT_day_change_pre1'] > 0 and regression_data['forecast_day_PCT_change'] < 0)
            )
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0
        #and regression_data['forecast_day_PCT10_change'] > -15
        and regression_data['month3HighChange'] < -10
        and regression_data['month3LowChange'] > 10
        #and regression_data['trend'] != 'down'
        ):    
        if(regression_data['SMA9'] < -1):
            add_in_csv(regression_data, regressionResult, ws, '##Common:sellNotM3HighLow-0(SMA9LT-1)') 
        elif(regression_data['SMA25'] < 0):
            add_in_csv(regression_data, regressionResult, ws, '##(Test):sellNotM3HighLow-0(SMA25LT0)') 
        else:
            add_in_csv(regression_data, regressionResult, ws, None) 
                
    return False

def sell_all_common_High_Low(regression_data, regressionResult, reg, ws):
    if( ("downTrend" in regression_data['series_trend'])
        and (0 < regression_data['PCT_day_change'] < 1) and (0 < regression_data['PCT_change'] < 1)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:LastDayUpInDowntrend')
    
    if((0 < regression_data['PCT_day_change'] < 2) and (0 < regression_data['PCT_change'] < 2.5)
        and (1 <= high_tail_pct(regression_data) < 3)
        ):
        if(regression_data['bar_low'] > regression_data['bar_low_pre1']
            and regression_data['bar_low_pre1'] < regression_data['bar_low_pre2']
            and regression_data['low'] > regression_data['low_pre1']
            and regression_data['low_pre1'] < regression_data['low_pre2'] < regression_data['low_pre3']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySellContinueHighTail-uptrendStart(CheckChart-Risky)') 
        elif(regression_data['high'] > regression_data['high_pre1']
            and regression_data['high_pre1'] > regression_data['high_pre2']
            and 0 < regression_data['PCT_day_change_pre1']
            and 0 < regression_data['PCT_day_change_pre2']
            and 0 < regression_data['PCT_day_change_pre3']
            and (regression_data['forecast_day_PCT3_change'] > 0)
            and (regression_data['forecast_day_PCT5_change'] > 5)
            and (regression_data['forecast_day_PCT7_change'] > 10 or regression_data['forecast_day_PCT7_change'] > 10)
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySellStartHighTail-uptrend(CheckChart-risky)') 
        elif((low_tail_pct(regression_data) < 1 and 2.5 > high_tail_pct(regression_data) >= 1.5)
            and regression_data['PCT_day_change'] > 0
            and regression_data['PCT_change'] > 0  
            ):
            if(regression_data['PCT_day_change_pre1'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySell-CheckChart-LastDayDown(|/mayFail(|before10AM))')
            else:      
                add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySell-CheckChart(|\mayFail(|before10AM))')
                
    if((-2 < regression_data['PCT_day_change'] < -0.3) and (-2 < regression_data['PCT_change'] < -0.3)
        and (1 <= low_tail_pct(regression_data) < 2)
        ):
        if(regression_data['low'] < regression_data['low_pre1']
            and regression_data['low_pre1'] < regression_data['low_pre2']
            and -5 < regression_data['PCT_day_change_pre1'] < 0
            and -5 < regression_data['PCT_day_change_pre2'] < 1
            and (regression_data['forecast_day_PCT3_change'] > -5)
            and (regression_data['forecast_day_PCT5_change'] > -5)
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySellContinueLowTail-downtrend(CheckChart-Risky)')
        elif(regression_data['low'] < regression_data['low_pre1']
            and regression_data['low_pre1'] < regression_data['low_pre2']
            and 0 < regression_data['PCT_day_change_pre2'] < 1
            and (regression_data['forecast_day_PCT3_change'] > -5)
            and (regression_data['forecast_day_PCT5_change'] > -5)
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySellContinueLowTail-downtrend-0(CheckChart-Risky)')    
        elif(((regression_data['low'] < regression_data['low_pre1']
              and regression_data['forecast_day_PCT2_change'] < 0
              )
              or (regression_data['PCT_day_change_pre1'] > 5)
              )
              and regression_data['forecast_day_PCT_change'] < 0 
              and regression_data['forecast_day_PCT3_change'] > 0
              and regression_data['forecast_day_PCT5_change'] > 0 
              and regression_data['forecast_day_PCT7_change'] > 0 
              and regression_data['forecast_day_PCT10_change'] > 0
            ): 
            add_in_csv(regression_data, regressionResult, ws, 'CommonHL:MaySellContinueLowTail-downtrend-1(CheckChart-Risky)')
        
    if((0 < regression_data['PCT_day_change'] < 0.75) and (0 < regression_data['PCT_change'] < 0.75)
        and (regression_data['SMA4_2daysBack'] < 0 or regression_data['SMA9_2daysBack'] < 0)
        and regression_data['SMA4'] > 0
        and regression_data['PCT_day_change_pre1'] > 0.5
        and regression_data['PCT_day_change_pre2'] > 0.5
        and (regression_data['PCT_day_change_pre1'] > 1.5 or regression_data['PCT_day_change_pre2'] > 1.5)
        ):   
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:sellSMA4Reversal')
    if('DowningMA-Risky' not in regression_data['filter4'] 
        and 'DowningMA' in regression_data['filter4']
        and (regression_data['PCT_day_change'] > 0.2 
             or (high_tail_pct(regression_data) > 1) and (high_tail_pct(regression_data) > low_tail_pct(regression_data))
            ) 
        and (-0.5 < regression_data['PCT_day_change'] < 1) and (-0.75 < regression_data['PCT_change'] < 1.5)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:sellDowningMA')
    elif(regression_data['SMA4'] < 0 and regression_data['SMA4_2daysBack'] < 0
        and regression_data['SMA4'] < regression_data['SMA4_2daysBack'] 
        and regression_data['SMA4'] < regression_data['SMA9'] < regression_data['SMA25']
        and (0 < regression_data['PCT_day_change'] < 1) and (-0.5 < regression_data['PCT_change'] < 1.5)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:sellDowningMAShortTerm')
    elif(regression_data['SMA4'] < regression_data['SMA4_2daysBack']  
        and regression_data['SMA9'] > 2 > regression_data['SMA4'] > 0
        and regression_data['SMA25'] > 2 > regression_data['SMA4'] > 0 
        and (0.5 < regression_data['PCT_day_change'] < 3) and (0.5 < regression_data['PCT_change'] < 3)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'CommonHL:sellRisingMAShortTerm-Risky')
        
    if(len(regression_data['filter']) > 9
        and -4 < regression_data['PCT_day_change'] < -2
        and -4 < regression_data['PCT_change'] < -2
        and low_tail_pct(regression_data) < 1.3
        and is_sell_from_filter_all_filter_relaxed(regression_data)
        ):
        if(regression_data['PCT_day_change_pre1'] > 1
           and regression_data['PCT_day_change_pre2'] < -1
           ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Sell-Relaxed-01')
        elif(-1 < regression_data['PCT_day_change_pre1'] < 1
            and regression_data['PCT_day_change_pre2'] > 1
            and regression_data['PCT_day_change_pre3'] > 1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Sell-Relaxed-02')
        elif(regression_data['PCT_day_change_pre1'] < -1
            and regression_data['PCT_day_change_pre2'] < -1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None)  
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None)

def sell_other_indicator(regression_data, regressionResult, reg, ws):
    tail_pct_filter(regression_data, regressionResult)
    base_line(regression_data, regressionResult, reg, ws)
    filterMA(regression_data, regressionResult)
    tail_reversal_filter(regression_data, regressionResult)
    if(regression_data['close'] > 50):
        sell_up_trend(regression_data, regressionResult, reg, ws)
        sell_down_trend(regression_data, regressionResult, reg, ws)
        sell_final(regression_data, regressionResult, reg, ws, ws)
        sell_pattern(regression_data, regressionResult, reg, ws, ws)
        sell_base_line_buy(regression_data, regressionResult, reg, ws)
        sell_morning_star_buy(regression_data, regressionResult, reg, ws)
        sell_evening_star_sell(regression_data, regressionResult, reg, ws)
        sell_day_high(regression_data, regressionResult, reg, ws)
        sell_trend_reversal(regression_data, regressionResult, reg, ws)
        sell_trend_break(regression_data, regressionResult, reg, ws)
        sell_consolidation_breakdown(regression_data, regressionResult, reg, ws)
        sell_final_candidate(regression_data, regressionResult, reg, ws)
        sell_oi(regression_data, regressionResult, reg, ws)
        sell_downingMA(regression_data, regressionResult, reg, ws)
        sell_study_downingMA(regression_data, regressionResult, reg, ws)
        sell_market_downtrend(regression_data, regressionResult, reg, ws)
        sell_supertrend(regression_data, regressionResult, reg, ws)
        sell_heavy_downtrend(regression_data, regressionResult, reg, ws)
        sell_check_chart(regression_data, regressionResult, reg, ws)
        sell_random_filter(regression_data, regressionResult, reg, ws)
        sell_tail_reversal_filter(regression_data, regressionResult, reg, ws)
        
        buy_year_high(regression_data, regressionResult, reg, ws)
        buy_year_low(regression_data, regressionResult, reg, ws, ws)
        buy_down_trend(regression_data, regressionResult, reg, ws)
        buy_final(regression_data, regressionResult, reg, ws, ws)
        buy_pattern(regression_data, regressionResult, reg, ws, ws)
        buy_morning_star_buy(regression_data, regressionResult, reg, ws)
        buy_evening_star_sell(regression_data, regressionResult, reg, ws)
        buy_day_low(regression_data, regressionResult, reg, ws)
        buy_trend_reversal(regression_data, regressionResult, reg, ws)
        buy_trend_break(regression_data, regressionResult, reg, ws)
        buy_consolidation_breakout(regression_data, regressionResult, reg, ws)
        buy_final_candidate(regression_data, regressionResult, reg, ws)
        buy_oi(regression_data, regressionResult, reg, ws)
        buy_up_trend(regression_data, regressionResult, reg, ws)
        buy_market_uptrend(regression_data, regressionResult, reg, ws)
        buy_check_chart(regression_data, regressionResult, reg, ws)
        buy_month3_high_continue(regression_data, regressionResult, reg, ws)
        buy_heavy_uptrend_reversal(regression_data, regressionResult, reg, ws)
        buy_supertrend(regression_data, regressionResult, reg, ws)
        buy_risingMA(regression_data, regressionResult, reg, ws)
        buy_study_risingMA(regression_data, regressionResult, reg, ws)
        buy_random_filters(regression_data, regressionResult, reg, ws)
        buy_tail_reversal_filter(regression_data, regressionResult, reg, ws)
        return True
    if(sell_skip_close_lt_50(regression_data, regressionResult, reg, ws)):
        return True
    return False

def sell_indicator_after_filter_accuracy(regression_data, regressionResult, reg, ws):
    if(is_buy_from_filter_all_filter(regression_data) != True):
        sell_af_high_indicators(regression_data, regressionResult, reg, ws)
        sell_af_down_continued(regression_data, regressionResult, reg, ws)
        sell_af_high_volatility(regression_data, regressionResult, reg, ws)
        sell_af_oi_negative(regression_data, regressionResult, reg, ws)
        sell_af_vol_contract(regression_data, regressionResult, reg, ws)
        sell_af_vol_contract_contrarian(regression_data, regressionResult, reg, ws)
        sell_af_others(regression_data, regressionResult, reg, ws)
    sell_af_high_tail(regression_data, regressionResult, reg, ws)
    #buy_af_low_tail(regression_data, regressionResult, reg, ws)
        
def sell_tail_reversal_filter(regression_data, regressionResult, reg, ws):
    if('MaySell-CheckChart' in regression_data['filter1']):
        if(regression_data['PCT_day_change_pre1'] < -2
            and regression_data['PCT_day_change_pre2'] > 0
            and is_ema14_sliding_down(regression_data)
            and (last_5_day_all_down_except_today(regression_data) != True)
            and regression_data['bar_high'] <  regression_data['bar_high_pre1']
            and regression_data['high'] <  regression_data['high_pre1']
            ):
            #add_in_csv(regression_data, regressionResult, ws, 'MaySell-CheckChart(downTrend-lastDayUp)')
            add_in_csv(regression_data, regressionResult, ws, None)
    if(('MaySell-CheckChart' in regression_data['filter1']) or ('MaySellCheckChart' in regression_data['filter1'])):
        if(1 < regression_data['PCT_day_change'] < 2
            and 1 < regression_data['PCT_change'] < 2
            and regression_data['PCT_day_change_pre1'] > 1
            and regression_data['PCT_day_change_pre2'] > 0
            and regression_data['PCT_day_change_pre3'] > 0
            and regression_data['monthHigh'] <= regression_data['high']
            and (regression_data['forecast_day_PCT7_change'] > 10 or regression_data['forecast_day_PCT10_change'] > 10)
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MaySell-CheckChart(downTrend-mayReverseLast4DaysUp)')
        elif(1 < regression_data['PCT_day_change'] < 2
            and 1 < regression_data['PCT_change'] < 2
            and regression_data['PCT_day_change_pre1'] > 1
            and regression_data['PCT_day_change_pre2'] > 0
            and regression_data['PCT_day_change_pre3'] > 0
            and regression_data['monthHigh'] <= regression_data['high']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MaySell-CheckChart(downTrend-mayReverseLast4DaysUp)-Risky')
        elif(1 < regression_data['PCT_day_change'] < 3.5
            and 1 < regression_data['PCT_change'] < 3.5
            and 1 < regression_data['PCT_day_change_pre1'] 
            and high_tail_pct_pre1(regression_data) > 1
            and regression_data['yearHighChange'] > -5
            and high_tail_pct(regression_data) < 2.5
            #and regression_data['PCT_day_change_pre3'] < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MaySell-CheckChart(yearHigh-minimumLast2DayUp)')
    elif(('MaySell-CheckChart' in regression_data['filter1']) or ('MaySellCheckChart' in regression_data['filter1'])):
        if(1 < regression_data['PCT_day_change'] < 3.5
            and 1 < regression_data['PCT_change'] < 3.5
            and 1 < regression_data['PCT_day_change_pre1'] 
            and regression_data['monthHighChange'] > -1
            and high_tail_pct(regression_data) < 2.5
            #and regression_data['PCT_day_change_pre3'] < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MaySell-CheckChart(monthHigh-minimumLast2DayUp)')
    
    if((-0.5 < regression_data['PCT_change'] < 0.5)
        and (0 < regression_data['PCT_day_change'] < 1)
        and abs(regression_data['PCT_change'] - regression_data['PCT_day_change']) < 0.8
        and low_tail_pct(regression_data) < 0.5
        and (0.6 < high_tail_pct(regression_data) < 1.5)
        and (high_tail_pct(regression_data) - low_tail_pct(regression_data)) > 0.6
        and regression_data['PCT_day_change_pre1'] < -1.8
        ):
        if(abs_month3High_less_than_month3Low(regression_data)
            ):
            add_in_csv(regression_data, regressionResult, ws, '$$MaySellCheckChart-(|_|`|)')
        else:
            add_in_csv(regression_data, regressionResult, ws, '$$MaySellCheckChart-(Risky)-(|_|`|)')
    elif(is_algo_sell(regression_data)
        and (1 <= regression_data['PCT_day_change'] < 2.5)
        and (1 <= regression_data['PCT_change'] < 4)
        ):
        if(('MaySellCheckChart' in regression_data['filter1'])
            and low_tail_pct(regression_data) < 0.8
            and (1 < regression_data['forecast_day_PCT5_change'] < 15)
            and (1 < regression_data['forecast_day_PCT7_change'] < 15)
            and (1 < regression_data['forecast_day_PCT10_change'] < 15)
            ):
            if(1.8 < high_tail_pct(regression_data) < 2.5):
                add_in_csv(regression_data, regressionResult, ws, '$$MaySellCheckChart-(IndexNotDownInSecondHalf)')
    
    if(is_algo_sell(regression_data)):        
        if(('MaySellCheckChart' in regression_data['filter1'])
            and ('Reversal' not in regression_data['filter3'])
            and regression_data['year2LowChange'] > 10
            and low_tail_pct(regression_data) < 0.5
            ):
            if(((1 < regression_data['PCT_day_change'] <= 2) and (0 < regression_data['PCT_change'] <= 2))
                and 3 > high_tail_pct(regression_data) > 1.8
                and regression_data['PCT_day_change_pre1'] > 0
                and regression_data['PCT_day_change_pre2'] > 0
                and regression_data['PCT_day_change_pre3'] > 0
                ):
                add_in_csv(regression_data, regressionResult, ws, "$$ML:MaySellCheckChart-last3DayUp")
            elif(((1 < regression_data['PCT_day_change'] <= 2) and (0 < regression_data['PCT_change'] <= 2))
                and 3 > high_tail_pct(regression_data) > 1.8):
                add_in_csv(regression_data, regressionResult, ws, "$$ML:(check-chart-2-3MidCapCross)MaySellCheckChart-Risky")
    elif(('MaySellCheckChart' in regression_data['filter1'])
        and ('Reversal' not in regression_data['filter3'])
        and regression_data['year2LowChange'] > 10
        and low_tail_pct(regression_data) < 0.5
        ):
        if((3 < regression_data['PCT_day_change'] <= 5) and (2 < regression_data['PCT_change'] <= 6)
            and 5 > high_tail_pct(regression_data) > 2.8
            ):
            add_in_csv(regression_data, regressionResult, ws, "$$ML:(Test)MaySellCheckChart-PCTDayChangeGT(3)BigLowTail-Risky")
    if(regression_data['year2HighChange'] < -3
        and regression_data['high'] == regression_data['weekHigh'] 
        and high_tail_pct(regression_data) > 1.5
        and (regression_data['PCT_day_change'] > -2 or is_algo_sell(regression_data))
        ):
        if(regression_data['weekHigh'] > regression_data['week2High']):
            add_in_csv(regression_data, regressionResult, ws, "weekHighGTweek2")
        if(regression_data['monthHighChange'] > -3
            and regression_data['monthHigh'] != regression_data['weekHigh']
            and regression_data['monthHigh'] == regression_data['month2High']
            #and (is_algo_sell(regression_data)
                #or ((regression_data['bar_high'] - regression_data['month2BarHigh'])/regression_data['month2BarHigh'])*100 > -1.5
                #or regression_data['PCT_day_change'] < 0
                #)
            ):
            if(regression_data['monthHighChange'] > 5):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT5")
            elif(regression_data['monthHighChange'] > 4):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT4")
            elif(regression_data['monthHighChange'] > 3):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT3")
            elif(regression_data['monthHighChange'] > 2):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT2")
            elif(regression_data['monthHighChange'] > 1):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT1")
            elif(regression_data['monthHighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT0")
            elif(regression_data['monthHighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-1")
            elif(regression_data['monthHighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-2")
            elif(regression_data['monthHighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-3")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-monthHighReversal")
        elif(regression_data['monthHighChange'] > -3
            and regression_data['monthHigh'] != regression_data['weekHigh']
            ):
            if(regression_data['monthHighChange'] > 5):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT5")
            elif(regression_data['monthHighChange'] > 4):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT4")
            elif(regression_data['monthHighChange'] > 3):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT3")
            elif(regression_data['monthHighChange'] > 2):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT2")
            elif(regression_data['monthHighChange'] > 1):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT1")
            elif(regression_data['monthHighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT0")
            elif(regression_data['monthHighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-1")
            elif(regression_data['monthHighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-2")
            elif(regression_data['monthHighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-3")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-monthHighReversal-risky")
        elif((regression_data['monthHighChange'] > -5 and regression_data['SMA9'] < 0)
            and regression_data['monthHigh'] != regression_data['weekHigh']
            and regression_data['monthHigh'] == regression_data['month2High']
            ):
            if(regression_data['monthHighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT0")
            elif(regression_data['monthHighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-1")
            elif(regression_data['monthHighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-2")
            elif(regression_data['monthHighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-3")
            elif(regression_data['monthHighChange'] > -4):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-4")
            elif(regression_data['monthHighChange'] > -5):
                add_in_csv(regression_data, regressionResult, ws, "monthHighChangeGT-5")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month2HighReversal-SMA9LT0")
        elif(regression_data['month2HighChange'] > -3
            and regression_data['month2High'] != regression_data['weekHigh']
            and regression_data['month2High'] == regression_data['month3High']
            #and (is_algo_sell(regression_data)
                #or ((regression_data['bar_high'] - regression_data['month3BarHigh'])/regression_data['month3BarHigh'])*100 > -1.5
                #or regression_data['PCT_day_change'] < 0
                #)
            ):
            if(regression_data['month2HighChange'] > 5):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT5")
            elif(regression_data['month2HighChange'] > 4):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT4")
            elif(regression_data['month2HighChange'] > 3):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT3")
            elif(regression_data['month2HighChange'] > 2):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT2")
            elif(regression_data['month2HighChange'] > 1):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT1")
            elif(regression_data['month2HighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT0")
            elif(regression_data['month2HighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-1")
            elif(regression_data['month2HighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-2")
            elif(regression_data['month2HighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-3")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month2HighReversal")
        elif(regression_data['month2HighChange'] > -3
            and regression_data['month2High'] != regression_data['weekHigh']
            ):
            if(regression_data['month2HighChange'] > 5):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT5")
            elif(regression_data['month2HighChange'] > 4):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT4")
            elif(regression_data['month2HighChange'] > 3):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT3")
            elif(regression_data['month2HighChange'] > 2):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT2")
            elif(regression_data['month2HighChange'] > 1):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT1")
            elif(regression_data['month2HighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT0")
            elif(regression_data['month2HighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-1")
            elif(regression_data['month2HighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-2")
            elif(regression_data['month2HighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-3")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month2HighReversal-risky")
        elif((regression_data['month2HighChange'] > -5 and regression_data['SMA9'] < 0)
            and regression_data['month2High'] != regression_data['weekHigh']
            and regression_data['month2High'] == regression_data['month3High']
            ):
            if(regression_data['month2HighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT0")
            elif(regression_data['month2HighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-1")
            elif(regression_data['month2HighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-2")
            elif(regression_data['month2HighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-3")
            elif(regression_data['month2HighChange'] > -4):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-4")
            elif(regression_data['month2HighChange'] > -5):
                add_in_csv(regression_data, regressionResult, ws, "month2HighChangeGT-5")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month2HighReversal-SMA9LT0")
        elif(regression_data['month3HighChange'] > -3
            and regression_data['month3High'] != regression_data['weekHigh']
            #and regression_data['month3High'] != regression_data['high_month3'] 
            and regression_data['month6High'] == regression_data['yearHigh'] 
            ):
            if(regression_data['month3HighChange'] > 5):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT5")
            elif(regression_data['month3HighChange'] > 4):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT4")
            elif(regression_data['month3HighChange'] > 3):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT3")
            elif(regression_data['month3HighChange'] > 2):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT2")
            elif(regression_data['month3HighChange'] > 1):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT1")
            elif(regression_data['month3HighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT0")
            elif(regression_data['month3HighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT-1")
            elif(regression_data['month3HighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT-2")
            elif(regression_data['month3HighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "month3HighChangeGT-3")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month3HighReversal")
        elif(regression_data['month3HighChange'] > 0
            and regression_data['month3High'] != regression_data['weekHigh']
            and regression_data['month3High'] == regression_data['month6High'] 
            ):
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month3HighBreakReversal")
        elif(regression_data['month6HighChange'] > -3
            and regression_data['month6High'] != regression_data['weekHigh']
            #and regression_data['month6High'] != regression_data['high_month6']
            and regression_data['yearHigh'] == regression_data['year2High']
            ):
            if(regression_data['month6HighChange'] > 5):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT5")
            elif(regression_data['month6HighChange'] > 4):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT4")
            elif(regression_data['month6HighChange'] > 3):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT3")
            elif(regression_data['month6HighChange'] > 2):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT2")
            elif(regression_data['month6HighChange'] > 1):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT1")
            elif(regression_data['month6HighChange'] > 0):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT0")
            elif(regression_data['month6HighChange'] > -1):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT-1")
            elif(regression_data['month6HighChange'] > -2):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT-2")
            elif(regression_data['month6HighChange'] > -3):
                add_in_csv(regression_data, regressionResult, ws, "month6HighChangeGT-3")
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month6HighReversal")
        elif(regression_data['month6HighChange'] > 0
            and regression_data['month6High'] != regression_data['weekHigh']
            and regression_data['month6High'] == regression_data['yearHigh'] 
            ):
            add_in_csv(regression_data, regressionResult, ws, "(Test)MaySellCheckChart-month6HighBreakReversal")
    
def sell_year_high(regression_data, regressionResult, reg, ws, ws1):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(-10 < regression_data['yearHighChange'] < -1 and regression_data['yearLowChange'] > 30 and -5 < regression_data['PCT_day_change'] < -0.75 
        and ten_days_more_than_ten(regression_data) and regression_data['forecast_day_PCT7_change'] > 5 and regression_data['forecast_day_PCT5_change'] > -0.5 and regression_data['forecast_day_PCT4_change'] > -0.5
        and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0
        and float(regression_data['forecast_day_VOL_change']) > 0
        ):
        add_in_csv(regression_data, regressionResult, ws, 'sellYearHigh')
        return True
    elif(-10 < regression_data['yearHighChange'] < 0 and regression_data['yearLowChange'] > 30 
        and (regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0.5)
        and regression_data['forecast_day_PCT10_change'] > 5 and regression_data['forecast_day_PCT7_change'] > 3 and regression_data['forecast_day_PCT5_change'] > -0.5
        and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0
        and float(regression_data['forecast_day_VOL_change']) > 0
        ):
        add_in_csv(regression_data, regressionResult, ws1, 'sellYearHigh1')
        return True
    elif(-10 < regression_data['yearHighChange'] < 0 and regression_data['yearLowChange'] > 30 
        and (regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0.5)
        and regression_data['forecast_day_PCT10_change'] > 0 and regression_data['forecast_day_PCT7_change'] > 0 and regression_data['forecast_day_PCT5_change'] > 0
        and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0
        and float(regression_data['forecast_day_VOL_change']) > 0
        ):
        add_in_csv(regression_data, regressionResult, ws1, 'sellYearHigh1')
        return True   
    return False

def sell_year_low(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(0 < regression_data['yearLowChange'] < 2 and regression_data['yearHighChange'] < -30 
        and -2 < regression_data['PCT_day_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0
        and (str(regression_data['score']) != '1-1' or str(regression_data['score']) != '10')
        and all_day_pct_change_negative(regression_data) and no_doji_or_spinning_sell_india(regression_data)
        and float(regression_data['forecast_day_VOL_change']) > 30
        and regression_data['PCT_day_change_pre1'] < 0.5
        ):
        add_in_csv(regression_data, regressionResult, ws, 'sellYearLow')
        return True
    return False

def sell_up_trend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(all_day_pct_change_positive_except_today(regression_data)
       and regression_data['forecast_day_PCT_change'] < 0 
       and -4 < regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0
       and regression_data['yearLowChange'] > 30
       and low_tail_pct(regression_data) < 0.5
       ):
        if (last_5_day_all_up_except_today(regression_data)
            and ten_days_more_than_ten(regression_data)
            and regression_data['forecast_day_PCT7_change'] > 5
            and regression_data['forecast_day_PCT10_change'] > 10
            ):
            add_in_csv(regression_data, regressionResult, ws, 'sellUpTrend-0(Risky)')
            return True
        elif(last_5_day_all_up_except_today(regression_data)
            and ten_days_more_than_seven(regression_data)
            and regression_data['forecast_day_PCT7_change'] > 4
            and regression_data['forecast_day_PCT10_change'] > 4
            ):
            add_in_csv(regression_data, regressionResult, ws, 'sellUpTrend-1')
            return True
        elif(regression_data['forecast_day_PCT7_change'] > 4
            and regression_data['forecast_day_PCT10_change'] > 4
            ):
            add_in_csv(regression_data, regressionResult, ws, 'sellUpTrend-2(Risky)')
            return True
        
    if(ten_days_more_than_ten(regression_data)
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and regression_data['forecast_day_PCT5_change'] > 0
        and regression_data['forecast_day_PCT7_change'] > 0
        and regression_data['forecast_day_PCT7_change'] > 0
        and -5 < regression_data['PCT_day_change'] < -2
        and -5 < regression_data['PCT_change'] < -2
        and regression_data['PCT_day_change_pre2'] > 0
        and regression_data['PCT_day_change_pre3'] > 0
        ):
        add_in_csv(regression_data, regressionResult, ws, '##sellTenDaysMoreThanTen')
            
    return False

def sell_down_trend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['forecast_day_PCT10_change'] < -10
        and regression_data['year2HighChange'] < -60
        and regression_data['month3LowChange'] < 10
        and (regression_data['forecast_day_PCT_change'] >= -1.5)
        and 3 < regression_data['PCT_day_change'] < 7
        and 2 < regression_data['PCT_change'] < 8
        and (regression_data['PCT_day_change_pre1'] < -4 or regression_data['PCT_day_change_pre2'] < -4)
        ):
        add_in_csv(regression_data, regressionResult, ws, 'sellYear2LowContinue')
        return True
    elif((regression_data['yearHighChange'] < -20 and regression_data['month3HighChange'] < -15)
       and (regression_data['yearLowChange'] > 15 or regression_data['month3LowChange'] > 10)
       and regression_data['forecast_day_PCT_change'] < 0
       and regression_data['forecast_day_PCT2_change'] < 0
       and regression_data['forecast_day_PCT3_change'] < 0
       and (regression_data['PCT_day_change_pre1'] > 0 or regression_data['PCT_day_change_pre2'] > 0)
       ):
        if(abs_yearHigh_less_than_yearLow(regression_data)
           and -4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] < -2
           and regression_data['PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['series_trend'] != 'upTrend'
           and str(regression_data['score']) != '10' 
           and ten_days_more_than_ten(regression_data)
           ):
            add_in_csv(regression_data, regressionResult, None, 'sellDownTrend-0')
            return True
        elif(abs_yearHigh_less_than_yearLow(regression_data)
           and -4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] < -2
           and regression_data['PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['series_trend'] != 'upTrend'
           and str(regression_data['score']) != '10' 
           and (regression_data['forecast_day_PCT7_change'] > 0 and regression_data['forecast_day_PCT10_change'] > 0)
           ):
            add_in_csv(regression_data, regressionResult, None, 'sellDownTrend-00')
            return True
        elif(abs_yearHigh_less_than_yearLow(regression_data)
           and regression_data['yearLowChange'] > 10 
           and -4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] < -2
           and regression_data['forecast_day_PCT_change'] > regression_data['PCT_change']
           and regression_data['PCT_day_change_pre1'] > 0
           and regression_data['series_trend'] != 'upTrend'
           and str(regression_data['score']) != '10'
           and regression_data['forecast_day_PCT5_change'] > -3
           and (regression_data['forecast_day_PCT7_change'] > -5 and regression_data['forecast_day_PCT10_change'] > -7) 
           ):
            add_in_csv(regression_data, regressionResult, None, '##Test:longDownTrend-0-IndexNotDownLastDay(checkBase)')
            return True
        elif(abs_yearHigh_less_than_yearLow(regression_data)
           and regression_data['yearLowChange'] > 10 
           and -4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] 
           and regression_data['forecast_day_PCT_change'] > regression_data['PCT_change']
           and regression_data['series_trend'] != 'upTrend'
           and str(regression_data['score']) != '10'
           and (regression_data['forecast_day_PCT7_change'] > -5 and regression_data['forecast_day_PCT10_change'] > -7) 
           ):
            add_in_csv(regression_data, regressionResult, None, '##Test:longDownTrend-1-IndexNotDownLastDay(checkBase)')
            return True
        elif(-4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] 
           and regression_data['forecast_day_PCT_change'] > regression_data['PCT_change']
           and regression_data['series_trend'] != 'upTrend'
           and str(regression_data['score']) != '10'
           and (regression_data['forecast_day_PCT7_change'] > -5 and regression_data['forecast_day_PCT10_change'] > -7) 
           ):
            add_in_csv(regression_data, regressionResult, None, '##Test:longDownTrend-2-IndexNotDownLastDay(checkBase)')
            return True
        elif(-4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] 
           and regression_data['forecast_day_PCT_change'] > regression_data['PCT_change']
           and regression_data['series_trend'] != 'upTrend'
           and str(regression_data['score']) != '10'
           and (regression_data['forecast_day_PCT7_change'] > -10 and regression_data['forecast_day_PCT10_change'] > -10) 
           ):
            #add_in_csv(regression_data, regressionResult, None, 'sellDownTrend-2-onlyNews')
            return False
    elif(all_day_pct_change_negative(regression_data) and -4 < regression_data['PCT_day_change'] < -2 and regression_data['yearLowChange'] > 30
        and regression_data['forecast_day_PCT10_change'] <= regression_data['PCT_change'] - 2
        and regression_data['forecast_day_PCT10_change'] <= regression_data['PCT_day_change'] - 2
        and float(regression_data['forecast_day_VOL_change']) > 30
        and regression_data['PCT_day_change_pre1'] < 0.5
        and float(regression_data['contract']) > 10
        and no_doji_or_spinning_sell_india(regression_data)):
        add_in_csv(regression_data, regressionResult, ws, '##Test:longDownTrend-Risky-IndexNotDownLastDay(checkBase)')
        return True
    return False

def sell_final(regression_data, regressionResult, reg, ws, ws1):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['yearLowChange'] > 10
       and regression_data['yearHighChange'] < -5
       and regression_data['month3HighChange'] < -3
       and -4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1
       and regression_data['forecast_day_VOL_change'] > 0
       and abs(regression_data['PCT_day_LC']) < 0.3
       and low_tail_pct(regression_data) < 1
       and str(regression_data['score']) != '10'
       ):
        if(str(regression_data['buyIndia']) == '' and -90 < regression_data['yearHighChange'] < -10
            and (ten_days_more_than_ten(regression_data)
                 or (last_5_day_all_up_except_today(regression_data)
                     and ten_days_more_than_seven(regression_data)
                    )
            ) 
            and regression_data['forecast_day_PCT7_change'] > 5 
            and regression_data['forecast_day_PCT5_change'] > -0.5 
            and regression_data['forecast_day_PCT4_change'] > -0.5
            and regression_data['forecast_day_PCT2_change'] < 0 
            and regression_data['forecast_day_PCT_change'] < 0):
            add_in_csv(regression_data, regressionResult, ws, 'sellFinal')
            return True
        elif(str(regression_data['buyIndia']) == '' and -90 < regression_data['yearHighChange'] < -10
            and (ten_days_more_than_ten(regression_data)
                 or (last_5_day_all_up_except_today(regression_data)
                     and ten_days_more_than_seven(regression_data)
                    )
            ) 
            and regression_data['forecast_day_PCT7_change'] >= 1 
            and regression_data['forecast_day_PCT5_change'] >= -1
            and regression_data['forecast_day_PCT2_change'] < 0 
            and regression_data['forecast_day_PCT_change'] < 0
            and regression_data['trend'] != 'down'
            ):
            add_in_csv(regression_data, regressionResult, ws1, 'sellFinal1')
            return True
    elif(regression_data['forecast_day_PCT4_change'] >= -0.5
        and regression_data['forecast_day_PCT5_change'] >= -0.5
        and regression_data['forecast_day_PCT7_change'] >= -0.5
        and regression_data['yearLowChange'] > 5 and regression_data['yearHighChange'] < -5
        and regression_data['month3HighChange'] < -3
        and (ten_days_more_than_ten(regression_data)
             or (last_5_day_all_up_except_today(regression_data)
                 and ten_days_more_than_seven(regression_data)
                )
             )
        and low_tail_pct(regression_data) < 1
        and regression_data['SMA50'] > 0
        ):  
        if(regression_data['forecast_day_PCT_change'] < 0
            and regression_data['bar_low'] < regression_data['bar_low_pre1']
            and regression_data['forecast_day_VOL_change'] > 0
            ):
            if(-4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['PCT_change'] < -2
               and regression_data['PCT_day_change_pre1'] > 0
                ):
                #add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-0')
                add_in_csv(regression_data, regressionResult, ws, None)
                return True
            elif(-5 < regression_data['PCT_day_change'] < -2 and -5 < regression_data['PCT_change'] < -2
               and regression_data['PCT_day_change_pre1'] > 0
               and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-00')
                return True
            elif(-6.5 < regression_data['PCT_day_change'] < -2 and -6.5 < regression_data['PCT_change'] < -2
               and regression_data['PCT_day_change_pre1'] > 0
               and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-00HighChange')
                return True
            elif(-2.5 < regression_data['PCT_day_change'] < -1 and -2.5 < regression_data['PCT_change'] < -1
                and regression_data['PCT_day_change_pre1'] < 0
                and (mlpValue <= -1 or kNeighboursValue <= -1)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-1')
                return True
            elif(-4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data)
                and regression_data['PCT_day_change_pre1'] < 0
                and (mlpValue <= -1 or kNeighboursValue <= -1)
                ):   
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-2')
                return True
            elif(-2.5 < regression_data['PCT_day_change'] < -0.5 and -2.5 < regression_data['PCT_change'] < -0.5
                and regression_data['PCT_day_change_pre1'] < 0
                and (mlpValue <= -1 or kNeighboursValue <= -1)
                ):
                add_in_csv(regression_data, regressionResult, ws, '##sellFinalCandidate-2-test')
                return True
        elif((((regression_data['open'] - regression_data['close']) * 1.5 > regression_data['high'] - regression_data['low']) or (regression_data['forecast_day_PCT_change'] < 0 and regression_data['PCT_day_change'] < -1))
            and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
            ):
            if(-2.5 < regression_data['PCT_day_change'] < -1 and -2.5 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] > -20):
                    if(((mlpValue <= -0.5 and kNeighboursValue <= -0.5) or is_algo_sell(regression_data))
                        ):    
                        add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalCandidate-3')
                        return True
                    return False
            elif(-5 < regression_data['PCT_day_change'] < -1 and -5 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] > 0):
                    if(((mlpValue <= -0.5 and kNeighboursValue <= -0.5) or is_algo_sell(regression_data))
                        ):    
                        add_in_csv(regression_data, regressionResult, ws, '##ML:sellFinalCandidate-4')
                        return True
                    return False
#         if(-1 < regression_data['PCT_day_change'] < 0 and -1 < regression_data['PCT_change'] < 0 
#             and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
#             and (mlpValue <= -1 or kNeighboursValue <= -1)
#             ):   
#             add_in_csv(regression_data, regressionResult, ws, '##sellFinalCandidate-5-(upLastDayOrDown2to3)')
#             return True
    return False

def sell_af_high_indicators(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_cla, kNeighboursValue_cla = get_reg_or_cla(regression_data, False)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if(regression_data['PCT_day_change'] > -5 and regression_data['PCT_change'] > -5 and abs(regression_data['PCT_day_change']) > 1.5 or (high_tail_pct(regression_data) > 1.5 or low_tail_pct(regression_data) < 1)):
        if(is_algo_sell(regression_data, True)
            and ((mlpValue_other <= -1 or kNeighboursValue_other <= -1) or regression_data['PCT_day_change'] > 0)
            and (abs(regression_data['filter_345_avg']) == 0
                 or (float(regression_data['filter_345_pct']) < 1.5 
                      and abs(regression_data['filter_345_avg']) < 50)
                )
            and (abs(regression_data['filter_all_avg']) == 0
                 or (float(regression_data['filter_all_pct']) < 1.5
                     and abs(regression_data['filter_all_avg']) < 50)
                )
            and (abs(regression_data['filter_pct_change_avg']) == 0
                 or (float(regression_data['filter_pct_change_pct']) < 1.5
                     and abs(regression_data['filter_pct_change_avg']) < 50)
                )
            and (abs(regression_data['filter_avg']) == 0
                 or (float(regression_data['filter_pct']) < 1.5
                     and abs(regression_data['filter_avg']) < 50)
                )
            ):
            if(2 < low_tail_pct(regression_data) < 4
                and (regression_data['PCT_day_change'] - low_tail_pct(regression_data)) < -6.5):
                return False
            if(mlpValue < -3 and mlpValue_other < -3
                and regression_data['PCT_day_change'] < -5
                and regression_data['PCT_change'] < -5
                ):
                if(-35 < regression_data['forecast_day_PCT10_change'] < -10):
                    add_in_csv(regression_data, regressionResult, ws, 'sellHighMLP-StockDowntrend(Risky)')
                elif(regression_data['forecast_day_PCT10_change'] > -10
                    and -10 < regression_data['PCT_day_change'] < -5
                    and -10 < regression_data['PCT_day_change'] < -5
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellHighMLP(Risky)')
            if(mlpValue <= -2.0 and kNeighboursValue <= -2.0 
               and regression_data['yearHighChange'] < -10 
               and regression_data['yearLowChange'] > 10
               and (low_tail_pct(regression_data) < 1.5 and (low_tail_pct(regression_data) < high_tail_pct(regression_data)))
               ):
                if(-4 < regression_data['PCT_day_change'] < -1.5 and -4 < regression_data['PCT_change'] < 0.5
                   and regression_data['forecast_day_PCT_change'] < 0
                   and low_tail_pct(regression_data) < 1
                   ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellHighIndicators')
                    return True
                if(-1 < regression_data['PCT_day_change'] < 0.5 and -2.5 < regression_data['PCT_change'] < 0.5):
                    #add_in_csv(regression_data, regressionResult, ws, '(longUpTrend)sellHighIndicators')
                    add_in_csv(regression_data, regressionResult, ws, None)
                    return True
            if(is_algo_sell(regression_data)
                and (regression_data['weekLowChange'] > 15) 
                and (-10 < regression_data['PCT_change'] < -4)
                and (-10 < regression_data['PCT_day_change'] < -4)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**(HighBothSellAll)sellUptrendReversal')
                
            if(is_algo_sell(regression_data)
                and (mlpValue_other <= 0 and kNeighboursValue_other <= 0)
                and ((-10 < mlpValue <= -4.0 and -10 < kNeighboursValue <= -2.0) or (-10 < mlpValue <= -2.0 and -10 < kNeighboursValue <= -4.0))
                ):
                if(mlpValue_cla < 0 or kNeighboursValue_cla < 0):
                    add_in_csv(regression_data, regressionResult, ws, '**sellHighIndicators-0')
            elif(is_algo_sell(regression_data)
                and (mlpValue_other <= 0 and kNeighboursValue_other <= 0)
                #and (mlpValue_cla < 0 or kNeighboursValue_cla < 0)
                and (-10 < mlpValue <= -2.0 and -10 < kNeighboursValue <= -1.5)
                and ((mlpValue + kNeighboursValue) < -5)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**sellHighIndicators-Risky')
            elif(is_algo_sell(regression_data)
                and (mlpValue_other <= 0 and kNeighboursValue_other <= 0)
                and (-5 <= mlpValue_cla < -1)
                and (-10 < mlpValue <= 0 and -10 < kNeighboursValue <= 0)
                and (mlpValue < -1 or kNeighboursValue < -1)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**(downtrend)sellHighMLPClaIndicators')
            elif((-15 < mlpValue < -5 and mlpValue_other < -2)
                and (-15 <= kNeighboursValue < 1)
                ):
                add_in_csv(regression_data, regressionResult, ws, '**(Test)sellHighMLPClaIndicators-Risky')
        elif((-15 < mlpValue < -5 and mlpValue_other < -2)
            and (-15 <= kNeighboursValue < 1)
            ):
            add_in_csv(regression_data, regressionResult, ws, '**(Test)sellHighMLPClaIndicators-Risky')        
    elif(regression_data['PCT_day_change'] > -5 and regression_data['PCT_change'] > -5
        and (-15 < mlpValue < -5 and mlpValue_other < -2)
        and (-15 <= kNeighboursValue < 1)
        ):
        add_in_csv(regression_data, regressionResult, ws, '**(Test)sellHighMLPClaIndicators-Risky')
    return False

def sell_af_high_tail(regression_data, regressionResult, reg, ws):
    if(low_tail_pct(regression_data) <= 1 and 1.3 <= high_tail_pct(regression_data) <= 2
       and high_tail_pct(regression_data) > (low_tail_pct(regression_data) + 0.5)
       ):
        if(0.5 < regression_data['PCT_day_change'] < 3 and 0.5 < regression_data['PCT_change'] < 3
           and (1 < regression_data['PCT_day_change'] or 1 < regression_data['PCT_change'])
           ):
           add_in_csv(regression_data, regressionResult, ws, '%%maySellTail')
        elif(-1 < regression_data['PCT_day_change'] < 3 and -1 < regression_data['PCT_change'] < 3 and 1.4 <= high_tail_pct(regression_data)):
           add_in_csv(regression_data, regressionResult, ws, '%%maySellTail-Risky')
    elif(low_tail_pct(regression_data) <= 1 and 2 <= high_tail_pct(regression_data) <= 4):
        if(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4):
           add_in_csv(regression_data, regressionResult, ws, '%%maySellTail')
               
def sell_af_down_continued(regression_data, regressionResult, reg, ws):
    if(high_tail_pct(regression_data) < 2 and low_tail_pct(regression_data) < 1.1):
        if(-3 < regression_data['PCT_day_change'] < -1.9 and -3 < regression_data['PCT_change'] < -1):
            if(regression_data['PCT_day_change_pre1'] > -0.75 and regression_data['PCT_change_pre1'] > -1):
                add_in_csv(regression_data, regressionResult, ws, '%%maySellDownContinueGT-3')
            else:
                add_in_csv(regression_data, regressionResult, ws, '%%maySellDownContinueGT-3-Risky')
        elif(-4 < regression_data['PCT_day_change'] < -3.25 and -4 < regression_data['PCT_change'] < -3):
            if(regression_data['PCT_day_change_pre1'] > -0.75 and regression_data['PCT_change_pre1'] > -1):
                add_in_csv(regression_data, regressionResult, ws, '%%maySellDownContinueLT-3')
            else:
                add_in_csv(regression_data, regressionResult, ws, '%%maySellDownContinueLT-3-Risky')

def sell_af_high_volatility(regression_data, regressionResult, reg, ws):
    if(high_tail_pct(regression_data) < 2 and low_tail_pct(regression_data) < 1.5):
        if(regression_data['PCT_day_change_pre1'] > -5
            and regression_data['PCT_change_pre1'] > -5
            and -20 < regression_data['PCT_day_change'] < -8
            and -20 < regression_data['PCT_change'] < -5
            and regression_data['forecast_day_PCT4_change'] > -20
            ):
            if(regression_data['PCT_day_change_pre1'] > -3 
                and regression_data['PCT_change_pre1'] > -3
                and -15 < regression_data['PCT_day_change']
                and -15 < regression_data['PCT_change']
                ):
                add_in_csv(regression_data, regressionResult, ws, None, '%%maySellHighVolatileDownContinueLT-8')
                return True
            else:
                add_in_csv(regression_data, regressionResult, ws, None, '%%maySellAfter10:30HighVolatileDownContinueLT-8-Risky')
                return True
    if(regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0
        and regression_data['forecast_day_PCT7_change'] > 0
        and regression_data['forecast_day_PCT10_change'] > 0
        and (regression_data['PCT_day_change'] > 15 and regression_data['PCT_change'] > 10
            or (regression_data['PCT_day_change_pre1'] > 5 and regression_data['forecast_day_PCT2_change'] > 20)
            )
        ):
        add_in_csv(regression_data, regressionResult, ws, None, '%%maySellAfter10:20HighVolatileLastDayUp-GT10')
        return True
    elif(
        regression_data['forecast_day_PCT7_change'] < 0
        and regression_data['forecast_day_PCT10_change'] < 0
        and 8 < regression_data['PCT_day_change'] < 20 and 8 < regression_data['PCT_change'] < 20
        ):
        add_in_csv(regression_data, regressionResult, ws, None, '%%mayBuyContinueHighVolatileLastDayUp-GT8')
        return True
    elif('%%mayBuyTail' in regression_data['filter']
        and (('downTrend' in regression_data['series_trend'])
             or pct_day_change_trend(regression_data) <= -3
            )
        ):
        countGt, countLt = pct_day_change_counter(regression_data)
        if(countGt < countLt):
            add_in_csv(regression_data, regressionResult, ws, None, '%%SellDownTrend-lowTail')
        elif(countGt == countLt
            and regression_data['PCT_day_change'] < 0
            ):
            add_in_csv(regression_data, regressionResult, ws, None, '%%BuyDownTrend-lowTail-alternateUpDown')
        return True    
    else:
        return False

def sell_af_oi_negative(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['greentrend'] == 1
        and regression_data['forecast_day_PCT_change'] > 0.5
        and regression_data['forecast_day_PCT2_change'] > 0.5
        and regression_data['forecast_day_PCT3_change'] > 0.5
        and regression_data['forecast_day_PCT4_change'] > 0.5
        and regression_data['forecast_day_PCT5_change'] > 0.5
        and regression_data['forecast_day_PCT7_change'] > 0.5
        and ten_days_more_than_ten(regression_data)
        and float(regression_data['forecast_day_VOL_change']) < -40 
        and regression_data['PCT_day_change_pre1'] > regression_data['PCT_day_change']
        and float(regression_data['contract']) < 0
        and float(regression_data['oi']) < 5
        and (regression_data['yearLowChange'] > 15 or regression_data['yearHighChange'] > -15)
        ):
        if(is_algo_sell(regression_data)):
            if(((0 < regression_data['PCT_day_change'] < 1 and 0.5 < regression_data['PCT_change'] < 1)
                or (0.5 < regression_data['PCT_day_change'] < 1 and 0 < regression_data['PCT_change'] < 1))
               ):
                add_in_csv(regression_data, regressionResult, ws, None, 'ML:sellNegativeOI-0-checkBase(1%up)')
                return True
            if(0 < regression_data['PCT_day_change'] < 2 and 0 < regression_data['PCT_change'] < 2 
                ):
                add_in_csv(regression_data, regressionResult, ws, None, 'ML:sellNegativeOI-1-checkBase(1%up)')
                return True
        return False
    return False

def sell_af_vol_contract(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if((ten_days_less_than_minus_ten(regression_data) == False
        or ten_days_less_than_minus_fifteen(regression_data) == True
        or (regression_data['forecast_day_PCT5_change'] > -5 and regression_data['forecast_day_PCT7_change'] > -5))
        #and ((mlpValue <= -0.3 and kNeighboursValue <= -0.3) or is_algo_sell(regression_data))
        and (float(regression_data['contract']) != 0 or float(regression_data['oi']) != 0)
        and float(regression_data['contract']) > 10
        and ((regression_data['PCT_day_change'] < -1 or regression_data['PCT_change'] < -1)
             or
             (regression_data['PCT_day_change'] < -0.75 and regression_data['PCT_change'] < -0.75)
             and regression_data['forecast_day_PCT_change'] > -2
             and regression_data['forecast_day_PCT2_change'] > -2
             and regression_data['forecast_day_PCT3_change'] > -2
             and regression_data['forecast_day_PCT4_change'] > -2
             and regression_data['forecast_day_PCT5_change'] > -2
             and regression_data['forecast_day_PCT7_change'] > -2
             and regression_data['forecast_day_PCT10_change'] > -2
            )
        and regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0
        and -5 < regression_data['forecast_day_PCT4_change'] < 0
        and regression_data['forecast_day_PCT5_change'] > -5
        and regression_data['forecast_day_PCT7_change'] > -5
        and regression_data['forecast_day_PCT10_change'] > -5
        and (regression_data['PCT_day_change'] > 0
            or regression_data['PCT_day_change_pre1'] > 0
            or regression_data['PCT_day_change_pre2'] > 0
            or regression_data['PCT_day_change_pre3'] > 0
            or regression_data['PCT_day_change_pre4'] > 0
            )
        and preDayPctChangeDown_orVolHigh(regression_data)
        and regression_data['open'] > 50
        and last_4_day_all_down(regression_data) == False
        and low_tail_pct(regression_data) < 1.5
        and regression_data['month3LowChange'] > 7.5
        and (regression_data['forecast_day_VOL_change'] > 150
            or (regression_data['PCT_day_change_pre2'] > 0
                and (((regression_data['volume'] - regression_data['volume_pre2'])*100)/regression_data['volume_pre2']) > 100
                and (((regression_data['volume'] - regression_data['volume_pre3'])*100)/regression_data['volume_pre3']) > 100
               )
           )
        ):
        if((regression_data['forecast_day_VOL_change'] > 70 and -2 < regression_data['PCT_day_change'] < -0.5 and -2 < regression_data['PCT_change'] < -1)
            and float(regression_data['contract']) > 10
            ):
            if((mlpValue <= -0.3 and kNeighboursValue <= -0.3) or is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, None, 'ML:oiSell-0')
                return True
            return False
        elif((regression_data['forecast_day_VOL_change'] > 35 and -2 < regression_data['PCT_day_change'] < -0.5 and -2 < regression_data['PCT_change'] < -1)
            and float(regression_data['contract']) > 20
            ):
            if((mlpValue <= -0.3 and kNeighboursValue <= -0.3) or is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, None, 'ML:oiSell-1')
                return True
            return False
        elif((regression_data['forecast_day_VOL_change'] > 150 and -3 < regression_data['PCT_day_change'] < -1 and -3 < regression_data['PCT_change'] < -1)
            ):
            if((mlpValue <= -0.3 and kNeighboursValue <= -0.3) or is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, None, 'ML:oiSell-2-checkBase')
                return True
            return False
        elif(((regression_data['forecast_day_VOL_change'] > 400 and -3.5 < regression_data['PCT_day_change'] < -1 and -3.5 < regression_data['PCT_change'] < -1)
            or (regression_data['forecast_day_VOL_change'] > 500 and -4.5 < regression_data['PCT_day_change'] < -1 and -4.5 < regression_data['PCT_change'] < -1)
            )
            ):
            if((mlpValue <= -0.3 and kNeighboursValue <= -0.3) or is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, None, 'oiSell-3-checkBase')
                return True
            return False
        elif((regression_data['forecast_day_VOL_change'] > 500 and -5 < regression_data['PCT_day_change'] < -1 and -5 < regression_data['PCT_change'] < -1)
            and float(regression_data['contract']) > 50
            and (regression_data['forecast_day_PCT10_change'] > 8 or regression_data['forecast_day_PCT7_change'] > 8)
            ):
            if((mlpValue <= -0.3 and kNeighboursValue <= -0.3) or is_algo_sell(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, None, 'oiSell-4-checkBase')
                return True
            return False
    return False

def sell_af_vol_contract_contrarian(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(
        #(float(regression_data['contract']) != 0 or float(regression_data['oi']) != 0)
        (ten_days_less_than_minus_ten(regression_data) == True and ten_days_less_than_minus_fifteen(regression_data) == False
             and (regression_data['forecast_day_PCT5_change'] < -5 and regression_data['forecast_day_PCT7_change'] < -5))
        and float(regression_data['contract']) > 10
        and (regression_data['PCT_day_change'] < -1 or regression_data['PCT_change'] < -1)
        and regression_data['forecast_day_PCT_change'] < -0.5
        and regression_data['forecast_day_PCT2_change'] < -0.5
        and regression_data['forecast_day_PCT3_change'] < 0
        and regression_data['forecast_day_PCT4_change'] < 0
        and preDayPctChangeDown_orVolHigh(regression_data)
        and regression_data['open'] > 50
        and last_7_day_all_down(regression_data) == False
        ):
        if((regression_data['forecast_day_VOL_change'] > 70 and -2 < regression_data['PCT_day_change'] < -0.75 and -2 < regression_data['PCT_change'] < -0.5)
            and float(regression_data['contract']) > 10
            ):
            if(('P@[') not in regression_data['sellIndia']):
                add_in_csv(regression_data, regressionResult, ws, None, 'Reversal(Test):buyReversalOI-0')
                return True
            return False
#         elif((regression_data['forecast_day_VOL_change'] > 35 and -2 < regression_data['PCT_day_change'] < -0.75 and -2 < regression_data['PCT_change'] < -0.5)
#             and float(regression_data['contract']) > 20
#             ):
#             if(('P@[') not in regression_data['sellIndia']):
#                 add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):buyReversalOI-1')
#                 return True
#             return False
#         elif((regression_data['forecast_day_VOL_change'] > 150 and -3 < regression_data['PCT_day_change'] < -0.75 and -3 < regression_data['PCT_change'] < -0.5)
#             ):
#             add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):buyReversalOI-2')
#             return True
#             elif(((regression_data['forecast_day_VOL_change'] > 300 and -3.5 < regression_data['PCT_day_change'] < -0.75 and -3.5 < regression_data['PCT_change'] < -0.5)
#                 or (regression_data['forecast_day_VOL_change'] > 400 and -4.5 < regression_data['PCT_day_change'] < -0.75 and -4.5 < regression_data['PCT_change'] < -0.5)
#                 )
#                 and regression_data['forecast_day_PCT10_change'] > -10
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):buyReversalOI-3-checkBase-(openAroundLastCloseAnd5MinuteChart)')
#                 return True
#             elif((regression_data['forecast_day_VOL_change'] > 500 and -5 < regression_data['PCT_day_change'] < -0.75 and -5 < regression_data['PCT_change'] < -0.5)
#                 #and float(regression_data['contract']) > 50
#                 and regression_data['forecast_day_PCT10_change'] > -10
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):buyReversalOI-4-checkBase-(openAroundLastCloseAnd5MinuteChart)')
#                 return True
        
#     if((('P@[' not in str(regression_data['buyIndia'])) and ('P@[' not in str(regression_data['sellIndia'])))
#         and (0 < regression_data['PCT_day_change'] < 1.5 and 0 < regression_data['PCT_change'] < 1.5)
#         and kNeighboursValue <= -1
#         and mlpValue <= 0
#         and is_recent_consolidation(regression_data) == False
#         and str(regression_data['score']) != '10'
#         and (low_tail_pct(regression_data) < 1.5 or (low_tail_pct(regression_data) < high_tail_pct(regression_data)))
#         ):
#             add_in_csv(regression_data, regressionResult, ws, 'Reversal(Test):sellReversalKNeighbours(upTrend)(upLastDay)')
#             return True        
    return False

def sell_af_others(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, True)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, True)
    mlpValue_cla, kNeighboursValue_cla = get_reg_or_cla(regression_data, False)
    mlpValue_other_cla, kNeighboursValue_other_cla = get_reg_or_cla_other(regression_data, False)
    
    if(len(regression_data['filter']) > 9
        and -4 < regression_data['PCT_day_change'] < -2
        and -4 < regression_data['PCT_change'] < -2
        and is_sell_from_filter_all_filter_relaxed(regression_data)
        ):
        if(regression_data['PCT_day_change_pre1'] > 1
           and regression_data['PCT_day_change_pre2'] < -1
           ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Sell-Relaxed-01')
        elif(-1 < regression_data['PCT_day_change_pre1'] < 1
            and regression_data['PCT_day_change_pre2'] > 1
            and regression_data['PCT_day_change_pre3'] > 1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Sell-Relaxed-02')
        elif(regression_data['PCT_day_change_pre1'] < -1
            and regression_data['PCT_day_change_pre2'] < -1
            ):
            add_in_csv(regression_data, regressionResult, ws, None, None)  
        else:
            add_in_csv(regression_data, regressionResult, ws, None, None)
            
    if('NA$NA:NA$(shortDownTrend)' in regression_data['series_trend']):
        add_in_csv(regression_data, regressionResult, ws, None, 'shortDownTrend')
                
    return False
    
def sell_pattern(regression_data, regressionResult, reg, ws, ws1):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    sell_pattern_without_mlalgo(regression_data, regressionResult)
    if(regression_data['yearHighChange'] < -10):
        if(is_algo_sell(regression_data)
            and regression_data['sellIndia_avg'] < -1.5 
            and low_tail_pct(regression_data) < 1.5
            ):
            add_in_csv(regression_data, regressionResult, ws, 'sellPatternsML')
            return True
    
def sell_base_line_buy(regression_data, regressionResult, reg, ws):
    return False
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if(low_tail_pct(regression_data) > 2
        and (mlpValue > 0 and kNeighboursValue > 0
             and (mlpValue_other > 0 or kNeighboursValue_other > 0)
            )
        ):
        if(-4 < regression_data['PCT_day_change'] < 0 and -4 < regression_data['PCT_change'] < 0
            ):
            if(-10 < regression_data['year2LowChange'] < 5):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYear2LowBreak-0')
                return True
            if(-10 < regression_data['yearLowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYearLowBreak-0')
                return True
            if(-10 < regression_data['month6LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth6LowBreak-0')
                return True
            if(-6.5 < regression_data['yearHighChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYearHighReversal-0-(downTrend)(checkBase)')
                return True
            if(-6.5 < regression_data['month6HighChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth6HighReversal-0-(downTrend)(checkBase)')
                return True
            if(-6.5 < regression_data['month3HighChange'] < 0
                and low_tail_pct(regression_data) > 2
                ):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth3HighReversal-0-(downTrend)(checkBase)')
                return True
    
    if(low_tail_pct(regression_data) > 2
        and (regression_data['close'] - regression_data['low']) > (regression_data['open'] - regression_data['close'])
        and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
        and (mlpValue > 0 and kNeighboursValue > 0)
        ):
        if(-4 < regression_data['PCT_day_change'] < 0 and -4 < regression_data['PCT_change'] < 0
            and -4 < regression_data['forecast_day_PCT_change'] < 0
            ):
            if(-10 < regression_data['year2LowChange'] < 5):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYear2LowBreak-1')
                return True
            if(-10 < regression_data['yearLowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYearLowBreak-1')
                return True
            if(-10 < regression_data['month6LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth6LowBreak-1')
                return True
            if(-6.5 < regression_data['yearHighChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYearHighReversal-1-(downTrend)(checkBase)')
                return True
            if(-6.5 < regression_data['month6HighChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth6HighReversal-1-(downTrend)(checkBase)')
                return True
            if(-6.5 < regression_data['month3HighChange'] < 0
                and low_tail_pct(regression_data) > 2
                ):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth3HighReversal-1-(downTrend)(checkBase)')
                return True
    
    if(('M@[,RSI]' in regression_data['buyIndia'])
        and (mlpValue > 0 and kNeighboursValue > 0)
        ):
        if(-4 < regression_data['PCT_day_change'] < 0 and -4 < regression_data['PCT_change'] < 0
            and -4 < regression_data['forecast_day_PCT_change'] < 0
            ):
            if(-5 < regression_data['year2LowChange'] < 5):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYear2LowBreak-2')
                return True
            if(-5 < regression_data['yearLowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyYearLowBreak-2')
                return True
            if(-5 < regression_data['month6LowChange'] < 0):
                add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth6LowBreak-2')
                return True
            
#     if(regression_data['forecast_day_PCT_change'] > 0
#         and regression_data['forecast_day_PCT2_change'] >= 0
#         and (mlpValue > 0 and kNeighboursValue > 0)
#         ):
#         if(-4 < regression_data['PCT_day_change'] < 0 and -4 < regression_data['PCT_change'] < 0
#             ):
#             if(-5 < regression_data['year2LowChange'] < 5):
#                 add_in_csv(regression_data, regressionResult, ws, 'Test:buyYear2LowBreak-3')
#                 return True
#             if(-5 < regression_data['yearLowChange'] < 0):
#                 add_in_csv(regression_data, regressionResult, ws, 'Test:buyYearLowBreak-3')
#                 return True
#             if(-5 < regression_data['month6LowChange'] < 0):
#                 add_in_csv(regression_data, regressionResult, ws, 'Test:buyMonth6LowBreak-3')
#                 return True
    
    return False
    
def sell_morning_star_buy(regression_data, regressionResult, reg, ws):
    return False
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(-8 < regression_data['forecast_day_PCT_change'] < -3
        and ((regression_data['yearLowChange'] > 15) or (regression_data['yearLowChange'] < 2))
        and high_tail_pct(regression_data) < 0.5
        and low_tail_pct(regression_data) > 3
        and regression_data['forecast_day_PCT10_change'] < -10
        ):
        if(-6 < regression_data['PCT_day_change'] < -2 and -6 < regression_data['PCT_change'] < -1
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
            and regression_data['forecast_day_PCT10_change'] < -10
            and regression_data['SMA25'] < -8.5
            ):
            add_in_csv(regression_data, regressionResult, ws, None)
            return True
    if(-10 < regression_data['forecast_day_PCT_change'] < -2
        and regression_data['PCT_day_change_pre1'] < 0
        and (ten_days_less_than_minus_seven(regression_data))
        and high_tail_pct(regression_data) < 0.5
        and 1.5 < low_tail_pct(regression_data) < 4
        ):
        if(-3 < regression_data['PCT_day_change'] < 0 and -3 < regression_data['PCT_change'] < 0
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
            and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
            and regression_data['forecast_day_PCT10_change'] < -10
            and regression_data['SMA25'] < -8.5
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-dayDown')
            return True
        if(0.3 < regression_data['PCT_day_change'] < 1
            and (regression_data['open'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
            and (regression_data['open'] - regression_data['low']) >= ((regression_data['high'] - regression_data['close']) * 3)
            and 2 < low_tail_pct(regression_data) < 3.5
            ):
            if(is_algo_buy(regression_data)):
                add_in_csv(regression_data, regressionResult, ws, 'ML:buyMorningStar-Risky-dayUp')
                return True
            return False
    if(-10 < regression_data['forecast_day_PCT_change'] < -2
        and regression_data['PCT_day_change_pre1'] < 0
        and (ten_days_less_than_minus_seven(regression_data))
        and high_tail_pct(regression_data) < 0.5
        and 1.5 < low_tail_pct(regression_data) < 7
        ):
        if(0.3 < regression_data['PCT_day_change'] < 1
            and (regression_data['open'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
            and (regression_data['open'] - regression_data['low']) >= ((regression_data['high'] - regression_data['close']) * 3)
            and regression_data['year2HighChange'] > -50
            ):
            add_in_csv(regression_data, regressionResult, ws, 'sellMorningStar-dayUp')
            return True
#         if(high_tail_pct(regression_data) < 1.5
#            and low_tail_pct(regression_data) > 1.5
#            ):
#             if(0 < regression_data['PCT_day_change'] < 1 and 0 < regression_data['PCT_change'] < 1 
#                 and kNeighboursValue >= 0
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##buyMorningStar-0-NotUpSecondHalfAndUp2to3')
#                 return True
#             if(-1.5 < regression_data['PCT_day_change'] < 0 and -1.5 < regression_data['PCT_change'] < 0
#                 and (regression_data['close'] - regression_data['low']) >= ((regression_data['open'] - regression_data['close']) * 3)
#                 and (regression_data['close'] - regression_data['low']) >= ((regression_data['high'] - regression_data['open']) * 3)
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##buyMorningStar-1-NotUpSecondHalfAndUp2to3')
#                 return True
    return False

def sell_evening_star_sell(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(8 > regression_data['forecast_day_PCT_change'] > 4
        and regression_data['month3HighChange'] > -50
        and low_tail_pct(regression_data) < 0.8
        and high_tail_pct(regression_data) > 2.5
        and regression_data['yearHighChange'] < -4
        ):
        if(6 > regression_data['PCT_day_change'] > 1.5 and 6 > regression_data['PCT_change'] > 1
            and (regression_data['high']-regression_data['close']) > ((regression_data['close']-regression_data['open']))
            and (regression_data['high']-regression_data['close']) >= ((regression_data['open']-regression_data['low'])*3)
            #and ((regression_data['yearHighChange'] < -15) or (regression_data['yearHighChange'] > -2))
            #and regression_data['forecast_day_PCT5_change'] < 10
            ):
            if(regression_data['year2HighChange'] < -80
            ):
                add_in_csv(regression_data, regressionResult, ws, 'buyEveningStar-0(Buy-After-1pc-down)')
                return True
            else:
                add_in_csv(regression_data, regressionResult, ws, 'sellEveningStar-0')
                return True
    elif(
        regression_data['PCT_day_change_pre1'] > 0
        and (ten_days_more_than_seven(regression_data))
        ):
        if(low_tail_pct(regression_data) < 1.5
            and high_tail_pct(regression_data) > 1.5
            ):
            if(1.5 > regression_data['PCT_day_change'] > 0.5 and 1.5 > regression_data['PCT_change'] > 0
                and (regression_data['high']-regression_data['close']) >= ((regression_data['close']-regression_data['open']) * 3)
                and (regression_data['high']-regression_data['close']) >= ((regression_data['open']-regression_data['low'])*3)
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellEveningStar-2(Check2-3MidCapCross)')
                    return True
                return False
        elif((low_tail_pct(regression_data) < 0.5 or (regression_data['forecast_day_PCT_change'] > 6 and low_tail_pct(regression_data) < 1))
            and 2 < high_tail_pct(regression_data) < 3.5
            and low_tail_pct(regression_data) < 0.5
            ):
            if(2 > regression_data['PCT_day_change'] > 0.5 and 2 > regression_data['PCT_change'] > 0
                ):
                if(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellEveningStar-Risky-3(Check2-3MidCapCross)')
                    return True
                return False
    elif(0 < regression_data['forecast_day_PCT_change'] < 8
        and low_tail_pct(regression_data) < 1
        and high_tail_pct(regression_data) > 2.5
        and regression_data['forecast_day_PCT10_change'] > (regression_data['forecast_day_PCT_change'] + 5)
        ):
        if(0 < regression_data['PCT_day_change'] < 6 and 0 < regression_data['PCT_change'] < 6
            and (regression_data['high'] - regression_data['bar_high']) >= ((regression_data['bar_high'] - regression_data['bar_low']) * 3)
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-HighUpperTail-(checkChart)')
        elif(-2 < regression_data['PCT_day_change'] < 6 and -2 < regression_data['PCT_change'] < 6
            and regression_data['PCT_day_change_pre1'] > 0
            and regression_data['PCT_day_change_pre2'] > 0
            and regression_data['PCT_day_change_pre3'] > 0
            and (regression_data['high'] - regression_data['bar_high']) >= ((regression_data['bar_high'] - regression_data['bar_low']) * 3)
            ):
            add_in_csv(regression_data, regressionResult, ws, 'buyMorningStar-(checkChart)')
            return True
    return False

def sell_day_high(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if((regression_data['PCT_day_change'] > 15 and regression_data['PCT_change'] > 10)
        or regression_data['forecast_day_PCT2_change'] > 20
        ):
        add_in_csv(regression_data, regressionResult, ws, 'maySellAfter10:20HighVolatileLastDayUp-GT10')
    elif(
        regression_data['PCT_day_change'] > 9 and regression_data['PCT_change'] > 7
        and regression_data['PCT_day_change_pre1'] < 0
        ):
        add_in_csv(regression_data, regressionResult, ws, None)
        
    if(regression_data['PCT_day_change_pre1'] > 1.5
       and regression_data['high'] > regression_data['high_pre1']
       and regression_data['bar_high'] > regression_data['bar_high_pre1']
       ):
        if((regression_data['PCT_day_change'] > 5 or regression_data['PCT_change'] > 5)
           and float(regression_data['forecast_day_VOL_change']) < -50
           and regression_data['PCT_day_change'] > regression_data['PCT_day_change_pre1'] > 0
           and regression_data['PCT_day_change_pre1'] > 2
           and regression_data['forecast_day_PCT_change'] > 5
           and regression_data['forecast_day_PCT10_change'] < -5
           ):
            add_in_csv(regression_data, regressionResult, ws, 'sellDayHighVolLow-0')
            return True
        elif((regression_data['PCT_day_change'] > 5 or regression_data['PCT_change'] > 5)
           and float(regression_data['forecast_day_VOL_change']) < -20
           and regression_data['PCT_day_change'] > regression_data['PCT_day_change_pre1'] > 0
           and regression_data['PCT_day_change_pre1'] > 1
           and regression_data['forecast_day_PCT_change'] > 4
           and regression_data['yearLowChange'] < 80
           ):
            add_in_csv(regression_data, regressionResult, ws, 'buyDayHighVolLow-0')
            return True
        elif((regression_data['PCT_day_change'] > 5 and regression_data['PCT_change'] > 4)
           and float(regression_data['forecast_day_VOL_change']) < -20
           and regression_data['PCT_day_change'] > regression_data['PCT_day_change_pre1'] > 0
           and regression_data['PCT_day_change_pre1'] > 1.5
           and regression_data['month3LowChange'] < 30
           ):
            add_in_csv(regression_data, regressionResult, ws, 'sellDayHighVolLow-01-checkMorningTrend(.5SL)')
            return True
#         elif((regression_data['PCT_day_change'] > 3 and regression_data['PCT_change'] > 3) 
#            and abs_yearHigh_more_than_yearLow(regression_data)
#            and float(regression_data['forecast_day_VOL_change']) < -50  
#            and regression_data['PCT_day_change'] > regression_data['PCT_day_change_pre1'] > 0
#            and regression_data['PCT_day_change_pre2'] < -1
#            ):
#             add_in_csv(regression_data, regressionResult, ws, 'sellDayHighVolLow-1-checkMorningTrend(.5SL)-NotYearHigh')
#             return True
#     if((regression_data['PCT_day_change'] > 2 and regression_data['PCT_change'] > 2) 
#        and float(regression_data['forecast_day_VOL_change']) < -50  
#        and regression_data['PCT_day_change_pre1'] > 0
#        ):
#         add_in_csv(regression_data, regressionResult, ws, '##sellDayHighVolLow-2')
#         return True
    return False

def sell_trend_reversal(regression_data, regressionResult, reg, ws):    
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['yearLowChange'] > 5 and regression_data['yearHighChange'] < -5
        and (ten_days_more_than_ten(regression_data)
             or (last_5_day_all_up_except_today(regression_data)
                 and ten_days_more_than_seven(regression_data)
                )
             )
        ):  
        if((((regression_data['open'] - regression_data['close']) * 1.5 > regression_data['high'] - regression_data['low']) or (regression_data['forecast_day_PCT_change'] < 0 and regression_data['PCT_day_change'] < -1))
            and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
            ):
            if(-2.5 < regression_data['PCT_day_change'] < -1 and -2.5 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] <= -30
                    and ('P@[' not in regression_data['buyIndia'])
                    and (regression_data['PCT_day_change_pre1'] < -1)
                    and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None)
                    return True
            if(-5 < regression_data['PCT_day_change'] < -1 and -5 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] <= -30
                    and (regression_data['PCT_day_change_pre1'] < -1)
                    and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
                    and regression_data['weekLowChange'] > 2
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None)
                    return False
                elif(regression_data['forecast_day_VOL_change'] <= -10
                    and (regression_data['PCT_day_change_pre1'] < -1)
                    and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
                    and regression_data['weekLowChange'] > 2
                    ):
                    add_in_csv(regression_data, regressionResult, ws, '##buyFinalContinue-2')
                    return True
    if(regression_data['yearLowChange'] > 5 and regression_data['yearHighChange'] < -5
        and ten_days_more_than_five(regression_data)
        ):
        if(regression_data['forecast_day_PCT_change'] < 0 and regression_data['PCT_day_change'] < 0 
            and (regression_data['PCT_day_change_pre1'] < -1)
            and regression_data['PCT_day_change'] < regression_data['PCT_day_change_pre1'] < 0
            and regression_data['forecast_day_VOL_change'] <= -20
            ):
            if(-5 < regression_data['PCT_day_change'] < -2 and -5 < regression_data['PCT_change'] < -2):
                if(is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalContinue-3')
                    return True 
                return False

#     if((-0.5 < regression_data['PCT_day_change'] < 0.5)
#         and (-0.5 < regression_data['PCT_change'] < 0.5)
#         and low_tail_pct(regression_data) < 0.3
#         and (0.6 < high_tail_pct(regression_data) < 0.9)
#         ):
#         add_in_csv(regression_data, regressionResult, ws, "Test:sellLowTailLessThan0.3-3(checkHillUp)")
    return False   

def sell_trend_break(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
#     if(regression_data['SMA200'] == 1
#        and regression_data['PCT_day_change'] < -2 and regression_data['PCT_change'] < -2
#        ):
#         add_in_csv(regression_data, regressionResult, ws, '##TestBreakOutSellConsolidate-0')
#         return True
    
    flag = False
    if(NIFTY_HIGH and high_tail_pct(regression_data) > 1.5
        and -1 < regression_data['PCT_change'] < 1 and -1 < regression_data['PCT_day_change'] < 1
        and regression_data['forecast_day_PCT10_change'] > -10
        #and ((mlpValue <= -0.5 and kNeighboursValue <= -0.5) or is_algo_sell(regression_data))
        ):
        add_in_csv(regression_data, regressionResult, ws, None, '##TEST:sellCheckLastHourDown')
        flag = True
    
    if(regression_data['yearLowChange'] < 5
       and 'M@[,RSI' in str(regression_data['buyIndia'])
       and 'P@[' not in str(regression_data['sellIndia'])
       and (regression_data['PCT_day_change'] < -5 and regression_data['PCT_change'] < 0)
       #and mlpValue >= 1
       ):
        add_in_csv(regression_data, regressionResult, ws, 'RSISellCandidate(yearLow)-lessThanMinus5')
        flag = True
    
    if(ten_days_more_than_five(regression_data)
        and regression_data['yearLowChange'] > 15
        and regression_data['weekLowChange'] > 5
        and -2 > regression_data['month3HighChange'] > -10
        ):
        if(regression_data['forecast_day_PCT_change'] < -1.5 and regression_data['PCT_day_change'] < 0 and regression_data['PCT_day_change_pre1'] > 0
            and abs(regression_data['PCT_day_change']) > abs(regression_data['PCT_day_change_pre1'])
            #and regression_data['open'] == regression_data['high']
            and regression_data['forecast_day_VOL_change'] >= -20
            ):
                add_in_csv(regression_data, regressionResult, ws, 'finalBreakOutSell-0')
                flag = True
    return flag

def sell_consolidation_breakdown(regression_data, regressionResult, reg, ws):
    week2BarLowChange = ((regression_data['bar_low'] - regression_data['week2BarLow'])/regression_data['bar_low'])*100
    if(-6 < regression_data['PCT_day_change'] < 0
        and -6 < regression_data['PCT_change']
        and (regression_data['PCT_day_change_pre1'] > 0 
             or regression_data['PCT_day_change_pre2'] > 0
             or regression_data['PCT_day_change_pre3'] > 0
            )
        #and regression_data['low'] < regression_data['high_pre3']
        and regression_data['low'] < regression_data['high_pre2']
        and regression_data['low'] < regression_data['high_pre1']
        and -3 < regression_data['forecast_day_PCT4_change'] < -0.5
        and -4 < regression_data['forecast_day_PCT3_change'] < 0
        and -4 < regression_data['forecast_day_PCT2_change'] < 0
        and -3 < regression_data['forecast_day_PCT_change'] < 0
        and -3 <= regression_data['forecast_day_PCT10_change'] <= 0
        and regression_data['bar_low'] <= regression_data['week2BarLow']
        ):
        if(regression_data['weekBarLow'] < regression_data['bar_low_pre1'] 
            and regression_data['week2BarLow'] < regression_data['bar_low_pre1']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'brokenToday')
        if(regression_data['weekLow'] < regression_data['week2Low']):
            add_in_csv(regression_data, regressionResult, ws, 'weekLowLTweek2Low')
        if(week2BarLowChange < -5):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarLowChangeLT-5')
        elif(week2BarLowChange < -4):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarLowChangeLT-4')
        elif(week2BarLowChange < -3):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarLowChangeLT-3')
        elif(week2BarLowChange < -2):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarLowChangeLT-2')
        elif(week2BarLowChange < -1):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarLowChangeLT-1')
        elif(week2BarLowChange < 0):
            add_in_csv(regression_data, regressionResult, ws, 'week2BarLowChangeLT0')
        if(regression_data['PCT_day_change'] < -1
            and regression_data['year2LowChange'] > 5
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakDown-2week')
        elif(regression_data['PCT_day_change'] > -1
            and regression_data['year2LowChange'] > 5
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakDown-2week-(risky)-PCTDayChangeGT-1')
        elif(regression_data['year2LowChange'] <= 5
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakDown-2week-(risky)-year2Low')
    elif(-6 < regression_data['PCT_day_change'] < -1.5
        and -6 < regression_data['PCT_change']
        and (regression_data['PCT_day_change_pre1'] > 0 or regression_data['PCT_day_change_pre2'] > 0)
        #and regression_data['low'] < regression_data['high_pre3']
        and regression_data['low'] < regression_data['high_pre2']
        and regression_data['low'] < regression_data['high_pre1']
        and -3 < regression_data['forecast_day_PCT4_change'] < -0.5
        and -4 < regression_data['forecast_day_PCT3_change'] < 0
        and -4 < regression_data['forecast_day_PCT2_change'] < 0
        and -3 < regression_data['forecast_day_PCT_change'] < 0
        and ((abs(regression_data['monthHighChange']) > abs(regression_data['monthLowChange']))
            )
        and regression_data['SMA4'] < regression_data['SMA4_2daysBack']
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakDown-month3HighGTMonth3Low')
    elif(-6 < regression_data['PCT_day_change'] < -1.5
        and -6 < regression_data['PCT_change']
        and (regression_data['PCT_day_change_pre1'] > 0 or regression_data['PCT_day_change_pre2'] > 0)
        #and regression_data['low'] < regression_data['high_pre3']
        and regression_data['low'] < regression_data['high_pre2']
        and regression_data['low'] < regression_data['high_pre1']
        and -3 < regression_data['forecast_day_PCT4_change'] < -0.5
        and -4 < regression_data['forecast_day_PCT3_change'] < 0
        and -4 < regression_data['forecast_day_PCT2_change'] < 0
        and -3 < regression_data['forecast_day_PCT_change'] < 0
        and ((-4 < regression_data['forecast_day_PCT5_change'] < 0
                and -4 < regression_data['forecast_day_PCT7_change'] < 0
               )
            )
        and regression_data['SMA4'] < regression_data['SMA4_2daysBack']
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)checkConsolidationBreakDown-forecastDayPCT7LT0')
   
def sell_final_candidate(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(regression_data['forecast_day_PCT4_change'] >= -0.5
        and regression_data['forecast_day_PCT5_change'] >= -0.5
        and regression_data['forecast_day_PCT7_change'] >= -0.5
        and regression_data['yearLowChange'] > 15 and regression_data['yearHighChange'] < -15
        and regression_data['month3HighChange'] < -3
        and (ten_days_more_than_ten(regression_data)
             or (last_5_day_all_up_except_today(regression_data)
                and ten_days_more_than_seven(regression_data)
                )
             )
        ):  
        if(regression_data['forecast_day_PCT_change'] < 0
            #and regression_data['bar_low'] < regression_data['bar_low_pre1']
            #and regression_data['forecast_day_VOL_change'] > 0
            ):
            if(-4 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['PCT_change'] < -2
                and regression_data['PCT_day_change_pre1'] > 0
                ):
                if(regression_data['SMA25'] > 0 and regression_data['SMA50'] > 0
                    and is_algo_sell(regression_data)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalCandidate-(lastDayUp)')
                    return True
                elif(regression_data['SMA25'] < 0 and regression_data['SMA50'] < 0
                    and is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalContinue-(lastDayUp)')
                    return True
                elif(is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalCandidate-1-(lastDayUp)')
                    return True
                return False
            if(-5 < regression_data['PCT_day_change'] < -2 and -5 < regression_data['PCT_change'] < -2
                and regression_data['PCT_day_change_pre1'] > 0
                and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
                and regression_data['forecast_day_PCT7_change'] > 5
                and regression_data['forecast_day_PCT10_change'] > 10
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-(lastDayUp)-(highChange)')
                return True
            elif(regression_data['PCT_day_change'] < -2 and regression_data['PCT_change'] < -2
                and regression_data['PCT_day_change_pre1'] > 0
                and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
                and (regression_data['forecast_day_PCT7_change'] < 5 or regression_data['forecast_day_PCT10_change'] < 10)
                ):
                if(('P@[' not in regression_data['sellIndia'])):
                    add_in_csv(regression_data, regressionResult, ws, 'buyFinalContinue-(lastDayUp)-(highChange)')
                    return True
                else:
                    add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-(lastDayUp)-(highChange)')
                    return True
            if(-4 < regression_data['PCT_day_change'] < -0.5 and -4 < regression_data['PCT_change'] < -0.5
                and regression_data['PCT_day_change_pre1'] < 0
                and 'P@[' in regression_data['buyIndia']
                ):
                add_in_csv(regression_data, regressionResult, ws, 'sellFinalCandidate-(lastDayDown)-(buyPattern)')
                return True
        if((((regression_data['open'] - regression_data['close']) * 1.5 > regression_data['high'] - regression_data['low']) or (regression_data['forecast_day_PCT_change'] < 0 and regression_data['PCT_day_change'] < -1))
            and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
            ):
            if(-2.5 < regression_data['PCT_day_change'] < -1 and -2.5 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] > -20
                    and is_algo_sell(regression_data)
                    ):    
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalCandidate-3')
                    return True
                elif(regression_data['forecast_day_VOL_change'] < -30
                    and is_algo_buy(regression_data)
                    and ('[' in regression_data['sellIndia'])
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalContinue-3')
                    return True
                return False
            if(-5 < regression_data['PCT_day_change'] < -1 and -5 < regression_data['PCT_change'] < -1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                if(regression_data['forecast_day_VOL_change'] < -30
                    and is_algo_sell(regression_data)
                    ):    
                    add_in_csv(regression_data, regressionResult, ws, 'ML:sellFinalCandidate-4')
                    return True
                elif(regression_data['forecast_day_VOL_change'] > -20
                    and is_algo_buy(regression_data)
                    and ('[' in regression_data['sellIndia'])
                    ):    
                    add_in_csv(regression_data, regressionResult, ws, 'ML:buyFinalContinue-4')
                    return True
                return False
    return False            

def sell_oi(regression_data, regressionResult, reg, ws):
    return False
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    if(((regression_data['PCT_day_change'] < -1 or regression_data['PCT_change'] < -1)
             or
             (regression_data['PCT_day_change'] < -0.75 and regression_data['PCT_change'] < -0.75)
             and regression_data['forecast_day_PCT_change'] > -2
             and regression_data['forecast_day_PCT2_change'] > -2
             and regression_data['forecast_day_PCT3_change'] > -2
             and regression_data['forecast_day_PCT4_change'] > -2
             and regression_data['forecast_day_PCT5_change'] > -2
             and regression_data['forecast_day_PCT7_change'] > -2
             and regression_data['forecast_day_PCT10_change'] > -2
            )
        and regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0
        and -5 < regression_data['forecast_day_PCT4_change'] < 0
        and regression_data['forecast_day_PCT5_change'] > -5
        and regression_data['forecast_day_PCT7_change'] > -5
        and regression_data['forecast_day_PCT10_change'] > -5
        and (regression_data['PCT_day_change'] > 0
            or regression_data['PCT_day_change_pre1'] > 0
            or regression_data['PCT_day_change_pre2'] > 0
            or regression_data['PCT_day_change_pre3'] > 0
            or regression_data['PCT_day_change_pre4'] > 0
            )
        and (regression_data['forecast_day_VOL_change'] > 150
            or (regression_data['PCT_day_change_pre2'] < 0
                and (((regression_data['volume'] - regression_data['volume_pre2'])*100)/regression_data['volume_pre2']) > 100
                and (((regression_data['volume'] - regression_data['volume_pre3'])*100)/regression_data['volume_pre3']) > 100
               )
           )
        and float(regression_data['contract']) > 100
        and(regression_data['PCT_day_change_pre1'] < 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
            )
        and regression_data['open'] > 50
        and (last_4_day_all_down(regression_data) == False) #Uncomment0 If very less data
        and (low_tail_pct(regression_data) < 1)
        and (low_tail_pct(regression_data) < 1.5 or (low_tail_pct(regression_data) < high_tail_pct(regression_data)))
        and regression_data['month3LowChange'] > 7.5
        ):
        if(-3 < regression_data['PCT_day_change'] < -1 and -3 < regression_data['PCT_change'] < -1 
            ):
            if(regression_data['forecast_day_PCT10_change'] > 0 or regression_data['forecast_day_PCT7_change'] > 0):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-0')
                return True
            elif(regression_data['forecast_day_PCT10_change'] > -10 or (regression_data['forecast_day_PCT5_change'] > -5 and regression_data['forecast_day_PCT7_change'] > -5)):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-1')
                return True
            else:
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-1-Risky')
                return True
        if(-6 < regression_data['PCT_day_change'] < -1 and -6 < regression_data['PCT_change'] < -1 
            and float(regression_data['forecast_day_VOL_change']) > 300 
            ):
            if(regression_data['forecast_day_PCT10_change'] > 0 or regression_data['forecast_day_PCT7_change'] > 0
               ):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-0-checkConsolidation')
                return True
            elif(regression_data['forecast_day_PCT10_change'] > -10 or (regression_data['forecast_day_PCT5_change'] > -5 and regression_data['forecast_day_PCT7_change'] > -5)
                -4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1
                ):
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-1-checkConsolidation')
                return True
            else:
                add_in_csv(regression_data, regressionResult, ws, 'openInterest-1-Risky')
                return True 
    return False
    
def sell_market_downtrend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    return False
#     if(('P@' not in regression_data['buyIndia'])
#         and ((-6 < regression_data['PCT_change'] < -2) 
#             and (-6 < regression_data['PCT_day_change'] < -2)
#             and regression_data['close'] < regression_data['bar_low_pre1']
#             )
#         #and regression_data['trend'] == 'down'
#         ):
#         if(regression_data['trend'] == 'down'):
#             if(('ReversalHighYear2' in regression_data['filter3'])):
#                 add_in_csv(regression_data, regressionResult, ws, "##DOWNTREND:sellYear2HighReversal(InDownTrend)")
#             if(('ReversalHighYear' in regression_data['filter3'])):
#                 add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellYearHighReversal(InDownTrend)')
#             if(('ReversalHighMonth6' in regression_data['filter3'])):
#                 add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellMonth6HighReversal(InDownTrend)')
# #             if(('ReversalHighMonth3' in regression_data['filter3'])):
# #                 add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellMonth3HighReversal(InDownTrend)')
# #             if(regression_data['month3LowChange'] > 15 and (regression_data['month6LowChange'] > 20 or regression_data['yearLowChange'] > 30)
# #                 ):
# #                 if(('NearHighYear2' in regression_data['filter3'])):
# #                     add_in_csv(regression_data, regressionResult, ws, None)
# #                 if(('NearHighYear' in regression_data['filter3'])):
# #                     add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellYearHigh(InDownTrend)')
# #                 if(('NearHighMonth6' in regression_data['filter3'])):
# #                     add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellMonth6High(InDownTrend)')
# #                 if(('NearHighMonth3' in regression_data['filter3'])):
# #                     add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellMonth3High(InDownTrend)')
#         if(((regression_data['PCT_day_change'] < -2) or (regression_data['PCT_change'] < -2) or ('MaySellCheckChart' in regression_data['filter1']))):
#             if('BreakLowYear' in regression_data['filter3']
#                 and regression_data['year2LowChange'] > 5
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellYearLowBreak')
#             if('BreakLowMonth6' in regression_data['filter3']
#                 and regression_data['yearLowChange'] > 5
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellMonth6LowBreak')
#             if('BreakLowMonth3' in regression_data['filter3']
#                 and regression_data['month6LowChange'] > 0
#                 ):
#                 add_in_csv(regression_data, regressionResult, ws, '##DOWNTREND:sellMonth3LowBreak')

def sell_supertrend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    
    if(regression_data['close'] > 50
        ):
        if(0 > regression_data['forecast_day_PCT2_change'] > regression_data['forecast_day_PCT3_change'] > regression_data['forecast_day_PCT4_change'] > regression_data['forecast_day_PCT5_change']
            and regression_data['forecast_day_PCT5_change'] < -5
            and -1.5 < regression_data['PCT_day_change_pre1'] < 0.5
            and -0.5 < regression_data['PCT_day_change'] < 0.5
            and regression_data['yearHighChange'] < -5
            and regression_data['yearLowChange'] > 5
            and regression_data['month3LowChange'] > 5
            #and regression_data['high'] < regression_data['high_pre1']
            ):
            if(regression_data['PCT_day_change_pre1'] < 0
                and -0.75 < regression_data['PCT_day_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, '--')
            elif(regression_data['PCT_day_change_pre1'] > 0
                and 0 < regression_data['PCT_day_change'] < 0.75):
                add_in_csv(regression_data, regressionResult, ws, '++')
            elif(regression_data['PCT_day_change_pre1'] < 0
                and 0 < regression_data['PCT_day_change'] < 0.75):
                add_in_csv(regression_data, regressionResult, ws, '-+')
            elif(regression_data['PCT_day_change_pre1'] > 0
                and -0.75 < regression_data['PCT_day_change'] < 0):
                add_in_csv(regression_data, regressionResult, ws, '+-')
#             if(regression_data['low_month6'] >= regression_data['low']):
#                 add_in_csv(regression_data, regressionResult, ws, 'LessThanM6Low')
#             else:
#                 add_in_csv(regression_data, regressionResult, ws, 'MoreThanM6Low')
            if(abs_month3High_more_than_month3Low(regression_data)
#                 and regression_data['yearHighChange'] < -10
#                 and regression_data['yearLowChange'] > 10
#                 and regression_data['month3LowChange'] > 10
                ):
                if(regression_data['PCT_day_change'] > 0
                    and -1 < regression_data['PCT_day_change_pre1'] < 0
                    and ('[' in regression_data['buyIndia'])
                    and low_tail_pct(regression_data) < 2
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-(Risky)')
                elif((regression_data['PCT_day_change'] > 0
                    or regression_data['forecast_day_PCT_change'] > 0)
                    and (('[' not in regression_data['sellIndia']) and ('[' in regression_data['buyIndia']))
                    ):
                    if(is_algo_buy(regression_data)):
                        add_in_csv(regression_data, regressionResult, ws, 'ML:buySuperTrend-0')
                        return True
                    return False
            elif(abs_month3High_less_than_month3Low(regression_data)):
                if(regression_data['PCT_day_change'] > 0
                    and -1 < regression_data['PCT_day_change_pre1'] < 0
                    and low_tail_pct(regression_data) < 2.5
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'sellSuperTrend-0')
            return True
        return False

def sell_heavy_downtrend(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    return False
#         if((('NearHighYear2' in regression_data['filter3']) 
#                 or ('NearHighYear' in regression_data['filter3'])
#                 or ('NearHighMonth6' in regression_data['filter3'])
#                 or ('ReversalHighYear2' in regression_data['filter3'])
#                 or ('ReversalHighYear' in regression_data['filter3'])
#                 or ('ReversalHighMonth6' in regression_data['filter3'])
#                 or ('BreakHighYear2' in regression_data['filter3'])
#                 or ('BreakHighYear' in regression_data['filter3'])
#                 or ('BreakHighMonth6' in regression_data['filter3'])
#             )
#             and regression_data['SMA200'] > regression_data['SMA100'] > regression_data['SMA50'] > regression_data['SMA25'] > regression_data['SMA9']
#             and  ("downTrend" in regression_data['series_trend'])
#             and ((abs(regression_data['PCT_day_change']) > abs(regression_data['PCT_day_change_pre1']))
#                 or regression_data['PCT_change'] > 2
#                 )
#             ):
#             if(((0 < regression_data['PCT_change'] < 3) and (0 < regression_data['PCT_day_change'] < 3))
#             ):
#                 add_in_csv(regression_data, regressionResult, ws, '##(Test)sellSeilingReversal-0')
#             elif(regression_data['PCT_day_change_pre2'] > 0 and regression_data['PCT_day_change'] < -1.5):
#                 add_in_csv(regression_data, regressionResult, ws, '##(Test)sellSeilingReversal-1')
                
#To-Do        
#         if( ("downTrend" in regression_data['series_trend'])
#             and (-15 < regression_data['forecast_day_PCT4_change'] < 0)
#             and (-15 < regression_data['forecast_day_PCT5_change'] < 0)
#             and (-15 < regression_data['forecast_day_PCT7_change'] < 0)
#             and (-15 < regression_data['forecast_day_PCT10_change'])
#             ):
#             if((regression_data['month3LowChange'] > 5)
#                 and ((-5 < regression_data['PCT_change'] < 0) and (-5 < regression_data['PCT_day_change'] < 0))
#                 and ('[') not in regression_data['buyIndia']
#                 and regression_data['close_pre1'] < regression_data['close_pre2']
#                 ):
#                 if((-5 < regression_data['PCT_change'] < -4) and (-5 < regression_data['PCT_day_change'] < -4)
#                 ):
#                     add_in_csv(regression_data, regressionResult, ws, '##sellHeavyDownTrend-0-Continue')
#                 elif(ten_days_less_than_minus_seven(regression_data)
#                      and(('NearLowMonth3' in regression_data['filter3'])
#                         or ('NearLowMonth6' in regression_data['filter3'])
#                         )
#                 ):
#                     add_in_csv(regression_data, regressionResult, ws, '##sellHeavyDownTrend-1-Continue')
#             if((regression_data['month3LowChange'] < 4)
#                 and ((-5 < regression_data['PCT_change'] < 0) and (-5 < regression_data['PCT_day_change'] < 0))
#                 ):
#                 if((-5 < regression_data['PCT_change'] < -2) and (-5 < regression_data['PCT_day_change'] < -1)
#                     and ten_days_less_than_minus_seven(regression_data)
#                     and(('NearLowMonth3' in regression_data['filter3'])
#                         or ('BreakLowMonth3' in regression_data['filter3'])
#                         or ('NearLowMonth6' in regression_data['filter3'])
#                         )
#                 ):
#                     add_in_csv(regression_data, regressionResult, ws, '##buyHeavyDownTrend-0-Reversal')

def sell_check_chart(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    if(('NearLowMonth3' in regression_data['filter3']) 
            or ('BreakLowMonth3' in regression_data['filter3'])
            or ('ReversalLowMonth3' in regression_data['filter3'])
        ):
        if((regression_data['forecast_day_PCT4_change'] > 0)
            and (regression_data['forecast_day_PCT5_change'] > 0)
            and (regression_data['forecast_day_PCT7_change'] > 0)
            and ((regression_data['forecast_day_PCT_change'] < 0 and (-4 < regression_data['PCT_change'] < -2) and (-4 < regression_data['PCT_day_change'] < -2))
                 or ( ("upTrend" in regression_data['series_trend']) and (-4 < regression_data['PCT_change'] < -1) and (-4 < regression_data['PCT_day_change'] < 1)))
            and ('[') not in regression_data['buyIndia']
            and regression_data['low'] < regression_data['low_pre1']
            and regression_data['bar_low'] < regression_data['bar_low_pre1']
            ):
            if('NearLowMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearLowMonth3')
            elif('ReversalLowMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalLowMonth3')
            elif('BreakLowMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'BreakLowMonth3')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-sell)Month3Low-Continue(InUpTrend)')
            
    #Check for last 5 from latest down should crossover
    if(('NearLowMonth3' in regression_data['filter3']) 
        or ('BreakLowMonth3' in regression_data['filter3'])
        or ('ReversalLowMonth3' in regression_data['filter3'])
        ):
        if(regression_data['month3HighChange'] < -15
            and ((1 < regression_data['PCT_change'] < 5) and (2 < regression_data['PCT_day_change'] < 5))
            and regression_data['close'] > regression_data['bar_high_pre1']
            and regression_data['close'] > regression_data['bar_high_pre2']
            ):
            if('NearLowMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearLowMonth3')
            elif('ReversalLowMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalLowMonth3')
            elif('BreakLowMonth3' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'BreakLowMonth3')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-sell):month3Low-InPlus')
            
    elif(('NearLowMonth6' in regression_data['filter3']) 
        or ('ReversalLowMonth6' in regression_data['filter3'])
        ):
        if(regression_data['month6LowChange'] < -20
            and ((1 < regression_data['PCT_change'] < 5) and (2 < regression_data['PCT_day_change'] < 5))
            and regression_data['close'] > regression_data['bar_high_pre1']
            and regression_data['close'] > regression_data['bar_high_pre2']
            ):
            if('NearLowMonth6' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearLowMonth6')
            elif('ReversalLowMonth6' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalLowMonth6')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-sell):month6Low-InPlus')
                
    elif(('NearLowYear' in regression_data['filter3']) 
        or ('ReversalLowYear' in regression_data['filter3'])
        ):
        if(regression_data['month6LowChange'] < -20
            and ((1 < regression_data['PCT_change'] < 5) and (2 < regression_data['PCT_day_change'] < 5))
            and regression_data['close'] > regression_data['bar_high_pre1']
            and regression_data['close'] > regression_data['bar_high_pre2']
            ):
            if('NearLowYear' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'NearLowYear')
            elif('ReversalLowYear' in regression_data['filter3']):
                add_in_csv(regression_data, regressionResult, ws, 'ReversalLowYear')
            add_in_csv(regression_data, regressionResult, ws, '##(Check-chart-sell):yearLow-InPlus')
                        
#         if((('NearLowYear2' in regression_data['filter3']) 
#             or ('ReversalLowYear2' in regression_data['filter3'])
#             or ('BreakLowYear2' in regression_data['filter3'])
#             )):    
#             if(regression_data['year2HighChange'] < -50
#                 and regression_data['year2LowChange'] < 1
#                 and ((regression_data['PCT_change'] < -1) and (regression_data['PCT_day_change'] < -1))
#                 and regression_data['close'] < regression_data['bar_low_pre1']
#                 ):
#                 if(regression_data['PCT_change'] < -4
#                     and ('BreakLowYear2' in regression_data['filter3'])
#                     ):
#                     add_in_csv(regression_data, regressionResult, ws, '##(Test)(Check-chart-buy):year2Low-InMinus')
#                 elif(regression_data['PCT_change'] < -2.5 and regression_data['PCT_day_change'] < -2.5
#                     and ('NearLowYear2' in regression_data['filter3'])
#                     ):
#                     add_in_csv(regression_data, regressionResult, ws, '##(Test)(Check-chart-sell):year2Low-InMinus')
#                 else:
#                     add_in_csv(regression_data, regressionResult, ws, '##(Test)(Check-chart):year2Low-InMinus')
#         elif((('NearLowYear' in regression_data['filter3'])
#             or ('ReversalLowYear' in regression_data['filter3'])
#             or ('BreakLowYear' in regression_data['filter3'])
#             )):
#             if(regression_data['yearHighChange'] < -40
#                 and regression_data['yearLowChange'] < 1
#                 and ((regression_data['PCT_change'] < -1) and (regression_data['PCT_day_change'] < -1))
#                 and regression_data['close'] < regression_data['bar_low_pre1']
#                 ):
#                 if(regression_data['PCT_change'] < -4
#                     and ('BreakLowYear' in regression_data['filter3'])
#                     ):
#                     add_in_csv(regression_data, regressionResult, ws, '##(Test)(Check-chart-buy):yearLow-InMinus')
#                 elif(regression_data['PCT_change'] < -2.5 and regression_data['PCT_day_change'] < -2.5
#                     and ('NearLowYear' in regression_data['filter3'])
#                     ):
#                     add_in_csv(regression_data, regressionResult, ws, '##(Test)(Check-chart-sell):yearLow-InMinus')
#                 else:
#                     add_in_csv(regression_data, regressionResult, ws, '##(Test)(Check-chart):yearLow-InMinus

def sell_downingMA(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    
    if(is_algo_sell(regression_data)
        and regression_data['SMA4'] < 0
        and regression_data['SMA9'] < 0
        and regression_data['SMA25'] < 0
        and regression_data['SMA100'] < -10
        and regression_data['SMA200'] < -10
        and regression_data['year2HighChange'] > 5
        and regression_data['yearHighChange'] > 5
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)ML:AllMANegative-Downtrend')
        return True
    
def sell_study_downingMA(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    mlpValue_other, kNeighboursValue_other = get_reg_or_cla_other(regression_data, reg)
    
    if(('(Confirmed)EMA6<EMA14' in regression_data['filter4'])
        and -3 > regression_data['PCT_day_change'] 
        and 0 > regression_data['PCT_change']
        and regression_data['PCT_day_change_pre1'] > 1
        and (regression_data['PCT_day_change'] < -3 or regression_data['PCT_change'] < -3)
        ):
        if(regression_data['month3HighChange'] > -10):
            add_in_csv(regression_data, regressionResult, ws, '(check-chart)buyEMA6<EMA14')
        elif(regression_data['PCT_day_change'] < -3.5 or regression_data['PCT_change'] < -3.5):
            add_in_csv(regression_data, regressionResult, ws, '(check-chart)buyEMA6<EMA14-Risky')
        
    if(regression_data['SMA9'] > 0 
        and regression_data['SMA4'] > 0
        and regression_data['forecast_day_PCT_change'] > 0.5
        and regression_data['forecast_day_PCT2_change'] > 0.5
        and regression_data['forecast_day_PCT5_change'] > 0.5
        and regression_data['forecast_day_PCT5_change'] < 5
        and regression_data['forecast_day_PCT7_change'] < 10
        and abs(regression_data['PCT_day_change_pre1']) > 2* (regression_data['PCT_day_change'])
        and ((-1.5 < regression_data['PCT_day_change'] < 0
            and regression_data['PCT_day_change_pre1'] > 2
            and regression_data['PCT_day_change_pre2'] > -1.5
            and regression_data['PCT_day_change_pre3'] > -1.5
            and (
                regression_data['PCT_day_change_pre2'] > 1
                or regression_data['PCT_day_change_pre3'] > 1
                )
            )
        )
        and low_tail_pct(regression_data) > 1
        and (high_tail_pct(regression_data) < 1.5)
        ):  
        add_in_csv(regression_data, regressionResult, ws, "##(Test)(check-chart)-buySMAUpTrend")
    elif((regression_data['SMA9'] > 0 or regression_data['SMA25'] > 0)
        and -0.5 < regression_data['weekLowChange'] < 0
        and regression_data['close'] > regression_data['low_week']
        and regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < 0
        and (regression_data['forecast_day_PCT5_change'] > 0
            or regression_data['forecast_day_PCT7_change'] > 0
            ) 
        and regression_data['forecast_day_PCT10_change'] > 0
        and ((0 < regression_data['PCT_day_change'] < -1
            and regression_data['PCT_day_change_pre1'] < -1
            and (
                regression_data['PCT_day_change_pre2'] > 0
                or regression_data['PCT_day_change_pre3'] > 0
                )
            )
        )
        and low_tail_pct(regression_data) > 1 
        and (high_tail_pct(regression_data) < 1.5)
        ):  
        add_in_csv(regression_data, regressionResult, ws, "##(Test)(check-chart)-buySMAUpTrend-2DayDown-weekLowTouch")
    return True

def sell_random_filter(regression_data, regressionResult, reg, ws):
    if(-5 < regression_data['PCT_day_change'] < -1
        and (regression_data['PCT_day_change'] - 0.5) < regression_data['PCT_change'] < (regression_data['PCT_day_change'] + 0.5)
        and (regression_data['forecast_day_PCT10_change'] > -1)
        and (regression_data['forecast_day_PCT7_change'] > 0
            or regression_data['forecast_day_PCT10_change'] > 0
            )
        and (regression_data['forecast_day_PCT_change'] < 0
            or regression_data['forecast_day_PCT2_change'] < 0
            or regression_data['forecast_day_PCT3_change'] < 0
            )
        and (((regression_data['high'] - regression_data['high_pre1'])/regression_data['high_pre1'])*100) < 0
        ):
        if(regression_data['week2LowChange'] < 0.5):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-near2WeekDown-mayReveresalBuy')
        elif(regression_data['PCT_day_change_pre1'] < -2
            and regression_data['PCT_day_change_pre2'] < -2
            and regression_data['forecast_day_PCT5_change'] < 0
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-3dayDown-mayReversalBuy')
        elif(regression_data['PCT_day_change'] < -2
            and regression_data['PCT_day_change_pre1'] < -2
            and regression_data['forecast_day_PCT5_change'] < 0
            ):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-2dayDown-mayReversalBuy')
        elif(-3 < regression_data['PCT_day_change'] < -2
            and regression_data['week2LowChange'] > 2
            ):
            if(is_algo_sell(regression_data)
                and -2 < regression_data['PCT_day_change_pre1'] < -0.5
                and regression_data['PCT_day_change_pre2'] < -0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-3dayDown')
            elif(is_algo_sell(regression_data)
                and -2 < regression_data['PCT_day_change_pre1'] < -0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-2dayDown')
            elif(regression_data['forecast_day_PCT7_change'] > -1
                and regression_data['PCT_day_change_pre1'] > 0 
                and regression_data['PCT_day_change_pre2'] > 0
                and regression_data['PCT_day_change_pre3'] > 0
                and regression_data['forecast_day_PCT_change'] < 0
                and regression_data['forecast_day_PCT2_change'] < 0
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-upTrendLastDayDown-mayReversalBuy') 
            elif(is_algo_sell(regression_data)
                and regression_data['forecast_day_PCT7_change'] > -1
                and (regression_data['PCT_day_change_pre2'] > 0 
                    or regression_data['bar_low'] < regression_data['bar_low_pre1']
                    )
                and abs_month3High_more_than_month3Low(regression_data)
                ):
                if(regression_data['bar_low'] < regression_data['bar_low_pre1']):
                    add_in_csv(regression_data, regressionResult, ws, 'bar_low')
                if(regression_data['forecast_day_PCT2_change'] < -1
                    and regression_data['forecast_day_PCT3_change'] < -1
                    ):
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown')
                else:
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-inDownTrend')
            elif(is_algo_sell(regression_data)
                and regression_data['forecast_day_PCT7_change'] > -1
                and (regression_data['PCT_day_change_pre2'] > 0 
                    or regression_data['bar_low'] < regression_data['bar_low_pre1']
                    )
                ):
                if(regression_data['bar_low'] < regression_data['bar_low_pre1']):
                    add_in_csv(regression_data, regressionResult, ws, 'bar_low')
                if(regression_data['forecast_day_PCT2_change'] < -1
                    and regression_data['forecast_day_PCT3_change'] < -1
                    ):
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-month3HighLTmonth3Low')
                else:
                    add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-inDownTrend-month3HighLTmonth3Low')
            elif(regression_data['forecast_day_PCT7_change'] > -1
                and regression_data['PCT_day_change_pre1'] > 0 
                and regression_data['PCT_day_change_pre2'] < 0
                and regression_data['bar_low'] > regression_data['bar_low_pre1']
                ):
                add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-mayReversalBuy')     
    if(1 < regression_data['PCT_day_change'] < 2.5
        and 1 < regression_data['PCT_change'] < 3.5
        and regression_data['PCT_day_change_pre1'] < -1.5
        and (regression_data['PCT_day_change_pre2'] < -1
            or regression_data['PCT_day_change_pre3'] < -1
            )
        #and (regression_data['forecast_day_PCT5_change'] > -2)
        and (regression_data['forecast_day_PCT7_change'] > -2)
        and (regression_data['forecast_day_PCT10_change'] > -2)
        and (regression_data['forecast_day_PCT5_change'] > 0
            and regression_data['forecast_day_PCT7_change'] > 0
            and regression_data['forecast_day_PCT10_change'] > 0
            )
        and (regression_data['forecast_day_PCT2_change'] < 0
            and regression_data['forecast_day_PCT3_change'] < 0
            )
        and (((regression_data['high'] - regression_data['high_pre1'])/regression_data['high_pre1'])*100) < 0
        and (((regression_data['high'] - regression_data['bar_high_pre1'])/regression_data['bar_high_pre1'])*100) < 0
        ):
        if(regression_data['week2LowChange'] < 0.5):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-lastDayUp-near2WeekDown-mayReveresalBuy')
        elif(regression_data['week2LowChange'] > 1):
            add_in_csv(regression_data, regressionResult, ws, '(Test)checkCupDown-lastDayUp')
        
    if((-0.5 < regression_data['forecast_day_PCT4_change'] < 0.5
        or -0.5 < regression_data['forecast_day_PCT5_change'] < 0.5
        )
        and 
        (regression_data['forecast_day_PCT4_change'] < -1
        or regression_data['forecast_day_PCT5_change'] < -1
        )
        and regression_data['forecast_day_PCT_change'] < -1
        and regression_data['forecast_day_PCT2_change'] < -2
        and regression_data['forecast_day_PCT3_change'] < -3
#         and regression_data['forecast_day_PCT7_change'] < -3
#         and regression_data['forecast_day_PCT10_change'] < -3
        and regression_data['PCT_day_change'] < -1
        and regression_data['PCT_day_change_pre1'] < -1
        and regression_data['PCT_day_change_pre2'] < -1
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)buyMay5DayFloorReversal')
        
#     if(-6 < regression_data['PCT_day_change'] < -1.5
#         and -6 < regression_data['PCT_change'] 
#         and (regression_data['PCT_day_change_pre2'] > 0 or regression_data['PCT_day_change_pre3'] > 0)
#         and regression_data['low'] < regression_data['low_pre2']
#         and regression_data['forecast_day_PCT3_change'] < 0
#         and regression_data['forecast_day_PCT2_change'] < 0
#         and regression_data['forecast_day_PCT_change'] < 0
#         and regression_data['SMA4'] < regression_data['SMA4_2daysBack']
#         and regression_data['forecast_day_PCT10_change'] > 0
#         and ((regression_data['forecast_day_PCT4_change'] < 0
#               and regression_data['forecast_day_PCT7_change'] > 0
#                 and (regression_data['forecast_day_PCT7_change'] > 5
#                     or regression_data['forecast_day_PCT10_change'] > 5
#                     )
#                 )
#             ) 
#         ):
#         add_in_csv(regression_data, regressionResult, ws, '(Test)check10DayHighReversal')
    
    if('DowningMA-Risky' not in regression_data['filter4'] 
        and 'DowningMA' in regression_data['filter4']
        and (regression_data['high'] < regression_data['high_pre1']) 
        and (0 < regression_data['PCT_day_change'] < 2) and (0 < regression_data['PCT_change'] < 2)
        and (regression_data['forecast_day_PCT5_change'] < -3
             or regression_data['forecast_day_PCT7_change'] < -3
             or regression_data['forecast_day_PCT10_change'] < -3
            )
        ):
        if(-0.5 < regression_data['week2LowChange'] < 0.5 
            and regression_data['week2Low'] != regression_data['low_pre1']
            and regression_data['week2Low'] != regression_data['low_pre2']
            #and (0 < regression_data['PCT_day_change'] < 1) and (0 < regression_data['PCT_change'] < 1)
            ):
            add_in_csv(regression_data, regressionResult, ws, None)
        elif(regression_data['week2LowChange'] > 1):
            add_in_csv(regression_data, regressionResult, ws, '(Test)sellDowningMA-(check5MinuteDownTrendAndSellSupertrend)')
        elif(is_algo_sell(regression_data)):
            add_in_csv(regression_data, regressionResult, ws, '(Test)sellDowningMA-(check5MinuteDownTrendAndSellSupertrend)-Already10DayDown')
    
    if((0 < regression_data['PCT_day_change'] < 0.75) and (0 < regression_data['PCT_change'] < 0.75)
        and (regression_data['SMA4_2daysBack'] < 0 or regression_data['SMA9_2daysBack'] < 0)
        and regression_data['SMA4'] > 0
        and regression_data['PCT_day_change_pre1'] > 0.5
        and regression_data['PCT_day_change_pre2'] > 0.5
        and (regression_data['PCT_day_change_pre1'] > 1.5 or regression_data['PCT_day_change_pre2'] > 1.5)
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)sellSMA4Reversal')
        
    if((2 < regression_data['PCT_day_change']) and (1.5 < regression_data['PCT_change'])
        and regression_data['bar_high'] > regression_data['bar_high_pre1']
        ):
        if(regression_data['year2HighChange'] < -40
            and regression_data['yearHighChange'] < -20
            and regression_data['PCT_day_change_pre1'] < 0
            and ((regression_data['bar_low_pre1'] < regression_data['bar_low_pre2'])
                or regression_data['low_pre1'] < regression_data['low_pre2']
                )
            and regression_data['SMA25'] < 0
            and regression_data['SMA50'] < 0
            and regression_data['SMA100'] < 0
            ):
            if(regression_data['high_pre1'] < regression_data['high_pre2'] < regression_data['high_pre3'] < regression_data['high_pre4']): 
                add_in_csv(regression_data, regressionResult, ws, 'sellYear2LowLT-40-last4DayDown(triggerAfter-9:15)')
            elif(regression_data['high_pre1'] < regression_data['high_pre2'] < regression_data['high_pre3']): 
                add_in_csv(regression_data, regressionResult, ws, 'sellYear2LowLT-40-last3DayDown(triggerAfter-9:15)')
                
    if((-2 < regression_data['PCT_day_change'] < -0.5) and (4 < regression_data['PCT_change'] < 10)
        and high_tail_pct(regression_data) > 1
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)sellLastDayHighUpReversal')
        
def sell_skip_close_lt_50(regression_data, regressionResult, reg, ws):
    if((regression_data['weekLowChange'] > 15) 
        and (-10 < regression_data['PCT_change'] < -4)
        and (-10 < regression_data['PCT_day_change'] < -4)
        ):
        add_in_csv(regression_data, regressionResult, ws, '(Test)sellUptrendReversal')
        return True
    return False

def sell_test_345(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = sell_other_indicator(regression_data, regressionResult, reg, ws)
    filterName = pct_change_filter(regression_data, regressionResult, False)
    regression_data['filterTest'] = filterName + ',' \
                                    + regression_data['series_trend'] + ',' \
                                    + regression_data['filter2'] + ',' \
                                    + regression_data['filter3'] + ',' \
                                    + regression_data['filter4'] + ',' \
                                    + regression_data['filter5'] 
    if regression_data['filterTest'] != '':
        return True  
    return False
        
def sell_test(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = sell_other_indicator(regression_data, regressionResult, reg, ws)
    if regression_data['filterTest'] != '':
        regression_data['filterTest'] = regression_data['series_trend'] + ',' \
                                        + regression_data['filterTest'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    
    if regression_data['filterTest'] != '':
        return True
        
    return False

def sell_test_pct_change(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = sell_other_indicator(regression_data, regressionResult, reg, ws)
    if regression_data['filterTest'] != '':
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        regression_data['filterTest'] = filterName + ',' \
                                        + filterNameTail + ',' \
                                        + regression_data['series_trend'] + ',' \
                                        + regression_data['filterTest'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    
    if regression_data['filterTest'] != '':
        return True
        
    return False

def sell_test_all(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = sell_other_indicator(regression_data, regressionResult, reg, ws)
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    regression_data['filterTest'] = filterName + ',' \
                                    + filterNameTail + ',' \
                                    + regression_data['series_trend'] + ',' \
                                    + regression_data['filter2'] + ',' \
                                    + regression_data['filter3'] + ',' \
                                    + regression_data['filter4'] + ',' \
                                    + regression_data['filter5'] + ',' \
                                    + regression_data['filterTest'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    if regression_data['filterTest'] != '':
        return True  
    return False

def sell_test_tech(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    if (regression_data['buyIndia'] != '' or regression_data['sellIndia'] != ''):
        regression_data['filterTest'] = 'B@{' + regression_data['buyIndia'] + '}' \
                                        + 'S@{' + regression_data['sellIndia'] + '}' 
        return True
    
    return False

def sell_test_tech_all(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    if (regression_data['buyIndia'] != '' or regression_data['sellIndia'] != ''):
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        regression_data['filterTest'] = filterNameTail + ',' \
                                        + regression_data['series_trend'] + ',' \
                                        + 'B@{' + regression_data['buyIndia'] + '}' \
                                        + 'S@{' + regression_data['sellIndia'] + '}' \
                                        + regression_data['filter3'] + ',' \
                                        + regression_data['filter4'] + ',' 
        return True
    
    return False

def sell_test_tech_all_pct_change(regression_data, regressionResult, reg, ws):
    regression_data['filter'] = " "
    regression_data['filter1'] = " "
    regression_data['filter2'] = " "
    regression_data['filter3'] = " "
    regression_data['filter4'] = " "
    regression_data['filter5'] = " "
    regression_data['filter6'] = " "
    regression_data['series_trend'] = trend_calculator(regression_data)
    flag = buy_other_indicator(regression_data, regressionResult, reg, ws)
    if (regression_data['buyIndia'] != '' or regression_data['sellIndia'] != ''):
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        regression_data['filterTest'] = filterName + ',' \
                                        + filterNameTail + ',' \
                                        + regression_data['series_trend'] + ',' \
                                        + 'B@{' + regression_data['buyIndia'] + '}' \
                                        + 'S@{' + regression_data['sellIndia'] + '}' \
                                        + regression_data['filter3'] + ',' \
                                        + regression_data['filter4'] + ',' 
        return True
    
    return False
      
def sell_oi_candidate(regression_data, regressionResult, reg, ws):
    mlpValue, kNeighboursValue = get_reg_or_cla(regression_data, reg)
    tail_pct_filter(regression_data, regressionResult)
    tail_reversal_filter(regression_data, regressionResult)
    flag = False
    if(regression_data['close'] > 50):
        if sell_base_line_buy(regression_data, regressionResult, reg, ws):
            flag = True
    if(sell_trend_reversal(regression_data, regressionResult, reg, ws)
        and regression_data['close'] > 50
        ):
        flag = True
    if(breakout_or_no_consolidation(regression_data) == True):
        if sell_morning_star_buy(regression_data, regressionResult, reg, ws): 
            flag = True
        if sell_evening_star_sell(regression_data, regressionResult, reg, ws): 
            flag = True
        if sell_day_high(regression_data, regressionResult, reg, ws):
            flag = True
        if sell_trend_break(regression_data, regressionResult, reg, ws):
            flag = True
        if sell_final_candidate(regression_data, regressionResult, reg, ws):
            flag = True
    return flag

def sell_all_filter(regression_data, regressionResult, reg, ws):
    flag = False
    if sell_year_high(regression_data, regressionResult, reg, None, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_year_low(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_up_trend(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_down_trend(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_final(regression_data, regressionResult, reg, None, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_af_high_indicators(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_pattern(regression_data, regressionResult, reg, None, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    if sell_oi(regression_data, regressionResult, reg, None):
        add_in_csv(regression_data, regressionResult, ws, None)
        flag = True
    return flag

def sell_filter_345_accuracy(regression_data, regressionResult):
    filtersDict = filter345sell
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filter = filterName + ',' \
                + regression_data['series_trend'] + ',' \
                + regression_data['filter2'] + ',' \
                + regression_data['filter3'] + ',' \
                + regression_data['filter4'] + ',' \
                + regression_data['filter5']
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_345_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_345_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_345_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_345_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])
                
def sell_filter_accuracy(regression_data, regressionResult):
    filtersDict=filtersell
    if regression_data['filter'] != '':
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        filter =regression_data['series_trend'] + ',' \
                + regression_data['filter'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
        if filter != '' and filter in filtersDict:
            regression_data['filter_avg'] = float(filtersDict[filter]['avg'])
            regression_data['filter_count'] = float(filtersDict[filter]['count'])
            if float(filtersDict[filter]['count']) >= 2:
                if float(filtersDict[filter]['avg']) >= 0:
                    regression_data['filter_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
                else:
                    regression_data['filter_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])
                     
def sell_filter_pct_change_accuracy(regression_data, regressionResult):
    filtersDict=filterpctchangesell
    if regression_data['filter'] != '':
        filterName = pct_change_filter(regression_data, regressionResult, False)
        filterNameTail = tail2_change_filter(regression_data, regressionResult, False)
        filter = filterName + ',' \
                + filterNameTail + ',' \
                + regression_data['series_trend'] + ',' \
                + regression_data['filter'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
        if filter != '' and filter in filtersDict:
            regression_data['filter_pct_change_avg'] = float(filtersDict[filter]['avg'])
            regression_data['filter_pct_change_count'] = float(filtersDict[filter]['count'])
            if float(filtersDict[filter]['count']) >= 2:
                if float(filtersDict[filter]['avg']) >= 0:
                    regression_data['filter_pct_change_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
                else:
                    regression_data['filter_pct_change_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])
                        
def sell_filter_345_all_accuracy(regression_data, regressionResult):
    filtersDict=filterallsell
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    filter = filterName + ',' \
                + filterNameTail + ',' \
                + regression_data['series_trend'] + ',' \
                + regression_data['filter2'] + ',' \
                + regression_data['filter3'] + ',' \
                + regression_data['filter4'] + ',' \
                + regression_data['filter5'] + ',' \
                + regression_data['filter'].replace("[MLBuy]:", "").replace("[MLSell]:", "").strip()
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_all_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_all_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_all_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_all_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def sell_filter_tech_accuracy(regression_data, regressionResult):
    filtersDict=filtertechsell
    filter ='B@{' + regression_data['buyIndia'] + '}' \
            + 'S@{' + regression_data['sellIndia'] + '}'
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_tech_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_tech_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_tech_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_tech_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def sell_filter_tech_all_accuracy(regression_data, regressionResult):
    filtersDict=filtertechallsell
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    filter = filterNameTail + ',' \
            + regression_data['series_trend'] + ',' \
            + 'B@{' + regression_data['buyIndia'] + '}' \
            + 'S@{' + regression_data['sellIndia'] + '}' \
            + regression_data['filter3'] + ',' \
            + regression_data['filter4'] + ','
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_tech_all_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_tech_all_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_tech_all_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_tech_all_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def sell_filter_tech_all_pct_change_accuracy(regression_data, regressionResult):
    filtersDict=filtertechallpctchangesell
    filterName = pct_change_filter(regression_data, regressionResult, False)
    filterNameTail = tail_change_filter(regression_data, regressionResult, False)
    filter = filterName + ',' \
            + filterNameTail + ',' \
            + regression_data['series_trend'] + ',' \
            + 'B@{' + regression_data['buyIndia'] + '}' \
            + 'S@{' + regression_data['sellIndia'] + '}' \
            + regression_data['filter3'] + ',' \
            + regression_data['filter4'] + ','
    if (filter != '') and (filter in filtersDict):
        regression_data['filter_tech_all_pct_change_avg'] = float(filtersDict[filter]['avg'])
        regression_data['filter_tech_all_pct_change_count'] = float(filtersDict[filter]['count'])
        if float(filtersDict[filter]['count']) >= 2:
            if float(filtersDict[filter]['avg']) >= 0:
                regression_data['filter_tech_all_pct_change_pct'] = (float(filtersDict[filter]['countgt'])*100)/float(filtersDict[filter]['count'])
            else:
                regression_data['filter_tech_all_pct_change_pct'] = -(float(filtersDict[filter]['countlt'])*100)/float(filtersDict[filter]['count'])

def sell_filter_all_accuracy(regression_data, regressionResult):
    sell_filter_345_accuracy(regression_data, regressionResult)
    sell_filter_accuracy(regression_data, regressionResult)
    sell_filter_pct_change_accuracy(regression_data, regressionResult) 
    sell_filter_345_all_accuracy(regression_data, regressionResult)
    sell_filter_tech_accuracy(regression_data, regressionResult)
    sell_filter_tech_all_accuracy(regression_data, regressionResult)
    sell_filter_tech_all_pct_change_accuracy(regression_data, regressionResult)

def filter_accuracy_finder_regression(regression_data, regressionResult, ws):
    Flag = False
    if("MLBuy" in regression_data['filter'] and len(regression_data['filter']) > 9):
        if(is_reg_buy_from_filter(regression_data, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
            or is_reg_buy_from_filter(regression_data, 'filter_avg', 'filter_count', 'filter_pct')
            or is_reg_buy_from_filter(regression_data, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
            or is_reg_buy_from_filter(regression_data, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
            or is_reg_buy_from_filter(regression_data, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
            or is_reg_buy_from_filter(regression_data, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
            or (regression_data['filter_tech_all_pct_change_avg'] > 3 and is_reg_buy_from_filter(regression_data, 'filter_tech_all_pct_change_avg', 'filter_tech_all_pct_change_count', 'filter_tech_all_pct_change_pct'))
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Buy-Reg-GT0.75')
            Flag = True
            
            
            
    if("MLSell" in regression_data['filter'] and len(regression_data['filter']) > 9):
        if(is_reg_buy_from_filter(regression_data, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
            or is_reg_sell_from_filter(regression_data, 'filter_avg', 'filter_count', 'filter_pct')
            or is_reg_sell_from_filter(regression_data, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
            or is_reg_sell_from_filter(regression_data, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
            or is_reg_sell_from_filter(regression_data, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
            or is_reg_sell_from_filter(regression_data, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
            or (regression_data['filter_tech_all_pct_change_avg'] < -3 and is_reg_sell_from_filter(regression_data, 'filter_tech_all_pct_change_avg', 'filter_tech_all_pct_change_count', 'filter_tech_all_pct_change_pct'))
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Sell-Reg-LT-0.75')
            Flag = True
    return Flag
        
def is_reg_buy_from_filter(regression_data, filter_avg, filter_count, filter_pct):
    if(regression_data[filter_avg] > 0.75
        and regression_data[filter_count] > 1
        and regression_data[filter_pct] >= 90
        ):
        return True
    else:
        return False
   
def is_reg_sell_from_filter(regression_data, filter_avg, filter_count, filter_pct):
    if(regression_data[filter_avg] < -0.75
        and regression_data[filter_count] > 1
        and regression_data[filter_pct] <= -90
        ):
        return True
    else:
        return False
        
def is_filter_all_accuracy(regression_data, regression_high, regression_low, regressionResult, reg, ws):
    superFlag = False
    if(abs(regression_data['filter_345_avg']) > 2):
        flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
        if(flag):
            superFlag = True
    flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
    if(flag):
        superFlag = True
    flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_avg', 'filter_count', 'filter_pct')
    if(flag):
        superFlag = True
    flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
    if(flag):
        superFlag = True
    if(abs(regression_data['filter_tech_avg']) > 2):
        flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
        if(flag):
            superFlag = True
    if(abs(regression_data['filter_tech_all_avg']) > 2):
        flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
        if(flag):
            superFlag = True
    if(abs(regression_data['filter_tech_all_pct_change_avg']) > 5):
        flag = filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, 'filter_tech_all_pct_change_avg', 'filter_tech_all_pct_change_count', 'filter_tech_all_pct_change_pct')
        if(flag):
            superFlag = True
    
    flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
    if(flag):
        superFlag = True
    flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
    if(flag):
        superFlag = True
    flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_avg', 'filter_count', 'filter_pct')
    if(flag):
        superFlag = True
    flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
    if(flag):
        superFlag = True
    if(abs(regression_data['filter_tech_avg']) > 2):    
        flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
        if(flag):
            superFlag = True
    if(abs(regression_data['filter_tech_all_avg']) > 2): 
        flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
        if(flag):
            superFlag = True
    if(abs(regression_data['filter_tech_all_pct_change_avg']) > 5):
        flag = filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, 'filter_tech_all_pct_change_avg', 'filter_tech_all_pct_change_count', 'filter_tech_all_pct_change_pct')
        if(flag):
            superFlag = True
    
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_avg', 'filter_count', 'filter_pct')
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
    is_filter_risky(regression_data, regressionResult, reg, ws, 'filter_tech_all_pct_change_avg', 'filter_tech_all_pct_change_count', 'filter_tech_all_pct_change_pct')
    
    
    if(buy_af_high_volatility(regression_data, regressionResult, reg, ws)):
        superFlag = True
    if(sell_af_high_volatility(regression_data, regressionResult, reg, ws)):
        superFlag = True 
        
    if(superFlag):
        return superFlag       
    
def is_buy_from_filter_all_filter_relaxed(regression_data):
    if(is_buy_from_filter_relaxed(regression_data, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
        or is_buy_from_filter_relaxed(regression_data, 'filter_avg', 'filter_count', 'filter_pct')
        or is_buy_from_filter_relaxed(regression_data, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
        or is_buy_from_filter_relaxed(regression_data, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
        or is_buy_from_filter_relaxed(regression_data, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
        or is_buy_from_filter_relaxed(regression_data, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
        ):
        return True
    else:
        return False
       
def is_sell_from_filter_all_filter_relaxed(regression_data):
    if(is_sell_from_filter_relaxed(regression_data, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
        or is_sell_from_filter_relaxed(regression_data, 'filter_avg', 'filter_count', 'filter_pct')
        or is_sell_from_filter_relaxed(regression_data, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
        or is_sell_from_filter_relaxed(regression_data, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
        or is_sell_from_filter_relaxed(regression_data, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
        or is_sell_from_filter_relaxed(regression_data, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
        ):
        return True
    else:
        return False                      
                       
def is_buy_from_filter_relaxed(regression_data, filter_avg, filter_count, filter_pct):
    if((regression_data[filter_avg] > 1 and regression_data[filter_pct] == 0)
       or (regression_data[filter_avg] > 1 and regression_data[filter_pct] >= 80)
       or (regression_data[filter_avg] > 0.75 and regression_data[filter_pct] >= 70 and regression_data[filter_count] >= 3)
       ):
        return True
    else:
        return False
   
def is_sell_from_filter_relaxed(regression_data, filter_avg, filter_count, filter_pct):
    if((regression_data[filter_avg] < -1 and regression_data[filter_pct] == 0)
       or (regression_data[filter_avg] < -1 and regression_data[filter_pct] <= -80)
       or (regression_data[filter_avg] < -0.75 and regression_data[filter_pct] <= -70 and regression_data[filter_count] >= 3)
       ):
        return True
    else:
        return False

def is_buy_from_filter_all_filter(regression_data):
    if(is_buy_from_filter(regression_data, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
        or is_buy_from_filter(regression_data, 'filter_avg', 'filter_count', 'filter_pct')
        or is_buy_from_filter(regression_data, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
        or is_buy_from_filter(regression_data, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
        or is_buy_from_filter(regression_data, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
        or is_buy_from_filter(regression_data, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
        ):
        return True
    else:
        return False
       
def is_sell_from_filter_all_filter(regression_data):
    if(is_sell_from_filter(regression_data, 'filter_345_avg', 'filter_345_count', 'filter_345_pct')
        or is_sell_from_filter(regression_data, 'filter_avg', 'filter_count', 'filter_pct')
        or is_sell_from_filter(regression_data, 'filter_pct_change_avg', 'filter_pct_change_count', 'filter_pct_change_pct')
        or is_sell_from_filter(regression_data, 'filter_all_avg', 'filter_all_count', 'filter_all_pct')
        or is_sell_from_filter(regression_data, 'filter_tech_avg', 'filter_tech_count', 'filter_tech_pct')
        or is_sell_from_filter(regression_data, 'filter_tech_all_avg', 'filter_tech_all_count', 'filter_tech_all_pct')
        ):
        return True
    else:
        return False                      
                       
def is_buy_from_filter(regression_data, filter_avg, filter_count, filter_pct):
    if((regression_data[filter_avg] > 3 and regression_data[filter_pct] == 0)
       or (regression_data[filter_avg] > 1 and regression_data[filter_pct] >= 80)
       or (regression_data[filter_avg] > 0.75 and regression_data[filter_pct] >= 70 and regression_data[filter_count] >= 3)
       ):
        return True
    else:
        return False
   
def is_sell_from_filter(regression_data, filter_avg, filter_count, filter_pct):
    if((regression_data[filter_avg] < -3 and regression_data[filter_pct] == 0)
       or (regression_data[filter_avg] < -1 and regression_data[filter_pct] <= -80)
       or (regression_data[filter_avg] < -0.75 and regression_data[filter_pct] <= -70 and regression_data[filter_count] >= 3)
       ):
        return True
    else:
        return False
        
def is_buy_filter_not_risky(regression_data):
    if((high_tail_pct(regression_data) < 2 or high_tail_pct(regression_data) > 4
            or (regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0)
        )
#         and (high_tail_pct(regression_data) < 1 
#             or (regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0)
#             or ("MLBuy" in regression_data['filter'])
#         )
       ):
        return True

def is_sell_filter_not_risky(regression_data):
    if((low_tail_pct(regression_data) < 2 or low_tail_pct(regression_data) > 4
             or (regression_data['PCT_day_change'] > 0 and regression_data['PCT_change'] > 0)
             )
#         and (low_tail_pct(regression_data) < 1
#              or (regression_data['PCT_day_change'] > 0 and regression_data['PCT_change'] > 0)
#              or ("MLSell" in regression_data['filter'])
#              )
       ):
        return True
     
def filter_accuracy_finder(regression_data, regression_high, regression_low, regressionResult, reg, ws, filter_avg, filter_count, filter_pct):
    if(abs(regression_data[filter_avg]) > 0.5
        and regression_data[filter_count] >= 1
        #and (regression_data[filter_count_oth] >= 2
        #    or (regression_data[filter_count_oth] >= 1 and abs(regression_data[filter_avg_oth]) > 2))
        ):
        if((("MLSell" in regression_data['filter']) and float(regression_data[filter_avg]) > 0 and abs(float(regression_data[filter_pct])) < 80)
            or (("MLBuy" in regression_data['filter']) and float(regression_data[filter_avg]) < 0 and abs(float(regression_data[filter_pct])) < 80)
            ):
            return False
                     
        
        if((low_tail_pct(regression_data) < 1.5 and high_tail_pct(regression_data) < 2.5 and regression_data[filter_avg] < 0)
            or (high_tail_pct(regression_data) < 1.5 and low_tail_pct(regression_data) < 2.5 and regression_data[filter_avg] > 0)
            ):
            if(regression_data[filter_count] >= 5
                and abs(regression_data[filter_pct]) >= 100
                and abs(regression_data[filter_avg]) > 1.5
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-0-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-0-Sell')
            elif(regression_data[filter_count] >= 3
                and abs(regression_data[filter_pct]) >= 100
                and abs(regression_data[filter_avg]) > 2.5
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-0-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-0-Sell')
            elif(regression_data[filter_count] >= 2
                and abs(regression_data[filter_pct]) >= 100
                and abs(regression_data[filter_avg]) > 5
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-0-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-0-Sell')
                    
            if(regression_data[filter_count] >= 3
                and abs(regression_data[filter_pct]) >= 100
                and abs(regression_data[filter_avg]) > 2
                ):
                if(regression_data[filter_avg] >= 0
                    and is_algo_sell(regression_high) != True
                    and is_algo_sell(regression_low) != True
                    #and (is_algo_buy(regression_high) == True or is_algo_buy(regression_low) == True)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-1-Buy')
                elif(regression_data[filter_avg] < 0
                    and is_algo_buy(regression_high) != True
                    and is_algo_buy(regression_low) != True
                    #and (is_algo_sell(regression_high) == True or is_algo_sell(regression_low) == True)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-1-Sell')
                    
            if(regression_data[filter_count] >= 3
                and abs(regression_data[filter_pct]) > 80
                and abs(regression_data[filter_avg]) > 3
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-2-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-2-Sell')   
            elif(regression_data[filter_count] >= 3
                and abs(regression_data[filter_pct]) > 85
                and abs(regression_data[filter_avg]) > 2.5
                and abs(regression_data[filter_count] * regression_data[filter_avg]) > 9
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-3-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-3-Sell')
            elif(regression_data[filter_count] >= 3
                and abs(regression_data[filter_pct]) >= 85
                and abs(regression_data[filter_avg]) >= 2.0
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-4-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-4-Sell')
            elif(regression_data[filter_count] >= 2
                and abs(regression_data[filter_pct]) >= 85
                and abs(regression_data[filter_avg]) >= 3
                ):
                if(regression_data[filter_avg] >= 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-5-Buy')
                elif(regression_data[filter_avg] < 0):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-5-Sell')
            elif(regression_data[filter_count] >= 2
                and((abs(regression_data[filter_pct]) >= 100 and abs(regression_data[filter_avg]) >= 2.5)
                     or (abs(regression_data[filter_pct]) >= 70 and abs(regression_data[filter_avg]) >= 3.5)
                     )
                ):
                if(regression_data[filter_avg] >= 0 and is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-6-Buy')
                elif(regression_data[filter_avg] < 0 and is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-6-Sell')
            elif(regression_data[filter_count] > 5
                and abs(regression_data[filter_pct]) > 60
                and abs(regression_data[filter_avg]) > 1.5
                ):
                if(regression_data[filter_avg] >= 0
                    and ((abs(regression_data[filter_avg]) > 2 and abs(regression_data[filter_pct]) > 80) or is_algo_buy(regression_data))
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-Risky-Buy')
                elif(regression_data[filter_avg] < 0
                    and ((abs(regression_data[filter_avg]) > 2 and abs(regression_data[filter_pct]) > 80) or is_algo_sell(regression_data))
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-Risky-Sell')
            elif(regression_data[filter_count] >= 2
                and abs(regression_data[filter_pct]) >= 70
                and abs(regression_data[filter_avg]) >= 2.0
                ):
                if(regression_data[filter_avg] >= 0 and is_algo_buy(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-7-Buy')
                elif(regression_data[filter_avg] < 0 and is_algo_sell(regression_data)):
                    add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-7-Sell')       
                
        Flag = False
        if(abs(float(regression_data[filter_avg])) > 1.5 
            and abs(regression_data[filter_pct]) > 50
            ):
            Flag = True
            
        if(("MLBuy" in regression_data['filter']) 
            and regression_data[filter_avg] >= 1.5
            and (regression_data[filter_pct] >= 80 or regression_data[filter_pct] == 0)
            ):
            Flag = True
            
        if(("MLSell" in regression_data['filter']) 
            and regression_data[filter_avg] <= -1.5
            and (regression_data[filter_pct] <= -80 or regression_data[filter_pct] == 0)
            ):
            Flag = True
                    
        #is_filter_risky(regression_data, regressionResult, reg, ws, filter_avg, filter_count, filter_pct)
        return Flag            

def filter_accuracy_finder_risky(regression_data, regressionResult, reg, ws, filter_avg, filter_count, filter_pct):
    if(abs(regression_data[filter_avg]) >= 0.5
        and regression_data[filter_count] >= 1
        #and (regression_data[filter_count_oth] >= 2
        #    or (regression_data[filter_count_oth] >= 1 and abs(regression_data[filter_avg_oth]) > 2))
        ):
        Flag = False
        if(regression_data[filter_count] >= 3
            and abs(regression_data[filter_pct]) >= 85
            and abs(regression_data[filter_avg]) > 5
            ):
            if(regression_data[filter_avg] >= 0):
                add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-SUPER-Buy')
            elif(regression_data[filter_avg] < 0):
                add_in_csv(regression_data, regressionResult, ws, None, 'STRONG-SUPER-Sell')
            Flag = True
            
        if((regression_data[filter_count] >= 3
            and abs(regression_data[filter_pct]) >= 80
            and abs(regression_data[filter_avg]) >= 0.6
            )
            or 
            (regression_data[filter_count] >= 2
            and abs(regression_data[filter_pct]) == 100
            and abs(regression_data[filter_avg]) >= 1
            )
            or 
            (regression_data[filter_count] >= 2
            and abs(regression_data[filter_pct]) == 100
            and regression_data[filter_avg] >= 0.5
            and "MLBuy" in regression_data['filter']
            )
            or 
            (regression_data[filter_count] >= 2
            and abs(regression_data[filter_pct]) == 100
            and regression_data[filter_avg] <= -0.5
            and "MLSell" in regression_data['filter'])
            ):
            if('%%mayBuyTail' in regression_data['filter']):
                if(regression_data[filter_avg] < 0 
                    and ("MLBuy" not in regression_data['filter'])
                    and (('upTrend' in regression_data['series_trend'])
                        or ('downTrend' in regression_data['series_trend'])
                        or (pct_day_change_trend(regression_data) <= 0))
                    and abs(regression_data[filter_pct]) >= 90
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'Sell-%%mayBuyTail')
                elif(regression_data[filter_avg] > 0 and ("MLSell" not in regression_data['filter'])
                    and (pct_day_change_trend(regression_data) >= 0)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'Buy-%%mayBuyTail')
                elif((('upTrend-min3Day$NA:shortUpTrend-min3Day$NA' in regression_data['series_trend'])
                      or pct_day_change_trend(regression_data) >= 3
                    )
                    and regression_data['high'] < regression_data['high_pre1']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'BuyUpTrend-%%mayBuyTail')
                Flag = True
                
            if('%%maySellTail' in regression_data['filter']):
                if(regression_data[filter_avg] > 0 
                    and ("MLSell" not in regression_data['filter'])
                    and (('downTrend' not in regression_data['series_trend'])
                        or ('upTrend' in regression_data['series_trend'])
                        or (pct_day_change_trend(regression_data) >= 0))
                    and abs(regression_data[filter_pct]) >= 90
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'Buy-%%maySellTail')
                elif(regression_data[filter_avg] < 0 and ("MLBuy" not in regression_data['filter'])
                     and (pct_day_change_trend(regression_data) <= 0)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'Sell-%%maySellTail')
                elif((('downTrend-min3Day$NA:shortDownTrend-min3Day$NA' in regression_data['series_trend'])
                      or pct_day_change_trend(regression_data) <= -3
                      )
                    and regression_data['low'] > regression_data['low_pre1']
                    ):
                    add_in_csv(regression_data, regressionResult, ws, None, 'SellDownTrend-%%maySellTail')
                Flag = True                
        if(abs(float(regression_data[filter_avg])) > 1.5 
            and abs(regression_data[filter_pct]) > 50
            ):
            Flag = True
            
        if(("MLBuy" in regression_data['filter']) 
            and regression_data[filter_avg] >= 1.5
            and (regression_data[filter_pct] >= 80 or regression_data[filter_pct] == 0)
            ):
            Flag = True
            
        if(("MLSell" in regression_data['filter']) 
            and regression_data[filter_avg] <= -1.5
            and (regression_data[filter_pct] <= -80 or regression_data[filter_pct] == 0)
            ):
            Flag = True
        
        #is_filter_risky(regression_data, regressionResult, reg, ws, filter_avg, filter_count, filter_pct)  
        buyRisky, sellRisky =  is_filter_risky(regression_data, regressionResult, reg, ws, filter_avg, filter_count, filter_pct, False)
        if(len(regression_data['filter']) > 9 
            and ((regression_data[filter_avg] >= 0.75 and regression_data[filter_count] >= 3 and regression_data[filter_pct] > 100 and regression_data['PCT_day_change'] < 2)
                 or (regression_data[filter_avg] >= 1.5 and regression_data[filter_count] >= 5 and regression_data[filter_pct] > 70)
                 or (regression_data[filter_avg] >= 2 and regression_data[filter_count] >= 4 and regression_data[filter_pct] > 80)
                 or (regression_data[filter_avg] >= 2.5 and regression_data[filter_count] >= 2 and regression_data[filter_pct] >= 90))
            and ("MLSell" not in regression_data['filter'])
            and (buyRisky == False or ("MLBuy" in regression_data['filter']))
            and is_buy_filter_not_risky(regression_data)
            and high_tail_pct(regression_data) < 1.5 and low_tail_pct(regression_data) < 2.5
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Buy')
            Flag = True
        
        if(len(regression_data['filter']) > 9
            and ((regression_data[filter_avg] <= -0.75 and regression_data[filter_count] >= 3 and regression_data[filter_pct] < -100 and regression_data['PCT_day_change'] > -2)
                 or (regression_data[filter_avg] <= -1.5 and regression_data[filter_count] >= 5 and regression_data[filter_pct] < -70)
                 or (regression_data[filter_avg] <= -2 and regression_data[filter_count] >= 4 and regression_data[filter_pct] < -80)
                 or (regression_data[filter_avg] <= -2.5 and regression_data[filter_count] >= 2 and regression_data[filter_pct] <= -90))
            and ("MLBuy" not in regression_data['filter'])
            and (sellRisky == False or ("MLSell" in regression_data['filter']))
            and is_sell_filter_not_risky(regression_data)
            and low_tail_pct(regression_data) < 1.5 and high_tail_pct(regression_data) < 2.5
            ):
            add_in_csv(regression_data, regressionResult, ws, None, 'Filter-Sell')
            Flag = True  
        return Flag
          
def is_filter_risky(regression_data, regressionResult, reg, ws, filter_avg, filter_count, filter_pct, update=True): 
    BuyRisky = False
    SellRisky = False
    if(0 < float(regression_data['PCT_day_change']) < 1
        and -3 < float(regression_data['PCT_change']) < 3
        and ("shortDownTrend" in regression_data['series_trend']
             or "downTrend" in regression_data['series_trend']
            )
        and float(regression_data[filter_avg]) > 0.5
        ):
        if update:
            add_in_csv(regression_data, regressionResult, ws, None, 'RISKYDOJI-Buy')
        BuyRisky = True
        
    if(-1 < float(regression_data['PCT_day_change']) < 0
        and -3 < float(regression_data['PCT_change']) < 3 
        and ("downTrend" in regression_data['series_trend'])
        and float(regression_data[filter_avg]) > 0.5
        ):
        if update:
            add_in_csv(regression_data, regressionResult, ws, None, 'RISKYDOJI-Buy')
        BuyRisky = True
    
    if(-1 < float(regression_data['PCT_day_change']) < 0
        and -3 < float(regression_data['PCT_change']) < 3 
        and (("shortUpTrend" in regression_data['series_trend'])
             or ("upTrend" in regression_data['series_trend'])
            )
        and float(regression_data[filter_avg]) < -0.5
        ):
        if update:
            add_in_csv(regression_data, regressionResult, ws, None, 'RISKYDOJI-Sell')
        SellRisky = True
    
    if(0 < float(regression_data['PCT_day_change']) < 1
        and -3 < float(regression_data['PCT_change']) < 3 
        and ("upTrend" in regression_data['series_trend'])
        and float(regression_data[filter_avg]) < -0.5
        ):
        if update:
            add_in_csv(regression_data, regressionResult, ws, None, 'RISKYDOJI-Sell')
        SellRisky = True
        
    if(abs(float(regression_data['PCT_day_change'])) < 0.7
        and -3 < float(regression_data['PCT_change']) < 3
        and regression_data['series_trend'] == "NA$NA:NA$NA"
        ):
        if update:
            add_in_csv(regression_data, regressionResult, ws, None, 'RISKYDOJI')
        BuyRisky = True
        SellRisky = True
    
    if(((regression_data['filter'] == " " and abs(float(regression_data[filter_avg])) > 0.75)
            or ("MLBuy" in regression_data['filter'] and float(regression_data[filter_avg]) < -0.5)
            or ("MLSell" in regression_data['filter'] and float(regression_data[filter_avg]) > 0.5)
            )
        and regression_data['filter4'].startswith('@s@')
        and regression_data['filter4'].endswith('@e@,') 
        ):
        if update:
            add_in_csv(regression_data, regressionResult, ws, None, 'RISKYFilter')  
        
    return BuyRisky, SellRisky  
    
        
    
        
        
        
        
        

            