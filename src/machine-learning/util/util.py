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

buyMLP = 1
buyMLP_MIN = 0
buyKN = 1
buyKN_MIN = 0
sellMLP = -1
sellMLP_MIN = 0
sellKN = -1
sellKN_MIN = 0

def add_in_csv(regression_data, regressionResult, ws, filter):
    if ((filter is not None) and (filter not in regression_data['filter'])):
        regression_data['filter'] = regression_data['filter'] + filter + ','
    tempRegressionResult = regressionResult.copy() 
    tempRegressionResult.append(regression_data['filter'])
    ws.append(tempRegressionResult) if (ws is not None) else False
    if(db.resultScripFutures.find_one({'scrip':regression_data['scrip']}) is None):
        db.resultScripFutures.insert_one({
            "scrip": regression_data['scrip'],
            "date": regression_data['date']
            })

def add_in_csv_hist_pattern(regression_data, regressionResult, ws, filter, avg, count):
    if ((filter is not None) and (filter not in regression_data['filter'])):
        regression_data['filter'] = regression_data['filter'] + filter + ','
    tempRegressionResult = regressionResult.copy() 
    tempRegressionResult.append(regression_data['filter'])
    tempRegressionResult.append(avg)
    tempRegressionResult.append(count)
    ws.append(tempRegressionResult) if (ws is not None) else False

def is_algo_buy(regression_data):
    if((regression_data['mlpValue'] >= buyMLP and regression_data['kNeighboursValue'] >= buyKN_MIN) 
        or (regression_data['mlpValue'] >= buyMLP_MIN and regression_data['kNeighboursValue'] >= buyKN)):
        return True
    else:
        return False   
    
def is_algo_sell(regression_data):
    if((regression_data['mlpValue'] <= sellMLP and regression_data['kNeighboursValue'] <= sellKN_MIN)
       or (regression_data['mlpValue'] <= sellMLP_MIN and regression_data['kNeighboursValue'] <= sellKN)):   
        return True
    else:
        return False
    
def is_algo_buy_classifier(regression_data):
    if((regression_data['mlpValue'] >= 0 and regression_data['kNeighboursValue'] >= 3) 
        or (regression_data['mlpValue'] >= 3 and regression_data['kNeighboursValue'] >= 0)
        or (regression_data['mlpValue'] >= 1 and regression_data['kNeighboursValue'] >= 1)
        ):
        return True
    else:
        return False   
    
def is_algo_sell_classifier(regression_data):
    if((regression_data['mlpValue'] <= 0 and regression_data['kNeighboursValue'] <= -3) 
        or (regression_data['mlpValue'] <= -3 and regression_data['kNeighboursValue'] <= 0)
        or (regression_data['mlpValue'] <= -1 and regression_data['kNeighboursValue'] <= -1)
        ):   
        return True
    else:
        return False    

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
    
def all_day_pct_change_positive(regression_data):
    if(regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > -0.5
        and regression_data['forecast_day_PCT3_change'] > -0.5
        and regression_data['forecast_day_PCT4_change'] > -0.5
        and regression_data['forecast_day_PCT5_change'] > -0.5
        and regression_data['forecast_day_PCT7_change'] > -0.5
        and regression_data['forecast_day_PCT10_change'] > -0.5):
        return True; 
    
def pct_change_negative_trend(regression_data):
    if (regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < 0
        and regression_data['forecast_day_PCT3_change'] < 0):
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
        return trend            
    return False           
    
def pct_change_positive_trend(regression_data):
    if (regression_data['forecast_day_PCT_change'] > 0
        and regression_data['forecast_day_PCT2_change'] > 0
        and regression_data['forecast_day_PCT3_change'] > 0):
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
        return trend
    return False       
        
def historical_data_OI(data):
    ardate = np.array([str(x) for x in (np.array(data['data'])[:,3][::-1]).tolist()])
    aroipctchange = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,6][::-1]).tolist()])
    arcontractpctchange = np.array([float(x.encode('UTF8')) for x in (np.array(data['data'])[:,7][::-1]).tolist()])
    return ardate, aroipctchange, arcontractpctchange

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

def get_regressionResult(regression_data, scrip, db, mlp_o, kneighbours_o):
    regression_data['filter'] = " "
    regression_data['series_trend'] = "NA"
    if pct_change_negative_trend(regression_data):
        regression_data['series_trend'] = "downTrend"
    if pct_change_positive_trend(regression_data):
         regression_data['series_trend'] = "upTrend"    
    
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
        result_data = db.scrip_result.find_one({'scrip':scrip.replace('&','').replace('-','_')})
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
    regressionResult.append(regression_data['scrip'])
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
    regressionResult.append(regression_data['PCT_day_change'])
    regressionResult.append(regression_data['PCT_change'])
    regressionResult.append(regression_data['score'])
    regressionResult.append(regression_data['mlpValue'])
    regressionResult.append(regression_data['kNeighboursValue'])
    regressionResult.append(mlp_o)
    regressionResult.append(kneighbours_o)
    regressionResult.append(regression_data['trend'])
    regressionResult.append(regression_data['yearHighChange'])
    regressionResult.append(regression_data['yearLowChange'])
    regressionResult.append(regression_data['series_trend'])
    regressionResult.append(resultDate)
    regressionResult.append(resultDeclared)
    regressionResult.append(resultSentiment)
    regressionResult.append(resultComment)
    return regressionResult

def all_withoutml(regression_data, regressionResult, ws):
    add_in_csv(regression_data, regressionResult, ws, '')

def buy_pattern_without_mlalgo(regression_data, regressionResult, ws_buyPattern2, ws_sellPattern2):
    buyPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/all-buy-filter-by-PCT-Change.csv')
    if regression_data['buyIndia'] != '' and regression_data['buyIndia'] in buyPatternsDict:
        if (abs(float(buyPatternsDict[regression_data['buyIndia']]['avg'])) >= .1 and float(buyPatternsDict[regression_data['buyIndia']]['count']) >= 2):
            if(-0.5 < regression_data['PCT_day_change'] < 3 and float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 1):
                avg = buyPatternsDict[regression_data['buyIndia']]['avg']
                count = buyPatternsDict[regression_data['buyIndia']]['count']
                add_in_csv_hist_pattern(regression_data, regressionResult, ws_buyPattern2, 'wml_buy', avg, count)
            elif(-3 < regression_data['PCT_day_change'] < 0.5 and float(buyPatternsDict[regression_data['buyIndia']]['avg']) < -1):
                avg = buyPatternsDict[regression_data['buyIndia']]['avg']
                count = buyPatternsDict[regression_data['buyIndia']]['count']
                add_in_csv_hist_pattern(regression_data, regressionResult, ws_sellPattern2, 'wml_buy', avg, count)

def buy_pattern_from_history(regression_data, regressionResult, ws_buyPattern2):
    buyPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/patterns-buy.csv')
    buyIndiaAvg = 0
    flag = False
    if regression_data['buyIndia'] != '' and regression_data['buyIndia'] in buyPatternsDict:
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
                       add_in_csv_hist_pattern(regression_data, regressionResult, ws_buyPattern2, 'buyPattern2', avg, count)
                    elif(float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 0.5 
                       or (float(buyPatternsDict[regression_data['buyIndia']]['avg']) > 0.3 and (regression_data['forecast_day_PCT10_change'] < -10 or regression_data['yearHighChange'] < -40))):
                        if(regression_data['forecast_day_PCT10_change'] < 0 and regression_data['forecast_day_PCT_change'] >= 0):
                            flag = True
                            add_in_csv_hist_pattern(regression_data, regressionResult, ws_buyPattern2, 'buyPattern2', avg, count)
                        elif(regression_data['forecast_day_PCT10_change'] > 0):    
                            flag = True
                            add_in_csv_hist_pattern(regression_data, regressionResult, ws_buyPattern2, 'buyPattern2', avg, count)     
    return buyIndiaAvg, flag

def buy_all_rule(regression_data, regressionResult, buyIndiaAvg, ws_buyAll):
    if(is_algo_buy(regression_data)
        and (regression_data['high']-regression_data['bar_high']) < (regression_data['bar_high']-regression_data['bar_low'])
        and 'P@[' not in str(regression_data['sellIndia'])
        and buyIndiaAvg >= -.70):
        add_in_csv(regression_data, regressionResult, ws_buyAll, None)
        if(0 < regression_data['yearLowChange'] < 10):
            if(regression_data['PCT_day_change'] < -1):
                return True
        else:
            return True
    return False

def buy_all_rule_classifier(regression_data, regressionResult, buyIndiaAvg, ws_buyAll):
    if(is_algo_buy_classifier(regression_data)
        and (regression_data['high']-regression_data['bar_high']) < (regression_data['bar_high']-regression_data['bar_low'])
        and 'P@[' not in str(regression_data['sellIndia'])
        and buyIndiaAvg >= -.5):
        add_in_csv(regression_data, regressionResult, ws_buyAll, None)
        if(0 < regression_data['yearLowChange'] < 10):
            if(regression_data['PCT_day_change'] < -1):
                return True
        else:
            return True
    return False

def buy_year_high(regression_data, regressionResult, ws_buyYearHigh):
    if(float(regression_data['forecast_day_VOL_change']) > 50
       and(regression_data['PCT_day_change_pre'] > 0 
           or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
           )
       and no_doji_or_spinning_buy_india(regression_data)
       ):
        if(-5 <= regression_data['yearHighChange'] < -1 and regression_data['yearLowChange'] > 30 
            and 0 < regression_data['PCT_day_change'] < 2 and regression_data['forecast_day_PCT2_change'] <= 2
            ):
            add_in_csv(regression_data, regressionResult, ws_buyYearHigh, 'buyYearHigh')
            return True
    if(float(regression_data['forecast_day_VOL_change']) > 50
       and(regression_data['PCT_day_change_pre'] > 0 
           or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
           )
       and no_doji_or_spinning_buy_india(regression_data)
       ):
        if(-15 < regression_data['yearHighChange'] < -5 and regression_data['yearLowChange'] > 30
            and 0 < regression_data['PCT_day_change'] < 2 and regression_data['forecast_day_PCT2_change'] <= 2
            ):
            add_in_csv(regression_data, regressionResult, ws_buyYearHigh, 'buyYearHigh1')
            return True
    return False

def buy_year_low(regression_data, regressionResult, ws_buyYearLow, ws_buyYearLow1):
    if(1 < regression_data['yearLowChange'] < 10 and regression_data['yearHighChange'] < -30 
        and 0.75 < regression_data['PCT_day_change'] < 5
        and regression_data['forecast_day_PCT10_change'] <= -10 and regression_data['forecast_day_PCT7_change'] < -5 and regression_data['forecast_day_PCT5_change'] < 0.5 and regression_data['forecast_day_PCT4_change'] < 0.5 
        and regression_data['forecast_day_PCT2_change'] > 0 and regression_data['forecast_day_PCT_change'] > 0):
        add_in_csv(regression_data, regressionResult, ws_buyYearLow, 'buyYearLow')
        return True
    elif(0 < regression_data['yearLowChange'] < 15 and regression_data['yearHighChange'] < -25 
        and (5 > regression_data['PCT_day_change'] > 0.75 and regression_data['PCT_change'] > -0.5)
        and regression_data['forecast_day_PCT10_change'] <= -5 and regression_data['forecast_day_PCT7_change'] < -3 and regression_data['forecast_day_PCT5_change'] < 0.5
        and 5 > regression_data['forecast_day_PCT2_change'] > -0.5 and regression_data['forecast_day_PCT_change'] > 0):
        add_in_csv(regression_data, regressionResult, ws_buyYearLow1, 'buyYearLow1')
        return True
    elif(0 < regression_data['yearLowChange'] < 15 and regression_data['yearHighChange'] < -25 
        and (5 > regression_data['PCT_day_change'] > 0.75 and regression_data['PCT_change'] > -0.5 and regression_data['PCT_day_change'] >= regression_data['PCT_change'])
        and regression_data['forecast_day_PCT10_change'] < -5 and regression_data['forecast_day_PCT7_change'] < 0 and regression_data['forecast_day_PCT5_change'] < 0
        and 5 > regression_data['forecast_day_PCT2_change'] > -0.5 and regression_data['forecast_day_PCT_change'] > 0):
        add_in_csv(regression_data, regressionResult, ws_buyYearLow1, 'buyYearLow1')
        return True
    return False

def buy_up_trend(regression_data, regressionResult, ws_buyUpTrend):
    if((regression_data['yearHighChange'] < -9 or regression_data['yearLowChange'] < 15)
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['forecast_day_PCT2_change'] > 0
       and regression_data['forecast_day_PCT3_change'] > 0    
       ):
        if(2 < regression_data['PCT_day_change'] < 3 and 2 < regression_data['forecast_day_PCT_change'] < 3
           and regression_data['PCT_change'] < regression_data['PCT_day_change'] + .2
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] + .2
           and regression_data['series_trend'] != 'downTrend'
           and regression_data['score'] != '0-1' 
           and regression_data['forecast_day_PCT7_change'] < 0 and regression_data['forecast_day_PCT10_change'] < 0
           ):
            add_in_csv(regression_data, regressionResult, None, 'upTrendCandidate-0')
            return True
        if(2 < regression_data['PCT_day_change'] < 3 and 2 < regression_data['forecast_day_PCT_change'] < 3
           and regression_data['PCT_change'] < regression_data['PCT_day_change'] + .2
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] + .2
           and regression_data['series_trend'] != 'downTrend'
           and regression_data['score'] != '0-1' 
           and regression_data['forecast_day_PCT7_change'] < 10 and regression_data['forecast_day_PCT10_change'] < 10
           ):
            add_in_csv(regression_data, regressionResult, None, 'upTrendCandidate-00')
            return True
        if(2 < regression_data['PCT_day_change'] < 3 and regression_data['PCT_change'] < 4
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_change']
           and regression_data['series_trend'] != 'downTrend'
           and regression_data['score'] != '0-1'
           and (regression_data['forecast_day_PCT7_change'] < 8 or regression_data['forecast_day_PCT10_change'] < 8)
           ):
            add_in_csv(regression_data, regressionResult, None, 'upTrendCandidate-1')
            return True
    elif((regression_data['yearHighChange'] < -9 or regression_data['yearLowChange'] < 15)
       and regression_data['forecast_day_PCT_change'] > 0
       and regression_data['forecast_day_PCT2_change'] > 0
       and regression_data['forecast_day_PCT3_change'] > 0    
       ):
        if(2 < regression_data['PCT_day_change'] < 3 and regression_data['PCT_change'] < 3
               and regression_data['forecast_day_PCT_change'] < regression_data['PCT_change']
               and regression_data['series_trend'] != 'downTrend'
               and regression_data['score'] != '0-1'
               and (regression_data['forecast_day_PCT7_change'] < 5 or regression_data['forecast_day_PCT10_change'] < 7)
               ):
                add_in_csv(regression_data, regressionResult, None, 'upTrendCandidate-1-YearHigh')
                return True    
    if(all_day_pct_change_positive(regression_data) and 0 < regression_data['PCT_day_change'] < 5 and regression_data['yearHighChange'] < -10
        and regression_data['forecast_day_PCT10_change'] >= regression_data['PCT_change'] + 2
        and regression_data['forecast_day_PCT10_change'] >= regression_data['PCT_day_change'] + 2
        and float(regression_data['forecast_day_VOL_change']) > 30
        and regression_data['PCT_day_change_pre'] > -0.5 
        and float(regression_data['contract']) > 10
        and no_doji_or_spinning_buy_india(regression_data)):
        add_in_csv(regression_data, regressionResult, ws_buyUpTrend, 'buyUpTrend')
        return True
    return False

def buy_down_trend(regression_data, regressionResult, ws_buyDownTrend):
    if(all_day_pct_change_negative(regression_data) 
       and 0 < regression_data['PCT_day_change'] < 5 
       and regression_data['forecast_day_PCT10_change'] < -5
       and (
           regression_data['forecast_day_PCT5_change'] < -10
           or regression_data['forecast_day_PCT7_change'] < -10
           or regression_data['forecast_day_PCT10_change'] < -10
           )
       and regression_data['yearHighChange'] < -10):
        add_in_csv(regression_data, regressionResult, ws_buyDownTrend, 'buyDownTrend')
        return True
    return False

def buy_final(regression_data, regressionResult, ws_buyFinal, ws_buyFinal1):
    if(regression_data['yearHighChange'] < -10 and regression_data['score'] != '0-1'
        and 4 > regression_data['PCT_day_change'] > 1 and 4 > regression_data['PCT_change'] > 1):   
        if( str(regression_data['sellIndia']) == '' and -90 < regression_data['yearHighChange'] < -10
            and regression_data['forecast_day_PCT10_change'] <= -10 and regression_data['forecast_day_PCT7_change'] < -5 and regression_data['forecast_day_PCT5_change'] < 0.5 and regression_data['forecast_day_PCT4_change'] < 0.5 
            and 5 > regression_data['forecast_day_PCT2_change'] > -0.5 and regression_data['forecast_day_PCT_change'] > 0):
            add_in_csv(regression_data, regressionResult, ws_buyFinal, 'buyFinal')
            return True
        elif(regression_data['forecast_day_PCT5_change'] <= 1 and regression_data['forecast_day_PCT7_change'] <= -1 and regression_data['forecast_day_PCT10_change'] <= -10
            and regression_data['forecast_day_PCT2_change'] > 0 and regression_data['forecast_day_PCT_change'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws_buyFinal, 'buyFinal1')
            return True
    return False

def buy_high_indicators(regression_data, regressionResult, ws_buyHighIndicators):
    if(regression_data['mlpValue'] > 1.0 and regression_data['kNeighboursValue'] > 1.0 and regression_data['yearLowChange'] < 16 and regression_data['yearHighChange'] < -35
        and (2.5 > regression_data['PCT_day_change'] > 0 and 2.5 > regression_data['PCT_change'] > -0.5)):
        add_in_csv(regression_data, regressionResult, ws_buyHighIndicators, 'buyHighIndicators')
        return True
    return False
              
def buy_pattern(regression_data, regressionResult, ws_buyPattern, ws_buyPattern1):
    score = ''
    if(regression_data['score'] == '10' or regression_data['score'] == '1-1'):
        score = 'up'
    if(-1 < regression_data['PCT_day_change'] < 4 and regression_data['yearLowChange'] > 5 and regression_data['score'] != '0-1'
        #and regression_data['trend'] != 'up'
    ):
        if(('MARUBOZU' in str(regression_data['buyIndia']) and regression_data['forecast_day_PCT5_change'] <= 0 and regression_data['forecast_day_PCT10_change'] <= -5)
           or ('HAMMER' in str(regression_data['buyIndia']) and regression_data['PCT_day_change'] > 0)
           #or 'ENGULFING' in str(regression_data['buyIndia'])
           #or 'PIERCING' in str(regression_data['buyIndia'])
           or ('MORNINGSTAR' in str(regression_data['buyIndia']) and regression_data['forecast_day_PCT5_change'] <= 0 and regression_data['forecast_day_PCT10_change'] <= -5)
           #or ':DOJISTAR' in str(regression_data['buyIndia'])
           #or 'MORNINGDOJISTAR' in str(regression_data['buyIndia'])
           or 'ABANDONEDBABY' in str(regression_data['buyIndia'])
           or 'COUNTERATTACK' in str(regression_data['buyIndia'])
           or 'KICKING' in str(regression_data['buyIndia'])
           or 'BREAKAWAY' in str(regression_data['buyIndia'])
           #or 'TRISTAR' in str(regression_data['buyIndia'])
           #or '3WHITESOLDIERS' in str(regression_data['buyIndia'])
           #or '3INSIDE' in str(regression_data['buyIndia'])
           ):
            add_in_csv(regression_data, regressionResult, ws_buyPattern, 'buyPattern')
            return True
        elif(
           ('CCI:BOP' in str(regression_data['buyIndia']) and 'BELTHOLD' in str(regression_data['buyIndia']))
           or ('AROON:BOP' in str(regression_data['buyIndia']) and 'BELTHOLD' in str(regression_data['buyIndia']) and 'ENGULFING' in str(regression_data['buyIndia']))
           or ('BELTHOLD' == str(regression_data['buyIndia']) and score == 'up')
           #or ('3OUTSIDE' in str(regression_data['buyIndia']) and regression_data['forecast_day_PCT5_change'] <= 0 and score == 'up')
           #or ('HARAMI' in str(regression_data['buyIndia']) and regression_data['forecast_day_PCT5_change'] <= 0 and score == 'up')
           #or (regression_data['yearHighChange'] <= -35 and 'HARAMI' in str(regression_data['buyIndia']) and 'SHORTLINE' in str(regression_data['buyIndia']) and regression_data['PCT_day_change'] > 0)
           or ('DOJI' in str(regression_data['buyIndia']) and 'GRAVESTONEDOJI' in str(regression_data['buyIndia']) and 'LONGLEGGEDDOJI' in str(regression_data['buyIndia']) and regression_data['PCT_day_change'] > 0)
           #or ('P@[,HIKKAKE]' == str(regression_data['buyIndia']) and regression_data['PCT_day_change'] < 0)
           #or (regression_data['yearHighChange'] <= -35 and 'BELTHOLD' in str(regression_data['buyIndia']) and 'LONGLINE' in str(regression_data['buyIndia']))
           #or (regression_data['yearHighChange'] <= -35 and ',CCI:BOP' in str(regression_data['buyIndia']) and 'LONGLINE' in str(regression_data['buyIndia']))
           ) and ((regression_data['forecast_day_PCT5_change'] <= -5) or regression_data['yearHighChange'] < -50):
            add_in_csv(regression_data, regressionResult, ws_buyPattern1, 'buyPattern1')
            return True
        elif(
           ('MARUBOZU' in str(regression_data['buyIndia']) and regression_data['forecast_day_PCT5_change'] <= 0 and regression_data['forecast_day_PCT10_change'] <= 1)
           or ('HAMMER' in str(regression_data['buyIndia']) and regression_data['PCT_day_change'] > 0)
           or 'ENGULFING' in str(regression_data['buyIndia'])
           or 'PIERCING' in str(regression_data['buyIndia'])
           or ('MORNINGSTAR' in str(regression_data['buyIndia']) and regression_data['forecast_day_PCT5_change'] <= 0 and regression_data['forecast_day_PCT10_change'] <= 1)
           #or ':DOJISTAR' in str(regression_data['buyIndia'])
           or 'MORNINGDOJISTAR' in str(regression_data['buyIndia'])
           or 'ABANDONEDBABY' in str(regression_data['buyIndia'])
           or 'COUNTERATTACK' in str(regression_data['buyIndia'])
           or 'KICKING' in str(regression_data['buyIndia'])
           or 'BREAKAWAY' in str(regression_data['buyIndia'])
           or 'TRISTAR' in str(regression_data['buyIndia'])
           or '3WHITESOLDIERS' in str(regression_data['buyIndia'])
           or '3INSIDE' in str(regression_data['buyIndia'])
           ) and 'DOJI' not in str(regression_data['buyIndia']) and ((regression_data['forecast_day_PCT5_change'] <= -5) or regression_data['yearHighChange'] < -50): 
            add_in_csv(regression_data, regressionResult, ws_buyPattern1, 'buyPattern1')
            return True
    return False

def morning_star_sell(regression_data, regressionResult, ws):
    if(5 > regression_data['forecast_day_PCT_change'] > 2
        and regression_data['PCT_day_change_pre'] > 0
        ):
        if(-1 < regression_data['PCT_day_change'] < 0 and -1 < regression_data['PCT_change'] < 0 
            and regression_data['kNeighboursValue'] < 0
            ):
            add_in_csv(regression_data, regressionResult, ws, 'msSellCandidate-0')
            return True
        elif(1 > regression_data['PCT_day_change'] > 0 and 1 > regression_data['PCT_change'] > 0
            and (regression_data['high']-regression_data['open']) > regression_data['PCT_day_change'] * 3
            and (regression_data['open']-regression_data['low']) < regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'msSellCandidate-1')
            return True
        elif(1 > regression_data['PCT_day_change'] > 0 and 1 > regression_data['PCT_change'] > 0
            and regression_data['forecast_day_PCT_change'] > regression_data['PCT_day_change'] * 3
            and regression_data['forecast_day_PCT10_change'] >= 10
            ):
            add_in_csv(regression_data, regressionResult, ws, 'msSellCandidate-2')
            return True   
    return False

def buy_oi_negative(regression_data, regressionResult, ws):
    if(regression_data['redtrend'] == -1
        and regression_data['forecast_day_PCT_change'] < 0
        and regression_data['forecast_day_PCT2_change'] < -0.5
        and regression_data['forecast_day_PCT3_change'] < -0.5
        and regression_data['forecast_day_PCT4_change'] < -0.5
        and regression_data['forecast_day_PCT5_change'] < -0.5
        and regression_data['forecast_day_PCT7_change'] < -0.5
        and regression_data['forecast_day_PCT10_change'] < -5
        and (regression_data['forecast_day_PCT5_change'] < -10
             or regression_data['forecast_day_PCT7_change'] < -10
             or regression_data['forecast_day_PCT10_change'] < -10
        )
        and float(regression_data['forecast_day_VOL_change']) < -30
        and regression_data['PCT_day_change_pre'] < 0 
        and float(regression_data['contract']) < 0
        and float(regression_data['oi']) < 5
        and (regression_data['yearHighChange'] < -15 or regression_data['yearLowChange'] < 15)
        and ((regression_data['mlpValue'] > 0 and regression_data['kNeighboursValue'] > 0) or is_algo_buy(regression_data))
        ):
        if(((-1 < regression_data['PCT_day_change'] < 0 and -1 < regression_data['PCT_change'] < -0.5)
            or (-1 < regression_data['PCT_day_change'] < -0.5 and -1 < regression_data['PCT_change'] < 0))
            ):
            add_in_csv(regression_data, regressionResult, ws, 'negativeOI-0')
            return True
        if(-2 < regression_data['PCT_day_change'] < 0 and -2 < regression_data['PCT_change'] < 0
            ):
            add_in_csv(regression_data, regressionResult, ws, 'negativeOI-1')
            return True
    return False

def buy_day_low(regression_data, regressionResult, ws):
    if((regression_data['PCT_day_change'] < -6 or regression_data['PCT_change'] < -6)
       and float(regression_data['forecast_day_VOL_change']) < -30
       and ((regression_data['mlpValue'] > 0.3 and regression_data['kNeighboursValue'] > 0.3) or is_algo_buy(regression_data))
       and regression_data['PCT_day_change_pre'] < 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayLow-ML')
        return True
    if((regression_data['PCT_day_change'] < -6 or regression_data['PCT_change'] < -6)
       and float(regression_data['forecast_day_VOL_change']) < -30
       and regression_data['PCT_day_change_pre'] < 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayLowVolLow-0')
        return True
    if((regression_data['PCT_day_change'] < -3 and regression_data['PCT_change'] < -3)
       and float(regression_data['forecast_day_VOL_change']) < 0
       and regression_data['PCT_day_change_pre'] < 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayLowVolLow-1')
        return True
    if((regression_data['PCT_day_change'] < -2.5 and regression_data['PCT_change'] < -2)
       and float(regression_data['forecast_day_VOL_change']) < 0
       and regression_data['PCT_day_change_pre'] < 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayLowVolLow-2')
        return True
    return False

def buy_oi_candidate(regression_data, regressionResult, ws):
    bar = regression_data['close'] - regression_data['open']
    bar_high = regression_data['high'] - regression_data['close']
    if morning_star_sell(regression_data, regressionResult, ws):
        return True
    if buy_oi_negative(regression_data, regressionResult, ws):
        return True
    if buy_day_low(regression_data, regressionResult, ws):
        return True
#     if(bar > 0 and bar_high != 0):
#         bar_pct_change = float(((bar-bar_high)/bar_high)*100)
#         if(5 > regression_data['forecast_day_PCT_change'] > 2
#            and bar_pct_change < 10
#            ):
#             return False
    else:
        if(((regression_data['mlpValue'] > 0.3 and regression_data['kNeighboursValue'] > 0.3) or is_algo_buy(regression_data))   
           and (float(regression_data['contract']) != 0 or float(regression_data['oi']) != 0)
           and float(regression_data['contract']) > 10
           and (regression_data['PCT_day_change'] > 1 or regression_data['PCT_change'] > 1)
           and regression_data['forecast_day_PCT_change'] > 0.5
           and regression_data['forecast_day_PCT2_change'] > 0.5
           and regression_data['forecast_day_PCT3_change'] > 0
           and regression_data['forecast_day_PCT4_change'] > 0
           and regression_data['forecast_day_PCT10_change'] < 10
           and regression_data['yearHighChange'] < -2
           and(regression_data['PCT_day_change_pre'] > 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
               )
           ):
            if((regression_data['forecast_day_VOL_change'] > 70 and 0.75 < regression_data['PCT_day_change'] < 2 and 0.5 < regression_data['PCT_change'] < 2)
                and float(regression_data['contract']) > 10
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiBuyCandidate-0')
                return True
            elif((regression_data['forecast_day_VOL_change'] > 35 and 0.75 < regression_data['PCT_day_change'] < 2 and 0.5 < regression_data['PCT_change'] < 2)
                and float(regression_data['contract']) > 20
                #and regression_data['PCT_day_change_pre'] > -0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiBuyCandidate-1')
                return True
            elif((regression_data['forecast_day_VOL_change'] > 100 and 0.75 < regression_data['PCT_day_change'] < 3 and 0.5 < regression_data['PCT_change'] < 3)
                #and regression_data['PCT_day_change_pre'] > -0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiBuyCandidate-2')
                return True
            elif(((regression_data['forecast_day_VOL_change'] > 400 and 0.75 < regression_data['PCT_day_change'] < 3.5 and 0.5 < regression_data['PCT_change'] < 3.5)
                or (regression_data['forecast_day_VOL_change'] > 500 and 0.75 < regression_data['PCT_day_change'] < 4.5 and 0.5 < regression_data['PCT_change'] < 4.5)
                )
                #and regression_data['PCT_day_change_pre'] > -0.5
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiBuyCandidate-3')
                return True
            elif((regression_data['forecast_day_VOL_change'] > 50 and 0.75 < regression_data['PCT_day_change'] < 5 and 0.5 < regression_data['PCT_change'] < 5)
                and float(regression_data['contract']) > 50 
                and regression_data['forecast_day_PCT10_change'] < -8 or regression_data['forecast_day_PCT7_change'] < -8
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiBuyCandidate-4')
                return True
            
        if(regression_data['forecast_day_PCT4_change'] <= 0.5
            and regression_data['forecast_day_PCT5_change'] <= 0.5
            and regression_data['forecast_day_PCT7_change'] <= 0.5
            and regression_data['forecast_day_PCT10_change'] < -10
            and ((regression_data['mlpValue'] > 0.5 and regression_data['kNeighboursValue'] > 0.5) or is_algo_buy(regression_data))
        ):
            if(regression_data['forecast_day_PCT_change'] > 0
               and regression_data['bar_high'] > regression_data['bar_high_pre']
                ):
                if(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4
                   and regression_data['PCT_day_change_pre'] < 0
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-0')
                    return True
                if(1 < regression_data['PCT_day_change'] < 5 and 1 < regression_data['PCT_change'] < 5
                   and regression_data['PCT_day_change_pre'] < 0
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-00')
                    return True
                if(1 < regression_data['PCT_day_change'] < 2.5 and 1 < regression_data['PCT_change'] < 2.5 
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-1')
                    return True
                if(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4
                    and no_doji_or_spinning_buy_india(regression_data)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-2')
                    return True
            if((((regression_data['close'] - regression_data['open']) * 1.5 > regression_data['high'] - regression_data['low']) or (regression_data['forecast_day_PCT_change'] > 0 and regression_data['PCT_day_change'] > 1))
                and regression_data['yearHighChange'] < -30 or regression_data['yearLowChange'] < 30
                ):
                if(1 < regression_data['PCT_day_change'] < 2.5 and 1 < regression_data['PCT_change'] < 2.5 
                    and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-3')
                    return True
                elif(1 < regression_data['PCT_day_change'] < 5 and 1 < regression_data['PCT_change'] < 5 
                    and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-4')
                    return True
            if(-0.3 < regression_data['PCT_day_change'] < 1 and -0.3 < regression_data['PCT_change'] < 1 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'finalBuyCandidate-5')
                return True    
    return False

def buy_oi(regression_data, regressionResult, ws):
    if(1 < regression_data['PCT_day_change'] < 4 and 1 < regression_data['PCT_change'] < 4 
        and regression_data['forecast_day_PCT_change'] > 0.5
        and regression_data['forecast_day_PCT2_change'] > 0.5
        and float(regression_data['forecast_day_VOL_change']) > 50 
        and float(regression_data['contract']) > 50
        and(regression_data['PCT_day_change_pre'] > 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
            )
        and regression_data['yearHighChange'] < -5
        ):
        if(regression_data['forecast_day_PCT10_change'] < 0 or regression_data['forecast_day_PCT7_change'] < 0):
            add_in_csv(regression_data, regressionResult, ws, 'openInterest-0')
            return True
        else:
            add_in_csv(regression_data, regressionResult, ws, 'openInterest-1')
            return True
    return False

def buy_all_common(regression_data, classification_data, regressionResult, ws_buyAllCommon):
    if((regression_data['kNeighboursValue'] >= 1 or (regression_data['mlpValue'] >= 1.5 and classification_data['mlpValue'] >= 1)) 
        #and regression_data['trend'] != 'up'
        and -1 < regression_data['PCT_change'] < 4
    ):
        add_in_csv(regression_data, regressionResult, ws_buyAllCommon, None)
        return True
    return False

def buy_all_filter(regression_data, regressionResult, ws_buyAllFilter):
    if buy_year_high(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
    if buy_year_low(regression_data, regressionResult, None, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
    if buy_up_trend(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
    if buy_down_trend(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
    if buy_final(regression_data, regressionResult, None, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
    if buy_high_indicators(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
#     if buy_pattern(regression_data, regressionResult, None, None):
#         ws_buyAllFilter.append(regressionResult) if (ws_buyAllFilter is not None) else False
#         return True
    if buy_oi(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_buyAllFilter, None)
        return True
    return False

def sell_pattern_without_mlalgo(regression_data, regressionResult, ws_buyPattern2, ws_sellPattern2):
    sellPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/all-buy-filter-by-PCT-Change.csv')
    if regression_data['sellIndia'] != '' and regression_data['sellIndia'] in sellPatternsDict:
        if (abs(float(sellPatternsDict[regression_data['sellIndia']]['avg'])) >= .1 and float(sellPatternsDict[regression_data['sellIndia']]['count']) >= 2):
            if(-3 < regression_data['PCT_day_change'] < 0.5 and float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -1):
                avg = sellPatternsDict[regression_data['sellIndia']]['avg']
                count = sellPatternsDict[regression_data['sellIndia']]['count']
                add_in_csv_hist_pattern(regression_data, regressionResult, ws_sellPattern2, 'wml_sell', avg, count)
            if(-0.5 < regression_data['PCT_day_change'] < 3 and float(sellPatternsDict[regression_data['sellIndia']]['avg']) > 1): 
                avg = sellPatternsDict[regression_data['sellIndia']]['avg']
                count = sellPatternsDict[regression_data['sellIndia']]['count']
                add_in_csv_hist_pattern(regression_data, regressionResult, ws_buyPattern2, 'wml_sell', avg, count)

def sell_pattern_from_history(regression_data, regressionResult, ws_sellPattern2):
    sellPatternsDict=scrip_patterns_to_dict('../../data-import/nselist/patterns-sell.csv')
    sellIndiaAvg = 0
    flag = False
    if regression_data['sellIndia'] != '' and regression_data['sellIndia'] in sellPatternsDict: 
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
                        add_in_csv_hist_pattern(regression_data, regressionResult, ws_sellPattern2, 'sellPattern2', avg, count) 
                    if(float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -0.5 
                        or (float(sellPatternsDict[regression_data['sellIndia']]['avg']) < -0.3 and (regression_data['forecast_day_PCT10_change'] > 10 or regression_data['yearLowChange'] > 40))):
                        if(regression_data['forecast_day_PCT10_change'] > 0 and regression_data['forecast_day_PCT_change'] <= 0):
                            flag = True
                            add_in_csv_hist_pattern(regression_data, regressionResult, ws_sellPattern2, 'sellPattern2', avg, count)
                        elif(regression_data['forecast_day_PCT10_change'] < 0):    
                            flag = True
                            add_in_csv_hist_pattern(regression_data, regressionResult, ws_sellPattern2, 'sellPattern2', avg, count)
    return sellIndiaAvg, flag

def sell_all_rule(regression_data, regressionResult, sellIndiaAvg, ws_sellAll):
    if(is_algo_sell(regression_data)
        and (regression_data['bar_low']-regression_data['low']) < (regression_data['bar_high']-regression_data['bar_low'])
        and 'P@[' not in str(regression_data['buyIndia'])
        and sellIndiaAvg <= 0.70):
        add_in_csv(regression_data, regressionResult, ws_sellAll, None)
        if(-10 < regression_data['yearHighChange'] < 0):
            if(regression_data['PCT_day_change'] > 1):
                return True
        else:
            return True
    return False

def sell_all_rule_classifier(regression_data, regressionResult, sellIndiaAvg, ws_sellAll):
    if(is_algo_sell_classifier(regression_data)
        and (regression_data['bar_low']-regression_data['low']) < (regression_data['bar_high']-regression_data['bar_low'])
        and 'P@[' not in str(regression_data['buyIndia'])
        and sellIndiaAvg <= 0.5):
        add_in_csv(regression_data, regressionResult, ws_sellAll, None)
        if(-10 < regression_data['yearHighChange'] < 0):
            if(regression_data['PCT_day_change'] > 1):
                return True
        else:
            return True
    return False
        
def sell_year_high(regression_data, regressionResult, ws_sellYearHigh, ws_sellYearHigh1):
    if(-10 < regression_data['yearHighChange'] < -1 and regression_data['yearLowChange'] > 30 and -5 < regression_data['PCT_day_change'] < -0.75 
        and regression_data['forecast_day_PCT10_change'] > 10 and regression_data['forecast_day_PCT7_change'] > 5 and regression_data['forecast_day_PCT5_change'] > -0.5 and regression_data['forecast_day_PCT4_change'] > -0.5
        and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0):
        add_in_csv(regression_data, regressionResult, ws_sellYearHigh, 'sellYearHigh')
        return True
    elif(-10 < regression_data['yearHighChange'] < 0 and regression_data['yearLowChange'] > 30 
        and (regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0.5)
        and regression_data['forecast_day_PCT10_change'] > 5 and regression_data['forecast_day_PCT7_change'] > 3 and regression_data['forecast_day_PCT5_change'] > -0.5
        and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0):
        add_in_csv(regression_data, regressionResult, ws_sellYearHigh1, 'sellYearHigh1')
        return True
    elif(-10 < regression_data['yearHighChange'] < 0 and regression_data['yearLowChange'] > 30 
        and (regression_data['PCT_day_change'] < 0 and regression_data['PCT_change'] < 0.5)
        and regression_data['forecast_day_PCT10_change'] > 0 and regression_data['forecast_day_PCT7_change'] > 0 and regression_data['forecast_day_PCT5_change'] > 0
        and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0):
        add_in_csv(regression_data, regressionResult, ws_sellYearHigh1, 'sellYearHigh1')
        return True   
    return False

def sell_year_low(regression_data, regressionResult, ws_sellYearLow):
    if(0 < regression_data['yearLowChange'] < 15 and regression_data['yearHighChange'] < -30 
        and -2 < regression_data['PCT_day_change'] < 0 and -2 < regression_data['forecast_day_PCT_change'] < 0
        and (regression_data['score'] != '1-1' or regression_data['score'] != '10')
        and all_day_pct_change_negative(regression_data) and no_doji_or_spinning_sell_india(regression_data)
        and float(regression_data['forecast_day_VOL_change']) > 50
        and(regression_data['PCT_day_change_pre'] < 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
            )
        ):
        add_in_csv(regression_data, regressionResult, ws_sellYearLow, 'sellYearLow')
        return True
    return False

def sell_up_trend(regression_data, regressionResult, ws_sellUpTrend):
    if(all_day_pct_change_positive(regression_data) 
       and -5 < regression_data['PCT_day_change'] < 0 
       and regression_data['forecast_day_PCT10_change'] > 5
       and (regression_data['forecast_day_PCT5_change'] > 10
            or regression_data['forecast_day_PCT7_change'] > 10
            or regression_data['forecast_day_PCT10_change'] > 10
           )
       and regression_data['yearLowChange'] > 30
       ):
        add_in_csv(regression_data, regressionResult, ws_sellUpTrend, 'sellUpTrend')
        return True
    return False

def sell_down_trend(regression_data, regressionResult, ws_sellDownTrend):
    if((regression_data['yearLowChange'] > 15 or regression_data['yearHighChange'] > -15)
       and regression_data['forecast_day_PCT_change'] < 0
       and regression_data['forecast_day_PCT2_change'] < 0
       and regression_data['forecast_day_PCT3_change'] < 0
       ):
        if(-3 < regression_data['PCT_day_change'] < -2 and -3 < regression_data['forecast_day_PCT_change'] < -2
           and regression_data['PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['series_trend'] != 'upTrend'
           and regression_data['score'] != '10' 
           and (regression_data['forecast_day_PCT7_change'] > 10 or regression_data['forecast_day_PCT10_change'] > 10)
           ):
            add_in_csv(regression_data, regressionResult, None, 'downTrendCandidate-0')
            return True
        if(-3 < regression_data['PCT_day_change'] < -2 and -3 < regression_data['forecast_day_PCT_change'] < -2
           and regression_data['PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] - 0.2
           and regression_data['series_trend'] != 'upTrend'
           and regression_data['score'] != '10' 
           and (regression_data['forecast_day_PCT7_change'] > 0 and regression_data['forecast_day_PCT10_change'] > 0)
           ):
            add_in_csv(regression_data, regressionResult, None, 'downTrendCandidate-00')
            return True
        elif(-3 < regression_data['PCT_day_change'] < -2 and -4 < regression_data['forecast_day_PCT_change'] 
           and regression_data['forecast_day_PCT_change'] > regression_data['PCT_change']
           and regression_data['series_trend'] != 'upTrend'
           and regression_data['score'] != '10'
           and (regression_data['forecast_day_PCT7_change'] > -5 and regression_data['forecast_day_PCT10_change'] > -7) 
           ):
            add_in_csv(regression_data, regressionResult, None, 'downTrendCandidate-1')
            return True
#     if(all_day_pct_change_negative(regression_data) and -5 < regression_data['PCT_day_change'] < 0 and regression_data['yearLowChange'] > 30
#         and regression_data['forecast_day_PCT10_change'] <= regression_data['PCT_change'] - 2
#         and regression_data['forecast_day_PCT10_change'] <= regression_data['PCT_day_change'] - 2
#         and float(regression_data['forecast_day_VOL_change']) > 30
#         and regression_data['PCT_day_change_pre'] < 0.5
#         and float(regression_data['contract']) > 10
#         and no_doji_or_spinning_sell_india(regression_data)):
#         add_in_csv(regression_data, regressionResult, ws_sellDownTrend, 'sellDownTrend')
#         return True
    return False

def sell_final(regression_data, regressionResult, ws_sellFinal, ws_sellFinal1):
    if(regression_data['yearLowChange'] > 10 and regression_data['score'] != '10'
       and -4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1):
        if(str(regression_data['buyIndia']) == '' and -90 < regression_data['yearHighChange'] < -10
            and regression_data['forecast_day_PCT10_change'] > 10 and regression_data['forecast_day_PCT7_change'] > 5 and regression_data['forecast_day_PCT5_change'] > -0.5 and regression_data['forecast_day_PCT4_change'] > -0.5
            and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0):
            add_in_csv(regression_data, regressionResult, ws_sellFinal, 'sellFinal')
            return True
        elif(regression_data['forecast_day_PCT10_change'] >= 10 and regression_data['forecast_day_PCT7_change'] >= 1 and regression_data['forecast_day_PCT5_change'] >= -1
            and regression_data['forecast_day_PCT2_change'] < 0 and regression_data['forecast_day_PCT_change'] < 0
            and regression_data['trend'] != 'down'
            ):
            add_in_csv(regression_data, regressionResult, ws_sellFinal1, 'sellFinal1')
            return True 
    return False

def sell_high_indicators(regression_data, regressionResult, ws_sellHighIndicators):
    if(regression_data['mlpValue'] < -1.0 and regression_data['kNeighboursValue'] < -1.0 and regression_data['yearHighChange'] > -10 and regression_data['yearLowChange'] > 30
        and (-2.5 < regression_data['PCT_day_change'] < 0 and -2.5 < regression_data['PCT_change'] < 0.5)):
        add_in_csv(regression_data, regressionResult, ws_sellHighIndicators, 'sellHighIndicators')
        return True         
    return False

def sell_pattern(regression_data, regressionResult, ws_sellPattern, ws_sellPattern1):
    score = ''
    if(regression_data['score'] == '1-1' or regression_data['score'] == '0-1'):
        score = 'down'
    if(-4 < regression_data['PCT_day_change'] < 1 and regression_data['yearHighChange'] < -5 and regression_data['score'] != '10'
        and regression_data['trend'] != 'down'
        ):
        if(('HANGINGMAN' in str(regression_data['sellIndia'])
           #or 'MARUBOZU' in str(regression_data['sellIndia'])
           #or 'ENGULFING' in str(regression_data['sellIndia'])
           or 'EVENINGSTAR' in str(regression_data['sellIndia'])
           #or ':DOJISTAR' in str(regression_data['sellIndia'])
           #or 'EVENINGDOJISTAR' in str(regression_data['sellIndia'])
           or 'ABANDONEDBABY' in str(regression_data['sellIndia'])
           or 'COUNTERATTACK' in str(regression_data['sellIndia'])
           or 'KICKING' in str(regression_data['sellIndia'])
           or 'BREAKAWAY' in str(regression_data['sellIndia'])
           #or 'TRISTAR' in str(regression_data['sellIndia'])
           or ('SHOOTINGSTAR' in str(regression_data['sellIndia']) and regression_data['PCT_day_change'] < 0)
           or 'DARKCLOUDCOVER' in str(regression_data['sellIndia'])
           #or '3INSIDE' in str(regression_data['sellIndia'])
           #or '3OUTSIDE' in str(regression_data['sellIndia'])
           #or '2CROWS' in str(regression_data['sellIndia'])
           #or '3BLACKCROWS' in str(regression_data['sellIndia'])
           ) and (regression_data['forecast_day_PCT5_change'] >= 0)):
            add_in_csv(regression_data, regressionResult, ws_sellPattern, 'sellPattern')
            return True
        elif(
           ('HARAMI' in str(regression_data['sellIndia']) and regression_data['forecast_day_PCT5_change'] >= 0 and score == 'down')
           or ('ENGULFING' in str(regression_data['sellIndia']) and 'LONGLINE' in str(regression_data['sellIndia']) and score == 'down')
           ) and ((regression_data['forecast_day_PCT5_change'] >= 5) or regression_data['yearLowChange'] > 50):
            add_in_csv(regression_data, regressionResult, ws_sellPattern, 'sellPattern')
            return True
        elif(
           'HANGINGMAN' in str(regression_data['sellIndia'])
           or 'MARUBOZU' in str(regression_data['sellIndia'])
           #or 'ENGULFING' in str(regression_data['sellIndia'])
           or 'EVENINGSTAR' in str(regression_data['sellIndia'])
           #or ':DOJISTAR' in str(regression_data['sellIndia'])
           #or 'EVENINGDOJISTAR' in str(regression_data['sellIndia'])
           or 'ABANDONEDBABY' in str(regression_data['sellIndia'])
           or 'COUNTERATTACK' in str(regression_data['sellIndia'])
           or 'KICKING' in str(regression_data['sellIndia'])
           or 'BREAKAWAY' in str(regression_data['sellIndia'])
           or 'TRISTAR' in str(regression_data['sellIndia'])
           or ('SHOOTINGSTAR' in str(regression_data['sellIndia']) and regression_data['PCT_day_change'] < 0)
           or 'DARKCLOUDCOVER' in str(regression_data['sellIndia'])
           #or '3INSIDE' in str(regression_data['sellIndia'])
           #or '3OUTSIDE' in str(regression_data['sellIndia'])
           or '2CROWS' in str(regression_data['sellIndia'])
           or '3BLACKCROWS' in str(regression_data['sellIndia'])
           or ('CLOSINGMARUBOZU' in str(regression_data['sellIndia']) and 'LONGLINE' in str(regression_data['sellIndia']))
           or ('M@[,CROSSOVER-MACD]' in str(regression_data['sellIndia']) and 'LONGLINE' in str(regression_data['sellIndia']))
           or ('3OUTSIDE' in str(regression_data['sellIndia']) and 'SPINNINGTOP' not in str(regression_data['sellIndia']) and 'LONGLINE' not in str(regression_data['sellIndia']))
           ) and ((regression_data['forecast_day_PCT5_change'] >= 5) or regression_data['yearLowChange'] > 50):
            add_in_csv(regression_data, regressionResult, ws_sellPattern1, 'sellPattern1')
            return True
    return False

def morning_star_buy(regression_data, regressionResult, ws):
    if(-5 < regression_data['forecast_day_PCT_change'] < -2
        and regression_data['PCT_day_change_pre'] < 0
        ):
        if(0 < regression_data['PCT_day_change'] < 1 and 0 < regression_data['PCT_change'] < 1 
            and regression_data['kNeighboursValue'] > 0
            ):
            add_in_csv(regression_data, regressionResult, ws, 'msBuyCandidate-0')
            return True
        elif(-1 < regression_data['PCT_day_change'] < 0 and -1 < regression_data['PCT_change'] < 0
            and (regression_data['low'] - regression_data['open']) < regression_data['PCT_day_change'] * 3
            and (regression_data['open'] - regression_data['high']) > regression_data['PCT_day_change']
            ):
            add_in_csv(regression_data, regressionResult, ws, 'msBuyCandidate-1')
            return True
        elif( -1 < regression_data['PCT_day_change'] < 0 and -1 < regression_data['PCT_change'] < 0
            and regression_data['forecast_day_PCT_change'] < regression_data['PCT_day_change'] * 3
            and regression_data['forecast_day_PCT10_change'] <= -10
            ):
            add_in_csv(regression_data, regressionResult, ws, 'msBuyCandidate-2')
            return True
    return False

def sell_oi_negative(regression_data, regressionResult, ws):
    if(regression_data['greentrend'] == 1
        and regression_data['forecast_day_PCT_change'] > 0.5
        and regression_data['forecast_day_PCT2_change'] > 0.5
        and regression_data['forecast_day_PCT3_change'] > 0.5
        and regression_data['forecast_day_PCT4_change'] > 0.5
        and regression_data['forecast_day_PCT5_change'] > 0.5
        and regression_data['forecast_day_PCT7_change'] > 0.5
        and regression_data['forecast_day_PCT10_change'] > 5
        and (regression_data['forecast_day_PCT5_change'] > 10
             or regression_data['forecast_day_PCT7_change'] > 10
             or regression_data['forecast_day_PCT10_change'] > 10
             )
        and float(regression_data['forecast_day_VOL_change']) < -30 
        and regression_data['PCT_day_change_pre'] > 0
        and float(regression_data['contract']) < 0
        and float(regression_data['oi']) < 5
        and (regression_data['yearLowChange'] > 15 or regression_data['yearHighChange'] > -15) 
        and ((regression_data['mlpValue'] < 0 and regression_data['kNeighboursValue'] < 0) or is_algo_sell(regression_data))
        ):
        if(((0 < regression_data['PCT_day_change'] < 1 and 0.5 < regression_data['PCT_change'] < 1)
            or (0.5 < regression_data['PCT_day_change'] < 1 and 0 < regression_data['PCT_change'] < 1))
           ):
            add_in_csv(regression_data, regressionResult, ws, 'negativeOI-0')
            return True
        if(0 < regression_data['PCT_day_change'] < 2 and 0 < regression_data['PCT_change'] < 2 
            ):
            add_in_csv(regression_data, regressionResult, ws, 'negativeOI-1')
            return True
    return False

def sell_day_high(regression_data, regressionResult, ws):
    if((regression_data['PCT_day_change'] > 6 or regression_data['PCT_change'] > 6) 
       and float(regression_data['forecast_day_VOL_change']) < -30  
       and ((regression_data['mlpValue'] < -0.3 and regression_data['kNeighboursValue'] < -0.3) or is_algo_sell(regression_data))
       and regression_data['PCT_day_change_pre'] > 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayHigh-ML')
        return True
    elif((regression_data['PCT_day_change'] > 6 or regression_data['PCT_change'] > 6)
       and float(regression_data['forecast_day_VOL_change']) < -30
       and regression_data['PCT_day_change_pre'] > 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayHighVolLow-0')
        return True
    if((regression_data['PCT_day_change'] > 3 and regression_data['PCT_change'] > 3) 
       and float(regression_data['forecast_day_VOL_change']) < 0  
       and regression_data['PCT_day_change_pre'] > 0
       ):
        add_in_csv(regression_data, regressionResult, ws, 'dayHighVolLow-1')
        return True
#     if((regression_data['PCT_day_change'] > 2.5 and regression_data['PCT_change'] > 2) 
#        and float(regression_data['forecast_day_VOL_change']) < 0  
#        and regression_data['PCT_day_change_pre'] > 0
#        ):
#         add_in_csv(regression_data, regressionResult, ws, 'dayHighVolLow-2')
#         return True
    return False

def sell_oi_candidate(regression_data, regressionResult, ws):
    bar = regression_data['open'] - regression_data['close']
    bar_low = regression_data['close'] - regression_data['low']
    if morning_star_buy(regression_data, regressionResult, ws): 
        return True
    if sell_oi_negative(regression_data, regressionResult, ws):
        return True
    if sell_day_high(regression_data, regressionResult, ws):
        return True
#     if(bar > 0 and bar_low != 0):
#         bar_pct_change = float(((bar - bar_low)/bar_low)*100)
#         if( -5 < regression_data['forecast_day_PCT_change'] < -2
#             and bar_pct_change < 10 ):
#             return False
    else:
        if(((regression_data['mlpValue'] < -0.3 and regression_data['kNeighboursValue'] < -0.3) or is_algo_sell(regression_data))
            and (float(regression_data['contract']) != 0 or float(regression_data['oi']) != 0)
            and float(regression_data['contract']) > 10
            and (regression_data['PCT_day_change'] < -1 or regression_data['PCT_change'] < -1)
            and regression_data['forecast_day_PCT_change'] < -0.5
            and regression_data['forecast_day_PCT2_change'] < -0.5
            and regression_data['forecast_day_PCT3_change'] < 0
            and regression_data['forecast_day_PCT4_change'] < 0
            and(regression_data['PCT_day_change_pre'] < 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
               )
            and (regression_data['forecast_day_PCT10_change'] > 0 or regression_data['forecast_day_PCT7_change'] > 0)
            ):
            if((regression_data['forecast_day_VOL_change'] > 70 and -2 < regression_data['PCT_day_change'] < -0.75 and -2 < regression_data['PCT_change'] < -0.5)
                and float(regression_data['contract']) > 10
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiSellCandidate-0')
                return True
            elif((regression_data['forecast_day_VOL_change'] > 35 and -2 < regression_data['PCT_day_change'] < -0.75 and -2 < regression_data['PCT_change'] < -0.5)
                and float(regression_data['contract']) > 20
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiSellCandidate-1')
                return True
            elif((regression_data['forecast_day_VOL_change'] > 100 and -3 < regression_data['PCT_day_change'] < -0.75 and -3 < regression_data['PCT_change'] < -0.5)
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiSellCandidate-2')
                return True
            elif(((regression_data['forecast_day_VOL_change'] > 400 and -3.5 < regression_data['PCT_day_change'] < -0.75 and -3.5 < regression_data['PCT_change'] < -0.5)
                or (regression_data['forecast_day_VOL_change'] > 500 and -4.5 < regression_data['PCT_day_change'] < -0.75 and -4.5 < regression_data['PCT_change'] < -0.5)
                )
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiSellCandidate-3')
                return True
            elif((regression_data['forecast_day_VOL_change'] > 50 and -5 < regression_data['PCT_day_change'] < -0.75 and -5 < regression_data['PCT_change'] < -0.5)
                and float(regression_data['contract']) > 50
                and regression_data['forecast_day_PCT10_change'] > 8 or regression_data['forecast_day_PCT7_change'] > 8
                ):
                add_in_csv(regression_data, regressionResult, ws, 'oiSellCandidate-4')
                return True
            
        
        if(regression_data['forecast_day_PCT4_change'] >= -0.5
            and regression_data['forecast_day_PCT5_change'] >= -0.5
            and regression_data['forecast_day_PCT7_change'] >= -0.5
            and regression_data['forecast_day_PCT10_change'] > 10
            and ((regression_data['mlpValue'] < -0.5 and regression_data['kNeighboursValue'] < -0.5) or is_algo_sell(regression_data))
            ):  
            if(regression_data['forecast_day_PCT_change'] < 0
               and regression_data['bar_low'] < regression_data['bar_low_pre']
                ):
                if(-4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1
                   and regression_data['PCT_day_change_pre'] > 0
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-0')
                    return True
                if(-5 < regression_data['PCT_day_change'] < -1 and -5 < regression_data['PCT_change'] < -1
                   and regression_data['PCT_day_change_pre'] > 0
                   and (regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30)
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-00')
                    return True
                if(-2.5 < regression_data['PCT_day_change'] < -1 and -2.5 < regression_data['PCT_change'] < -1
                    ):
                    add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-1')
                    return True
                if(-4 < regression_data['PCT_day_change'] < -1 and -4 < regression_data['PCT_change'] < -1 
                    and no_doji_or_spinning_sell_india(regression_data)
                    ):   
                    add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-2')
                    return True 
            if((((regression_data['open'] - regression_data['close']) * 1.5 > regression_data['high'] - regression_data['low']) or (regression_data['forecast_day_PCT_change'] < 0 and regression_data['PCT_day_change'] < -1))
                and regression_data['yearHighChange'] > -30 or regression_data['yearLowChange'] > 30
                ):
                if(-2.5 < regression_data['PCT_day_change'] < -1 and -2.5 < regression_data['PCT_change'] < -1 
                    and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                    ):   
                    add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-3')
                    return True
                elif(-5 < regression_data['PCT_day_change'] < -1 and -5 < regression_data['PCT_change'] < -1 
                    and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                    ):   
                    add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-4')
                    return True
            if(-1 < regression_data['PCT_day_change'] < 0.3 and -1 < regression_data['PCT_change'] < 0.3 
                and no_doji_or_spinning_sell_india(regression_data) and no_doji_or_spinning_buy_india(regression_data)
                ):   
                add_in_csv(regression_data, regressionResult, ws, 'finalSellCandidate-5')
                return True    
    return False

def sell_oi(regression_data, regressionResult, ws):
    if(-6 < regression_data['PCT_day_change'] < -1 and -6 < regression_data['PCT_change'] < -1 
        and regression_data['forecast_day_PCT_change'] < -0.5
        and regression_data['forecast_day_PCT2_change'] < -0.5
        and float(regression_data['forecast_day_VOL_change']) > 50 
        and float(regression_data['contract']) > 50
        and(regression_data['PCT_day_change_pre'] < 0 
               or (regression_data['volume'] > regression_data['volume_pre2'] and regression_data['volume'] > regression_data['volume_pre3'])
            )
        ):
        if(regression_data['forecast_day_PCT10_change'] > 0 or regression_data['forecast_day_PCT7_change'] > 0):
            add_in_csv(regression_data, regressionResult, ws, 'openInterest-0')
            return True
        else:
            add_in_csv(regression_data, regressionResult, ws, 'openInterest-1')
            return True
    return False

def sell_all_common(regression_data, classification_data, regressionResult, ws_sellAllCommon):
    if((regression_data['kNeighboursValue'] <= -1 or (regression_data['mlpValue'] <= -1.5 and classification_data['mlpValue'] <= 0)) 
        and regression_data['trend'] != 'down'
        #and -4 < regression_data['PCT_change'] < 1
        ):
        add_in_csv(regression_data, regressionResult, ws_sellAllCommon, None)
        return True
    return False

def sell_all_filter(regression_data, regressionResult, ws_sellAllFilter):
    if sell_year_high(regression_data, regressionResult, None, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
    if sell_year_low(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
    if sell_up_trend(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
    if sell_down_trend(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
    if sell_final(regression_data, regressionResult, None, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
    if sell_high_indicators(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
#     if sell_pattern(regression_data, regressionResult, None, None):
#         ws_sellAllFilter.append(regressionResult) if (ws_sellAllFilter is not None) else False
#         return True
    if sell_oi(regression_data, regressionResult, None):
        add_in_csv(regression_data, regressionResult, ws_sellAllFilter, None)
        return True
    return False
