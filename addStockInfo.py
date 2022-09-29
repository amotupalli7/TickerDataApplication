from logging.handlers import TimedRotatingFileHandler
from time import strptime
from stockDataApp import *
from openpyxl import *



############################################################################################################################################################
# Gets minute data for the specified date for specified ticker
############################################################################################################################################################

def getMinuteData(ticker,date):
    symbol = ticker
    sDate = strToMili(date)
    periodType = 'day'
    frequencyType = 'minute'
    eDate = dateToMili(strToDate(date) + timedelta(days=1))
    needExtendedHoursData = 'true'
    data = makeRequest(symbol=symbol,periodType=periodType,frequencyType=frequencyType,startDate=sDate,endDate=eDate, needExtendedHoursData=needExtendedHoursData)
    data = data["candles"]
    for day in data:
        day["datetime"] = datetime.fromtimestamp(int(day["datetime"]/1000))
        dtString = day["datetime"].strftime("%H:%M:%S")
        day["datetime"] = dtString
    
    return data

############################################################################################################################################################
# Gets regular market time data
############################################################################################################################################################
def getMarketHourData(ticker,date):
    data = getMinuteData(ticker,date)
    marketTimeData = []
    bool = False
    for day in data:
        if(day['datetime'] == '09:30:00'): bool= True
        if(day['datetime'] == '16:01:00'): bool = False
        if(bool): marketTimeData.append(day)

    return marketTimeData



############################################################################################################################################################
# Gets premarket minute data
############################################################################################################################################################

def getPremarketData(ticker,date):
    minuteData = getMinuteData(ticker,date)
    premarketData = []
    for day in minuteData:
        if(day['datetime'] == '09:30:00'): return premarketData
        premarketData.append(day)
    
    return(premarketData)

############################################################################################################################################################
# Gets premarket high information 
############################################################################################################################################################

def getPremarketHighInfo(ticker,date):
    data = getPremarketData(ticker,date)
    premarketHigh = 0
    premarketHighTime = ""
    for minute in data:
        if minute['high'] >= premarketHigh: 
            premarketHigh = minute['high']
            premarketHighTime = minute['datetime']
    
    #premarketHighTime = strptime(premarketHighTime,"%H:%M:%S")
    return (premarketHigh,premarketHighTime)

############################################################################################################################################################
# Gets HOD and LOD time information
############################################################################################################################################################

def getRegularHODandLODTime(ticker,date):
    data = getMarketHourData(ticker,date)
    HOD = 0
    LOD = 1000000000000000
    HighTime = ""
    LowTime = ""
    for minute in data:
        if minute['high'] >= HOD:
            HOD = minute['high']
            HighTime = minute['datetime']
        if minute['low'] <= LOD:
            LOD = minute['low']
            LowTime = minute['datetime']
    #HighTime = strptime(HighTime,"%H:%M:%S")
    #LowTime = strptime(LowTime,"%H:%M:%S")
    return(HighTime,LowTime)

############################################################################################################################################################
# Fill a sheet that is formatted with these headers Ticker, Date, Gap %, Premarket High, Premarket Time, HOD Time, LOD time
############################################################################################################################################################



############################################################################################################################################################
# main where everything gets run
############################################################################################################################################################
data = (getMinuteData('HYRE','09/06/2022'))

#print(getPremarketHighInfo('ABOS','09/28/2022'))
#print(getRegularHODandLODTime('ABOS','09/28/2022'))




# premarketData = list(getMinuteData('CYRN','03/21/2022'))
# timeData = (getRegularHODandLODTime('CYRN','03/21/2022'))


