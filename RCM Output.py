#!/usr/bin/env python
# coding: utf-8

# In[78]:


import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
from empyrical import omega_ratio,sortino_ratio,sharpe_ratio
import matplotlib.pyplot as plt
import more_itertools as mit
from operator import itemgetter

# File name
file = 'RCM Data - PSFL Dec19.xlsx'
# Column range for analysed fund in file
columns = 'Y:AC'
dateCol = 'A:A'
# Row start of file column headers
rows = 13
# Period of drawdown report
period = 15

'''
NOTE:'Date', 'NAV', and 'Monthly Return' column names REQUIRED to be in that exact format
'''

#Indices
ReturnReportIndex = ['1 Month','3 Months','6 Months','12 Months','2 Years','3 Years', '5 Years']

RiskStatisticsIndex = ['Maximum Drawdown','Annualized Std Dev','Losing Months (%)','Average Losing Month','Loss Std Dev']

# Dataframe of original fund
df = pd.read_excel(file,usecols=columns,skiprows = rows)
if 'Date' not in df.columns:
    dfDate = pd.read_excel(file,usecols=dateCol,skiprows = rows)
    df['Date'] = dfDate['Date']
    df.rename(columns=lambda c: c.split('.')[0], inplace=True)
df = df.dropna(subset=['NAV'])

'''
Below are the other 3 funds to derive the correlations, they get concatenated with the initial fund and the correlation is found
with indices picked hardcodingly.
'''
#Other funds frames for correlations.
SP500 = pd.read_excel(file,usecols='AD:AJ',skiprows=rows)
MSCI = pd.read_excel(file,usecols='AN:AV',skiprows=rows)
Bloomberg = pd.read_excel(file,usecols='BT:BW',skiprows=rows)

# Set date index for all funds
SP500['Date'] = df['Date']
MSCI['Date'] = df['Date']
Bloomberg['Date'] = df['Date']

# rename columns to remove forced .3, .4 , etc.. namings due to duplicates (can print to see output)
SP500.rename(columns=lambda c: c.split('.')[0], inplace=True)
MSCI.rename(columns=lambda c: c.split('.')[0], inplace=True)
Bloomberg.rename(columns=lambda c: c.split('.')[0], inplace=True)

# Filtered out to allow for correlation calculation
SP500 = SP500[(SP500['Date'].isnull()==False) & (SP500['NAV'].isnull()==False)]
MSCI = MSCI[(MSCI['Date'].isnull()==False) & (MSCI['NAV'].isnull()==False)]
Bloomberg = Bloomberg[(Bloomberg['Date'].isnull()==False) & (Bloomberg['NAV'].isnull()==False)]

#Indices (hardcode picked)
MSCIIndex = SP500['Monthly Return'].corr(df['Monthly Return'])
BloombergIndex = Bloomberg['Monthly Return'].corr(df['Monthly Return'])

# Returns the best rolling 3,6,12 months before the current NAV
def FindRate(df,NAV,month):
    try:
        return NAV / df.tail(month+1).iloc[[0]]['NAV'][0] - 1
    except:
        return 0

# Function just to colour the worksheets accordingly
def colour(df,worksheet,row,workbook,color):
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': color,
    'border': 1, 'font_color':'white'})
    header_format.set_center_across()
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(row, col_num + 1, value, header_format) 
        
df = df[df['Date'].notnull()].reset_index(drop=True)
df['Date'] = df['Date'].map(str)
df = df[df['Date'].str.contains(':',na=False)]
df.set_index(['Date'],inplace=True)
df = df.iloc[1:]

# Function that returns the risk/reward statistics using empyrical library
def RiskRewardStats(df):
    global RiskRewardList
    RiskRewardIndex = ['Sharpe Ratio','Sortino Ratio','Omega Ratio','Skewness','Kurtosis',
                      'Correlation vs MSCI World TR Index','Correlation vs Bloomberg Index']
    OmegaRatio = omega_ratio(df['Monthly Return'])
    Kurtosis = df['Monthly Return'].kurt()
    Skewness = df['Monthly Return'].skew()
    SharpeRatio = sharpe_ratio(df['Monthly Return'],period='monthly')
    SortinoRatio = sortino_ratio(df['Monthly Return'],period='monthly')
    RiskRewardList = [SharpeRatio,SortinoRatio,OmegaRatio,Skewness,Kurtosis,MSCIIndex,BloombergIndex]
    RiskRewardDf = pd.DataFrame(RiskRewardList,columns=['Value'],index=RiskRewardIndex)
    return RiskRewardDf

RiskReward = RiskRewardStats(df)

# Function that returns the Return Statistics 
def ReturnStatistics(df):
    global ReturnStats
    ReturnStatisticsIndex = ['AUM','NAV','Current Month','Last Month','Year to Date','3 Month ROR','12 Month ROR','24 Month ROR',
          '36 Month ROR','Total Return','Compound ROR','Best Month','Winning Months (%)']
    #Current AUM
    AUM = df.tail(1)['AUM'][0]
    #Current month NAV
    NAV = df.tail(1)['NAV'][0]
    #Current month return
    CurrMonth = df.tail(1)['Monthly Return'][0]
    #Previous Month Return
    PrevMonth = df.tail(2)['Monthly Return'][0]
    #Current_Date (NOT IN ENTRIES)
    CurrentDate = df.tail(1)
    #YeartoDate
    YeartoDate = (NAV/df.tail(datetime.strptime(CurrentDate.index[0],"%Y-%m-%d %H:%M:%S").month +1)['NAV'][0]) - 1
    #Three month rate of return
    ThreeMonth = FindRate(df,NAV,3)
    #6 month rate of return
    TwelveMonth = FindRate(df,NAV,12)
    #24 month rate of return
    TwentyFourMonth = FindRate(df,NAV,24)
    #36 month rate of return
    ThirtySixMonth = FindRate(df,NAV,36)
    #Current cumulative compound
    TotalReturn = df.tail(1)['Cumulative Compound'][0]
    #Compound rate of return from inception
    CompRor = (df.tail(1)['Cumulative Compound'][0] + 1)**(12/(df.shape[0]-2)) - 1
    # Best month 
    BestMonth = df['Monthly Return'].max()
    #WinningMonths
    WinningMonths = df[df['Monthly Return'] > 0]['Monthly Return'].shape[0]/(df.shape[0]-2)
    #Return stats list
    ReturnStats = [AUM,NAV,CurrMonth,PrevMonth,YeartoDate,ThreeMonth,TwelveMonth,TwentyFourMonth,ThirtySixMonth,
                   TotalReturn,CompRor,BestMonth,WinningMonths]
    Results =  pd.DataFrame(data=ReturnStats,columns=['Value'],index=ReturnStatisticsIndex)
    return Results

Return = ReturnStatistics(df)
Writer = pd.ExcelWriter('RCM Results.xlsx',engine='xlsxwriter')
Workbook = Writer.book
Return.to_excel(Writer,sheet_name = 'OUTPUT')
Worksheet = Writer.sheets['OUTPUT']
Worksheet.set_tab_color('#00B050')

#Formats for the excel sheet
Format = Workbook.add_format({'num_format':'$#,##0.00','border':1,'text_wrap': True,
                              'valign': 'top','font_color': '#FF5A34','bold':True})
Format.set_center_across()
Format2 = Workbook.add_format({'num_format':'#,##0.00','border':1,'text_wrap': True,
                              'valign': 'center','font_color': '#FF5A34','bold':True})
FormatTop = Workbook.add_format({'num_format':'#,##0.00','border':1,'text_wrap': True,
                              'valign': 'top','font_color': '#FF5A34','bold':True})
FormatTop.set_center_across()

percFormat = Workbook.add_format({'num_format':'#,##0.00%','text_wrap': True,'valign': 'center',
                                  'font_color': '#FF5A34','border':1,'bold':True})
TotalFormat = Workbook.add_format({'bold': True, 'bg_color':'#FF5A34','border': 1,'text_wrap': True,
                                       'font_color':'white','font_size':15})
TotalFormat.set_center_across()
formatBlue = Workbook.add_format({'bold': True, 'bg_color':'#122057','border': 1,'text_wrap': True,
                                   'font_color':'white'})
formatBlue.set_center_across()
Border = Workbook.add_format({'border':1})
colour(Return,Worksheet,0,Workbook,'#FF5A34')

Worksheet.write('A1','Polar Star Ltd Fund',TotalFormat)

#Return Report
maxmonth = df[['Monthly Return']].max()
minmonth = df[['Monthly Return']].min() 

df = df.reset_index()
df.index += 1 
df.rename(columns={'Date ':'Date'}, inplace=True)
Dates = df['Date'].tolist()
Variable = df['Monthly Return'].tolist()
dictionary = dict(zip(Dates,Variable))
d = Dates

'''
Create a cumulative max column that keeps track of the previous maximum NAV and a Drawdown columnn for which the drawdown
is calculated for each row, the minimum is then taken to find the maximum drawdown
'''

df['HighValue'] = df['NAV'].cummax()
df['Drawdown'] = df['NAV']/df['HighValue'] - 1

#Remove NAN entry
df = df.iloc[1:]

'''
Drawdown calculation. Essentially sort the dataset by all monthly returns < 0, followed by grouping the dataset
by consecutive indices to separate out all of the drawdowns
'''
def drawdown(df,period):
    #Convert drawdown column type to float
    df['Drawdown'] = df['Drawdown'].astype('float64')
    #Sort by all drawdowns < 0
    dfpart = df[df['Drawdown']<0]
    #Index of the dataframe
    index = dfpart.index.tolist()
    #Holds all drawdown groups
    grouplist = []
    #Holds all drawdown depths, starting, and ending dates
    drawdowns = []
    #Holds all recovered months
    RecoverMonths = []
    '''
    Transforms index list into a list of lists of consecutive indices, grouplist then holds the equivalent rows sampled 
    from dfpart
    '''
    for group in mit.consecutive_groups(index):
        groups = list(group)
        grouplist.append(dfpart.loc[groups[0]:groups[-1]])
    '''
    Populates drawdowns list with minimum values of each group (i.e max drawdowns or depth) as well as the equivalent 
    starting and ending dates which are found using the indices and refer back to the original dataframe
    '''
    for i in range (0,len(grouplist)):
        Groupers = list(grouplist[i]['Drawdown'])
        if i < len(grouplist)-1:
            RecoverMonths.append(len(Groupers[Groupers.index(min(Groupers))::]))
        else:
            RecoverMonths.append('RECOVERING')
        drawdowns.append([grouplist[i]['Drawdown'].min(),len(grouplist[i]['Drawdown'])+1,RecoverMonths[i],str(grouplist[i]['Date'].iloc[0]),
                          str(grouplist[i]['Date'].iloc[-1])])
        lastdate = grouplist[i]['Date'].iloc[-1]
        nextdateindex = dfpart[dfpart['Date'] == lastdate].index[0] + 1
        '''
        The try except statement below is to confirm the existence of a final date, else have it return "not found"
        '''
        try:
            drawdowns[i][4] = df['Date'].loc[nextdateindex]
        except:
            drawdowns[i][4] = 'RECOVERING'
    # Sort by reversed absolute values
    drawdowns = sorted(drawdowns, key=lambda x: abs(x[0]),reverse=True)
    #Pick until certain period
    drawdowns = drawdowns[:period]
    return drawdowns

Drawdowndf = pd.DataFrame(data = drawdown(df,period),columns=['Depth(%)','Length (Months)','Recovery Months',
                                                              'Starting Date','Ending Date'])

#Risk Statistics and how I calculated them
MaxDrawdown = df['Drawdown'].min()
AnnualizedStdDev = df.loc[:,"Monthly Return"].std() * np.sqrt(12)
LosingMonthsPercentage = df[(df['Monthly Return'] < 0) | (df['Monthly Return'] == 0) ].shape[0]/(df.shape[0]-1)
AvgLosingMonth = np.average(df[(df['Monthly Return'] < 0) | (df['Monthly Return'] == 0)]['Monthly Return'])
LossStdDev = np.std(df[(df['Monthly Return'] < 0) | (df['Monthly Return'] == 0)]['Monthly Return'])
Risks = [MaxDrawdown,AnnualizedStdDev,LosingMonthsPercentage,AvgLosingMonth,LossStdDev]
#,MSCIIndex,BloombergIndex]
RiskDf = pd.DataFrame(data = Risks,columns=['Values'],index=RiskStatisticsIndex)

# Convert to date object
def Dt(key):
    return datetime.strptime(key,"%Y-%m-%d %H:%M:%S")

'''
These functions are for calculating minimum,maximum, average and median return reports for each month
'''
def match(s):
    for x in s:
        for i,y in enumerate(x):
            if y in dictionary:
                x[i] = dictionary[y]
    return s

# Essentially groups dates into x month unique groupings
def Groups(months,Dates,stopper):
    start = 0
    groups = []
    while start < stopper:
        for i in range(start, start + months):
            groups.append(Dates[i])
        start+=1
    return groups

class subgroups:
    def __init__(self,months):
        self.months = months
    def Maxima(self,months):
        stopper = len(Dates) - (months - 1)
        try:
            groups = Groups(months,Dates,stopper)
            # Splits into subgroupings of x months (print(groups,months) to see)
            groups =  [groups[i:i+months] for i in range(0, len(groups), months)]
            numperiods = len(groups)
            # Returns mapping to monthly return instead of dates in groupings (print to see output)
            groups = match(groups)
            # Summ list basically holds the returns for each grouping
            Summ = [(np.prod([(element2)+1 for element2 in element]) -1) for element in groups if str(element)!='nan']
            Summ = [element for element in Summ if str(element)!='nan']
            Summ_pos = [element for element in Summ if element > 0]
            Summ_neg = [element for element in Summ if element < 0]
            Last = Summ[-1]
            Winning = len([element for element in Summ if element > 0])
            return np.nanmax(Summ),np.nanmin(Summ),np.average(Summ),np.median(Summ),Winning/len(Summ),Last,np.average(Summ_pos),np.average(Summ_neg),numperiods
        except:
            print('nope')
            return 0,0,0,0,0,0,0,0,0
        
oneMonth = subgroups(1)
threeMonth = subgroups(3)
sixMonth = subgroups(6)
twelveMonth = subgroups(12)
twoYear = subgroups(24)
threeYear = subgroups(36)
fiveYear = subgroups(60)

oneMonthmax = subgroups.Maxima(oneMonth,1)[0]
oneMonthmin = subgroups.Maxima(oneMonth,1)[1]
oneMonthAvg = subgroups.Maxima(oneMonth,1)[2]
oneMonthMed = subgroups.Maxima(oneMonth,1)[3]
oneMonthWin = subgroups.Maxima(oneMonth,1)[4]
oneMonthLast = subgroups.Maxima(oneMonth,1)[5]
oneMonthavgpos = subgroups.Maxima(oneMonth,1)[6]
oneMonthavgneg = subgroups.Maxima(oneMonth,1)[7]
oneMonthlen = subgroups.Maxima(oneMonth,1)[8]
threeMonthmax = subgroups.Maxima(threeMonth,3)[0]
threeMonthmin = subgroups.Maxima(threeMonth,3)[1]
threeMonthAvg = subgroups.Maxima(threeMonth,3)[2]
threeMonthMed = subgroups.Maxima(threeMonth,3)[3]
threeMonthWin = subgroups.Maxima(threeMonth,3)[4]
threeMonthLast = subgroups.Maxima(threeMonth,3)[5]
threeMonthavgpos = subgroups.Maxima(threeMonth,3)[6]
threeMonthavgneg = subgroups.Maxima(threeMonth,3)[7]
threeMonthlen = subgroups.Maxima(threeMonth,3)[8]
sixMonthmax = subgroups.Maxima(sixMonth,6)[0]
sixMonthmin = subgroups.Maxima(sixMonth,6)[1]
sixMonthAvg = subgroups.Maxima(sixMonth,6)[2]
sixMonthMed = subgroups.Maxima(sixMonth,6)[3]
sixMonthWin = subgroups.Maxima(sixMonth,6)[4]
sixMonthLast = subgroups.Maxima(sixMonth,6)[5]
sixMonthavgpos = subgroups.Maxima(sixMonth,6)[6]
sixMonthavgneg = subgroups.Maxima(sixMonth,6)[7]
sixMonthlen = subgroups.Maxima(sixMonth,6)[8]
twelveMonthmax = subgroups.Maxima(twelveMonth,12)[0]
twelveMonthmin = subgroups.Maxima(twelveMonth,12)[1]
twelveMonthAvg = subgroups.Maxima(twelveMonth,12)[2]
twelveMonthMed = subgroups.Maxima(twelveMonth,12)[3]
twelveMonthWin = subgroups.Maxima(twelveMonth,12)[4]
twelveMonthLast = subgroups.Maxima(twelveMonth,12)[5]
twelveMonthavgpos = subgroups.Maxima(twelveMonth,12)[6]
twelveMonthavgneg = subgroups.Maxima(twelveMonth,12)[7]
twelveMonthlen = subgroups.Maxima(twelveMonth,12)[8]
twoYearmax = subgroups.Maxima(twoYear,24)[0]
twoYearmin = subgroups.Maxima(twoYear,24)[1]
twoYearAvg = subgroups.Maxima(twoYear,24)[2]
twoYearMed = subgroups.Maxima(twoYear,24)[3]
twoYearWin = subgroups.Maxima(twoYear,24)[4]
twoYearLast = subgroups.Maxima(twoYear,24)[5]
twoYearavgpos = subgroups.Maxima(twoYear,24)[6]
twoYearavgneg = subgroups.Maxima(twoYear,24)[7]
twoYearlen = subgroups.Maxima(twoYear,24)[8]
threeYearmax = subgroups.Maxima(threeYear,36)[0]
threeYearmin = subgroups.Maxima(threeYear,36)[1]
threeYearAvg = subgroups.Maxima(threeYear,36)[2]
threeYearMed = subgroups.Maxima(threeYear,36)[3]
threeYearWin = subgroups.Maxima(threeYear,36)[4]
threeYearLast = subgroups.Maxima(threeYear,36)[5]
threeYearavgpos = subgroups.Maxima(threeYear,36)[6]
threeYearavgneg = subgroups.Maxima(threeYear,36)[7]
threeYearlen = subgroups.Maxima(threeYear,36)[8]
fiveYearmax = subgroups.Maxima(fiveYear,60)[0]
fiveYearmin = subgroups.Maxima(fiveYear,60)[1]
fiveYearAvg = subgroups.Maxima(fiveYear,60)[2]
fiveYearMed = subgroups.Maxima(fiveYear,60)[3]
fiveYearWin = subgroups.Maxima(fiveYear,60)[4]
fiveYearLast = subgroups.Maxima(fiveYear,60)[5]
fiveYearlen = subgroups.Maxima(fiveYear,60)[8]

ReturnReportmins = [oneMonthmin,threeMonthmin,sixMonthmin,twelveMonthmin,twoYearmin,threeYearmin,fiveYearmin]
ReturnReportmaxs = [oneMonthmax,threeMonthmax,sixMonthmax,twelveMonthmax,twoYearmax,threeYearmax,fiveYearmax]
ReturnReportavgs = [oneMonthAvg,threeMonthAvg,sixMonthAvg,twelveMonthAvg,twoYearAvg,threeYearAvg,fiveYearAvg]
ReturnReportmeds = [oneMonthMed,threeMonthMed,sixMonthMed,twelveMonthMed,twoYearMed,threeYearMed,fiveYearMed]
ReturnReportwins = [oneMonthWin,threeMonthWin,sixMonthWin,twelveMonthWin,twoYearWin,threeYearWin,fiveYearWin]
ReturnReportLast = [oneMonthLast,threeMonthLast,sixMonthLast,twelveMonthLast,twoYearLast,threeYearLast,fiveYearLast]
ReturnReportavgpos = [oneMonthavgpos,threeMonthavgpos,sixMonthavgpos,twelveMonthavgpos,twoYearavgpos,threeYearavgpos]
ReturnReportavgneg = [oneMonthavgneg,threeMonthavgneg,sixMonthavgneg,twelveMonthavgneg,twoYearavgneg,threeYearavgneg]
ReturnReportlen = [oneMonthlen,threeMonthlen,sixMonthlen,twelveMonthlen,twoYearlen,threeYearlen,fiveYearlen]

ReturnReport = pd.DataFrame(data= np.zeros((7,3)),columns=['Minimum','Maximum','Average'],index=ReturnReportIndex)

TimeWindow = pd.DataFrame(data = [ReturnReportmins,ReturnReportmaxs,ReturnReportavgs,ReturnReportmeds,
                                 ReturnReportwins,ReturnReportLast,ReturnReportavgpos,ReturnReportavgneg,ReturnReportlen],columns = ReturnReport.index,index = 
                         ReturnReport.columns.tolist() + 
                              ['Median','Winning','Last','Avg. Pos. Period','Avg. Neg. Period','# of Periods'])
TimeWindow = TimeWindow.replace(np.nan, '-', regex=True)

ReturnReport.to_excel(Writer,sheet_name = 'OUTPUT',startrow = Return.shape[0] + 4)
TimeWindow.to_excel(Writer,sheet_name = 'OUTPUT',startrow = Return.shape[0] + ReturnReport.shape[0] + 7)
RiskDf.to_excel(Writer,sheet_name = 'OUTPUT',startcol = Return.shape[1] + 4)
RiskReward.to_excel(Writer,sheet_name = 'OUTPUT',startcol = Return.shape[1] + 4,startrow = RiskDf.shape[0]+2)
Drawdowndf.index+=1
Drawdowndf.to_excel(Writer,sheet_name = 'OUTPUT',startrow=1,startcol= Return.shape[1] + RiskDf.shape[1] + 8)
Worksheet.set_column('A:A',18)
Worksheet.set_column('B:E',10)
Worksheet.set_column('F:F',30)
Worksheet.set_column('G:H',8)
Worksheet.set_column('I:O',8)
Worksheet.set_column('P:Q',10)

#Hardcoded formattings
for i,element in enumerate(ReturnStats):
    # Dollar or percent formatting.. hardcoded
    if i == 0 or i == 1:
        Worksheet.merge_range(i+1,1,i+1,2,element,Format)
    else:
        Worksheet.merge_range(i+1,1,i+1,2,element,percFormat)
        
#Return report
for i,element in enumerate(ReturnReportmins):
    Worksheet.write('B{}'.format(i+19),element,percFormat)
for i,element in enumerate(ReturnReportmaxs):
    Worksheet.write('C{}'.format(i+19),element,percFormat)    
for i,element in enumerate(ReturnReportavgs):
    Worksheet.write('D{}'.format(i+19),element,percFormat)   
for i,element in enumerate(RiskRewardList):
    Worksheet.merge_range(i+8,6,i+8,7,element,Format2) 
for i,element in enumerate(ReturnReportmeds):
    Worksheet.write('E{}'.format(i+19),element,percFormat)
for i,element in enumerate(ReturnReportwins):
    Worksheet.write('F{}'.format(i+19),element,percFormat)
for i,element in enumerate(ReturnReportLast):
    Worksheet.merge_range(i+18,6,i+18,7,element,percFormat)

#Time Window Analysis
for i,element in enumerate(TimeWindow.iloc[:,0].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
        Format = Format2
    Worksheet.merge_range(i+28,1,i+28,4,element,Format) 
for i,element in enumerate(TimeWindow.iloc[:,1].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
        Format = Format2
    Worksheet.merge_range(i+28,5,i+28,6,element,Format)   
for i,element in enumerate(TimeWindow.iloc[:,2].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
        Format = Format2
    Worksheet.merge_range(i+28,7,i+28,8,element,Format)   
for i,element in enumerate(TimeWindow.iloc[:,3].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
        Format = Format2
    Worksheet.merge_range(i+28,9,i+28,13,element,Format)
for i,element in enumerate(TimeWindow.iloc[:,4].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
         Format = Format2
    Worksheet.merge_range(i+28,14,i+28,15,element,Format)
for i,element in enumerate(TimeWindow.iloc[:,5].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
        Format = Format2
    Worksheet.merge_range(i+28,16,i+28,17,element,Format)
for i,element in enumerate(TimeWindow.iloc[:,6].tolist()):
    Format = percFormat
    if i == len(TimeWindow.iloc[:,0].tolist())-1:
         Format = Format2
    Worksheet.merge_range(i+28,18,i+28,20,element,Format)  
    
# #Risk stats
for i,element in enumerate(Risks):
    Worksheet.merge_range(i+1,6,i+1,7,element,percFormat)
#Drawdown stats
for i,element in enumerate(drawdown(df,period)):
    Worksheet.merge_range(i+2,11,i+2,12,element[0],percFormat)  
    Worksheet.merge_range(i+2,13,i+2,14,element[1],FormatTop)  
    Worksheet.merge_range(i+2,15,i+2,16,element[2],FormatTop) 
    Worksheet.merge_range(i+2,17,i+2,18,element[3],FormatTop) 
    Worksheet.merge_range(i+2,19,i+2,20,element[4],FormatTop) 
    
Worksheet.merge_range(0,0,0,2,'RETURN STATISTICS',TotalFormat)
Worksheet.merge_range(0,5,0,7,'RISK STATISTICS',TotalFormat)
Worksheet.merge_range(Return.shape[0]+3,0,Return.shape[0]+3,7,'RETURN REPORT',TotalFormat)
Worksheet.merge_range(RiskDf.shape[0]+2,Return.shape[1]+4,RiskDf.shape[0]+2,Return.shape[1]+6,'RISK/REWARD STATS',TotalFormat)
Worksheet.merge_range(0,10,0,20,'DRAWDOWN REPORT',TotalFormat)
Worksheet.merge_range(Return.shape[0] + ReturnReport.shape[0] + 6,0,Return.shape[0] + ReturnReport.shape[0] + 6
                      ,20,'TIME WINDOW ANALYSIS',TotalFormat)

#Blue format row
Worksheet.write('A18','PERIOD',formatBlue)
Worksheet.write('B18','MINIMUM',formatBlue)
Worksheet.write('C18','MAXIMUM',formatBlue)
Worksheet.write('D18','AVERAGE',formatBlue)
Worksheet.write('E18','MEDIAN',formatBlue)
Worksheet.write('F18','WINNING (%)',formatBlue)
Worksheet.merge_range(17,6,17,7,'LAST',formatBlue)
#Drawdown column names
Worksheet.write('K2','NUMBER',formatBlue)
Worksheet.merge_range(1,11,1,12,'DEPTH (%)',formatBlue)  
Worksheet.merge_range(1,13,1,14,'LENGTH (MONTHS)',formatBlue) 
Worksheet.merge_range(1,15,1,16,'RECOVERY (MONTHS)',formatBlue) 
Worksheet.merge_range(1,17,1,18,'STARTING DATE',formatBlue) 
Worksheet.merge_range(1,19,1,20,'ENDING DATE',formatBlue) 
Worksheet.set_zoom(55)
#Time window names
Worksheet.write('A28','STATISTIC',formatBlue)
Worksheet.merge_range(27,1,27,4,'1 MONTH',formatBlue) 
Worksheet.merge_range(27,5,27,6,'3 MONTH',formatBlue) 
Worksheet.merge_range(27,7,27,8,'6 MONTH',formatBlue) 
Worksheet.merge_range(27,9,27,13,'12 MONTH',formatBlue) 
Worksheet.merge_range(27,14,27,15,'2 YEARS',formatBlue) 
Worksheet.merge_range(27,16,27,17,'3 YEARS',formatBlue) 
Worksheet.merge_range(27,18,27,20,'5 YEARS',formatBlue) 

Writer.save()

# df.plot(kind='line',x='Date',y='Monthly Return',figsize=(30,15))
# df.plot(kind='line',x='Date',y='Drawdown',figsize=(30,15))

# plt.show()
print('Written.')


# In[ ]:




