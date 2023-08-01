import sys
import pandas as pd
import numpy as np
import datetime
import os
import pygsheets
import datetime as dt
import warnings
warnings.filterwarnings("ignore")
from pandas.tseries.offsets import BDay


def bday_counter_forward(date,n) :
    
    tw_holiday_list = ["2023-04-03","2023-04-04","2023-04-05","2023-05-01","2023-06-23","2023-06-22","2023-09-29","2023-10-09","2023-10-10"]
    date = str(date)[0:10]
    yesterday = str(pd.to_datetime(date)+BDay(n))[0:10]
    
    while yesterday in tw_holiday_list :
        yesterday = str(pd.to_datetime(yesterday)+BDay(n))[0:10]
    else :
        return yesterday
                

def third_wen(date_str):
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    
    day=21-(dt.date(y,m,1).weekday()+4)%7  
    return dt.date(y,m,day)  


def TWN_contract_month(date_str) :
    TWN_ex_list= ['2022-01-25','2022-02-24','2022-03-30','2022-04-28','2022-05-30','2022-06-29','2022-07-28','2022-08-30','2022-09-29','2022-10-28','2022-11-29','2022-12-29'
        ,'2023-01-30','2023-02-23','2023-03-30','2023-04-27','2023-05-30','2023-06-29','2023-07-28','2023-08-30','2023-09-28','2023-10-30','2023-11-29','2023-12-28']
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    yyyymm = date_str[0:7]
    
    結算日 = pd.to_datetime([x for x in TWN_ex_list if x.startswith(yyyymm)== True][0])
    
    if pd.to_datetime(date_str) <= 結算日 : # 開盤前除息 仍算在該月合約
        return date_str[0:7]+'合約'
    
    if pd.to_datetime(date_str) > pd.to_datetime(結算日) : 
        if m == 12 :
            return str(y+1)+'-01合約'
    
        else :
            return str(y)+'-0'+str(m+1)+'合約'

def TXF_contract_month(date_str):
    close_day = third_wen(date_str)
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    if pd.to_datetime(date_str) <= pd.to_datetime(close_day) : # 開盤前除息 仍算在該月合約
        return date_str[0:7]+'合約'
    
    if pd.to_datetime(date_str) > pd.to_datetime(close_day) : 

        if m == 12 :
            return str(y+1)+'-01合約'
    
        else :
            return str(y)+'-0'+str(m+1)+'合約'
        
        
today = str((dt.datetime.today().date()))    
path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+today+"/"
Total_dividend_df = pd.read_csv(path+"全除息個股序列總表_"+today+".csv")

date_list = []
date_list.append(str(dt.datetime.today().date())[:10])
for i in range(1,5):
    date_list.append(bday_counter_forward(date_list[i-1],1))
    
富台合約_list = []
台指合約_list = []
for i in ['2023-01-30','2023-02-23','2023-03-30','2023-04-27','2023-05-30','2023-06-29','2023-07-28','2023-08-30','2023-09-28','2023-10-30','2023-11-29','2023-12-28'] :
    富台合約_list.append(TWN_contract_month(i))
    台指合約_list.append(TXF_contract_month(i))
    
    
import sys
import pandas as pd
import numpy as np
import datetime
import os
import pygsheets
import datetime as dt
import warnings
warnings.filterwarnings("ignore")
from pandas.tseries.offsets import BDay


def bday_counter_forward(date,n) :
    
    tw_holiday_list = ["2023-04-03","2023-04-04","2023-04-05","2023-05-01","2023-06-23","2023-06-22","2023-09-29","2023-10-09","2023-10-10"]
    date = str(date)[0:10]
    yesterday = str(pd.to_datetime(date)+BDay(n))[0:10]
    
    while yesterday in tw_holiday_list :
        yesterday = str(pd.to_datetime(yesterday)+BDay(n))[0:10]
    else :
        return yesterday
                

def third_wen(date_str):
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    
    day=21-(dt.date(y,m,1).weekday()+4)%7  
    return dt.date(y,m,day)  


def TWN_contract_month(date_str) :
    TWN_ex_list= ['2022-01-25','2022-02-24','2022-03-30','2022-04-28','2022-05-30','2022-06-29','2022-07-28','2022-08-30','2022-09-29','2022-10-28','2022-11-29','2022-12-29'
        ,'2023-01-30','2023-02-23','2023-03-30','2023-04-27','2023-05-30','2023-06-29','2023-07-28','2023-08-30','2023-09-28','2023-10-30','2023-11-29','2023-12-28']
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    yyyymm = date_str[0:7]
    
    結算日 = pd.to_datetime([x for x in TWN_ex_list if x.startswith(yyyymm)== True][0])
    
    if pd.to_datetime(date_str) <= 結算日 : # 開盤前除息 仍算在該月合約
        return date_str[0:7]+'合約'
    
    if pd.to_datetime(date_str) > pd.to_datetime(結算日) : 
        if m == 12 :
            return str(y+1)+'-01合約'
    
        else :
            return str(y)+'-0'+str(m+1)+'合約'

def TXF_contract_month(date_str):
    close_day = third_wen(date_str)
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    if pd.to_datetime(date_str) <= pd.to_datetime(close_day) : # 開盤前除息 仍算在該月合約
        return date_str[0:7]+'合約'
    
    if pd.to_datetime(date_str) > pd.to_datetime(close_day) : 

        if m == 12 :
            return str(y+1)+'-01合約'
    
        else :
            return str(y)+'-0'+str(m+1)+'合約'
        
        
today = str((dt.datetime.today().date()))    
path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+today+"/"
Total_dividend_df = pd.read_csv(path+"全除息個股序列總表_"+today+".csv")

date_list = []
date_list.append(str(dt.datetime.today().date())[:10])
for i in range(1,5):
    date_list.append(bday_counter_forward(date_list[i-1],1))
    
富台合約_list = []
台指合約_list = []
for i in ['2023-01-30','2023-02-23','2023-03-30','2023-04-27','2023-05-30','2023-06-29','2023-07-28','2023-08-30','2023-09-28','2023-10-30','2023-11-29','2023-12-28'] :
    富台合約_list.append(TWN_contract_month(i))
    台指合約_list.append(TXF_contract_month(i))
    
    
date_dic = {}
for date in date_list :
    台指合約時間 = TXF_contract_month(date) 
    富台合約時間 = TWN_contract_month(date) 

    台指監督合約 = 台指合約_list[台指合約_list.index(台指合約時間):台指合約_list.index(台指合約時間)+3]
    富台監督合約 = 富台合約_list[富台合約_list.index(富台合約時間):富台合約_list.index(富台合約時間)+3]

    台指今日除息點數 = Total_dividend_df[(Total_dividend_df['除息日'] == date) & (Total_dividend_df['台指影響點數'] != 0)].台指影響點數.sum()
    富台今日除息點數 = Total_dividend_df[(Total_dividend_df['除息日'] == date) & (Total_dividend_df['富台影響點數'] != 0)].富台影響點數.sum()

    台指預計除息公司_dic = {}
    富台預計除息公司_dic = {}
    for 合約截止月份 in 台指監督合約 :
        台指總除息點數 = Total_dividend_df[(Total_dividend_df['除息日'].apply(lambda x : pd.to_datetime(x)) > pd.to_datetime(date)) & (Total_dividend_df['台指合約月份'] == 合約截止月份)].台指影響點數.sum()
        台指總除息點數 = 台指總除息點數
        台指預計除息公司_dic[合約截止月份] = 台指總除息點數

    for 合約截止月份 in 富台監督合約 :
        富台總除息點數 = Total_dividend_df[(Total_dividend_df['除息日'].apply(lambda x : pd.to_datetime(x)) > pd.to_datetime(date)) & (Total_dividend_df['富台合約月份'] == 合約截止月份)].富台影響點數.sum()
        富台總除息點數 = 富台總除息點數
        富台預計除息公司_dic[合約截止月份] = 富台總除息點數

    台指累積除息點數 = pd.DataFrame(台指預計除息公司_dic,index = ["台指累積除息點數"]).T.cumsum()
    台指累積除息點數.loc["今日除息點數"] = 台指今日除息點數
    台指累積除息點數 = 台指累積除息點數.append(台指累積除息點數).iloc[3:7]

    富台累積除息點數 = pd.DataFrame(富台預計除息公司_dic,index = ["富台累積除息點數"]).T.cumsum()
    富台累積除息點數.loc["今日除息點數"] = 富台今日除息點數
    富台累積除息點數 = 富台累積除息點數.append(富台累積除息點數).iloc[3:7]

    總除息列表 = pd.concat([台指累積除息點數,富台累積除息點數],axis = 1).fillna(0).reset_index().sort_values('index').set_index('index')
    總除息列表['sort'] = [(i+1) % len(總除息列表) for i in range(len(總除息列表))]
    總除息列表 = 總除息列表.sort_values('sort').drop('sort',axis = 1).reset_index()
    總除息列表 = 總除息列表.rename(columns = {'index':date})

    date_dic[date] = 總除息列表
    
gc = pygsheets.authorize(service_account_file='thi-warrant-62cee492afc0.json')

survey_url = 'https://docs.google.com/spreadsheets/d/17DZ6JQDPf8j1jFtGZ_1OLE1bLgxtBwrUrKSztMOWZ8A/edit#gid=0'
sh = gc.open_by_url(survey_url)

ws = sh.worksheet_by_title('五日內除息總表')

ws.clear()

for ind, cell in enumerate(['A1:C1',"E1:G1","A9:C9","E9:G9","A17:C17"]) :
    date = date_list[ind]
    ws.update_value(cell, " {} 預估除息總表".format(date))

for ind, cell in enumerate(["A2",'E2','A10','E10','A18']) :
    date = date_list[ind]
    ws.set_dataframe(date_dic[date].set_index(date).apply(lambda x : np.around(x,4)).replace("nan",""), cell, copy_index=True, nan='')
# 
ws.update_value("E17:G17", "更新時間 : {}".format(str((dt.datetime.today()))[:19]))