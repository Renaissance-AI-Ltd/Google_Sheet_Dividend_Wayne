#!/usr/bin/env python
# coding: utf-8

# In[1]:


import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime
import os
import PyWinDDE
import tejapi
plt.style.use('seaborn')
import datetime as dt
import ssl
import warnings
import string
warnings.filterwarnings("ignore")


ssl._create_default_https_context = ssl._create_unverified_context


# In[2]:


結算日月份=[7,8,9]
add_hours=12
print("add hours : %s" %(add_hours))
print("\n")
# 12 

print("連接DQ2報價 如超過10秒無連線完成文字即表示DQ2報價系統連線異常 請點Request_dq2.bat 並重新執行此檔案")
DQ_stock_dde=PyWinDDE.DDEClient("DQII", "TWSE")
DQ_index_dde=PyWinDDE.DDEClient("DQII", "INDX")
DQ_future_dde=PyWinDDE.DDEClient("DQII", "FUSA")
print("連線完成!")


TF=float(DQ_stock_dde.request('#010.129').decode()) # 金融股現貨昨收
TE=float(DQ_stock_dde.request('#020.129').decode())
XI=10000#float(DQ_stock_dde.request('#079.129').decode())
#XI=float(10000)

TX=float(DQ_future_dde.request('WTX&.129').decode())
TWN=float(DQ_future_dde.request('STWN&.129').decode())




TWII = float(input("輸入台灣加權指數 (TWII) 昨收 : "))
TWNI = float(input("輸入富台RIC指數現貨 (FTCRTWTN) 昨收 : "))
XI=10000

path='//192.168.60.83/Wellington/Allen/Index'
twn_path='//192.168.60.83/Wellington/Allen/Index/BloombergTWNI/'
index_list=os.listdir(r'\\192.168.60.83\Wellington\Allen\Index')

TWN_list=os.listdir(r'\\192.168.60.83\Wellington\Allen\Index\BloombergTWNI')


print("富台現貨昨收 : %s" %(TWNI))
print("台指現貨昨收 : %s" %(TWII))
print("\n")
print("金融股現貨昨收 : %s" %(TF))
print("電子股現貨昨收 : %s" %(TE))
print("富台期昨收 : %s" %(TWN))
print("台指期貨昨收 : %s" %(TX))


# In[3]:


今日資料日期=(datetime.datetime.today()).date()


today_readmode=今日資料日期.strftime('%b %d %Y')

TX_use_data=sorted([i for i in index_list if ('NU' in i and 'C.XLS' in i)],reverse=True)[0]
TWN_use_data=[i for i in TWN_list if today_readmode in i][0]
MTW_use_data=sorted([i for i in index_list if 'tw_per' in i],reverse=True)[0]

print('TX_Use:{} \nTWN_Use:{} \nMTW_Use:{}'.format(TX_use_data,TWN_use_data,MTW_use_data))


MTW=pd.read_excel(path+'/'+MTW_use_data,header=0,index_col=0).dropna(thresh=3,axis=0).iloc[1][1]


# In[4]:


def Init_constitute_data(path, TX_use_data, TWN_use_data, MTW_use_data): # 抓成分股有啥
    加權指數成分股=pd.read_html(path+'/'+TX_use_data,encoding='cp950',header=0)[-1]

    加權指數成分股.columns=加權指數成分股.iloc[0]
    加權指數成分股=加權指數成分股.iloc[1:]
    加權指數成分股['Sector Code']=加權指數成分股['Sector Code'].astype(int)
    加權指數成分股=加權指數成分股.set_index('Local Code')

    電子指數成分股=加權指數成分股[(加權指數成分股['Sector Code']>=24) & (加權指數成分股['Sector Code']<=31)]
    金融指數成分股=加權指數成分股[(加權指數成分股['Sector Code']==17)]

    非金電指數成分股=加權指數成分股[(加權指數成分股['Sector Code']<24) | (加權指數成分股['Sector Code']>31)]
    非金電指數成分股=非金電指數成分股[(非金電指數成分股['Sector Code']!=17)]

    摩根data=pd.read_excel(path+'/'+MTW_use_data,header=0,index_col=0).dropna(thresh=10,axis=0)
    摩根data.columns=摩根data.iloc[0]
    摩根data=摩根data.iloc[1:]
    摩根data['Reuters Code (RIC)']=摩根data['Reuters Code (RIC)'].apply(lambda x:x.split('.')[0])
    摩根data=摩根data.set_index('Reuters Code (RIC)')

    富台指數成分股=pd.read_excel(twn_path+'/'+TWN_use_data,header=0,index_col=0)[["股數","價格"]]
    富台指數成分股.index=[i.split(" ")[0] for i in 富台指數成分股.index]
    富台指數成分股.columns=['Shares',"Close"]
    return 加權指數成分股,電子指數成分股,金融指數成分股,非金電指數成分股,摩根data,富台指數成分股

加權指數成分股, 電子指數成分股, 金融指數成分股, 非金電指數成分股, 摩根data, 富台指數成分股 = Init_constitute_data(path, TX_use_data, TWN_use_data, MTW_use_data)


use_calc_data=datetime.datetime(year=datetime.datetime.today().year,month=datetime.datetime.today().month,day=1) # 計算年初時間

date_list=[pd.to_datetime(str((use_calc_data+datetime.timedelta(i)).date())) for i in range(0,200)] # 日期
date_ser=pd.Series(date_list) 




def Create_TX_closeday_func(date_ser):
    month_list=[]
    week_list=[]
    day_list=[]
    close_day_list=[]
    count=0
    week_num=1
    for i in date_ser:
        month_list.append(i.month)
        day_list.append(i.isoweekday())
        if len(day_list)>=2:
            if day_list[0]>3 and count==0:
                week_num=0
                count+=1
            if day_list[-1]<day_list[-2] and month_list[-1]==month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                week_num+=1
            elif month_list[-1]!=month_list[-2] and i.isoweekday()>3:
                week_num=0
            elif month_list[-1]!=month_list[-2] and i.isoweekday()<=3:
                week_num=1
            
            if week_num==3 and day_list[-1]==3:
                close_day_list.append(i)


    台指結算日=pd.DataFrame(index=close_day_list)
    台指結算日.index.name='結算日'
    return 台指結算日

台指結算日 = Create_TX_closeday_func(date_ser)

def Create_foreign_closeday_func(date_ser):
    month_list=[]
    week_list=[]
    day_list=[]
    close_day_list=[]
    week_num=1
    for i in date_ser:
        month_list.append(i.month)
        day_list.append(i.isoweekday())
        if len(day_list)>=2:
            if day_list[-2]==1 and day_list[-2]!=5 and month_list[-1]!=month_list[-2]:
                close_day_list.append(i-datetime.timedelta(4))
            
            elif day_list[-2]==2 and month_list[-1]!=month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                close_day_list.append(i-datetime.timedelta(2))
            
            elif day_list[-2]==3 and month_list[-1]!=month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                close_day_list.append(i-datetime.timedelta(2))
            
            elif day_list[-2]==4 and month_list[-1]!=month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                close_day_list.append(i-datetime.timedelta(2))
            
            elif day_list[-2]==5 and day_list[-1]!=1 and month_list[-1]!=month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                close_day_list.append(i-datetime.timedelta(2))
            
            elif day_list[-2]==6 and month_list[-1]!=month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                close_day_list.append(i-datetime.timedelta(3))

            elif day_list[-2]==7 and month_list[-1]!=month_list[-2]:
                week_list.append([i.year,month_list[-1],week_num])
                close_day_list.append(i-datetime.timedelta(4))
            elif day_list[-1]==1 and day_list[-2]==5 and month_list[-1]!=month_list[-2]:
                close_day_list.append(i-datetime.timedelta(4))

    海外結算日=pd.DataFrame(index=close_day_list)
    海外結算日.index.name='結算日'
    return 海外結算日

海外結算日 = Create_foreign_closeday_func(date_ser)

def third_wen(date_str):
    
    y = int(date_str[0:4])
    m = int(date_str[5:7])
    
    
    day=21-(dt.date(y,m,1).weekday()+4)%7  
    return dt.date(y,m,day)  

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


# In[5]:


import os
path = '//192.168.60.81/Wellington/Wayne/除息預估/加權指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)
path = '//192.168.60.81/Wellington/Wayne/除息預估/電子指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)
path = '//192.168.60.81/Wellington/Wayne/除息預估/金融指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)
path = '//192.168.60.81/Wellington/Wayne/除息預估/非金電指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)
path = '//192.168.60.81/Wellington/Wayne/除息預估/富台指/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)
path = '//192.168.60.81/Wellington/Wayne/除息預估/摩台指/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)
path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)

path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/全除息個股序列總表/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)

path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/預估除息總表/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)

path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/預估除息公司列表/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)

path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/前三十大權值股近五筆開會資訊/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())
if not os.path.isdir(path):
    os.makedirs(path)

# ## 抓TEJ資料

# In[6]:


tejapi.ApiConfig.api_key = "SAgEsjF1Z9hWTnPsVorVaKXA22VSfp"
info = tejapi.ApiConfig.info()
data=tejapi.get('TWN/AMT', paginate=True)
table_info = tejapi.table_info('TWN/AMT')
cname_mapping_dict=table_info['columns']
data.columns=[cname_mapping_dict[i]['cname'] for i in data.columns]

data["會議日期"]=pd.to_datetime(data["會議日期"].apply(lambda x:str(x)[0:10]))
data["除息日"]=pd.to_datetime(data["除息日"].apply(lambda x:str(x)[0:10]))
data["除權日(配股)"]=pd.to_datetime(data["除權日(配股)"].apply(lambda x:str(x)[0:10]))
data['公司']=data['公司'].astype(str)


# In[7]:


tejapi.ApiConfig.api_key = "SAgEsjF1Z9hWTnPsVorVaKXA22VSfp"
info = tejapi.ApiConfig.info()
下市data=tejapi.get('TWN/AIND', paginate=True)
table_info = tejapi.table_info('TWN/AIND')
cname_mapping_dict=table_info['columns']
下市data.columns=[cname_mapping_dict[i]['cname'] for i in 下市data.columns]


# In[8]:


下市股票 = 下市data[['公司簡稱','下市日期']].dropna().reset_index(drop = True)
下市股票['下市年分'] =下市股票.下市日期.apply(lambda x : int(str(x)[0:4]))


# In[9]:


delist_dic = {}
for ind, firm_name in enumerate(下市股票.公司簡稱.unique()) :
    delist_dic[firm_name] = 下市股票.下市年分[ind]


# In[10]:


data=data.sort_values(by='會議日期')
set_data=data.set_index(['會議日期','公司']).copy()
除息_use_data=set_data
除息_use_data=除息_use_data.sort_index()
除息_use_data=除息_use_data[除息_use_data['現金股利(元)'].fillna(0)!=0] # 把沒現金股利的抓出來
除息_use_data['開會年度']=除息_use_data['開會年度'].apply(lambda x:int(str(x)[0:4]))


# ### 處理預估的除息日

# In[11]:

開會條件 = (除息_use_data['開會年度']==2022) | (除息_use_data['開會年度']==2023) # 2023 目前尚無現金股利發放的公司 
 
確認除息_use_data=除息_use_data[((除息_use_data['除息日'].isna()==False) & (開會條件))] # 把確認除息日的抓出來

今年開會公司=除息_use_data[開會條件].reset_index().sort_values(['公司','會議日期']).reset_index(drop = True).groupby('公司').tail(1) # 把最後一筆抓出來 如果股票沒下市就預估除權息日

今年預估除息點數 = 今年開會公司[今年開會公司['除息日'].isna()].reset_index(drop = True) # 把除息日為空的值抓出來
今年預估除息點數['下市年度'] = 今年預估除息點數['公司'].map(delist_dic) # mapping 下市公司
今年預估除息點數 = 今年預估除息點數[今年預估除息點數['下市年度'].isna()] # 把非下市公司抓出來預估
今年預估除息點數 = 今年預估除息點數[今年預估除息點數['常會YN／董事會D'] != "D"] # 不估董事會，只估常會

去年除息日期=除息_use_data[(除息_use_data['開會年度']!=2023)] # 除了今年的都抓出來

# In[12]:

print("\n")

print("除息公司預估，若預估除息有三十大權值股會特別標註，請特別注意並檢查")
predict_exist_company = []
predict_exist_calculate_day = []

for i in 去年除息日期.index:
    if pd.isna(去年除息日期['除息日'].loc[i][0]): # 去年除息日沒資料的話拿除權日來補
        #print(i)
        去年除息日期['除息日'].loc[i][0]=去年除息日期['除權日(配股)'].loc[i][0]

去年除息日期.loc[:,'相差日期']=(去年除息日期.reset_index()['除息日']-去年除息日期.reset_index()['會議日期']).values
去年除息日期=去年除息日期.reset_index().set_index('公司')
去年除息日期 = 去年除息日期[去年除息日期['相差日期'].notna()]
今年預估除息點數=今年預估除息點數.reset_index().set_index('公司')

alphabet = [letter for letter in string.ascii_uppercase]
for i in 今年預估除息點數.index:
    if i[-1] in alphabet : # 去除特別股
        next
    elif i in 去年除息日期[去年除息日期['相差日期'].notna()].index :
        predict_exist_company.append(i)
        predict_exist_calculate_day.append(今年預估除息點數.loc[i]["會議日期"]) 

# In[13]:
alphabet = [letter for letter in string.ascii_uppercase]
predict_company=set(predict_exist_company) # 歷年有資料的在估計
predict_start = predict_exist_calculate_day
predict_date_list=[]
dividend_list=[]
total_firm_df = 去年除息日期['相差日期'].dropna()
calculate_df = 去年除息日期[去年除息日期['相差日期'].notna()]
#台指前30 = ["2330","2317","2454","2412","6505","2308","2881","2882","2303","1303","1301","2002","2886","3711","2891","1326","1216","5880","5871","2884","2892","2207","3045","2603","2880","2382","3008","2395","2885","2912"]
#富台前30 = ["2330","2317","2454","2308","2303","2881","2412","1303","2891","2882","2002","2886","3711","2884","1301","5871","1216","2892","2885","5880","1326","1101","2880","3034","3008","2327","2382","2883","2207","2357"]
前30大權值股 =["2330","2317","2454","2412","6505","2308","2881","2303","2882","1303","1301","2002","3711","2886","2891","1326","1216","5880","5871","2884","2892","3045","2207","2603","3008","2382","2880","2885","2912","2395","1101","3034","2327","2883","2357","2609","4904","5876","2615","6415","1590","2890","3037","2379","2887","1605","4938","2801","2408","2301","2345","2409","2474","3529","8069"]


前30大預估權值股 = []

int_list = []
predict_company_sort = []
predict_company_complete = []


for i in predict_company :
    if i[-1] in alphabet :
        next
    else :
        int_list.append(int(i))
    
int_list_filter = []

for i in int_list :
    if len(str(i))<4 :
        next
    else :
        int_list_filter.append(i)
        
int_list_filter.sort()

for i in int_list_filter :
    predict_company_sort.append(str(i))

for num , i in enumerate(list(predict_company_sort)):   
    ###################### 確認歷年除權息資料數 & 相差時間
    if i in total_firm_df.index: # 如果歷年公司有相差的時間
         # 印出歷年有時間的
        
        if type(total_firm_df.loc[i])==pd.core.series.Series: # 如果有歷年除權息資料大於一筆
            
            if len(total_firm_df.loc[i]) > 1 : # 如果歷年除權息資料 >1
                
                for j in range(len(去年除息日期.loc[i])): # DF長度
                    
                    if pd.isna(total_firm_df.loc[i][j])==False and total_firm_df.loc[i][j] >= pd.to_timedelta('1D'): # 如果第N筆資料有天數 且日期大於1天 (由上而下看，最舊到最新)
                        
                        if "原會議日" in calculate_df.loc[i].臨時會開會目的[j] :
                            next
                        else :
                            use_days=total_firm_df.loc[i][j] # 用第一筆資料
                            參考日期= str(去年除息日期[去年除息日期['相差日期'].notna()].loc[i].會議日期[j])[0:10]
                            除息日期= str(去年除息日期[去年除息日期['相差日期'].notna()].loc[i].除息日[j])[0:10]
                            # print(use_days)
                    
                    else : # 如果沒資料 或者少於1天
                        next
                    
        if type(total_firm_df.loc[i]) == pd._libs.tslibs.timedeltas.Timedelta : # 歷年只有一筆資料
                use_days=total_firm_df.loc[i]
                參考日期= str(去年除息日期[去年除息日期['相差日期'].notna()].loc[i].會議日期)[0:10]
                除息日期= str(去年除息日期[去年除息日期['相差日期'].notna()].loc[i].除息日)[0:10]

    ################### 預估今年除權息時間
            
    if type(今年預估除息點數.loc[i])==pd.core.frame.DataFrame and use_days >= pd.to_timedelta('1D'): # 如果今年董事會議數>1
        
        predict_date=今年預估除息點數.loc[i]['會議日期'][-1]+use_days # 預估為最新一筆董事會後N天 (取決於上次會議日期)
        
        dividend=今年預估除息點數.loc[i]['現金股利(元)'][-1]
        
    if type(今年預估除息點數.loc[i])==pd.core.series.Series and use_days >= pd.to_timedelta('1D'):
        
        predict_date=今年預估除息點數.loc[i]['會議日期']+use_days
        
        dividend=今年預估除息點數.loc[i]['現金股利(元)']
    
    ################## 濾掉除權日是負數的值
    if  use_days < pd.to_timedelta('1D'):
        print("%s 最新一筆除息日期為負，因此不納入除息預估" %(i))
        print("\n")
        
    if  use_days >= pd.to_timedelta('1D') : # 如果除息日為正再加入預估
    
    
        predict_date_list.append(predict_date)

        dividend_list.append(dividend)
        
        predict_company_complete.append(i) #有的公司再加進去預估

    ################## 輸出公司
    if i in 前30大權值股 and use_days >= pd.to_timedelta('1D'):
        print("預估公司 : %s (此股為台指/富台前三十大成分股), 預估除息日間隔 : %s , 今年會議日期 : %s , 今年預估除息日期 : %s, 參考日期 : %s, 參考除息日期 : %s" %(i,use_days,str(predict_start[num])[0:10],str(predict_date)[0:10],參考日期,除息日期))
        前30大預估權值股.append(i)    
    elif use_days >= pd.to_timedelta('1D'):
        print("預估公司 : %s , 預估除息日間隔 : %s , 今年會議日期 : %s, 今年預估除息日期 : %s, 參考日期 : %s, 參考除息日期 : %s" %(i,use_days,str(predict_start[num])[0:10],str(predict_date)[0:10],參考日期,除息日期))
    print("\n")

print("\n")
print("除息點數")
# In[14]:

predict_date_df=pd.DataFrame((predict_date_list,predict_company_complete,dividend_list)).T
predict_date_df.columns=['除息日','公司','現金股利(元)']
predict_date_df['確認/預估']='預估'

if "2308" in predict_date_df.公司.to_list() :
    predict_date_df.loc[predict_date_df[predict_date_df['公司'] =="2308"].index[0],"除息日"] = "2023-06-30"
    
    


predict_date_time_list=[]
for i in predict_date_df['除息日']:
    if i<datetime.datetime.today().date(): # 預估除息日小於今天
        predict_date_time_list.append(pd.to_datetime(datetime.datetime.today().date()+pd.to_timedelta('14D'))) #最短14天
    else:
        predict_date_time_list.append(i)

predict_date_df['除息日']=predict_date_time_list

add_days_list=[]
for i in predict_date_df['除息日']:
    if i.isoweekday()==6:
        add_day=2
    elif i.isoweekday()==7:
        add_day=1
    else:
        add_day=0
    add_days_list.append(pd.to_timedelta('{}D'.format(add_day)))


predict_date_df['除息日']=predict_date_df['除息日']+pd.Series(add_days_list)


確認除息公司=確認除息_use_data.reset_index()[['除息日','公司','現金股利(元)']]
確認除息公司['確認/預估']='確認'


Total_dividend_df=predict_date_df.append(確認除息公司).set_index(['除息日'])
Total_dividend_df=Total_dividend_df.sort_index()
Total_dividend_df['公司']=Total_dividend_df['公司'].apply(lambda x:x.split(' ')[0])

Total_dividend_df['除息日_str'] = Total_dividend_df.index
Total_dividend_df['台指合約月份'] = Total_dividend_df.除息日_str.apply(lambda x : TXF_contract_month(str(x)))
Total_dividend_df['富台合約月份'] = Total_dividend_df.除息日_str.apply(lambda x : TWN_contract_month(str(x)))
Total_dividend_df = Total_dividend_df.drop(['除息日_str'], axis=1)


# In[15]:


def sort_data(div_data,shares_data):
    use_list=[]
    shares_data=shares_data.astype(float)
    shares_data.index=shares_data.index.astype(str)
    for i in div_data['公司']:
        try:
            use_list.append(shares_data.loc[str(i)])
        except:
            use_list.append(0)
            pass
    return use_list





富台股數list=[]
台指股數list=[]
電子股數list=[]
金融股數list=[]
摩根股數list=[]
非金電股數list=[]

台指股數list=sort_data(Total_dividend_df,加權指數成分股['Number of Shares in Issue(Unit:1000 Shares)'])
電子股數list=sort_data(Total_dividend_df,電子指數成分股['Number of Shares in Issue(Unit:1000 Shares)'])
金融股數list=sort_data(Total_dividend_df,金融指數成分股['Number of Shares in Issue(Unit:1000 Shares)'])
非金電股數list=sort_data(Total_dividend_df,非金電指數成分股['Number of Shares in Issue(Unit:1000 Shares)'])
富台股數list=sort_data(Total_dividend_df,富台指數成分股['Shares'])
摩根股數list=sort_data(Total_dividend_df,摩根data['Shares FIF Adjusted'])
    

Total_dividend_df['台指股數']=台指股數list
Total_dividend_df['電子股數']=電子股數list
Total_dividend_df['金融股數']=金融股數list
Total_dividend_df['非金電股數']=非金電股數list
Total_dividend_df['富台股數']=富台股數list
Total_dividend_df['摩根股數']=摩根股數list


加權指數總市值=(加權指數成分股['Closing Price'].astype(float)*加權指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
電子指數總市值=(電子指數成分股['Closing Price'].astype(float)*電子指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
金融指數總市值=(金融指數成分股['Closing Price'].astype(float)*金融指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
非金電指數總市值=(非金電指數成分股['Closing Price'].astype(float)*非金電指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
富台指數總市值=(富台指數成分股.astype(float)['Shares']*富台指數成分股.astype(float)['Close']).sum()
摩根指數總市值=(摩根data['Price']*摩根data['Shares FIF Adjusted']).sum()



Total_dividend_df['台指影響點數']=(Total_dividend_df['現金股利(元)']*Total_dividend_df['台指股數'])/加權指數總市值*TWII
Total_dividend_df['電子影響點數']=(Total_dividend_df['現金股利(元)']*Total_dividend_df['電子股數'])/電子指數總市值*TE
Total_dividend_df['金融影響點數']=(Total_dividend_df['現金股利(元)']*Total_dividend_df['金融股數'])/金融指數總市值*TF
Total_dividend_df['非金電影響點數']=(Total_dividend_df['現金股利(元)']*Total_dividend_df['非金電股數'])/非金電指數總市值*XI
Total_dividend_df['富台影響點數']=(Total_dividend_df['現金股利(元)']*Total_dividend_df['富台股數'])/富台指數總市值*TWNI
Total_dividend_df['摩根影響點數']=(Total_dividend_df['現金股利(元)']*Total_dividend_df['摩根股數'])/摩根指數總市值*MTW

原始ratio=(TWN/(TWNI)/(TX/(TWII)))*10000-10000
new_ratio=((TWN/(TWNI-Total_dividend_df['富台影響點數']))/(TX/(TWII-Total_dividend_df['台指影響點數'])))*10000-10000
影響ratio=(new_ratio-原始ratio)


Total_dividend_df['台富影響Ratio']=影響ratio



台指影響市值=(Total_dividend_df['現金股利(元)']*Total_dividend_df['台指股數']).reset_index().groupby('除息日')[0].sum()
電子影響市值=(Total_dividend_df['現金股利(元)']*Total_dividend_df['電子股數']).reset_index().groupby('除息日')[0].sum()
金融影響市值=(Total_dividend_df['現金股利(元)']*Total_dividend_df['金融股數']).reset_index().groupby('除息日')[0].sum()
非金電影響市值=(Total_dividend_df['現金股利(元)']*Total_dividend_df['非金電股數']).reset_index().groupby('除息日')[0].sum()
富台影響市值=(Total_dividend_df['現金股利(元)']*Total_dividend_df['富台股數']).reset_index().groupby('除息日')[0].sum()
摩根影響市值=(Total_dividend_df['現金股利(元)']*Total_dividend_df['摩根股數']).reset_index().groupby('除息日')[0].sum()



台指除息家數=(Total_dividend_df['現金股利(元)']*Total_dividend_df['台指股數']).astype(bool).astype(int).reset_index().groupby('除息日')[0].sum()
電子除息家數=(Total_dividend_df['現金股利(元)']*Total_dividend_df['電子股數']).astype(bool).astype(int).reset_index().groupby('除息日')[0].sum()
金融除息家數=(Total_dividend_df['現金股利(元)']*Total_dividend_df['金融股數']).astype(bool).astype(int).reset_index().groupby('除息日')[0].sum()
非金電除息家數=(Total_dividend_df['現金股利(元)']*Total_dividend_df['非金電股數']).astype(bool).astype(int).reset_index().groupby('除息日')[0].sum()
富台除息家數=(Total_dividend_df['現金股利(元)']*Total_dividend_df['富台股數']).astype(bool).astype(int).reset_index().groupby('除息日')[0].sum()
摩根除息家數=(Total_dividend_df['現金股利(元)']*Total_dividend_df['摩根股數']).astype(bool).astype(int).reset_index().groupby('除息日')[0].sum()


台指除息點數=pd.Series((台指影響市值/加權指數總市值)*TWII,index=date_ser).fillna(0)
電子除息點數=pd.Series((電子影響市值/電子指數總市值)*TE,index=date_ser).fillna(0)
金融除息點數=pd.Series((金融影響市值/金融指數總市值)*TF,index=date_ser).fillna(0)
非金電除息點數=pd.Series((非金電影響市值/非金電指數總市值)*XI,index=date_ser).fillna(0)
富台除息點數=pd.Series((富台影響市值/富台指數總市值)*TWNI,index=date_ser).fillna(0)
摩根除息點數=pd.Series((摩根影響市值/摩根指數總市值)*MTW,index=date_ser).fillna(0)


台指除息家數=pd.Series(台指除息家數,index=date_ser).fillna(0)
電子除息家數=pd.Series(電子除息家數,index=date_ser).fillna(0)
金融除息家數=pd.Series(金融除息家數,index=date_ser).fillna(0)
非金電除息家數=pd.Series(非金電除息家數,index=date_ser).fillna(0)
富台除息家數=pd.Series(富台除息家數,index=date_ser).fillna(0)
摩根除息家數=pd.Series(摩根除息家數,index=date_ser).fillna(0)




富台前十大成分股=round(100*((富台指數成分股['Shares']*富台指數成分股['Close']).sort_values(ascending=False).head(10)/富台指數總市值),2)
台指前十大成分股=round(100*((加權指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)*加權指數成分股['Closing Price'].astype(float)).sort_values(ascending=False).head(10)/加權指數總市值),2)



加權指數總市值=(加權指數成分股['Closing Price'].astype(float)*加權指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
電子指數總市值=(電子指數成分股['Closing Price'].astype(float)*電子指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
金融指數總市值=(金融指數成分股['Closing Price'].astype(float)*金融指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
非金電指數總市值=(非金電指數成分股['Closing Price'].astype(float)*非金電指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)).sum()
富台指數總市值=(富台指數成分股.astype(float)['Shares']*富台指數成分股.astype(float)['Close']).sum()
摩根指數總市值=(摩根data['Price']*摩根data['Shares FIF Adjusted']).sum()


富台權重佔比=round(100*((富台指數成分股['Shares']*富台指數成分股['Close']).sort_values(ascending=False)/富台指數總市值),2)
台指權重佔比=round(100*((加權指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)*加權指數成分股['Closing Price'].astype(float)).sort_values(ascending=False)/加權指數總市值),2)
電子權重佔比=round(100*((電子指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)*電子指數成分股['Closing Price'].astype(float)).sort_values(ascending=False)/電子指數總市值),2)
金融權重佔比=round(100*((金融指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)*金融指數成分股['Closing Price'].astype(float)).sort_values(ascending=False)/金融指數總市值),2)
非金電權重佔比=round(100*((非金電指數成分股['Number of Shares in Issue(Unit:1000 Shares)'].astype(float)*非金電指數成分股['Closing Price'].astype(float)).sort_values(ascending=False)/非金電指數總市值),2)
摩根權重佔比=round(100*((摩根data['Price']*摩根data['Shares FIF Adjusted']).astype(float).sort_values(ascending=False)/摩根指數總市值),2)


future_代碼=pd.read_html('https://www.taifex.com.tw/cht/2/stockLists')[1].set_index('證券代號')['股票期貨、選擇權商品代碼']
future_代碼=future_代碼.dropna()+'F'
future_代碼.index=future_代碼.index.astype(int).astype(str)
future_代碼=future_代碼.sort_index()
future_代碼=future_代碼[future_代碼.index.duplicated('last')==False]
future_代碼='W'+future_代碼+'&'

new_future_index=pd.Series(list(future_代碼.index)).replace('50','0050').replace('56','0056')
future_代碼.index=new_future_index.astype(str)
future_代碼=future_代碼[future_代碼.index.duplicated('first')==False]


future_代碼_list=[]
for i in Total_dividend_df['公司'].values:
    try:
        future_代碼_list.append(future_代碼.loc[i])
    except:
        future_代碼_list.append(np.nan)


Total_dividend_df['期貨代碼']=future_代碼_list

stock_last_price_list=[]
for i in Total_dividend_df['公司']:
    if i in 加權指數成分股.index:
        stock_last_price_list.append(加權指數成分股['Closing Price'].loc[i])
    elif i in 摩根data.index:
        stock_last_price_list.append(摩根data['Price'].loc[i])
    elif i in 富台指數成分股.index:
        stock_last_price_list.append(富台指數成分股['Close'].loc[i])
    else:
        stock_last_price_list.append(DQ_stock_dde.request('{}.129'.format(i)).decode())


# In[16]:


Total_dividend_df.to_csv('//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/全除息個股序列總表/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/全除息個股序列總表_公槽_備援_'+
                         str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+" "+
                         str(datetime.datetime.today())[11:19].replace(":",'-')+'.csv',encoding='utf_8_sig')


Total_dividend_df.to_csv('//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/全除息個股序列總表_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')




富台權重佔比.index=富台權重佔比.index.astype(str)


Total_dividend_df.loc[str(datetime.datetime.today()):].dropna()



Total_dividend_df[Total_dividend_df['確認/預估']=='預估']['台指影響點數'].sum()#.loc[str(datetime.datetime.today())[:10]:]



權重data=pd.concat([台指權重佔比,電子權重佔比,金融權重佔比,非金電權重佔比,富台權重佔比,摩根權重佔比,future_代碼],axis=1,
                 keys=['TX weight','TE weight','TF weight','XIF weight','TWN weight','MTW weight','future code']).fillna(0)





total_list=[]
item=['101','102','125','113','114']
for i in 權重data['future code']:
    use_list=[]
    for t in item:
        if i !=0:
            use_list.append(f"=DQII|FUSA!'{i}.{t}'")
        else:
            use_list.append("")
    total_list.append(pd.Series(use_list))

future_data=pd.concat(total_list,axis=1).T



權重data['F Bid']=future_data[0].values
權重data['F Ask']=future_data[1].values
權重data['F Trade']=future_data[2].values
權重data['F Bid Volume']=future_data[3].values
權重data['F Ask Volume']=future_data[4].values


前十大用df=Total_dividend_df.loc[str(datetime.datetime.today().year)+'-'+str(datetime.datetime.today().month):].reset_index().set_index('公司')


# In[17]:


國內除息_list=[]
for data in zip([台指除息點數,電子除息點數,金融除息點數,非金電除息點數],[台指除息家數,電子除息家數,金融除息家數,非金電除息家數]):
    use_dict={}
    for i in 台指結算日.iloc[0:3].index:
        month_data_calc=(data[0].loc[:i].loc[::-1].cumsum()[::-1])
        use_month=str(i.month)
        if len(use_month)==1:
            use_month='0'+use_month
        use_dict[f'{use_month}月合約加總剩餘影響點數']=month_data_calc-data[0]
    加總影響點數=data[0]
    加總除權息家數=data[1]
    加總累計影響點數=data[0].cumsum()
    
    除息_dataframe=pd.DataFrame([加總影響點數,加總除權息家數,加總累計影響點數]).T
    除息_dataframe.columns=['加總影響點數','加總除權息家數','加總累計影響點數']
    國內除息_list.append(除息_dataframe.join(pd.DataFrame(use_dict)))



國內除息_list[0].to_csv('//192.168.60.81/Wellington/Wayne/除息預估/加權指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/加權指數_加總除權息總表_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')
國內除息_list[1].to_csv('//192.168.60.81/Wellington/Wayne/除息預估/電子指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/電子指數_加總除權息總表_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')
國內除息_list[2].to_csv('//192.168.60.81/Wellington/Wayne/除息預估/金融指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/金融指數_加總除權息總表_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')
國內除息_list[3].to_csv('//192.168.60.81/Wellington/Wayne/除息預估/非金電指數/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/金融指數_加總除權息總表_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')



國外除息_list=[]
for data in zip([富台除息點數,摩根除息點數],[富台除息家數,摩根除息家數]):
    use_dict={}
    for i in 海外結算日.iloc[0:3].index:
        month_data_calc=(data[0].loc[:i].loc[::-1].cumsum()[::-1])
        use_month=str(i.month)
        if len(use_month)==1:
            use_month='0'+use_month
        use_dict[f'{use_month}月合約加總剩餘影響點數']=month_data_calc-data[0]
    加總影響點數=data[0]
    加總除權息家數=data[1]
    加總累計影響點數=data[0].cumsum()
    
    除息_dataframe=pd.DataFrame([加總影響點數,加總除權息家數,加總累計影響點數]).T
    除息_dataframe.columns=['加總影響點數','加總除權息家數','加總累計影響點數']
    國外除息_list.append(除息_dataframe.join(pd.DataFrame(use_dict)))



國外除息_list[0].to_csv('//192.168.60.81/Wellington/Wayne/除息預估/富台指/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/富台指_加總除權息總表_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')
國外除息_list[1].to_csv('//192.168.60.81/Wellington/Wayne/除息預估/摩台指/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/摩台指_加總除權息總表_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')

總表=pd.DataFrame(index=國外除息_list[0].columns)
總表.index.name=str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())

總表['加權指數']=國內除息_list[0].loc[str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())]
總表['電子指數']=國內除息_list[1].loc[str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())]
總表['金融指數']=國內除息_list[2].loc[str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())]
總表['非金電指數']=國內除息_list[3].loc[str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())]
總表['摩台指']=國外除息_list[1].loc[str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())]
總表['富台指']=國外除息_list[0].loc[str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())]



總表.drop('加總累計影響點數',inplace=True)
總表=總表.round(4)
總表.loc['加總除權息家數']=總表.loc['加總除權息家數'].astype(int)

總表.to_csv('//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/預估除息總表/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/預估除息總表_公槽_備援_'+
str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+" "+
str(datetime.datetime.today())[11:19].replace(":",'-')+'.csv',encoding='utf_8_sig')


總表.to_csv('//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/預估除息總表_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.csv',encoding='utf_8_sig')





print(總表)


# In[18]:


total_lst = predict_date_df.公司
str_list = [str(x) for x in total_lst]
# str_list.append('000930')
df_total_list= {}

data=tejapi.get('TWN/AMT', paginate=True)
table_info = tejapi.table_info('TWN/AMT')
cname_mapping_dict=table_info['columns']
data.columns=[cname_mapping_dict[i]['cname'] for i in data.columns]

data["會議日期"]=pd.to_datetime(data["會議日期"].apply(lambda x:str(x)[0:10]))
data["除息日"]=pd.to_datetime(data["除息日"].apply(lambda x:str(x)[0:10]))
data["除權日(配股)"]=pd.to_datetime(data["除權日(配股)"].apply(lambda x:str(x)[0:10]))
data['公司']=data['公司'].astype(str)



# In[19]:

for i in total_lst : 

    df_total_list[i] = data.reset_index()[data.reset_index()['公司'] == i][['會議日期','公司',
                                                                                       '常會YN／董事會D','董事會日期','現金股利(元)',"除息日",'股息分配型態','臨時會開會目的']].tail(5)
    df_total_list[i]['下市年分'] = df_total_list[i]['公司'].map(delist_dic)
    df_total_list[i]['相差日期'] = (df_total_list[i]['除息日']-df_total_list[i]['會議日期']).values
    df_total_list[i]['會議日期'] = df_total_list[i]['會議日期'].apply(lambda x : str(x))
    df_total_list[i]['除息日'] = df_total_list[i]['除息日'].apply(lambda x : str(x))
    df_total_list[i]['董事會日期'] = df_total_list[i]['董事會日期'].apply(lambda x : str(x))

path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/預估除息公司列表_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.xlsx'

with pd.ExcelWriter(path) as writer:  
    
    for ind ,firm_name in enumerate(df_total_list.keys()) :
        
        df_total_list[firm_name].to_excel(writer,startrow = ind*7+ind+1)

path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/預估除息公司列表/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/預估除息公司列表_公槽_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.xlsx'

with pd.ExcelWriter(path) as writer:  
    
    for ind ,firm_name in enumerate(df_total_list.keys()) :
        
        df_total_list[firm_name].to_excel(writer,startrow = ind*7+ind+1)


print("\n")
print("輸出台指/富台 前30大權值股近五筆開會資料")

int_list = []
前30大權值股_sort = []

for i in 前30大權值股 :
    int_list.append(int(i))
    
int_list.sort()
for i in int_list :
    前30大權值股_sort.append(str(i))


#前30大權值股_int = [2330,2317,2454,2412,6505,2308,2881,2303,2882,1303,1301,2002,3711,2886,2891,1326,1216,5880,5871,2884,2892,3045,2207,2603,3008,2382,2880,2885,2912,2395,1101,3034,2327,2883,2357]

前30大權值股_str = ["2330","2317","2454","2412","6505","2308","2881","2303","2882","1303","1301","2002","3711","2886","2891","1326","1216","5880","5871","2884","2892","3045","2207","2603","3008","2382","2880","2885","2912","2395","1101","3034","2327","2883","2357","2609","4904","5876","2615","6415","1590","2890","3037","2379","2887","1605","4938","2801","2408","2301","2345","2409","2474","3529","8069"]


前30大預估權值股.sort(key=lambda x: 前30大權值股_str.index(x))

df_total_list = {}
for i in 前30大權值股_str : 

    df_total_list[i] = data.reset_index()[data.reset_index()['公司'] == i][['會議日期','公司',
                                                                                       '常會YN／董事會D','董事會日期','現金股利(元)',"除息日",'股息分配型態','臨時會開會目的']].tail(5)
    df_total_list[i]['相差日期'] = (df_total_list[i]['除息日']-df_total_list[i]['會議日期']).values
    df_total_list[i]['會議日期'] = df_total_list[i]['會議日期'].apply(lambda x : str(x))
    df_total_list[i]['除息日'] = df_total_list[i]['除息日'].apply(lambda x : str(x))
    df_total_list[i]['董事會日期'] = df_total_list[i]['董事會日期'].apply(lambda x : str(x))
    
path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/前三十大權值股近五筆開會資訊_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.xlsx'
with pd.ExcelWriter(path) as writer:  
    
    for ind ,firm_name in enumerate(df_total_list.keys()) :
        
        df_total_list[firm_name].to_excel(writer,startrow = ind*7+ind+1)


path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/log/前三十大權值股近五筆開會資訊/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/前三十大權值股近五筆開會資訊_公槽_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.xlsx'

with pd.ExcelWriter(path) as writer:  
    
    for ind ,firm_name in enumerate(df_total_list.keys()) :
        
        df_total_list[firm_name].to_excel(writer,startrow = ind*7+ind+1)

# In[53]:


df_total_list = {}
for j in 前30大預估權值股 : 

    df_total_list[j] = data.reset_index()[data.reset_index()['公司'] == j][['會議日期','公司',
                                                                                       '常會YN／董事會D','董事會日期','現金股利(元)',"除息日",'股息分配型態','臨時會開會目的']].tail(5)

    df_total_list[j]['相差日期'] = (df_total_list[j]['除息日']-df_total_list[j]['會議日期']).values
    df_total_list[j]['會議日期'] = df_total_list[j]['會議日期'].apply(lambda x : str(x))
    df_total_list[j]['除息日'] = df_total_list[j]['除息日'].apply(lambda x : str(x))
    df_total_list[j]['董事會日期'] = df_total_list[j]['董事會日期'].apply(lambda x : str(x))
    
path = '//192.168.60.81/Wellington/Wayne/除息預估/除息總整理/'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'/前30大預估權值股_備援_'+str((datetime.datetime.today()+datetime.timedelta(hours=add_hours)).date())+'.xlsx'
with pd.ExcelWriter(path) as writer:  
    
    for ind ,firm_name in enumerate(df_total_list.keys()) :
        
        df_total_list[firm_name].to_excel(writer,startrow = ind*7+ind+1)


