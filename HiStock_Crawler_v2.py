# -*- coding: utf-8 -*-
"""
Created on Mon Mar  1 15:12:06 2021
20210410_Change log : 
    1. 爬蟲加入爬取技術指標(MA、KD、RSI、MACD、OSC) 
    2. 加入站上均線、三大法人、成值選股策略
    3. 加入Line Message 提醒 
    4. Excel存入指定資料夾
    5. 加入輸出 LOG 機制
@author: BennyXu
"""

# Import Package
import re
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib
import logging

## 測時間
import time
## CWD
import os
print("Current Working Directory is : ",os.getcwd())

# Parameters
## 取得 HiStock URL
URL = "https://histock.tw"
## 創建 Excel, DataFrame 並加上抓取日期
TD = date.today().strftime("%Y%m%d")
PATH = os.path.join(os.getcwd(),'HiStock_Crawler\\HiStock_Crawler_%s.xlsx'%TD)
PATH_UPDN = os.path.join(os.getcwd(),'HiStock_Crawler_UpDn\\HiStock_Crawler_UpDn_%s.xlsx'%TD)
PATH_LOG = os.path.join(os.getcwd(),'LOG\\HiStock_Crawler_%s.log'%TD)
EW = pd.ExcelWriter (PATH)
EW2 = pd.ExcelWriter (PATH_UPDN)
## LOGGER
LOGGER = logging.getLogger()
LOGGER.setLevel(logging.DEBUG)
FORMATER = logging.Formatter(
    '[%(levelname)s-Line:%(lineno)d] %(asctime)s %(message)s',\
    datefmt='%Y%m%d %H:%M:%S')

# HiStock 爬蟲
def HiStock_Web_Crawler(URL = URL, EW = EW):
    url = URL + "/twstock"
    attempt0 = 0
    while attempt0 < 3:
        try:
            re1 = requests.get(url)
            break
        except:
            attempt0 += 1
            logging.debug(" is when Connection refused by the server..")
            time.sleep(5)
            continue
    soup = BeautifulSoup(re1.content, "html.parser")
    ## 上市公司產業類股表格 
    table = soup.find_all("table", {"id":"tb_list"})[0]
    
    for row in table.find_all("a"):
        ## 剔除金融股
        if 'A035' not in str(row):
            DICT = {"股票" : [], "代號" : [], "交易量(張)" : [], "開盤價" : [], "最高價" : [], "最低價" : [], \
                    "收盤價" : [], "EPS" : [], "本益比" : [], "股價淨值比" : [], "現金殖利率" : [],\
                    "外資日期" : [], "外資天數" : [], "外資張數" : [], "投信日期" : [], "投信天數" : [],\
                    "投信張數" : [], "自營日期" : [], "自營天數" : [], "自營張數" : [],"融資日期" : [],\
                    "融資天數" : [], "融資張數" : [], "融券日期" : [], "融券天數" : [], "融券張數" : [],\
                    "營收日期" : [], "營收天數" : [], "漲跌幅(點)" : [], "漲跌幅(%)" : [], "MV5" : [], \
                    "MV10" : [], "MV20" : [], "MV60" : [], "K9" : [], "D9" : [], "RSI6" : [], "RSI12" : [],\
                    "DIF" : [], "MACD" : [], "OSC" : []}
            # print('----------------------Now crawling %s 類股----------------------' %str(row['title']))
            logging.info(' is when crawling %s 類股'%str(row['title']))
            url2 = URL + row["href"]
            attempt1 = 0
            while attempt1 < 3:
                try:
                    re2 = requests.get(url2)
                    break
                except:
                    attempt1 += 1
                    logging.debug(" is when Connection refused by the server..")
                    time.sleep(5)
                    continue
            soup2 = BeautifulSoup(re2.content, "html.parser")
            
            ## 類股中的個股表格
            table2 = soup2.find_all("a", {"class":"link"})      
            ## 個股迴圈
            for col in table2:
                if '/twclass/' not in str(col):
                    # print('----------------Now crawling %s 個股----------------' %str(col['title']))
                    logging.info(' is when crawling %s 個股'%str(col['title']))
                    ## 個股中文,代號
                    DICT['股票'].append(col['title'])
                    DICT['代號'].append(col['href'].replace('/stock/', ''))
                    ticker = col['href'].replace('/stock/', '')
                    
                    # URL 
                    url3 = URL + col["href"]
                    urlmv5 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=mean5'
                    urlmv10 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=mean10'
                    urlmv20 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=mean20'
                    urlmv60 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=mean60'
                    urlk9 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=k9'
                    urld9 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=d9'
                    urlrsi6 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=rsi6'
                    urlrsi12 = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=rsi12'
                    urldif = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=rsi12'
                    urlmacd = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=macd'
                    urlosc = 'https://histock.tw/stock/chip/chartdata.aspx?no=' + ticker + '&days=80&m=osc'
                    
                    attempt2 = 0
                    while attempt2 < 3:
                        try:
                            re3 = requests.get(url3)
                            remv5 = requests.get(urlmv5)
                            remv10 = requests.get(urlmv10)
                            remv20 = requests.get(urlmv20)
                            remv60 = requests.get(urlmv60)
                            rek9 = requests.get(urlk9)
                            red9 = requests.get(urld9)
                            rersi6 = requests.get(urlrsi6)
                            rersi12 = requests.get(urlrsi12)
                            redif = requests.get(urldif)
                            remacd = requests.get(urlmacd)
                            reosc = requests.get(urlosc)
                            break
                        except:
                            attempt2 += 1
                            logging.debug(" is when Connection refused by the server..")
                            time.sleep(5)
                            continue
                    # Soup
                    soup3 = BeautifulSoup(re3.content, "html.parser")
                    soupmv5 = BeautifulSoup(remv5.content, "html.parser")
                    soupmv10 = BeautifulSoup(remv10.content, "html.parser")
                    soupmv20 = BeautifulSoup(remv20.content, "html.parser")
                    soupmv60 = BeautifulSoup(remv60.content, "html.parser")
                    soupk9 = BeautifulSoup(rek9.content, "html.parser")
                    soupd9 = BeautifulSoup(red9.content, "html.parser")
                    souprsi6 = BeautifulSoup(rersi6.content, "html.parser")
                    souprsi12 = BeautifulSoup(rersi12.content, "html.parser")
                    soupdif = BeautifulSoup(redif.content, "html.parser")
                    soupmacd = BeautifulSoup(remacd.content, "html.parser")
                    souposc = BeautifulSoup(reosc.content, "html.parser")
                    
                    # Data
                    data = soup3.find_all('script', {"type" :"text/javascript"})
                    data2 = soup3.find_all('table',{"class" : "tb-stock tbChip"})
                    data3 = soup3.find('ul', {'class' : 'priceinfo mt10'}).find_all('span', {'id' : 'Price1_lbTChange'})
                    data33 = soup3.find('ul', {'class' : 'priceinfo mt10'}).find_all('span', {'id' : 'Price1_lbTPercent'})
                    mv5 = str(soupmv5).split(',')[-1].split(']')[0]
                    mv10 = str(soupmv10).split(',')[-1].split(']')[0]
                    mv20 = str(soupmv20).split(',')[-1].split(']')[0]
                    mv60 = str(soupmv60).split(',')[-1].split(']')[0]
                    k9 = str(soupk9).split(',')[-1].split(']')[0]
                    d9 = str(soupd9).split(',')[-1].split(']')[0]
                    rsi6 = str(souprsi6).split(',')[-1].split(']')[0]
                    rsi12 = str(souprsi12).split(',')[-1].split(']')[0]
                    dif = str(soupdif).split(',')[-1].split(']')[0]
                    macd = str(soupmacd).split(',')[-1].split(']')[0]
                    osc = str(souposc).split(',')[-1].split(']')[0]
                    
                    # 價量資料
                    for string in re.split(r";|}|{",str(data)):
                        if ('candlestick' in string) and ('data:' in string):
                            ## 取最新開、高、低、收
                            list_1 = [float(i) for i in re.findall(r"\d+\.?\d*",string)[-4:]]
                            ## 取最新交易量(張)
                        if ('成交量(張)' in string) and ('data:' in string):
                            list_2 = [float(i) for i in re.findall(r"\d+\.?\d*",string)[-4:-3]]
                    ## 取 EPS、現金殖利率
                    list_3 = [content.text for content in soup3.find('table', {'class' : "tb-stock tbBasic"})\
                              .find_all('td')][4:8]
                    # 其他資料
                    for content in data2:
                        ## 取外資、自營、投信、融資、融券資料
                        Revenue = [content2.text for content2 in content.find_all('td')[-2:]]
                        Short = [content2.text for content2 in content.find_all('td')[-5:-2]]
                        Margin = [content2.text for content2 in content.find_all('td')[-8:-5]]
                        Trust = [content2.text for content2 in content.find_all('td')[-11:-8]]
                        Dealer = [content2.text for content2 in content.find_all('td')[-14:-11]]
                        Foreign = [content2.text for content2 in content.find_all('td')[-17:-14]]
                    ## 取漲跌幅(點數、趴數)
                    UpDn = [content3.text for content3 in data3]
                    UpDn_PC = [content3.text for content3 in data33]
                    
                    ## 爬不到或已下市股票儲存'Error'
                    if len(list_1) != 4:
                        list_1 = ['Error data'] * 4
                        logging.debug(" is when %s Price data Error"%str(col['title']))
                    if len(list_2) != 1:
                        list_2 = ['Error data']
                        logging.debug(" is when %s Volumn data Error"%str(col['title']))
                    if len(list_3) != 4:
                        list_3 = ['Error data'] * 4
                        logging.debug(" is when %s EPS data Error"%str(col['title']))
                    if len(Revenue) != 2:
                        Revenue = ['Error data'] * 2
                        logging.debug(" is when %s Revenue data Error"%str(col['title']))
                    if len(Short) != 3:
                        Short = ['Error data'] * 3
                        logging.debug(" is when %s Short data Error"%str(col['title']))
                    if len(Margin) != 3:
                        Margin = ['Error data'] * 3
                        logging.debug(" is when %s Margin data Error"%str(col['title']))
                    if len(Trust) != 3:
                        Trust = ['Error data'] * 3
                        logging.debug(" is when %s Trust data Error"%str(col['title']))
                    if len(Dealer) != 3:
                        Dealer = ['Error data'] * 3
                        logging.debug(" is when %s Dealer data Error"%str(col['title']))
                    if len(Foreign) != 3:
                        Foreign = ['Error data'] * 3
                        logging.debug(" is when %s Foreign data Error"%str(col['title']))
                    if len(UpDn) != 1:
                        UpDn = ['Error data']
                        logging.debug(" is when %s UpDn data Error"%str(col['title']))
                    if len(UpDn_PC) != 1:
                        UpDn_PC = ['Error data']
                        logging.debug(" is when %s UpDn_PC data Error"%str(col['title']))
                        
                    ## 將最新交易量(張)、開、高、低、收 裝入 DICT
                    DICT['交易量(張)'].append(list_2[0])
                    DICT['開盤價'].append(list_1[0])
                    DICT['最高價'].append(list_1[1])
                    DICT['最低價'].append(list_1[2])
                    DICT['收盤價'].append(list_1[3])
                    DICT['EPS'].append(list_3[0])
                    DICT['本益比'].append(list_3[1])
                    DICT['股價淨值比'].append(list_3[2])
                    DICT['現金殖利率'].append(list_3[3])
                    
                    ## 將外資、自營、投信、融資、融券資料裝入DICT
                    DICT['外資日期'].append(Foreign[0])
                    DICT['外資天數'].append(Foreign[1])
                    DICT['外資張數'].append(Foreign[2])
                    DICT['投信日期'].append(Trust[0])
                    DICT['投信天數'].append(Trust[1])
                    DICT['投信張數'].append(Trust[2])
                    DICT['自營日期'].append(Dealer[0])
                    DICT['自營天數'].append(Dealer[1])
                    DICT['自營張數'].append(Dealer[2])
                    DICT['融資日期'].append(Margin[0])
                    DICT['融資天數'].append(Margin[1])
                    DICT['融資張數'].append(Margin[2])
                    DICT['融券日期'].append(Short[0])
                    DICT['融券天數'].append(Short[1])
                    DICT['融券張數'].append(Short[2])
                    DICT['營收日期'].append(Revenue[0])
                    DICT['營收天數'].append(Revenue[1])
                    
                    ## 將漲跌幅(點數、趴數)裝入
                    DICT['漲跌幅(點)'].append(UpDn[0])
                    DICT['漲跌幅(%)'].append(UpDn_PC[0])
                    
                    ## 將技術指標裝入
                    DICT['MV5'].append(mv5)
                    DICT['MV10'].append(mv10)
                    DICT['MV20'].append(mv20)
                    DICT['MV60'].append(mv60)
                    DICT['K9'].append(k9)
                    DICT['D9'].append(d9)
                    DICT['RSI6'].append(rsi6)
                    DICT['RSI12'].append(rsi12)
                    DICT['DIF'].append(dif)
                    DICT['MACD'].append(macd)
                    DICT['OSC'].append(osc)
                    
        ## 類股 DataFrame
        Result = pd.DataFrame.from_dict(DICT)
        
        ## Sort
        Result = Result.sort_values(by = ['漲跌幅(%)', '交易量(張)'], ascending = False)
        Result.to_excel(excel_writer=EW, sheet_name=row['title'],index=None) 
    return Result

# 選股策略
def Pick_Strategy(xls):
    Dict = {'Name' : [], 'Ticker' : [], '5mv' : [], '10mv' : [], '20mv' : [],\
            'Open' : [], 'High' : [],'Close' : [], 'Low' : [],\
            'UpDn(%)' : [], 'UpDn(pt)' : [],'Volumn' : [], 'Condition' : [], 'Sort_key' : []}
    Dict1 = {'Name' : [], 'Ticker' : [],\
            'Open' : [], 'High' : [],'Close' : [], 'Low' : [],\
            'UpDn(%)' : [], 'UpDn(pt)' : [],'Volumn' : [], 'Sort_key' : []}
    for num, j in enumerate(xls.sheet_names):
        if j != '金融':
            df = pd.read_excel(xls, j)
            if not df.empty:
                # 取閥值
                df['成值'] = df['交易量(張)'].map(lambda x: int(x) if type(x) is int else 0) * \
                             df['收盤價'].map(lambda x: float(x) if type(x) is float else 0) 
                df['漲幅'] = [float(re.findall(r"\d+\.?\d*",df['漲跌幅(%)'][i])[0]) \
                                 if '-' not in df['漲跌幅(%)'][i] else 0 for i in df['漲跌幅(%)'].index] # 僅取正值
                df['MV5'] = df['MV5'].map(lambda x: float(x) if type(x) is float else 0)
                df['MV10'] = df['MV10'].map(lambda x: float(x) if type(x) is float else 0)
                df['MV20'] = df['MV20'].map(lambda x: float(x) if type(x) is float else 0)

                # 站上均線 and漲幅不超 3%
                df['con1'] = [0 if (type(df.loc[i,'收盤價']) is str or type(df.loc[i,'MV5']) is str) else 1 if\
                            df.loc[i,'開盤價'] <= df.loc[i,'MV5'] and df.loc[i,'收盤價'] >= df.loc[i,'MV5'] and \
                            (df.loc[i,'收盤價']-df.loc[i,'MV5'])/df.loc[i,'MV5'] <= 0.03 else 0 for i in df.index]
                df['con1_1'] = [0 if (type(df.loc[i,'收盤價']) is str or type(df.loc[i,'MV10']) is str) else 1 if\
                            df.loc[i,'開盤價'] <= df.loc[i,'MV10'] and df.loc[i,'收盤價'] >= df.loc[i,'MV10'] and \
                            (df.loc[i,'收盤價']-df.loc[i,'MV10'])/df.loc[i,'MV10'] <= 0.03 else 0 for i in df.index]
                df['con1_2'] = [0 if (type(df.loc[i,'收盤價']) is str or type(df.loc[i,'MV20']) is str) else 1 if\
                            df.loc[i,'開盤價'] <= df.loc[i,'MV20'] and df.loc[i,'收盤價'] >= df.loc[i,'MV20'] and \
                            (df.loc[i,'收盤價']-df.loc[i,'MV20'])/df.loc[i,'MV20'] <= 0.03 else 0 for i in df.index]
                # 成值在類股前 75%
                df['con2'] = [0 if (type(df.loc[i,'成值']) is str) else 1 if \
                              df.loc[i,'成值'] >= np.percentile(df['成值'],75) else 0 for i in df.index]
                # 外資買 or 外資、投信同買 or 三大法人同買
                df['con3'] = [1 if ('買' in df.loc[i,'外資天數'] and int(df.loc[i,'外資天數'].split('超')[1].split('天')[0]) > 0)\
                                else 0 for i in df.index]
                df['con4'] = [1 if ('買' in df.loc[i,'外資天數'] and int(df.loc[i,'外資天數'].split('超')[1].split('天')[0]) > 0 \
                                and '買' in df.loc[i,'投信天數'] and int(df.loc[i,'投信天數'].split('超')[1].split('天')[0]) > 0)\
                                else 0 for i in df.index]
                df['con5'] = [1 if ('買' in df.loc[i,'外資天數'] and int(df.loc[i,'外資天數'].split('超')[1].split('天')[0]) > 0 \
                                and '買' in df.loc[i,'投信天數'] and int(df.loc[i,'投信天數'].split('超')[1].split('天')[0]) > 0 \
                                and '買' in df.loc[i,'自營天數'] and int(df.loc[i,'自營天數'].split('超')[1].split('天')[0]) > 0)
                                else 0 for i in df.index]
                df['condition'] = ['站上5日線' if df.loc[i,'con1'] == 1 else '站上10日線' if df.loc[i,'con1_1'] == 1 \
                                   else '站上月線' if df.loc[i,'con1_2'] == 1 else '0' for i in df.index]
                condition = np.where(((df['con1'] == 1) | (df['con1_1'] == 1) | (df['con1_2'] == 1)) & (df['con2'] == 1) \
                                     & ((df['con3'] == 1) | (df['con4'] == 1) | (df['con5'] == 1)))

                
                for i in df.loc[condition].index:
                    Dict['Name'].append(df.loc[i,'股票'])
                    Dict['Ticker'].append(df.loc[i,'代號'])
                    Dict['5mv'].append(df.loc[i,'MV5'])
                    Dict['10mv'].append(df.loc[i,'MV10'])
                    Dict['20mv'].append(df.loc[i,'MV20'])
                    Dict['Open'].append(df.loc[i,'開盤價'])
                    Dict['High'].append(df.loc[i,'最高價'])
                    Dict['Close'].append(df.loc[i,'收盤價'])
                    Dict['Low'].append(df.loc[i,'最低價'])
                    Dict['UpDn(%)'].append(df.loc[i,'漲跌幅(%)'])
                    Dict['UpDn(pt)'].append(df.loc[i,'漲跌幅(點)'])
                    Dict['Volumn'].append(df.loc[i,'交易量(張)'])
                    Dict['Condition'].append(df.loc[i,'condition'])
                    Dict['Sort_key'].append(df.loc[i,'漲幅'])
                for i in df.index:
                    # 漲幅超過 8% 之飆股
                    if df.loc[i,'漲幅'] > 8:
                        Dict1['Name'].append(df.loc[i,'股票'])
                        Dict1['Ticker'].append(df.loc[i,'代號'])
                        Dict1['Open'].append(df.loc[i,'開盤價'])
                        Dict1['High'].append(df.loc[i,'最高價'])
                        Dict1['Close'].append(df.loc[i,'收盤價'])
                        Dict1['Low'].append(df.loc[i,'最低價'])
                        Dict1['UpDn(%)'].append(df.loc[i,'漲跌幅(%)'])
                        Dict1['UpDn(pt)'].append(df.loc[i,'漲跌幅(點)'])
                        Dict1['Volumn'].append(df.loc[i,'交易量(張)'])
                        Dict1['Sort_key'].append(df.loc[i,'漲幅'])
    # Strategy_UpDn
    Strategy_UpDn = pd.DataFrame.from_dict(Dict)    
    # Highest_UpDn
    Highest_UpDn = pd.DataFrame.from_dict(Dict1)
    # Remove Duplicates
    Strategy_UpDn = Strategy_UpDn.drop_duplicates(ignore_index=True)
    Strategy_UpDn = Strategy_UpDn.reset_index()
    Strategy_UpDn = Strategy_UpDn.sort_values(by = ['Condition','Sort_key'], ascending = False)
    Highest_UpDn = Highest_UpDn.drop_duplicates(ignore_index=True)
    Highest_UpDn = Highest_UpDn.reset_index()
    Highest_UpDn = Highest_UpDn.sort_values(by = ['Sort_key'], ascending = False)
    # 丟掉 Sort_key
    Highest_UpDn = Highest_UpDn.iloc[:,:-1]
    Strategy_UpDn = Strategy_UpDn.iloc[:,:-1]
    # 策略 Ticker
    Ticker_SU = [str(h) +','+ str(i) +','+ str(j) +','+ k \
                 for h,i,j,k in zip(Strategy_UpDn['Condition'],Strategy_UpDn['Name'],\
                                    Strategy_UpDn['Ticker'], Strategy_UpDn['UpDn(%)'])]
    Ticker_SU = pd.Series(Ticker_SU)
    Ticker_SU = Ticker_SU.drop_duplicates()
    # 漲幅 Ticker
    Ticker_HU = [str(i) +','+ str(j) +','+ k \
                 for i,j,k in zip(Highest_UpDn['Name'],Highest_UpDn['Ticker'], Highest_UpDn['UpDn(%)'])]
    Ticker_HU = pd.Series(Ticker_HU)
    Ticker_HU = Ticker_HU.drop_duplicates()
    # To Excel
    Strategy_UpDn.to_excel(excel_writer=EW2, sheet_name='Strategy_UpDn')
    Highest_UpDn.to_excel(excel_writer=EW2, sheet_name='Highest_UpDn')
    return Ticker_SU, Ticker_HU


# 寄信通知
def Send_Gmail(PATH_UPDN = PATH_UPDN, TD = TD):
    send_user = 'veryveryveryhandsome'   #發件人
    password = 'hobe820511'   #授權碼/密碼
    receive_users = ['hndsmhsu@gmail.com', 'f202925@gmail.com', \
                     'chien.chang91@gmail.com', 'linkhun428@gmail.com', 'hjhang777@gmail.com']  #收件人
    subject = 'HiStock_Crawler_UpDn_%s'%TD  #郵件主題
    email_text = '今日台股觀察列表'   #郵件正文
    server_address = 'smtp.gmail.com'   #伺服器地址
    
    #構造一個郵件體：正文 附件
    msg = MIMEMultipart()
    msg['Subject']=subject  #主題
    msg['From']=send_user   #發件人
    msg['To']=", ".join(receive_users) #收件人
    
    #構建正文
    part_text = MIMEText(email_text)
    msg.attach(part_text) 
    
    #構建郵件附件
    part_attach1 = MIMEApplication(open(PATH_UPDN,'rb').read())   #開啟附件
    part_attach1.add_header('Content-Disposition','attachment',filename = 'HiStock_Crawler_UpDn_%s.xlsx'%TD) #為附件命名
    msg.attach(part_attach1)   #新增附件
    
    # 傳送郵件 SMTP
    smtp= smtplib.SMTP(server_address, 587)  # 連線伺服器，SMTP_SSL是安全傳輸
    smtp.ehlo() #申請身分
    smtp.starttls() #加密文件，避免私密信息被截取
    smtp.login(send_user, password)
    smtp.sendmail(send_user, receive_users, msg.as_string())  # 傳送郵件
    return None
# 傳送訊息 with 圖片 with 貼圖
def lineNotifyMessage(token, msg):
    headers = {"Authorization": "Bearer " + token}
    payload = {'message': msg}
    # file = {'imageFile': open(picURL, 'rb')}
    r = requests.post("https://notify-api.line.me/api/notify", headers = headers, params = payload)
    if r.status_code == 200:
        return 'Send OK'
    else:
        return (r.status_code,'Send Error!')
    
# Run Module
## LOGGER 設置
CH = logging.StreamHandler()
CH.setLevel(logging.INFO)
CH.setFormatter(FORMATER)

FH = logging.FileHandler(PATH_LOG)
FH.setLevel(logging.INFO)
FH.setFormatter(FORMATER)

LOGGER.addHandler(CH)
LOGGER.addHandler(FH)
## 開始測量
start = time.time()
logging.info('is when Crawling Started')
HiStock_Web_Crawler()
logging.info('is when Crawling Ended')
## Excel 輸出
logging.info('is when Excel Exporting')
EW.save()
EW.close()
XLS = pd.ExcelFile(PATH)
logging.info('is when Excel finishe Exporting')
## 選股策略
logging.info('is when Pick_Strategy Started')
Ticker_SU, Ticker_HU = Pick_Strategy(xls = XLS)
EW2.save()
EW2.close()
logging.info('is when Pick_Strategy Ended')
logging.info('is when Line message Sending')
token = 'lBtxS0nMHzBI5FR5qzCHrWz951vgdPKiW3vHfCATGPh'
msg1 = ('\n 站上均線 & 外資買 & 漲幅不大潛力股 \n' + Ticker_SU.to_string(index=False))
msg2 = ('\n 漲幅 8% 以上飆股 \n' + Ticker_HU.to_string(index=False))
lineNotifyMessage(token, msg1)
lineNotifyMessage(token, msg2)
logging.info('is when Line message Sended')
## 寄信
logging.info('is when Email Sending')
Send_Gmail()
logging.info('is when Email Sended')
## 結束測量
end = time.time()
## 輸出結果
print("執行時間：%.3f 分" %((end - start)/60))