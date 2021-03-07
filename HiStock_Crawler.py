# -*- coding: utf-8 -*-
"""
Created on Mon Mar  1 15:12:06 2021

@author: BennyXu
"""

# Import Package
import re
import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib

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
PATH = os.path.join(os.getcwd(),'HiStock_Crawler_%s.xlsx'%TD)
PATH_UPDN = os.path.join(os.getcwd(),'HiStock_Crawler_UpDn_%s.xlsx'%TD)
EW = pd.ExcelWriter (PATH)
EW2 = pd.ExcelWriter (PATH_UPDN)


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
            print("Connection refused by the server..")
            print("Let me sleep for 5 seconds")
            print("ZZzzzz...")
            time.sleep(5)
            print("Was a nice sleep, now let me continue...")
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
                    "營收日期" : [], "營收天數" : [], "漲跌幅(點)" : [], "漲跌幅(%)" : []}
            url2 = URL + row["href"]
            attempt1 = 0
            while attempt1 < 3:
                try:
                    re2 = requests.get(url2)
                    break
                except:
                    attempt1 += 1
                    print("Connection refused by the server..")
                    print("Let me sleep for 5 seconds")
                    print("ZZzzzz...")
                    time.sleep(5)
                    print("Was a nice sleep, now let me continue...")
                    continue
            soup2 = BeautifulSoup(re2.content, "html.parser")
            print('----------------------Now crawling %s 類股----------------------' %str(row['title']))
            ## 類股中的個股表格
            table2 = soup2.find_all("a", {"class":"link"})      
            ## 個股迴圈
            for col in table2:
                if '/twclass/' not in str(col):
                    ## 個股中文,代號
                    DICT['股票'].append(col['title'])
                    DICT['代號'].append(col['href'].replace('/stock/', ''))
                    
                    url3 = URL + col["href"]
                    attempt2 = 0
                    while attempt2 < 3:
                        try:
                            re3 = requests.get(url3)
                            break
                        except:
                            attempt2 += 1
                            print("Connection refused by the server..")
                            print("Let me sleep for 5 seconds")
                            print("ZZzzzz...")
                            time.sleep(5)
                            print("Was a nice sleep, now let me continue...")
                            continue
                    soup3 = BeautifulSoup(re3.content, "html.parser")
                    data = soup3.find_all('script', {"type" :"text/javascript"})
                    data2 = soup3.find_all('table',{"class" : "tb-stock tbChip"})
                    data3 = soup3.find('ul', {'class' : 'priceinfo mt10'}).find_all('span', {'id' : 'Price1_lbTChange'})
                    data33 = soup3.find('ul', {'class' : 'priceinfo mt10'}).find_all('span', {'id' : 'Price1_lbTPercent'})
                    ## 取 EPS、現金殖利率
                    list_3 = [content.text for content in soup3.find('table', {'class' : "tb-stock tbBasic"})\
                              .find_all('td')][4:8]
                    print('----------------Now crawling %s 個股----------------' %str(col['title']))
                    list_1 = []
                    list_2 = []
                    for string in re.split(r";|}|{",str(data)):
                        if ('candlestick' in string) and ('data:' in string):
                            ## 取最新開、高、低、收
                            list_1 = [float(i) for i in re.findall(r"\d+\.?\d*",string)[-4:]]
                            ## 取最新交易量(張)
                        if ('成交量(張)' in string) and ('data:' in string):
                            list_2 = [float(i) for i in re.findall(r"\d+\.?\d*",string)[-4:-3]]
                    ## 爬不到或已下市股票儲存'Error'
                    if len(list_1) != 4:
                        list_1 = ['Error data'] * 4
                        print('list_1 Length :',len(list_1),'Im Error')
                    if len(list_2) != 1:
                        list_2 = ['Error data']
                        print('list_2 Length :',len(list_2),'Im Error')
                    if len(list_3) != 4:
                        list_3 = ['Error data'] * 4
                        print('list_3 Length :',len(list_3),'Im Error')
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
                    Revenue = []
                    Short = []
                    Margin = []
                    Trust = []
                    Dealer = []
                    Foreign = []
                    UpDn = []
                    UpDn_PC = []
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
                    ## 檢查資料長度，有錯誤者儲存'Error'
                    if len(Revenue) != 2:
                        Revenue = ['Error data'] * 2
                        print('Revenue Length :',len(Revenue),'Im Error')
                    if len(Short) != 3:
                        Short = ['Error data'] * 3
                        print('Short Length :',len(Short),'Im Error')
                    if len(Margin) != 3:
                        Margin = ['Error data'] * 3
                        print('Margin Length :',len(Margin),'Im Error')
                    if len(Trust) != 3:
                        Trust = ['Error data'] * 3
                        print('Trust Length :',len(Trust),'Im Error')
                    if len(Dealer) != 3:
                        Dealer = ['Error data'] * 3
                        print('Dealer Length :',len(Dealer),'Im Error')
                    if len(Foreign) != 3:
                        Foreign = ['Error data'] * 3
                        print('Foreign Length :',len(Foreign),'Im Error')
                    if len(UpDn) != 1:
                        UpDn = ['Error data']
                        print('UpDn Length :',len(UpDn),'Im Error')
                    if len(UpDn_PC) != 1:
                        UpDn_PC = ['Error data']
                        print('UpDn_PC Length :',len(UpDn_PC),'Im Error')
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
        else:
            continue
        ## 類股 DataFrame
        Result = pd.DataFrame.from_dict(DICT)
        ## Sort
        Result = Result.sort_values(by = ['漲跌幅(%)', '交易量(張)'], ascending = False)
        Result.to_excel(excel_writer=EW, sheet_name=row['title'],index=None) 
    return Result

# 選股策略
def Highest_UpDn(xls):
    Dict = {'Cat' : [], 'Name' : [], 'Ticker' : [],\
            'Open' : [], 'High' : [],'Close' : [], 'Low' : [],\
            'UpDn(%)' : [], 'UpDn(pt)' : [],'Volumn' : []}
    Dict1 = {'Name' : [], 'Ticker' : [],\
            'Open' : [], 'High' : [],'Close' : [], 'Low' : [],\
            'UpDn(%)' : [], 'UpDn(pt)' : [],'Volumn' : []}
    for num, j in enumerate(xls.sheet_names):
        if j != '金融':
            df = pd.read_excel(xls, j)
            df['交易量(張)'] = df['交易量(張)'].map(lambda x: int(x) if type(x) is int else 0)
            df['漲跌幅'] = [abs(float(re.findall(r"\d+\.?\d*",df['漲跌幅(%)'][i])[0])) \
                             for i in df['漲跌幅(%)'].index]
            condition1 = df['漲跌幅'] == df['漲跌幅'].max()
            condition2 = df['交易量(張)'] == df['交易量(張)'].max()
            if not df.empty:
                for con in [condition1, condition2]:
                    Dict['Cat'].append(str(j))
                    Dict['Name'].append(df.loc[con]['股票'].values[0])
                    Dict['Ticker'].append(df.loc[con]['代號'].values[0])
                    Dict['Open'].append(df.loc[con]['開盤價'].values[0])
                    Dict['High'].append(df.loc[con]['最高價'].values[0])
                    Dict['Close'].append(df.loc[con]['收盤價'].values[0])
                    Dict['Low'].append(df.loc[con]['最低價'].values[0])
                    Dict['UpDn(%)'].append(df.loc[con]['漲跌幅(%)'].values[0])
                    Dict['UpDn(pt)'].append(df.loc[con]['漲跌幅(點)'].values[0])
                    Dict['Volumn'].append(df.loc[con]['交易量(張)'].values[0])
                for i in df.index:
                    if df.loc[i,'漲跌幅'] >= 5:
                        Dict1['Name'].append(df.loc[i,'股票'])
                        Dict1['Ticker'].append(df.loc[i,'代號'])
                        Dict1['Open'].append(df.loc[i,'開盤價'])
                        Dict1['High'].append(df.loc[i,'最高價'])
                        Dict1['Close'].append(df.loc[i,'收盤價'])
                        Dict1['Low'].append(df.loc[i,'最低價'])
                        Dict1['UpDn(%)'].append(df.loc[i,'漲跌幅(%)'])
                        Dict1['UpDn(pt)'].append(df.loc[i,'漲跌幅(點)'])
                        Dict1['Volumn'].append(df.loc[i,'交易量(張)'])
            else: continue
    Cat_Highest_UpDn = pd.DataFrame.from_dict(Dict)
    Cat_Highest_UpDn.to_excel(excel_writer=EW2, sheet_name='Cat_Highest_UpDn')
    Highest_UpDn = pd.DataFrame.from_dict(Dict1)
    Highest_UpDn.to_excel(excel_writer=EW2, sheet_name='Highest_UpDn')
    return Cat_Highest_UpDn

# 寄信通知
def Send_Gmail(PATH_UPDN = PATH_UPDN, TD = TD):
    send_user = '*************'   #發件人
    password = '**************'   #授權碼/密碼
    receive_users = ['*****@gmail.com', '******@gmail.com']  #收件人
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
    print('--------------------郵件傳送成功！--------------------')
    return None

# Run Module
## 開始測量
start = time.time()
print('----------------------------------------Start Crawling--------------------------------------')
HiStock_Web_Crawler()
## Excel 輸出
print('-------------------------------------Export Excel-----------------------------------')
EW.save()
EW.close()
XLS = pd.ExcelFile(PATH)
## 選股策略
print('----------------------------------Export Highest_UpDn--------------------------------')
Highest_UpDn(xls = XLS)
EW2.save()
EW2.close()
print('----------------------------------Sending Email--------------------------------')
## 寄信
Send_Gmail()
## 結束測量
end = time.time()
## 輸出結果
print('----------------------------------------End Crawling--------------------------------------')
print("執行時間：%.3f 分" %((end - start)/60))