# HiStock_Crawler
Web crawler of Histock(Taiwan Stock) using Python

## Usage : 
### 1. 抓取 https://histock.tw/twstock 台股各類股之資訊，寫入'HiStock_Crawler_%s.xlsx'%今天日期 
![image](https://user-images.githubusercontent.com/49243751/110233731-03f6e700-7f61-11eb-85a3-9be862d14046.png)
#### 1_1 欄位 :
            -股票  
            -代號  
            -交易量(張)  
            -開盤價  
            -最高價  
            -最低價  
            -收盤價  
            -EPS  
            -本益比  
            -股價淨值比  
            -現金殖利率 
            -外資日期  
            -外資天數  
            -外資張數  
            -投信日期  
            -投信天數 
            -投信張數  
            -自營日期  
            -自營天數  
            -自營張數 
            -融資日期 
            -融資天數  
            -融資張數  
            -融券日期  
            -融券天數  
            -融券張數 
            -營收日期
            -營收天數
            -漲跌幅(點)
            -漲跌幅(%)
            
### 2. 將各類股中漲(跌)幅、交易量第一名拉出來，寫入 'HiStock_Crawler_UpDn_%s.xlsx'%今天日期 
![image](https://user-images.githubusercontent.com/49243751/110233871-f42bd280-7f61-11eb-897b-b6286be110fb.png)

### 3. 寄信將選股結果寄至gmail
![image](https://user-images.githubusercontent.com/49243751/110233981-d317b180-7f62-11eb-8941-c4b9a6278931.png)

## Step :
### 1. 下載HiStock_Crawler.py, HiStock_Crawler.bat

### 2. 改bat檔路徑

![image](https://user-images.githubusercontent.com/49243751/110234145-d52e4000-7f63-11eb-91c6-cd702397c8f1.png)

### 3. 改Email位置(寄件人、收件人)
![image](https://user-images.githubusercontent.com/49243751/110234261-656c8500-7f64-11eb-9ce9-a3f88eb2c4b7.png)

### 4. 把寄件人gmail之安全性調低
![image](https://user-images.githubusercontent.com/49243751/110234273-82a15380-7f64-11eb-909c-12563331dd62.png)

### 5. 執行 HiStock_Crawler.bat
![image](https://user-images.githubusercontent.com/49243751/110234322-d318b100-7f64-11eb-8e61-b1504fbf0d85.png)

## V2 Change Log : 
### 1. 爬蟲加入爬取技術指標(MA、KD、RSI、MACD、OSC) 
### 2. 加入站上均線、三大法人、成值選股策略
### 3. 加入Line Message 提醒 
### 4. Excel存入指定資料夾
### 5. 加入輸出 LOG 機制
