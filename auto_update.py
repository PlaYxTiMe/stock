import pandas as pd
import time
from datetime import datetime,timedelta
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas_datareader as pdr
import sqlite3
import threading
import sched


def get_chrom(url,driver=r'C:\webdriver\chromedriver'):
    try:
        options=webdriver.ChromeOptions()
        options.add_argument('--headless')
        chrom=webdriver.Chrome(driver,options=options)
        chrom.implicitly_wait(1)
        chrom.get(url)
        return chrom
    except Exception as e:
        return None

def new_stock_price(column,stock_new):
    try:
        with sqlite3.connect('stock.db') as conn:             
            stock=pd.read_sql_query(f'select * from "{column}"',con=conn)
        stock.index=pd.to_datetime(stock['date'])
    except Exception as e:
        print(e)
    try:
        if stock.index[-1]!=stock_new:
            try:
                try:
                    df2=pdr.DataReader(f'{column}.tw','yahoo',stock.index[-1]+timedelta(1))
                except Exception as e:
                    df2=pdr.DataReader(f'{column}.two','yahoo',stock.index[-1]+timedelta(1))
                df2['date']=df2.index.astype(str)
                df3=df2[['date','High','Low','Open','Close','Volume']]
                with sqlite3.connect('stock.db') as conn:                   
                    df3.to_sql(f'{column}',conn,if_exists='append',index=False)
            except Exception as e:
                print(e)
    except Exception as e:
        print(e)
    print(f'{column}<<Done!') 
    
def deposit():
    sqlstr='''
    create table if not exists "deposit" (
        id integer primary key autoincrement UNIQUE,
        "編號" text,
        "股名" text,
        "除息日" text,
        "現金股利" text,
        "除息參考價" text,
        "現金殖利率" REAL
    );
    '''
    url='https://goodinfo.tw/tw/StockDividendScheduleList.asp?MARKET_CAT=%E5%85%A8%E9%83%A8&YEAR=%E5%8D%B3%E5%B0%87%E9%99%A4%E6%AC%8A%E6%81%AF&INDUSTRY_CAT=%E5%85%A8%E9%83%A8'
    chrom=get_chrom(url)
    count=0
    page=4
    datas=[]
    while count<1:
        try:
            xpath=f'/html/body/table[2]/tbody/tr/td[3]/table/tbody/tr/td/div/form/nobr[3]/select/option[{page}]'
            chrom.find_element(By.XPATH,xpath).click()
            soup=BeautifulSoup(chrom.page_source,'lxml')
            test=soup.find('div',style="position:relative;width:811px;").find('tbody')
            for i in test.find_all('tr',align="center"):
                day=i.find_all('td')[4].text
                money=i.find('td',align="right").text
                price=i.find_all('td')[5].text
                if price !='':
                    yield_=round((eval(money)/eval(price))*100,2)
                else:
                    yield_=0
                datas.append([i.find('a',class_="link_black").text,i.find_all('td')[2].text,day,money,price,yield_])
            page+=1
        except Exception as e:
            count+=1
    df=pd.DataFrame(datas,columns=['編號','股名','除息日','現金股利','除息參考價','現金殖利率'])
    conn=sqlite3.connect('stock.db')
    cursor=conn.cursor()
    cursor.execute(sqlstr)
    conn.commit()
    with sqlite3.connect('stock.db') as conn:                   
        df.to_sql('deposit',conn,if_exists='replace',index=False)
    chrom.quit()
    
def auto_update():
    today_year=datetime.today().strftime('%Y-%m-%d')[0:4]   ##今年
    today_month=int(datetime.today().strftime('%Y-%m-%d')[5:7])  ##今月
    stock_new=pdr.DataReader('2356.tw','yahoo',f'{today_year}-{today_month}').loc[:datetime.today().strftime('%Y-%m-%d')].index[-1]
    with sqlite3.connect('stock.db') as conn:
        df=pd.read_sql_query('select * from "deposit"',con=conn)
    for value in df['除息日'].values:
        if '即將除息' in value:
            df=df.replace(value,value[0:-4])
        if '今日除息' in value:
            df=df.replace(value,value[0:-4])
    df.index=pd.to_datetime(df['除息日'],yearfirst=True)
    threads=[]
    count=0
    for column in set(df.loc['2021']['編號']):
        threads.append(threading.Thread(target=new_stock_price,args=(column,stock_new)))
        count+=1
        if count%35==0:
            for thread in threads:
                thread.start()
                time.sleep(0.4)  ##(30)0.4=763秒  (30)0.3==
            for thread in threads:
                thread.join()
            threads=[]
    for thread in threads:
        thread.start()
        time.sleep(0.4)
    for thread in threads:
        thread.join()    
    deposit()
if __name__=='__main__':
    s=sched.scheduler(time.time,time.sleep)
    #end_time=None
    s.enter(0,0,auto_update,)
    s.run()
            






