# -*- coding: utf-8 -*-
"""
Created on Wed Nov 17 20:16:23 2021

@author: USER
"""

import pandas as pd
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas_datareader as pdr
import sqlite3
import threading


url='https://goodinfo.tw/StockInfo/index.asp'
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
    get_chrom(url)
    xpath='/html/body/table[2]/tbody/tr/td[1]/div[1]/table/tbody/tr/td[1]/table/tbody/tr[10]/td/nobr/a'
    
    element=chrom.find_element(By.XPATH,xpath)
    element.click()
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


def get_chrom(url,driver=r'C:\webdriver\chromedriver'):
    try:
        options=webdriver.ChromeOptions()
        options.add_argument('--headless')
        chrom=webdriver.Chrome(driver,options=options)
        chrom.implicitly_wait(3)
        chrom.get(url)
        return chrom
    except Exception as e:
        return None
    
    
xpath='/html/body/table[2]/tbody/tr/td[1]/div[1]/table/tbody/tr/td[1]/table/tbody/tr[10]/td/nobr/a'
chrom=get_chrom(url)
chrom.find_element(By.XPATH,xpath).click()
xpath3='/html/body/table[2]/tbody/tr/td[3]/table/tbody/tr/td/div/form/nobr[3]/select/option[4]'
chrom.find_element(By.XPATH,xpath3).click()
soup=BeautifulSoup(chrom.page_source,'lxml')
test=soup.find('div',style="position:relative;width:811px;").find('tbody')
datas=[]
for i in test.find_all('tr',align="center"):
    #print(i.find('a',class_="link_black").text,i.find_all('td')[2].text,i.find_all('td')[4].text,i.find('td',align="right").text)
    datas.append([i.find('a',class_="link_black").text,i.find_all('td')[2].text])
df=pd.DataFrame(datas,columns=['編號','股名'])

def test(data):
    sqlstr=f'''
    create table if not exists "{data}" (
        id integer primary key autoincrement UNIQUE,
        date text,
        High integer,
        Low integer,
        Open integer,
        Close integer,
        Volume integer
    );
    '''
    try:
        conn=sqlite3.connect('stock.db')
        cursor=conn.cursor()
        cursor.execute(sqlstr)
        conn.commit()
        try:
            df=pdr.DataReader(f'{data}.tw','yahoo','2008')
        except Exception as e:
            df=pdr.DataReader(f'{data}.two','yahoo','2008')
        df['date']=df.index.astype(str)
        df1=df[['date','High','Low','Open','Close','Volume']]
        with sqlite3.connect('stock.db') as conn:                   
            df1.to_sql(f'{data}',conn,if_exists='replace',index=False)
    except Exception as e:
        print(e)

threads=[]
count=0
for data in set(df['編號']):
    threads.append(threading.Thread(target=test,args=(data,)))
    print(f'{data}獲取成功')
    count+=1
    if count%25==0:    
        for thread in threads:
            thread.start()
            time.sleep(0.5)     
        for thread in threads:
            thread.join()
        time.sleep(0.5)
        print('Done!')
        threads=[]
for thread in threads:
    thread.start()
    time.sleep(0.5)
for thread in threads:
    thread.join()
deposit()
print('Done!')   
