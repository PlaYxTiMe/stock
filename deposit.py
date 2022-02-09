import pandas as pd
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import sqlite3




url='https://goodinfo.tw/tw/StockDividendScheduleList.asp?MARKET_CAT=%E5%85%A8%E9%83%A8&YEAR=%E5%8D%B3%E5%B0%87%E9%99%A4%E6%AC%8A%E6%81%AF&INDUSTRY_CAT=%E5%85%A8%E9%83%A8'
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
    
    chrom=get_chrom(url)
    count=0
    page=4
    datas=[]
    while count<1:
        try:
            xpath=f'/html/body/table[2]/tbody/tr/td[3]/table/tbody/tr/td/div/form/nobr[3]/select/option[{page}]'
            chrom.find_element(By.XPATH,xpath).click()
            time.sleep(0.5)
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
    
    conn=sqlite3.connect('stock.db')
    cursor=conn.cursor
    conn.execute(sqlstr)
    conn.commit()
    conn.close()
    
    df3=pd.DataFrame(datas,columns=['編號','股名','除息日','現金股利','除息參考價','現金殖利率'])
    with sqlite3.connect('stock.db') as conn:                   
        return df3.to_sql('deposit',conn,if_exists='append',index=False)
        conn.close()
    chrom.quit()
deposit()   