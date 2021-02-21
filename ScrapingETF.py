from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd

opt = webdriver.ChromeOptions()
opt.add_argument('headless')

drv = webdriver.Chrome('C:\Users\이임주\Documents\Projects\Stock\chromedriver.exe', options=opt)
drv.implicitly_wait(3)
drv.get('https://finance.naver.com/sise/etf.nhn')

bs = BeautifulSoup(drv.page_source, 'lxml')
drv.quit()
table = bs.find_all('table', class_="type_1 type_etf")
df = pd.read_html(str(table), header=0)[0]

df = df.drop(columns=['Unnamed: 9'])
df = df.dropna()
df.index = range(1, len(df) + 1)
print(df)

etf_td = bs.find_all("td", class_="ctg")
etfs = {}
for td in etf_td:
    s = str(td.a["href"].split("="))
    code = s[-1]
    etfs[td.a.text] = code
print("etfs: ", etfs)
