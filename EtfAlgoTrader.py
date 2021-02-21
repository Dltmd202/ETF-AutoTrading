import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
from slacker import Slacker
import time, calendar
from bs4 import BeautifulSoup
from urllib.request import urlopen
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
token = "xoxb-1769457119238-1776238220130-udiBOYwElxAokfoBrMoefJM9"
slack = Slacker(token)


def dbgout(message):
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    slack.chat.post_message('#etf-algo-trading', strbuf)


def printlog(message, *args):
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)



cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')


def check_creon_system():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        dbgout("Login")
        return False

    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False

    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True


def get_current_price(code):
    cpStock.SetInputValue(0, code)
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)
    item['ask'] = cpStock.GetHeaderValue(16)
    item['bid'] = cpStock.GetHeaderValue(17)
    return item['cur_price'], item['ask'], item['bid']


def get_ohlc(code, qty):

    cpOhlc.SetInputValue(0, code)
    cpOhlc.SetInputValue(1, ord('2'))
    cpOhlc.SetInputValue(4, qty)
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5])
    cpOhlc.SetInputValue(6, ord('D'))
    cpOhlc.SetInputValue(9, ord('1'))
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count):
        index.append(cpOhlc.GetDataValue(0, i))
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
                     cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)])
    df = pd.DataFrame(rows, columns=columns, index=index)
    return df


def get_stock_balance(code):

    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]
    accFlag = cpTradeUtil.GoodsList(acc, 1)
    cpBalance.SetInputValue(0, acc)
    cpBalance.SetInputValue(1, accFlag[0])
    cpBalance.SetInputValue(2, 50)
    cpBalance.BlockRequest()
    if code == 'ALL':
        dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('종목수: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)
        stock_name = cpBalance.GetDataValue(0, i)
        stock_qty = cpBalance.GetDataValue(15, i)
        if code == 'ALL':
            dbgout(str(i + 1) + ' ' + stock_code + '(' + stock_name + ')'
                   + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name,
                           'qty': stock_qty})
        if stock_code == code:
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0


def get_current_cash():
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]
    accFlag = cpTradeUtil.GoodsList(acc, 1)
    cpCash.SetInputValue(0, acc)
    cpCash.SetInputValue(1, accFlag[0])
    cpCash.BlockRequest()
    return cpCash.GetHeaderValue(9)


def get_target_price(code):
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        target_price = today_open + (lastday_high - lastday_low) * 0.5
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None


def get_movingaverage(code, window):
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None


def buy_etf(code):
    try:
        global bought_list
        if code in bought_list:
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code)
        target_price = get_target_price(code)
        ma5_price = get_movingaverage(code, 5)
        ma10_price = get_movingaverage(code, 10)
        buy_qty = 0
        if ask_price > 0:
            buy_qty = buy_amount // ask_price
        stock_name, stock_qty = get_stock_balance(code)
        printlog('bought_list:', bought_list, 'len(bought_list):',
            len(bought_list), 'target_buy_count:', target_buy_count)
        if current_price > target_price and current_price > ma5_price \
                and current_price > ma10_price:
            printlog(stock_name + '(' + str(code) + ') ' + str(buy_qty) +
                     'EA : ' + str(current_price) + ' meets the buy condition!`')
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]
            accFlag = cpTradeUtil.GoodsList(acc, 1)
            cpOrder.SetInputValue(0, "2")
            cpOrder.SetInputValue(1, acc)
            cpOrder.SetInputValue(2, accFlag[0])
            cpOrder.SetInputValue(3, code)
            cpOrder.SetInputValue(4, buy_qty)
            cpOrder.SetInputValue(7, "2")
            cpOrder.SetInputValue(8, "12")

            ret = cpOrder.BlockRequest()
            printlog('최유리 FoK 매수 ->', stock_name, code, buy_qty, '->', ret)
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time / 1000)
                time.sleep(remain_time / 1000)
                return False
            time.sleep(2)
            printlog('현금주문 가능금액 :', buy_amount)
            stock_name, bought_qty = get_stock_balance(code)
            printlog('get_stock_balance :', stock_name, stock_qty)
            if bought_qty > 0:
                bought_list.append(code)
                dbgout("`buy_etf(" + str(stock_name) + ' : ' + str(code) +
                       ") -> " + str(bought_qty) + "EA bought!" + "`")
    except Exception as ex:
        dbgout("`buy_etf(" + str(code) + ") -> exception! " + str(ex) + "`")


def sell_all():
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]
        accFlag = cpTradeUtil.GoodsList(acc, 1)
        while True:
            stocks = get_stock_balance('ALL')
            total_qty = 0
            for s in stocks:
                total_qty += s['qty']
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:
                    cpOrder.SetInputValue(0, "1")
                    cpOrder.SetInputValue(1, acc)
                    cpOrder.SetInputValue(2, accFlag[0])
                    cpOrder.SetInputValue(3, s['code'])
                    cpOrder.SetInputValue(4, s['qty'])
                    cpOrder.SetInputValue(7, "1")
                    cpOrder.SetInputValue(8, "12")
                    ret = cpOrder.BlockRequest()
                    printlog('최유리 IOC 매도', s['code'], s['name'], s['qty'],
                             '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('주의: 연속 주문 제한, 대기시간:', remain_time / 1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))


if __name__ == '__main__':
    try:
        symbol_list = ['A122630', 'A252670', 'A233740', 'A250780', 'A225130',
                       'A280940', 'A261220', 'A217770', 'A295000', 'A176950']
        bought_list = []
        target_buy_count = 5
        buy_percent = 0.19
        printlog('check_creon_system() :', check_creon_system())
        stocks = get_stock_balance('ALL')
        total_cash = int(get_current_cash())
        buy_amount = total_cash * buy_percent
        printlog('100% 증거금 주문 가능 금액 :', total_cash)
        printlog('종목별 주문 비율 :', buy_percent)
        printlog('종목별 주문 금액 :', buy_amount)
        printlog('시작 시간 :', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False;

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)
            if t_9 < t_now < t_start and soldout == False:
                soldout = True
                sell_all()
            if t_start < t_now < t_sell:
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 5:
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:
                if sell_all() == True:
                    dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')