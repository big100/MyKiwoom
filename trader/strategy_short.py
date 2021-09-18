import os
import sys
import psutil
import sqlite3
import pandas as pd
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.setting import db_stg, ui_num
from utility.static import now, timedelta_sec, thread_decorator

TUJAGMDIVIDE = 5       # 종목당 투자금 분할계수


class StrategyShort:
    def __init__(self, qlist):
        self.windowQ = qlist[0]
        self.workerQ = qlist[1]
        self.stgsQ = qlist[3]

        self.list_buy = []
        self.list_sell = []
        self.df = None
        self.dict_intg = {
            '스레드': 0,
            '시피유': 0.,
            '메모리': 0.
        }
        self.dict_time = {
            '관심종목': now(),
            '부가정보': now()
        }

        self.Start()

    def Start(self):
        while True:
            data = self.stgsQ.get()
            if data == '데이터베이스로딩':
                self.DatabaseLoad()
            elif len(data) == 2:
                self.UpdateList(data[0], data[1])
            elif len(data) == 10:
                self.BuyStrategy(data[0], data[1], data[2], data[3], data[4], data[5],
                                 data[6], data[7], data[8], data[9])
            elif len(data) == 5:
                self.SellStrategy(data[0], data[1], data[2], data[3], data[4])

            if now() > self.dict_time['관심종목']:
                self.df.sort_values(by=['변동성'], ascending=False, inplace=True)
                self.windowQ.put([ui_num['short'], self.df])
                self.dict_time['관심종목'] = timedelta_sec(1)
            if now() > self.dict_time['부가정보']:
                self.UpdateInfo()
                self.dict_time['부가정보'] = timedelta_sec(2)

    def DatabaseLoad(self):
        con = sqlite3.connect(db_stg)
        df = pd.read_sql('SELECT * FROM short', con)
        con.close()
        df = df.set_index('index')
        df['등락율'] = 0
        df['현재가'] = 0
        df['시가'] = 0
        df['고가'] = 0
        df['저가'] = 0
        self.df = df[['등락율', '현재가', '시가', '고가', '저가', '변동성']].copy()
        self.df.sort_values(by=['변동성'], ascending=False, inplace=True)
        self.windowQ.put([ui_num['short'] + 100, self.df])

    def UpdateList(self, gubun, code):
        if gubun == '매수완료':
            if code in self.list_buy:
                self.list_buy.remove(code)
        elif gubun == '매도완료':
            if code in self.list_sell:
                self.list_sell.remove(code)

    def BuyStrategy(self, code, per, c, o, h, low, dict_name, intrade, injango, batting):
        prec = self.df['현재가'][code]
        self.df.at[code, ['등락율', '현재가', '시가', '고가', '저가']] = per, c, o, h, low

        if code in self.list_buy:
            return
        if intrade:
            return
        if injango:
            return
        if not (c >= o + self.df['변동성'][code] > prec != 0):
            return

        oc = int(batting / TUJAGMDIVIDE / c)
        if oc > 0:
            name = dict_name[code]
            self.list_buy.append(code)
            self.workerQ.put(['단기매수', code, name, c, oc])

    def SellStrategy(self, code, name, jc, c, o):
        if code in self.list_sell:
            return

        oc = 0
        if c >= o + self.df['변동성'][code] or c <= o - self.df['변동성'][code]:
            oc = jc
        if oc > 0:
            self.list_sell.append(code)
            self.workerQ.put(['단기매도', code, name, c, jc])

    @thread_decorator
    def UpdateInfo(self):
        info = [5, self.dict_intg['메모리'], self.dict_intg['스레드'], self.dict_intg['시피유']]
        self.windowQ.put(info)
        self.UpdateSysinfo()

    def UpdateSysinfo(self):
        p = psutil.Process(os.getpid())
        self.dict_intg['메모리'] = round(p.memory_info()[0] / 2 ** 20.86, 2)
        self.dict_intg['스레드'] = p.num_threads()
        self.dict_intg['시피유'] = round(p.cpu_percent(interval=2) / 2, 2)
