import os
import sys
import psutil
import sqlite3
import numpy as np
import pandas as pd
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.setting import columns_gj1, db_stg, ui_num
from utility.static import now, timedelta_sec, thread_decorator, strf_time

TUJAGMDIVIDE = 5    # 종목당 투자금 분할계수


class Strategy:
    def __init__(self, qlist):
        self.windowQ = qlist[0]
        self.workerQ = qlist[1]
        self.stgtQ = qlist[2]

        self.list_buy = []
        self.list_sell = []
        self.dict_gsjm = {}     # key: 종목코드, value: 10시이전 DataFrame, 10시이후 list
        self.dict_intg = {
            '체결강도차이1': 0.,
            '평균시간1': 0,
            '거래대금차이1': 0,
            '체결강도하한1': 0.,
            '누적거래대금하한1': 0,
            '등락율하한1': 0.,
            '등락율상한1': 0.,
            '청산수익률1': 0.,

            '체결강도차이2': 0.,
            '평균시간2': 0,
            '거래대금차이2': 0,
            '체결강도하한2': 0.,
            '누적거래대금하한2': 0,
            '등락율하한2': 0.,
            '등락율상한2': 0.,
            '청산수익률2': 0.,

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
        con = sqlite3.connect(db_stg)
        df = pd.read_sql('SELECT * FROM setting', con)
        df = df.set_index('index')
        self.dict_intg['체결강도차이1'] = df['체결강도차이1'][0]
        self.dict_intg['평균시간1'] = df['평균시간1'][0]
        self.dict_intg['거래대금차이1'] = df['거래대금차이1'][0]
        self.dict_intg['체결강도하한1'] = df['체결강도하한1'][0]
        self.dict_intg['누적거래대금하한1'] = df['누적거래대금하한1'][0]
        self.dict_intg['등락율하한1'] = df['등락율하한1'][0]
        self.dict_intg['등락율상한1'] = df['등락율상한1'][0]
        self.dict_intg['청산수익률1'] = df['청산수익률1'][0]
        self.dict_intg['체결강도차이2'] = df['체결강도차이2'][0]
        self.dict_intg['평균시간2'] = df['평균시간2'][0]
        self.dict_intg['거래대금차이2'] = df['거래대금차이2'][0]
        self.dict_intg['체결강도하한2'] = df['체결강도하한2'][0]
        self.dict_intg['누적거래대금하한2'] = df['누적거래대금하한2'][0]
        self.dict_intg['등락율하한2'] = df['등락율하한2'][0]
        self.dict_intg['등락율상한2'] = df['등락율상한2'][0]
        self.dict_intg['청산수익률2'] = df['청산수익률2'][0]
        con.close()
        while True:
            data = self.stgtQ.get()
            if len(data) == 2:
                self.UpdateList(data[0], data[1])
            elif len(data) == 14:
                self.BuyStrategy(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7],
                                 data[8], data[9], data[10], data[11], data[12], data[13])
            elif len(data) == 7:
                self.SellStrategy(data[0], data[1], data[2], data[3], data[4], data[5], data[6])

            if now() > self.dict_time['관심종목']:
                self.windowQ.put([ui_num['관심종목'], self.dict_gsjm])
                self.dict_time['관심종목'] = timedelta_sec(1)
            if now() > self.dict_time['부가정보']:
                self.UpdateInfo()
                self.dict_time['부가정보'] = timedelta_sec(2)

    def UpdateList(self, gubun, code):
        if '조건진입' in gubun:
            time = 1 if int(strf_time('%H%M%S')) <= 100000 else 2
            if code not in self.dict_gsjm.keys():
                data = np.zeros((self.dict_intg[f'평균시간{time}'] + 2, len(columns_gj1))).tolist()
                df = pd.DataFrame(data, columns=columns_gj1)
                df['체결시간'] = '090000'
                self.dict_gsjm[code] = df.copy()
            if gubun == '조건진입마지막':
                self.windowQ.put([ui_num['관심종목'] + 100, self.dict_gsjm])
        elif '관심종목리셋' in gubun:
            data = np.zeros((self.dict_intg['평균시간2'] + 2, len(columns_gj1))).tolist()
            df = pd.DataFrame(data, columns=columns_gj1)
            df['체결시간'] = '100000'
            for code in self.dict_gsjm.keys():
                self.dict_gsjm[code] = df.copy()
            self.windowQ.put([ui_num['관심종목'] + 100, self.dict_gsjm])
        elif gubun == '조건이탈':
            if code in self.dict_gsjm.keys():
                del self.dict_gsjm[code]
        elif gubun == '매수완료':
            if code in self.list_buy:
                self.list_buy.remove(code)
        elif gubun == '매도완료':
            if code in self.list_sell:
                self.list_sell.remove(code)

    def BuyStrategy(self, code, name, c, o, h, low, per, ch, dm, d, injango, vitimedown, vid5priceup, batting):
        if code not in self.dict_gsjm.keys():
            return

        time = 1 if int(strf_time('%H%M%S')) <= 100000 else 2
        hlm = round((h + low) / 2)
        hlmp = round((c / hlm - 1) * 100, 2)
        sm = int(dm - self.dict_gsjm[code]['누적거래대금'][1])
        self.dict_gsjm[code] = self.dict_gsjm[code].shift(1)
        if len(self.dict_gsjm[code]) == self.dict_intg[f'평균시간{time}'] + 2 and \
                self.dict_gsjm[code]['체결강도'][self.dict_intg[f'평균시간{time}']] != 0.:
            avg_sm = round(self.dict_gsjm[code]['거래대금'][1:self.dict_intg[f'평균시간{time}'] + 1].mean(), 2)
            avg_ch = round(self.dict_gsjm[code]['체결강도'][1:self.dict_intg[f'평균시간{time}'] + 1].mean(), 2)
            high_ch = round(self.dict_gsjm[code]['체결강도'][1:self.dict_intg[f'평균시간{time}'] + 1].max(), 2)
            self.dict_gsjm[code].at[self.dict_intg[f'평균시간{time}'] + 1] = 0., 0., avg_sm, 0, avg_ch, high_ch, d
        self.dict_gsjm[code].at[0] = per, hlmp, sm, dm, ch, 0., d

        if self.dict_gsjm[code]['체결강도'][self.dict_intg[f'평균시간{time}']] == 0:
            return
        if code in self.list_buy:
            return

        # 전략 비공개

        oc = int(batting / TUJAGMDIVIDE / c)
        if oc > 0:
            self.list_buy.append(code)
            self.workerQ.put(['매수', code, name, c, oc])

    def SellStrategy(self, code, name, per, sp, jc, ch, c):
        if code in self.list_sell:
            return

        oc = 0
        if per >= 29:
            oc = jc

        time = 1 if int(strf_time('%H%M%S')) <= 100000 else 2

        # 전략 비공개

        if oc > 0:
            self.list_sell.append(code)
            self.workerQ.put(['매도', code, name, c, oc])

    @thread_decorator
    def UpdateInfo(self):
        info = [4, self.dict_intg['메모리'], self.dict_intg['스레드'], self.dict_intg['시피유']]
        self.windowQ.put(info)
        self.UpdateSysinfo()

    def UpdateSysinfo(self):
        p = psutil.Process(os.getpid())
        self.dict_intg['메모리'] = round(p.memory_info()[0] / 2 ** 20.86, 2)
        self.dict_intg['스레드'] = p.num_threads()
        self.dict_intg['시피유'] = round(p.cpu_percent(interval=2) / 2, 2)
