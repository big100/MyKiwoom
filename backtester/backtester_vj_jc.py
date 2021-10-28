import os
import sys
import sqlite3
import datetime
import pandas as pd
from matplotlib import pyplot as plt
from multiprocessing import Process, Queue
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.setting import DB_TICK, DB_BACKTEST
from utility.static import strf_time, strp_time, timedelta_sec, timedelta_day

BETTING = 20000000     # 종목당 배팅금액
TESTPERIOD = 14        # 백테스팅 기간(14일 경우 과거 2주간의 데이터를 백테스팅한다)
TOTALTIME = 36000      # 백테스팅 기간 동안 9시부터 10시까지의 시간 총합, 단위 초
START_TIME = 90000
END_TIME = 100000
MULTI_COUNT = 6


class BackTesterVj:
    def __init__(self, q_, code_list_, num_, df2_):
        self.q = q_
        self.code_list = code_list_
        self.df_mt = df2_

        self.gap_ch = num_[0]
        self.avg_time = num_[1]
        self.gap_sm = num_[2]
        self.ch_low = num_[3]
        self.dm_low = num_[4]
        self.per_low = num_[5]
        self.per_high = num_[6]
        self.sell_ratio = num_[7]

        self.code = None
        self.df = None

        self.totalcount = 0
        self.totalcount_p = 0
        self.totalcount_m = 0
        self.totalholdday = 0
        self.totaleyun = 0
        self.totalper = 0.

        self.hold = False
        self.buytime = None
        self.buycount = 0
        self.buyprice = 0
        self.sellprice = 0
        self.highper = 0
        self.index = 0
        self.indexb = 0
        self.indexn = 0
        self.ccond = 0

        self.Start()

    def Start(self):
        conn = sqlite3.connect(DB_TICK)
        tcount = len(self.code_list)
        int_daylimit = int(strf_time('%Y%m%d', timedelta_day(-TESTPERIOD)))
        for k, code in enumerate(self.code_list):
            self.code = code
            self.df = pd.read_sql(f"SELECT * FROM '{code}'", conn).set_index('index')
            self.df['고저평균대비등락율'] = (self.df['현재가'] / ((self.df['고가'] + self.df['저가']) / 2) - 1) * 100
            self.df['고저평균대비등락율'] = self.df['고저평균대비등락율'].round(2)
            self.df['직전체결강도'] = self.df['체결강도'].shift(1)
            self.df['직전당일거래대금'] = self.df['당일거래대금'].shift(1)
            self.df = self.df.fillna(0)
            self.df['초당거래대금'] = self.df['당일거래대금'] - self.df['직전당일거래대금']
            self.df['직전초당거래대금'] = self.df['초당거래대금'].shift(1)
            self.df = self.df.fillna(0)
            self.df['초당거래대금평균'] = self.df['직전초당거래대금'].rolling(window=self.avg_time).mean()
            self.df['체결강도평균'] = self.df['직전체결강도'].rolling(window=self.avg_time).mean()
            self.df['최고체결강도'] = self.df['직전체결강도'].rolling(window=self.avg_time).max()
            self.df = self.df.fillna(0)

            self.totalcount = 0
            self.totalcount_p = 0
            self.totalcount_m = 0
            self.totalholdday = 0
            self.totaleyun = 0
            self.totalper = 0.

            self.hold = False
            self.buytime = None
            self.buycount = 0
            self.buyprice = 0
            self.sellprice = 0
            self.highper = 0
            self.index = 0
            self.indexb = 0
            self.indexn = 0
            self.ccond = 0

            lasth = len(self.df) - 1
            for h, index in enumerate(self.df.index):
                if h != 0 and index[:8] != self.df.index[h - 1][:8]:
                    self.ccond = 0
                if int(index[:8]) <= int_daylimit or \
                        (not self.hold and (END_TIME <= int(index[8:]) or int(index[8:]) < START_TIME)):
                    continue
                self.index = index
                self.indexn = h
                if not self.hold and START_TIME < int(index[8:]) < END_TIME and self.BuyTerm():
                    self.Buy()
                elif self.hold and START_TIME < int(index[8:]) < END_TIME and self.SellTerm():
                    self.Sell()
                elif self.hold and (h == lasth or int(index[8:]) >= END_TIME > int(self.df.index[h - 1][8:])):
                    self.Sell()
            self.Report(k + 1, tcount)
        conn.close()

    def BuyTerm(self):
        if type(self.df['현재가'][self.index]) == pd.Series:
            return False
        try:
            if self.code not in self.df_mt['거래대금순위'][self.index]:
                self.ccond = 0
            else:
                self.ccond += 1
        except KeyError:
            return False
        if self.ccond < self.avg_time + 1:
            return False

        # 전략 비공개

        return True

    def Buy(self):
        매도호가5 = self.df['매도호가5'][self.index]
        매도호가4 = self.df['매도호가4'][self.index]
        매도호가3 = self.df['매도호가3'][self.index]
        매도호가2 = self.df['매도호가2'][self.index]
        매도호가1 = self.df['매도호가1'][self.index]
        매도잔량5 = self.df['매도잔량5'][self.index]
        매도잔량4 = self.df['매도잔량4'][self.index]
        매도잔량3 = self.df['매도잔량3'][self.index]
        매도잔량2 = self.df['매도잔량2'][self.index]
        매도잔량1 = self.df['매도잔량1'][self.index]
        현재가 = self.df['현재가'][self.index]
        매수수량 = int(BETTING / 현재가)
        if 매수수량 > 0:
            남은수량 = 매수수량
            직전남은수량 = 매수수량
            매수금액 = 0
            호가정보 = {매도호가1: 매도잔량1}
            for 매도호가, 매도잔량 in 호가정보.items():
                남은수량 -= 매도잔량
                if 남은수량 <= 0:
                    매수금액 += 매도호가 * 직전남은수량
                    break
                else:
                    매수금액 += 매도호가 * 매도잔량
                    직전남은수량 = 남은수량
            if 남은수량 <= 0:
                예상체결가 = round(매수금액 / 매수수량, 2)
                self.buyprice = 예상체결가
                self.buycount = 매수수량
                self.hold = True
                self.indexb = self.indexn
                self.buytime = strp_time('%Y%m%d%H%M%S', self.index)

    def SellTerm(self):
        if type(self.df['현재가'][self.index]) == pd.Series:
            return False

        bg = self.buycount * self.buyprice
        cg = self.buycount * self.df['현재가'][self.index]
        eyun, per = self.GetEyunPer(bg, cg)
        if per > self.highper:
            self.highper = per

        # 전략 비공개

        return False

    def Sell(self):
        매수호가1 = self.df['매수호가1'][self.index]
        매수호가2 = self.df['매수호가2'][self.index]
        매수호가3 = self.df['매수호가3'][self.index]
        매수호가4 = self.df['매수호가4'][self.index]
        매수호가5 = self.df['매수호가5'][self.index]
        매수잔량1 = self.df['매수잔량1'][self.index]
        매수잔량2 = self.df['매수잔량2'][self.index]
        매수잔량3 = self.df['매수잔량3'][self.index]
        매수잔량4 = self.df['매수잔량4'][self.index]
        매수잔량5 = self.df['매수잔량5'][self.index]
        남은수량 = self.buycount
        직전남은수량 = 남은수량
        매도금액 = 0
        호가정보 = {매수호가1: 매수잔량1, 매수호가2: 매수잔량2, 매수호가3: 매수잔량3, 매수호가4: 매수잔량4, 매수호가5: 매수잔량5}
        for 매수호가, 매수잔량 in 호가정보.items():
            남은수량 -= 매수잔량
            if 남은수량 <= 0:
                매도금액 += 매수호가 * 직전남은수량
                break
            else:
                매도금액 += 매수호가 * 매수잔량
                직전남은수량 = 남은수량
        if 남은수량 <= 0:
            예상체결가 = round(매도금액 / self.buycount, 2)
            self.sellprice = 예상체결가
            self.hold = False
            self.CalculationEyun()
            self.highper = 0
            self.indexb = 0

    def CalculationEyun(self):
        self.totalcount += 1
        bg = self.buycount * self.buyprice
        cg = self.buycount * self.sellprice
        eyun, per = self.GetEyunPer(bg, cg)
        self.totalper = round(self.totalper + per, 2)
        self.totaleyun = int(self.totaleyun + eyun)
        self.totalholdday += self.indexn - self.indexb
        if per > 0:
            self.totalcount_p += 1
        else:
            self.totalcount_m += 1
        self.q.put([self.code, self.df.index[self.indexb], self.index, self.buyprice, self.sellprice, per, eyun])

    # noinspection PyMethodMayBeStatic
    def GetEyunPer(self, bg, cg):
        gtexs = cg * 0.0023
        gsfee = cg * 0.00015
        gbfee = bg * 0.00015
        texs = gtexs - (gtexs % 1)
        sfee = gsfee - (gsfee % 10)
        bfee = gbfee - (gbfee % 10)
        pg = int(cg - texs - sfee - bfee)
        eyun = pg - bg
        per = round(eyun / bg * 100, 2)
        return eyun, per

    def Report(self, count, tcount):
        if self.totalcount > 0:
            plus_per = round((self.totalcount_p / self.totalcount) * 100, 2)
            self.q.put([self.code, self.totalcount, self.totalholdday, self.totalcount_p, self.totalcount_m,
                        plus_per, self.totalper, self.totaleyun])
            totalcount, totalholdday, totalcount_p, totalcount_m, plus_per, totalper, totaleyun = \
                self.GetTotal(plus_per, self.totalholdday)
            print(f" 종목코드 {self.code} | 보유기간합계 {totalholdday}초 | 거래횟수 {totalcount}회 | "
                  f" 익절 {totalcount_p}회 | 손절 {totalcount_m}회 | 승률 {plus_per}% |"
                  f" 수익률 {totalper}% | 수익금 {totaleyun}원 [{count}/{tcount}]")
        else:
            self.q.put([self.code, 0, 0, 0, 0, 0., 0., 0])

    def GetTotal(self, plus_per, totalholdday):
        totalcount = str(self.totalcount)
        totalcount = '  ' + totalcount if len(totalcount) == 1 else totalcount
        totalcount = ' ' + totalcount if len(totalcount) == 2 else totalcount
        totalholdday = str(totalholdday)
        totalholdday = '   ' + totalholdday if len(totalholdday) == 1 else totalholdday
        totalholdday = '  ' + totalholdday if len(totalholdday) == 2 else totalholdday
        totalholdday = ' ' + totalholdday if len(totalholdday) == 3 else totalholdday
        totalholdday = totalholdday + '0' if len(totalholdday) == 1 else totalholdday
        totalcount_p = str(self.totalcount_p)
        totalcount_p = '  ' + totalcount_p if len(totalcount_p) == 1 else totalcount_p
        totalcount_p = ' ' + totalcount_p if len(totalcount_p) == 2 else totalcount_p
        totalcount_m = str(self.totalcount_m)
        totalcount_m = '  ' + totalcount_m if len(totalcount_m) == 1 else totalcount_m
        totalcount_m = ' ' + totalcount_m if len(totalcount_m) == 2 else totalcount_m
        plus_per = str(plus_per)
        plus_per = '  ' + plus_per if len(plus_per.split('.')[0]) == 1 else plus_per
        plus_per = ' ' + plus_per if len(plus_per.split('.')[0]) == 2 else plus_per
        plus_per = plus_per + '0' if len(plus_per.split('.')[1]) == 1 else plus_per
        totalper = str(self.totalper)
        totalper = '   ' + totalper if len(totalper.split('.')[0]) == 1 else totalper
        totalper = '  ' + totalper if len(totalper.split('.')[0]) == 2 else totalper
        totalper = ' ' + totalper if len(totalper.split('.')[0]) == 3 else totalper
        totalper = totalper + '0' if len(totalper.split('.')[1]) == 1 else totalper
        totaleyun = format(self.totaleyun, ',')
        if len(totaleyun.split(',')) == 1:
            totaleyun = '         ' + totaleyun if len(totaleyun.split(',')[0]) == 1 else totaleyun
            totaleyun = '        ' + totaleyun if len(totaleyun.split(',')[0]) == 2 else totaleyun
            totaleyun = '       ' + totaleyun if len(totaleyun.split(',')[0]) == 3 else totaleyun
            totaleyun = '      ' + totaleyun if len(totaleyun.split(',')[0]) == 4 else totaleyun
        elif len(totaleyun.split(',')) == 2:
            totaleyun = '     ' + totaleyun if len(totaleyun.split(',')[0]) == 1 else totaleyun
            totaleyun = '    ' + totaleyun if len(totaleyun.split(',')[0]) == 2 else totaleyun
            totaleyun = '   ' + totaleyun if len(totaleyun.split(',')[0]) == 3 else totaleyun
            totaleyun = '  ' + totaleyun if len(totaleyun.split(',')[0]) == 4 else totaleyun
        elif len(totaleyun.split(',')) == 3:
            totaleyun = ' ' + totaleyun if len(totaleyun.split(',')[0]) == 1 else totaleyun
        return totalcount, totalholdday, totalcount_p, totalcount_m, plus_per, totalper, totaleyun


class Total:
    def __init__(self, q_, last_, num_, df1_):
        super().__init__()
        self.q = q_
        self.last = last_
        self.name = df1_

        self.gap_ch = num_[0]
        self.avg_time = num_[1]
        self.gap_sm = num_[2]
        self.ch_low = num_[3]
        self.dm_low = num_[4]
        self.per_low = num_[5]
        self.per_high = num_[6]
        self.sell_ratio = num_[7]

        self.Start()

    def Start(self):
        columns = ['거래횟수', '보유기간합계', '익절', '손절', '승률', '수익률', '수익금']
        df_back = pd.DataFrame(columns=columns)
        df_tsg = pd.DataFrame(columns=['종목명', '매수시간', '매도시간', '매수가', '매도가', '수익률', 'sgm'])
        k = 0
        while True:
            data = self.q.get()
            if len(data) == 7:
                name = self.name['종목명'][data[0]]
                if data[2] in df_tsg.index:
                    df_tsg.at[data[2]] = df_tsg['종목명'][data[2]] + ';' + name, \
                                         df_tsg['매수시간'][data[2]] + ';' + data[1], \
                                         df_tsg['매도시간'][data[2]] + ';' + data[2], \
                                         df_tsg['매수가'][data[2]] + ';' + str(data[3]), \
                                         df_tsg['매도가'][data[2]] + ';' + str(data[4]), \
                                         df_tsg['수익률'][data[2]] + ';' + str(data[5]), \
                                         df_tsg['sgm'][data[2]] + data[6]
                else:
                    df_tsg.at[data[2]] = name, data[1], data[2], str(data[3]), str(data[4]), str(data[5]), data[6]
            else:
                df_back.at[data[0]] = data[1], data[2], data[3], data[4], data[5], data[6], data[7]
                k += 1
            if k == self.last:
                break

        if len(df_back) > 0:
            text = [self.gap_ch, self.avg_time, self.gap_sm, self.ch_low, self.dm_low, self.per_low, self.per_high, self.sell_ratio]
            print(f' {text}')
            df_back = df_back[df_back['거래횟수'] > 0]
            tc = df_back['거래횟수'].sum()
            if tc != 0:
                pc = df_back['익절'].sum()
                mc = df_back['손절'].sum()
                pper = round(pc / tc * 100, 2)
                avghold = round(df_back['보유기간합계'].sum() / tc, 2)
                avgsp = round(df_back['수익률'].sum() / tc, 2)
                tsg = int(df_back['수익금'].sum())
                onedaycount = round(tc / TOTALTIME, 4)
                onegm = int(BETTING * onedaycount * avghold)
                if onegm < BETTING:
                    onegm = BETTING
                tsp = round(tsg / onegm * 100, 4)
                text = f" 종목당 배팅금액 {format(BETTING, ',')}원, 필요자금 {format(onegm, ',')}원,"\
                       f" 종목출현빈도수 {onedaycount}개/초, 거래횟수 {tc}회, 평균보유기간 {avghold}초,\n 익절 {pc}회,"\
                       f" 손절 {mc}회, 승률 {pper}%, 평균수익률 {avgsp}%, 수익률합계 {tsp}%, 수익금합계 {format(tsg, ',')}원"
                print(text)
                conn = sqlite3.connect(DB_BACKTEST)
                df_back.to_sql(f"vj_jc_code_{strf_time('%Y%m%d')}", conn, if_exists='append', chunksize=1000)
                conn.close()

        if len(df_tsg) > 0:
            df_tsg.sort_values(by=['매도시간'], inplace=True)
            df_tsg['sgm_cumsum'] = df_tsg['sgm'].cumsum()
            df_tsg[['sgm', 'sgm_cumsum']] = df_tsg[['sgm', 'sgm_cumsum']].astype(int)
            conn = sqlite3.connect(DB_BACKTEST)
            df_tsg.to_sql(f"vj_jc_time_{strf_time('%Y%m%d')}", conn, if_exists='append', chunksize=1000)
            conn.close()
            df_tsg.plot(figsize=(12, 9), rot=45)
            plt.show()


if __name__ == "__main__":
    start = datetime.datetime.now()

    con = sqlite3.connect(DB_TICK)
    df = pd.read_sql("SELECT name FROM sqlite_master WHERE TYPE = 'table'", con)
    df1 = pd.read_sql('SELECT * FROM codename', con).set_index('index')
    df2 = pd.read_sql('SELECT * FROM moneytop', con).set_index('index')
    con.close()

    table_list = list(df['name'].values)
    table_list.remove('moneytop')
    table_list.remove('codename')
    if 'dist' in table_list:
        table_list.remove('dist')
    if 'dist_chk' in table_list:
        table_list.remove('dist_chk')
    if 'sqlite_sequence' in table_list:
        table_list.remove('sqlite_sequence')
    if 'temp' in table_list:
        table_list.remove('temp')
    last = len(table_list)

    if len(table_list) > 0:
        gap_ch = 5.0
        avg_time = 60
        gap_sm = 90
        ch_low = 90
        dm_low = 3000
        per_low = 5
        per_high = 25
        sell_ratio = 0.5
        num = [gap_ch, avg_time, gap_sm, ch_low, dm_low, per_low, per_high, sell_ratio]

        q = Queue()
        w = Process(target=Total, args=(q, last, num, df1))
        w.start()
        procs = []
        workcount = int(last / MULTI_COUNT) + 1
        for j in range(0, last, workcount):
            code_list = table_list[j:j + workcount]
            p = Process(target=BackTesterVj, args=(q, code_list, num, df2))
            procs.append(p)
            p.start()
        for p in procs:
            p.join()
        w.join()

    end = datetime.datetime.now()
    print(f" 백테스팅 소요시간 {end - start}")
