import os
import sys
import sqlite3
import pandas as pd
from matplotlib import pyplot as plt
from multiprocessing import Process, Queue
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import now, strf_time
from utility.setting import db_day, db_backtest, db_stg, graph_path

STARTDAY = 20150600     # 백테스팅 시작날짜
BATTING = 5000000       # 단위 원
PREMONEY = 300000       # 단위 백만
AVGPERIOD = 7           # 돌파계수 평균기간


class BackTesterShort:
    def __init__(self, q_, code_list_):
        self.q = q_
        self.code_list = code_list_
        self.code = None
        self.df = None

        self.totalcount = 0
        self.totalcount_p = 0
        self.totalcount_m = 0
        self.totalholdday = 0
        self.totaleyun = 0
        self.totalper = 0.

        self.hold = False
        self.buycount = 0
        self.buyprice = 0
        self.sellprice = 0
        self.index = 0
        self.indexb = 0
        self.indexn = 0

        self.Start()

    def Start(self):
        conn = sqlite3.connect(db_day)
        tcount = len(self.code_list)
        for k, code in enumerate(self.code_list):
            self.code = code
            self.df = pd.read_sql(f"SELECT * FROM '{code}'", conn)
            self.df = self.df.set_index('일자')
            self.df['고가저가폭'] = self.df['고가'] - self.df['저가']
            self.df['종가시가폭'] = self.df['현재가'] - self.df['시가']
            self.df[['종가시가폭']] = self.df[['종가시가폭']].abs()
            self.df['돌파계수'] = self.df['종가시가폭'] / self.df['고가저가폭']
            self.df[['돌파계수']] = self.df[['돌파계수']].astype(float).round(2)
            self.df['평균돌파계수'] = self.df['돌파계수'].rolling(window=AVGPERIOD).mean()
            self.df[['평균돌파계수']] = self.df[['평균돌파계수']].round(2)
            self.totalcount = 0
            self.totalcount_p = 0
            self.totalcount_m = 0
            self.totalholdday = 0
            self.totaleyun = 0
            self.totalper = 0.
            lasth = len(self.df) - 1
            for h, index in enumerate(self.df.index):
                if int(index) < STARTDAY or h < AVGPERIOD:
                    continue
                self.index = index
                self.indexn = h
                if not self.hold and self.BuyTerm():
                    self.Buy()
                elif self.hold and self.SellTerm():
                    self.Sell()
                if self.hold and h == lasth:
                    self.Sell(lastcandle=True)
            self.Report(k + 1, tcount)
        conn.close()

    def BuyTerm(self):
        if self.df['현재가'][self.index] == 0 or self.df['현재가'][self.indexn - 1] == 0 or \
                self.df['현재가'][self.indexn - 2] == 0:
            return False
        k = self.df['고가저가폭'][self.indexn - 1] * self.df['평균돌파계수'][self.indexn - 1]
        if self.df['거래대금'][self.indexn - 1] >= PREMONEY and \
                self.df['고가'][self.index] >= self.df['시가'][self.index] + k:
            return True
        return False

    def Buy(self):
        k = self.df['고가저가폭'][self.indexn - 1] * self.df['평균돌파계수'][self.indexn - 1]
        self.buyprice = self.df['시가'][self.index] + k
        self.buycount = int(BATTING / self.buyprice)
        if self.buycount == 0:
            return
        self.indexb = self.indexn
        self.hold = True

    def SellTerm(self):
        k = self.df['고가저가폭'][self.indexn - 1] * self.df['평균돌파계수'][self.indexn - 1]
        low = self.df['시가'][self.index] - k
        high = self.df['시가'][self.index] + k
        if self.df['저가'][self.index] <= low:
            self.sellprice = low
            return True
        if self.df['고가'][self.index] >= high:
            self.sellprice = high
            return True
        return False

    def Sell(self, lastcandle=False):
        if lastcandle:
            self.sellprice = self.df['현재가'][self.index]
        self.hold = False
        self.CalculationEyun()
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
        self.q.put([self.index, self.code, per, eyun])

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
            avgholdday = round(self.totalholdday / self.totalcount, 2)
            self.q.put([self.code, self.totalcount, avgholdday, self.totalcount_p, self.totalcount_m,
                        plus_per, self.totalper, self.totaleyun])
            totalcount, avgholdday, totalcount_p, totalcount_m, plus_per, totalper, totaleyun = \
                self.GetTotal(plus_per, avgholdday)
            print(f" 종목코드 {self.code} | 평균보유기간 {avgholdday}일 | 거래횟수 {totalcount}회 | "
                  f" 익절 {totalcount_p}회 | 손절 {totalcount_m}회 | 승률 {plus_per}% |"
                  f" 수익률 {totalper}% | 수익금 {totaleyun}원 [{count}/{tcount}]")
        else:
            self.q.put([self.code, 0, 0, 0, 0, 0., 0., 0])

    def GetTotal(self, plus_per, avgholdday):
        totalcount = str(self.totalcount)
        totalcount = '  ' + totalcount if len(totalcount) == 1 else totalcount
        totalcount = ' ' + totalcount if len(totalcount) == 2 else totalcount
        avgholdday = str(avgholdday)
        avgholdday = '  ' + avgholdday if len(avgholdday.split('.')[0]) == 1 else avgholdday
        avgholdday = ' ' + avgholdday if len(avgholdday.split('.')[0]) == 2 else avgholdday
        avgholdday = avgholdday + '0' if len(avgholdday.split('.')[1]) == 1 else avgholdday
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
        return totalcount, avgholdday, totalcount_p, totalcount_m, plus_per, totalper, totaleyun


class Total:
    def __init__(self, q_, last_, totalday_, df1_):
        super().__init__()
        self.q = q_
        self.last = last_
        self.name = df1_
        self.totalday = totalday_

        self.Start()

    def Start(self):
        columns = ['거래횟수', '평균보유기간', '익절', '손절', '승률', '수익률', '수익금']
        df_back = pd.DataFrame(columns=columns)
        df_tsg = pd.DataFrame(columns=['종목명', 'per', 'ttsg'])
        k = 0
        while True:
            data = self.q.get()
            if len(data) == 4:
                name = self.name['종목명'][data[1]]
                if data[0] in df_tsg.index:
                    df_tsg.at[data[0]] = df_tsg['종목명'][data[0]] + ';' + name, \
                                         df_tsg['per'][data[0]] + data[2], \
                                         df_tsg['ttsg'][data[0]] + data[3]
                else:
                    df_tsg.at[data[0]] = name, data[2], data[3]
            else:
                df_back.at[data[0]] = data[1], data[2], data[3], data[4], data[5], data[6], data[7]
                k += 1
            if k == self.last:
                break

        if len(df_back) > 0:
            tc = df_back['거래횟수'].sum()
            if tc != 0:
                pc = df_back['익절'].sum()
                mc = df_back['손절'].sum()
                pper = round(pc / tc * 100, 2)
                df_back_ = df_back[df_back['평균보유기간'] != 0]
                avghold = round(df_back_['평균보유기간'].sum() / len(df_back_), 2)
                avgsp = round(df_back['수익률'].sum() / tc, 2)
                tsg = int(df_back['수익금'].sum())
                onedaycount = round(tc / self.totalday, 2)
                onegm = int(BATTING * onedaycount * avghold)
                if onegm < BATTING:
                    onegm = BATTING
                tsp = round(tsg / onegm * 100, 2)
                text = f" 종목당 배팅금액 {format(BATTING, ',')}원, 필요자금 {format(onegm, ',')}원, "\
                       f" 종목출현빈도수 {onedaycount}개/일, 거래횟수 {tc}회, 평균보유기간 {avghold}일, 익절 {pc}회, "\
                       f" 손절 {mc}회, 승률 {pper}%, 평균수익률 {avgsp}%, 수익률합계 {tsp}%, 수익금합계 {format(tsg, ',')}원"
                print(text)
                conn = sqlite3.connect(db_backtest)
                df_back.to_sql(f"{strf_time('%Y%m%d')}_short", conn, if_exists='replace', chunksize=1000)
                conn.close()

        if len(df_tsg) > 0:
            df_tsg['일자'] = df_tsg.index
            df_tsg.sort_values(by=['일자'], inplace=True)
            df_tsg['ttsg_cumsum'] = df_tsg['ttsg'].cumsum()
            df_tsg[['ttsg', 'ttsg_cumsum']] = df_tsg[['ttsg', 'ttsg_cumsum']].astype(int)
            conn = sqlite3.connect(db_backtest)
            df_tsg.to_sql(f"{strf_time('%Y%m%d')}_short_day", conn, if_exists='replace', chunksize=1000)
            conn.close()
            df_tsg.plot(title='short', figsize=(12, 9), rot=45)
            plt.savefig(f"{graph_path}/{strf_time('%Y%m%d')}_short.png")
            plt.show()


if __name__ == "__main__":
    start = now()

    con = sqlite3.connect(db_stg)
    df1 = pd.read_sql('SELECT * FROM codename', con)
    df1 = df1.set_index('index')
    con.close()

    con = sqlite3.connect(db_day)
    df = pd.read_sql("SELECT name FROM sqlite_master WHERE TYPE = 'table'", con)
    table_list = list(df['name'].values)
    last = len(table_list)

    day_list = []
    for i, codee in enumerate(table_list):
        df = pd.read_sql(f"SELECT * FROM '{codee}'", con)
        df = df.set_index('일자')
        for indexx in df.index:
            if int(indexx) >= STARTDAY and indexx not in day_list:
                day_list.append(indexx)
        print(f' 백테스팅 총 거래일수 계산 중 ... {i + 1}/{last}')
    totalday = len(day_list)
    con.close()

    q = Queue()

    w = Process(target=Total, args=(q, last, totalday, df1))
    w.start()
    procs = []
    workcount = int(last / 6) + 1
    for j in range(0, last, workcount):
        code_list = table_list[j:j + workcount]
        p = Process(target=BackTesterShort, args=(q, code_list))
        procs.append(p)
        p.start()
    for p in procs:
        p.join()
    w.join()

    end = now()
    print(f' 백테스팅 소요시간 {end - start}')
