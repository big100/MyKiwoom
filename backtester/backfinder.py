import os
import sys
import sqlite3
import pandas as pd
from multiprocessing import Process, Queue
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.setting import DB_TICK, DB_BACKFIND
from utility.static import now, strf_time, timedelta_sec, strp_time

AVG_TICK_COUNT = 60     # 평균값계산틱수
PER_CALC_COUNT = 300    # 현재시점 이후 수익률 계산용 틱수
CALC_AFTER_PER = 1.05   # 현재시점 이후 PER_CALC_COUNT 틱수내 수익률


class BackFinder:
    def __init__(self, q_, code_list_, df_mt_):
        self.q = q_
        self.code_list = code_list_
        self.df_mt = df_mt_
        self.Start()

    def Start(self):
        conn = sqlite3.connect(DB_TICK)
        tcount = len(self.code_list)
        for k, code in enumerate(self.code_list):
            columns = ['등락율', '시가대비등락율', '고저평균대비등락율', '초당거래대금', '당일거래대금',
                       '체결강도', '체결강도차이', '초당거래대금차이']
            df_bf = pd.DataFrame(columns=columns)
            count_cond = 0
            df = pd.read_sql(f"SELECT * FROM '{code}'", conn)
            df = df.set_index('index')
            lasth = len(df) - 1

            for h, index in enumerate(df.index):
                try:
                    if code not in self.df_mt['거래대금순위'][index]:
                        count_cond = 0
                    else:
                        count_cond += 1
                except KeyError:
                    continue
                if count_cond < AVG_TICK_COUNT + 1:
                    continue
                if strp_time('%Y%m%d%H%M%S', index) < \
                        timedelta_sec(180, strp_time('%Y%m%d%H%M%S', df['VI발동시간'][index])):
                    continue
                if df['현재가'][index] >= df['VI아래5호가'][index]:
                    continue
                if h >= lasth - PER_CALC_COUNT:
                    break
                if df['현재가'][h:h + PER_CALC_COUNT].max() > df['현재가'][index] * CALC_AFTER_PER:
                    per = df['등락율'][index]
                    oper = round((df['현재가'][index] / df['시가'][index] - 1) * 100, 2)
                    hper = round((df['현재가'][index] / ((df['고가'][index] + df['저가'][index]) / 2) - 1) * 100, 2)
                    sm = int(df['초당거래대금'][index])
                    dm = int(df['당일거래대금'][index])
                    ch = df['체결강도'][index]
                    gap_ch = round(df['체결강도'][index] - df['체결강도'][h - AVG_TICK_COUNT:h].mean(), 2)
                    gap_sm = round(df['초당거래대금'][index] - df['초당거래대금'][h - AVG_TICK_COUNT:h].mean(), 2)
                    df_bf.at[code + index] = per, oper, hper, sm, dm, ch, gap_ch, gap_sm
            print(f' 백파인더 검색 중 ... [{k + 1}/{tcount}]')
            self.q.put(df_bf)
        conn.close()


class Total:
    def __init__(self, q_, last_):
        super().__init__()
        self.q = q_
        self.last = last_
        self.Start()

    def Start(self):
        df = []
        k = 0
        while True:
            data = self.q.get()
            df.append(data)
            k += 1
            if k == self.last:
                break
        if len(df) > 0:
            df = pd.concat(df)
            conn = sqlite3.connect(DB_BACKFIND)
            df.to_sql(f"{strf_time('%Y%m%d')}_tick", conn, if_exists='replace', chunksize=1000)
            conn.close()


if __name__ == "__main__":
    start = now()
    q = Queue()

    con = sqlite3.connect(DB_TICK)
    df_name = pd.read_sql("SELECT name FROM sqlite_master WHERE TYPE = 'table'", con)
    df_mt = pd.read_sql('SELECT * FROM moneytop', con).set_index('index')
    con.close()
    table_list = list(df_name['name'].values)
    table_list.remove('moneytop')
    table_list.remove('codename')
    last = len(table_list)

    w = Process(target=Total, args=(q, last))
    w.start()
    procs = []
    workcount = int(last / 6) + 1
    for j in range(0, last, workcount):
        code_list = table_list[j:j + workcount]
        p = Process(target=BackFinder, args=(q, code_list, df_mt))
        procs.append(p)
        p.start()
    for p in procs:
        p.join()
    w.join()

    end = now()
    print(f' 백파인더 소요시간 {end - start}')
