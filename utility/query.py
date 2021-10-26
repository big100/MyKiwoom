import sqlite3
import pandas as pd
from utility.setting import DB_STG, DB_TICK
from utility.static import now, float2str1p6


class Query:
    def __init__(self, windowQ, traderQ, queryQ):
        self.windowQ = windowQ
        self.traderQ = traderQ
        self.queryQ = queryQ
        self.con1 = sqlite3.connect(DB_STG)
        self.cur1 = self.con1.cursor()
        self.con2 = sqlite3.connect(DB_TICK)
        self.cur2 = self.con2.cursor()
        self.cur2.execute('pragma journal_mode=WAL')
        self.cur2.execute('pragma synchronous=normal')
        self.cur2.execute('pragma temp_store=memory')
        self.trigger = False
        self.remove_trigger()
        self.Start()

    def __del__(self):
        self.remove_trigger()
        self.con1.close()
        self.con2.close()

    def Start(self):
        k, j = 0, 0
        df = pd.DataFrame()
        while True:
            query = self.queryQ.get()
            if query == '디비트리거시작':
                self.create_trigger()
                self.trigger = True
            elif query[0] == 1:
                try:
                    if len(query) == 2:
                        self.cur1.execute(query[1])
                        self.con1.commit()
                    elif len(query) == 4:
                        query[1].to_sql(query[2], self.con1, if_exists=query[3], chunksize=1000, method='multi')
                except Exception as e:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - con1 {e}'])
            elif query[0] == 2:
                try:
                    if len(query) == 2:
                        if type(query[1]) == str:
                            self.cur2.execute(query[1])
                            self.con2.commit()
                        else:
                            k += 1
                            for code in list(query[1].keys()):
                                query[1][code]['종목코드'] = code
                                df = df.append(query[1][code])
                            if k == 4 and self.trigger:
                                start = now()
                                df.to_sql("temp", self.con2, if_exists='append', chunksize=1000, method='multi')
                                self.cur2.execute('insert into "dist" ("cnt") values (1);')
                                save_time = float2str1p6((now() - start).total_seconds())
                                text = f'시스템 명령 실행 알림 - 틱데이터 저장 쓰기소요시간은 [{save_time}]초입니다.'
                                self.windowQ.put([1, text])
                                k = 0
                                df = pd.DataFrame()
                    elif len(query) == 3:
                        start = now()
                        j += 1
                        last = len(list(query[1].keys()))
                        for i, code in enumerate(list(query[1].keys())):
                            text = f'시스템 명령 실행 알림 - 시스템 명령 실행 알림 - 틱데이터 저장 중 ... [{j}]{i+1}/{last}'
                            self.windowQ.put([1, text])
                            query[1][code].to_sql(code, self.con2, if_exists='append', chunksize=1000, method='multi')
                        save_time = float2str1p6((now() - start).total_seconds())
                        self.windowQ.put([1, f'시스템 명령 실행 알림 - 틱데이터 저장 쓰기소요시간은 [{save_time}]초입니다.'])
                        if j == 4:
                            self.traderQ.put('틱데이터저장완료')
                    elif len(query) == 4:
                        query[1].to_sql(query[2], self.con2, if_exists=query[3], chunksize=1000, method='multi')
                except Exception as e:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - con2 {e}'])

    def create_trigger(self):
        res = self.cur2.execute("SELECT name FROM sqlite_master WHERE type='table';")
        table_list = []
        for name in res.fetchall():
            table_list.append(name[0])

        const_str = '"index", 현재가, 시가, 고가, 저가, 등락율, 당일거래대금, 체결강도, 초당매수수량, 초당매도수량, VI해제시간,' \
                    'VI아래5호가, 매도총잔량, 매수총잔량, 매도호가5, 매도호가4, 매도호가3, 매도호가2, 매도호가1, 매수호가1, 매수호가2,' \
                    '매수호가3, 매수호가4, 매수호가5, 매도잔량5, 매도잔량4, 매도잔량3, 매도잔량2, 매도잔량1, 매수잔량1, 매수잔량2,' \
                    '매수잔량3, 매수잔량4, 매수잔량5'

        list_stock_table = []
        for table_name in table_list:
            if len(table_name) == 6:
                list_stock_table.append(table_name)

        query_create_temp = \
            'CREATE TABLE IF NOT EXISTS "temp" ("index" TEXT, "종목코드" TEXT, "현재가" REAL, "시가" REAL, "고가" REAL,' \
            '"저가" REAL, "등락율" REAL, "당일거래대금" REAL, "체결강도" REAL, "초당매수수량" REAL, "초당매도수량" REAL,' \
            '"VI해제시간" TEXT, "VI아래5호가" REAL, "매도총잔량" REAL, "매수총잔량" REAL, "매도호가5" REAL, "매도호가4" REAL,' \
            '"매도호가3" REAL, "매도호가2" REAL, "매도호가1" REAL, "매수호가1" REAL, "매수호가2" REAL, "매수호가3" REAL,' \
            '"매수호가4" REAL, "매수호가5" REAL, "매도잔량5" REAL, "매도잔량4" REAL, "매도잔량3" REAL, "매도잔량2" REAL,' \
            '"매도잔량1" REAL, "매수잔량1" REAL, "매수잔량2" REAL, "매수잔량3" REAL, "매수잔량4" REAL, "매수잔량5" REAL);'
        query_create_dist = \
            'CREATE TABLE IF NOT EXISTS "dist" (uid integer primary key autoincrement, cnt integer,' \
            ' reg_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL);'
        query_create_distchk = \
            'CREATE TABLE IF NOT EXISTS "dist_chk" (uid integer primary key autoincrement, cnt integer,' \
            'reg_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL);'

        s = 'CREATE TRIGGER IF NOT EXISTS "dist_trigger" INSERT ON "dist" BEGIN INSERT INTO "dist_chk" ("cnt") values (1);\n'
        for i in range(len(list_stock_table)):
            s += 'INSERT INTO "' + list_stock_table[i] + '" SELECT ' + const_str + ' FROM temp WHERE 종목코드 = "' + \
                list_stock_table[i] + '";\n'
        s += 'DELETE FROM temp;\n'
        s += 'INSERT INTO "dist_chk" ("cnt") values (2);\n'  # 디버깅 속도측정용
        s += 'END;\n'
        query_create_trigger = s

        self.cur2.execute(query_create_temp)
        self.cur2.execute(query_create_dist)
        self.cur2.execute(query_create_distchk)
        self.cur2.execute(query_create_trigger)

    def remove_trigger(self):
        try:
            self.cur2.execute('drop trigger dist_trigger;')
        except sqlite3.OperationalError:
            pass
