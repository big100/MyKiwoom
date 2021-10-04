import os
import sys
import time
import psutil
import pythoncom
from PyQt5 import QtWidgets
from threading import Timer, Lock
from PyQt5.QAxContainer import QAxWidget
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import *
from utility.setting import *

TUJAGMDIVIDE = 5    # 종목당 투자금 분할계수


class Trader:
    app = QtWidgets.QApplication(sys.argv)

    def __init__(self, windowQ, traderQ, stgQ, receivQ, soundQ, queryQ, teleQ, hoga1Q, hoga2Q,
                 chart1Q, chart2Q, chart3Q, chart4Q, chart5Q, chart6Q, chart7Q, chart8Q, chart9Q):
        self.windowQ = windowQ
        self.traderQ = traderQ
        self.stgQ = stgQ
        self.soundQ = soundQ
        self.receivQ = receivQ
        self.queryQ = queryQ
        self.teleQ = teleQ
        self.hoga1Q = hoga1Q
        self.hoga2Q = hoga2Q
        self.chart1Q = chart1Q
        self.chart2Q = chart2Q
        self.chart3Q = chart3Q
        self.chart4Q = chart4Q
        self.chart5Q = chart5Q
        self.chart6Q = chart6Q
        self.chart7Q = chart7Q
        self.chart8Q = chart8Q
        self.chart9Q = chart9Q
        self.lock = Lock()

        self.dict_name = {}     # key: 종목코드, value: 종목명
        self.dict_sghg = {}     # key: 종목코드, value: [상한가, 하한가]
        self.dict_vipr = {}     # key: 종목코드, value: [갱신여부, 발동시간+5초, 해제시간+180초, UVI, DVI, UVID5]
        self.dict_cond = {}     # key: 조건검색식번호, value: 조건검색식명
        self.dict_hoga = {}     # key: 호가창번호, value: [종목코드, 갱신여부, 호가잔고(DataFrame)]
        self.dict_chat = {}     # key: UI번호, value: 종목코드
        self.dict_gsjm = {}     # key: 종목코드, value: 마지막체결시간
        self.dict_df = {
            '실현손익': pd.DataFrame(columns=columns_tt),
            '거래목록': pd.DataFrame(columns=columns_td),
            '잔고평가': pd.DataFrame(columns=columns_tj),
            '잔고목록': pd.DataFrame(columns=columns_jg),
            '체결목록': pd.DataFrame(columns=columns_cj),
            'TRDF': pd.DataFrame(columns=[])
        }
        self.dict_intg = {
            '장운영상태': 1,
            '예수금': 0,
            '추정예수금': 0,
            '추정예탁자산': 0,
            '종목당투자금': 0,
            'TR제한수신횟수': 0,
            '스레드': 0,
            '시피유': 0.,
            '메모리': 0.
        }
        self.dict_strg = {
            '당일날짜': strf_time('%Y%m%d'),
            '계좌번호': '',
            'TR종목명': '',
            'TR명': ''
        }
        self.dict_bool = {
            '데이터베이스로딩': False,
            '계좌잔고조회': False,
            '업종차트조회': False,
            '업종지수등록': False,
            'VI발동해제등록': False,
            '잔고청산': False,
            '실시간데이터수신중단': False,
            '일별거래목록저장': False,

            '테스트': False,
            '모의투자': False,
            '알림소리': False,

            '로그인': False,
            'TR수신': False,
            'TR다음': False
        }
        remaintime = (strp_time('%Y%m%d%H%M%S', self.dict_strg['당일날짜'] + '090100') - now()).total_seconds()
        exittime = timedelta_sec(remaintime) if remaintime > 0 else timedelta_sec(600)
        self.dict_time = {
            '휴무종료': exittime,
            '잔고갱신': now(),
            '부가정보': now(),
            '호가잔고': now(),
            'TR시작': now(),
            'TR재개': now()
        }
        self.dict_item = None
        self.list_trcd = None
        self.list_kosd = None
        self.list_buy = []
        self.list_sell = []

        self.ocx = QAxWidget('KHOPENAPI.KHOpenAPICtrl.1')
        self.ocx.OnEventConnect.connect(self.OnEventConnect)
        self.ocx.OnReceiveTrData.connect(self.OnReceiveTrData)
        self.ocx.OnReceiveRealData.connect(self.OnReceiveRealData)
        self.ocx.OnReceiveChejanData.connect(self.OnReceiveChejanData)
        self.Start()

    def Start(self):
        self.CreateDatabase()
        self.LoadDatabase()
        self.CommConnect()
        self.EventLoop()

    def CreateDatabase(self):
        con = sqlite3.connect(db_stg)
        df = pd.read_sql("SELECT name FROM sqlite_master WHERE TYPE = 'table'", con)
        table_list = list(df['name'].values)
        con.close()

        if 'chegeollist' not in table_list:
            self.queryQ.put([1, "CREATE TABLE chegeollist ('index' TEXT, '종목명' TEXT, '주문구분' TEXT,"
                                "'주문수량' INTEGER, '미체결수량' INTEGER, '주문가격' INTEGER, '체결가' INTEGER,"
                                "'체결시간' TEXT)"])
            self.queryQ.put([1, "CREATE INDEX 'ix_chegeollist_index' ON 'chegeollist' ('index')"])
            self.windowQ.put([1, '시스템 명령 실행 알림 - 데이터베이스 chegeollist 테이블 생성 완료'])

        if 'jangolist' not in table_list:
            self.queryQ.put([1, "CREATE TABLE 'jangolist' ('index' TEXT, '종목명' TEXT, '매입가' INTEGER,"
                                "'현재가' INTEGER, '수익률' REAL, '평가손익' INTEGER, '매입금액' INTEGER, '평가금액' INTEGER,"
                                "'시가' INTEGER, '고가' INTEGER, '저가' INTEGER, '전일종가' INTEGER, '보유수량' INTEGER)"])
            self.queryQ.put([1, "CREATE INDEX 'ix_jangolist_index' ON 'jangolist' ('index')"])
            self.windowQ.put([1, '시스템 명령 실행 알림 - 데이터베이스 jangolist 테이블 생성 완료'])

        if 'tradelist' not in table_list:
            self.queryQ.put([1, "CREATE TABLE tradelist ('index' TEXT, '종목명' TEXT, '매수금액' INTEGER,"
                                "'매도금액' INTEGER, '주문수량' INTEGER, '수익률' REAL, '수익금' INTEGER, '체결시간' TEXT)"])
            self.queryQ.put([1, "CREATE INDEX 'ix_tradelist_index' ON 'tradelist' ('index')"])
            self.windowQ.put([1, '시스템 명령 실행 알림 - 데이터베이스 tradelist 테이블 생성 완료'])

        if 'totaltradelist' not in table_list:
            self.queryQ.put([1, "CREATE TABLE 'totaltradelist' ('index' TEXT, '총매수금액' INTEGER, '총매도금액' INTEGER,"
                                "'총수익금액' INTEGER, '총손실금액' INTEGER, '수익률' REAL, '수익금합계' INTEGER)"])
            self.queryQ.put([1, "CREATE INDEX 'ix_totaltradelist_index' ON 'totaltradelist' ('index')"])
            self.windowQ.put([1, '시스템 명령 실행 알림 - 데이터베이스 totaltradelist 테이블 생성 완료'])

        if 'setting' not in table_list:
            df = pd.DataFrame([[0, 1, 1,
                                10., 180, 200, 100, 2000, 0., 25., 3.]],
                              columns=['테스트', '모의투자', '알림소리',
                                       '체결강도차이', '평균시간', '거래대금차이',
                                       '체결강도하한', '누적거래대금하한', '등락율하한', '등락율상한', '청산수익률'],
                              index=[0])
            self.queryQ.put([1, df, 'setting', 'replace'])
            self.windowQ.put([1, '시스템 명령 실행 알림 - 데이터베이스 setting 테이블 생성 완료'])

        time.sleep(2)

    def LoadDatabase(self):
        self.dict_bool['데이터베이스로딩'] = True
        self.windowQ.put([2, '데이터베이스 불러오기'])
        con = sqlite3.connect(db_stg)
        df = pd.read_sql('SELECT * FROM setting', con)
        df = df.set_index('index')
        self.dict_bool['테스트'] = df['테스트'][0]
        self.dict_bool['모의투자'] = df['모의투자'][0]
        self.dict_bool['알림소리'] = df['알림소리'][0]
        self.windowQ.put([2, f"테스트모드 {self.dict_bool['테스트']}"])
        self.windowQ.put([2, f"모의투자 {self.dict_bool['모의투자']}"])
        self.windowQ.put([2, f"알림소리 {self.dict_bool['알림소리']}"])

        df = pd.read_sql(f"SELECT * FROM chegeollist WHERE 체결시간 LIKE '{self.dict_strg['당일날짜']}%'", con)
        self.dict_df['체결목록'] = df.set_index('index').sort_values(by=['체결시간'], ascending=False)

        df = pd.read_sql(f"SELECT * FROM tradelist WHERE 체결시간 LIKE '{self.dict_strg['당일날짜']}%'", con)
        self.dict_df['거래목록'] = df.set_index('index').sort_values(by=['체결시간'], ascending=False)

        df = pd.read_sql(f'SELECT * FROM jangolist', con)
        self.dict_df['잔고목록'] = df.set_index('index').sort_values(by=['매입금액'], ascending=False)

        if len(self.dict_df['체결목록']) > 0:
            self.windowQ.put([ui_num['체결목록'], self.dict_df['체결목록']])
        if len(self.dict_df['거래목록']) > 0:
            self.windowQ.put([ui_num['거래목록'], self.dict_df['거래목록']])
        if len(self.dict_df['잔고목록']) > 0:
            for code in self.dict_df['잔고목록'].index:
                self.traderQ.put([sn_jscg, code, '10;12;14;30;228', 1])
                self.receivQ.put(f'잔고편입 {code}')

        self.windowQ.put([1, '시스템 명령 실행 알림 - 데이터베이스 정보 불러오기 완료'])

    def CommConnect(self):
        self.windowQ.put([2, 'OPENAPI 로그인'])
        self.ocx.dynamicCall('CommConnect()')
        while not self.dict_bool['로그인']:
            pythoncom.PumpWaitingMessages()

        self.dict_strg['계좌번호'] = self.ocx.dynamicCall('GetLoginInfo(QString)', 'ACCNO').split(';')[0]

        self.list_kosd = self.GetCodeListByMarket('10')
        list_code = self.GetCodeListByMarket('0') + self.list_kosd
        dict_code = {}
        for code in list_code:
            name = self.GetMasterCodeName(code)
            self.dict_name[code] = name
            dict_code[name] = code

        self.chart9Q.put(self.dict_name)
        self.windowQ.put([3, dict_code])
        self.windowQ.put([4, self.dict_name])

        self.windowQ.put([1, '시스템 명령 실행 알림 - OpenAPI 로그인 완료'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('키움증권 오픈에이피아이에 로그인하였습니다.')

        if int(strf_time('%H%M%S')) > 90000:
            self.windowQ.put([2, '장운영상태'])
            self.dict_intg['장운영상태'] = 3

    def EventLoop(self):
        self.GetAccountjanGo()
        self.GetKospiKosdaqChart()
        self.OperationRealreg()
        self.UpjongjisuRealreg()
        self.ViRealreg()
        while True:
            if not self.traderQ.empty():
                work = self.traderQ.get()
                if type(work) == list:
                    if len(work) == 10:
                        self.SendOrder(work)
                    elif len(work) == 5:
                        self.BuySell(work[0], work[1], work[2], work[3], work[4])
                        continue
                    elif len(work) in [2, 4]:
                        self.UpdateRealreg(work)
                        continue
                elif type(work) == str:
                    self.RunWork(work)

            if self.dict_intg['장운영상태'] == 1 and now() > self.dict_time['휴무종료']:
                break

            if self.dict_intg['장운영상태'] == 2:
                if int(strf_time('%H%M%S')) > 152900 and not self.dict_bool['잔고청산']:
                    self.JangoChungsan()
            if self.dict_intg['장운영상태'] == 8:
                self.AllRemoveRealreg()
                self.SaveDatabase()
                break

            if now() > self.dict_time['호가잔고']:
                self.PutHogaJanngo()
                self.dict_time['호가잔고'] = timedelta_sec(0.25)
            if now() > self.dict_time['잔고갱신']:
                self.UpdateTotaljango()
                self.dict_time['잔고갱신'] = timedelta_sec(1)
            if now() > self.dict_time['부가정보']:
                self.UpdateInfo()
                self.dict_time['부가정보'] = timedelta_sec(2)

            time_loop = timedelta_sec(0.25)
            while now() < time_loop:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.0001)

        self.stgQ.put('전략연산프로세스종료')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 종료'])
        self.SysExit()

    def SendOrder(self, order):
        name = order[-1]
        del order[-1]
        ret = self.ocx.dynamicCall(
            'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)', order)
        if ret != 0:
            self.windowQ.put([1, f'시스템 명령 오류 알림 - {name} {order[5]}주 {order[0]} 주문 실패'])

    def BuySell(self, gubun, code, name, c, oc):
        if gubun == '매수' and code in self.dict_df['잔고목록'].index:
            self.stgQ.put(['매수취소', code])
            return
        elif gubun == '매도' and code not in self.dict_df['잔고목록'].index:
            self.stgQ.put(['매도취소', code])
            return

        if gubun == '매수' and code in self.list_sell:
            self.stgQ.put(['매수취소', code])
            self.windowQ.put([1, '매매 시스템 오류 알림 - 현재 매도 주문중인 종목입니다.'])
            return
        if gubun == '매도' and code in self.list_buy:
            self.stgQ.put(['매도취소', code])
            self.windowQ.put([1, '매매 시스템 오류 알림 - 현재 매수 주문중인 종목입니다.'])
            return

        if gubun == '매수':
            if self.dict_intg['추정예수금'] < oc * c:
                cond = (self.dict_df['체결목록']['주문구분'] == '시드부족') & (self.dict_df['체결목록']['종목명'] == name)
                df = self.dict_df['체결목록'][cond]
                if len(df) == 0 or \
                        (len(df) > 0 and now() > timedelta_sec(180, strp_time('%Y%m%d%H%M%S%f', df['체결시간'][0]))):
                    self.Order('시드부족', code, name, c, oc)
                self.stgQ.put(['매수취소', code])
                return

        self.Order(gubun, code, name, c, oc)

    def Order(self, gubun, code, name, c, oc):
        on = 0
        if gubun == '매수':
            self.dict_intg['추정예수금'] -= oc * c
            self.list_buy.append(code)
            on = 1
        elif gubun == '매도':
            self.list_sell.append(code)
            on = 2

        if self.dict_bool['모의투자'] or gubun == '시드부족':
            self.UpdateChejanData(code, name, '체결', gubun, c, c, oc, 0,
                                  strf_time('%Y%m%d%H%M%S%f'), strf_time('%Y%m%d%H%M%S%f'))
        else:
            self.traderQ.put([gubun, '4989', self.dict_strg['계좌번호'], on, code, oc, 0, '03', '', name])

    def UpdateRealreg(self, rreg):
        if len(rreg) == 2:
            self.ocx.dynamicCall('SetRealRemove(QString, QString)', rreg)
        elif len(rreg) == 4:
            self.ocx.dynamicCall('SetRealReg(QString, QString, QString, QString)', rreg)

    def RunWork(self, work):
        if '현재가' in work:
            gubun = int(work.split(' ')[0][3:5])
            code = work.split(' ')[-1]
            name = self.dict_name[code]
            if gubun == ui_num['차트P0']:
                gubun = ui_num['차트P1']
                if ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]:
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[work]).start()
                    return
                self.chart6Q.put('기업개요 ' + code)
                self.chart7Q.put('기업공시 ' + code)
                self.chart8Q.put('종목뉴스 ' + code)
                self.chart9Q.put('재무제표 ' + code)
                self.hoga1Q.put('초기화')
                self.traderQ.put([sn_cthg, code, '10;12;14;30;228;41;61;71;81', 1])
                if 0 in self.dict_hoga.keys():
                    self.traderQ.put([sn_cthg, self.dict_hoga[0][0]])
                self.dict_hoga[0] = [code, True, pd.DataFrame(columns=columns_hj)]
                self.GetChart(gubun, code, name)
                self.GetTujajaChegeolH(code)
            elif gubun == ui_num['차트P1']:
                if (ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]) or \
                        (ui_num['차트P3'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P3']]):
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[work]).start()
                    return
                self.hoga1Q.put('초기화')
                self.traderQ.put([sn_cthg, code, '10;12;14;30;228;41;61;71;81', 1])
                if 0 in self.dict_hoga.keys():
                    self.traderQ.put([sn_cthg, self.dict_hoga[0][0]])
                self.dict_hoga[0] = [code, True, pd.DataFrame(columns=columns_hj)]
                self.GetChart(gubun, code, name)
            elif gubun == ui_num['차트P3']:
                if (ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]) or \
                        (ui_num['차트P3'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P3']]):
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[work]).start()
                    return
                self.hoga2Q.put('초기화')
                self.traderQ.put([sn_cthg, code, '10;12;14;30;228;41;61;71;81', 1])
                if 1 in self.dict_hoga.keys():
                    self.traderQ.put([sn_cthg, self.dict_hoga[1][0]])
                self.dict_hoga[1] = [code, True, pd.DataFrame(columns=columns_hj)]
                self.GetChart(gubun, code, name)
            elif gubun == ui_num['차트P5']:
                tradeday = work.split(' ')[-2]
                if ui_num['차트P5'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P5']]:
                    return
                if int(tradeday) < int(strf_time('%Y%m%d', timedelta_day(-5))):
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 5일 이전의 체결정보는 조회할 수 없습니다.'])
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[work]).start()
                    return
                self.GetChart(gubun, code, name, tradeday)
        elif '매수취소' in work:
            code = work.split(' ')[1]
            name = self.dict_name[code]
            term = (self.dict_df['체결목록']['종목명'] == name) & (self.dict_df['체결목록']['미체결수량'] > 0) & \
                   (self.dict_df['체결목록']['주문구분'] == '매수')
            df = self.dict_df['체결목록'][term]
            if len(df) == 1:
                on = df.index[0]
                omc = df['미체결수량'][on]
                order = ['매수취소', '4989', self.dict_strg['계좌번호'], 3, code, omc, 0, '00', on, name]
                self.traderQ.put(order)
        elif '매도취소' in work:
            code = work.split(' ')[1]
            name = self.dict_name[code]
            term = (self.dict_df['체결목록']['종목명'] == name) & (self.dict_df['체결목록']['미체결수량'] > 0) & \
                   (self.dict_df['체결목록']['주문구분'] == '매도')
            df = self.dict_df['체결목록'][term]
            if len(df) == 1:
                on = df.index[0]
                omc = df['미체결수량'][on]
                order = ['매도취소', '4989', self.dict_strg['계좌번호'], 4, code, omc, 0, '00', on, name]
                self.traderQ.put(order)
        elif work == '데이터베이스 불러오기':
            if not self.dict_bool['데이터베이스 로딩']:
                self.LoadDatabase()
        elif work == 'OPENAPI 로그인':
            if self.ocx.dynamicCall('GetConnectState()') == 0:
                self.CommConnect()
        elif work == '계좌평가 및 잔고':
            if not self.dict_bool['계좌잔고조회']:
                self.GetAccountjanGo()
        elif work == '코스피 코스닥 차트':
            if not self.dict_bool['업종차트조회']:
                self.GetKospiKosdaqChart()
        elif work == '장운영시간 알림 등록':
            if not self.dict_bool['장운영시간등록']:
                self.OperationRealreg()
        elif work == '업종지수 주식체결 등록':
            if not self.dict_bool['업종지수등록']:
                self.UpjongjisuRealreg()
        elif work == 'VI발동해제 등록':
            if not self.dict_bool['VI발동해제등록']:
                self.ViRealreg()
        elif work == '장운영상태':
            if self.dict_intg['장운영상태'] != 3:
                self.windowQ.put([2, '장운영상태'])
                self.dict_intg['장운영상태'] = 3
        elif work == '실시간 조건검색식 등록':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 해당 명령은 콜렉터에서 자동실행됩니다.'])
        elif work == '잔고청산':
            if not self.dict_bool['잔고청산']:
                self.JangoChungsan()
        elif work == '실시간 데이터 수신 중단':
            if not self.dict_bool['실시간데이터수신중단']:
                self.AllRemoveRealreg()
        elif work == '틱데이터 저장':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 해당 명령은 콜렉터에서 자동실행됩니다.'])
        elif work == '시스템 종료':
            if not self.dict_bool['일별거래목록저장']:
                self.SaveDatabase()
            self.SysExit()
        elif work == '/당일체결목록':
            if len(self.dict_df['체결목록']) > 0:
                self.teleQ.put(self.dict_df['체결목록'])
            else:
                self.teleQ.put('현재는 거래목록이 없습니다.')
        elif work == '/당일거래목록':
            if len(self.dict_df['거래목록']) > 0:
                self.teleQ.put(self.dict_df['거래목록'])
            else:
                self.teleQ.put('현재는 거래목록이 없습니다.')
        elif work == '/계좌잔고평가':
            if len(self.dict_df['잔고목록']) > 0:
                self.teleQ.put(self.dict_df['잔고목록'])
            else:
                self.teleQ.put('현재는 잔고목록이 없습니다.')
        elif work == '/잔고청산주문':
            if not self.dict_bool['실시간데이터수신중단']:
                self.AllRemoveRealreg()
            if not self.dict_bool['잔고청산']:
                self.JangoChungsan()
        elif '설정' in work:
            bot_number = work.split(' ')[1]
            chat_id = int(work.split(' ')[2])
            self.queryQ.put([1, f"UPDATE telegram SET str_bot = '{bot_number}', int_id = '{chat_id}'"])
            if self.dict_bool['알림소리']:
                self.soundQ.put('텔레그램 봇넘버 및 아이디가 변경되었습니다.')
            else:
                self.windowQ.put([1, '시스템 명령 실행 알림 - 텔레그램 봇넘버 및 아이디 설정 완료'])
        elif work == '테스트모드 ON/OFF':
            if self.dict_bool['테스트']:
                self.dict_bool['테스트'] = False
                self.queryQ.put([1, 'UPDATE setting SET 테스트 = 0'])
                self.windowQ.put([2, '테스트모드 OFF'])
                if self.dict_bool['알림소리']:
                    self.soundQ.put('테스트모드 설정이 OFF로 변경되었습니다.')
            else:
                self.dict_bool['테스트'] = True
                self.queryQ.put([1, 'UPDATE setting SET 테스트 = 1'])
                self.windowQ.put([2, '테스트모드 ON'])
                if self.dict_bool['알림소리']:
                    self.soundQ.put('테스트모드 설정이 ON으로 변경되었습니다.')
        elif work == '모의투자 ON/OFF':
            if self.dict_bool['모의투자']:
                self.dict_bool['모의투자'] = False
                self.queryQ.put([1, 'UPDATE setting SET 모의투자 = 0'])
                self.windowQ.put([2, '모의투자 OFF'])
                if self.dict_bool['알림소리']:
                    self.soundQ.put('모의투자 설정이 OFF로 변경되었습니다.')
            else:
                self.dict_bool['모의투자'] = True
                self.queryQ.put([1, 'UPDATE setting SET 모의투자 = 1'])
                self.windowQ.put([2, '모의투자 ON'])
                if self.dict_bool['알림소리']:
                    self.soundQ.put('모의투자 설정이 ON으로 변경되었습니다.')
        elif work == '알림소리 ON/OFF':
            if self.dict_bool['알림소리']:
                self.dict_bool['알림소리'] = False
                self.queryQ.put([1, 'UPDATE setting SET 알림소리 = 0'])
                self.windowQ.put([2, '알림소리 OFF'])
                if self.dict_bool['알림소리']:
                    self.soundQ.put('알림소리 설정이 OFF로 변경되었습니다.')
            else:
                self.dict_bool['알림소리'] = True
                self.queryQ.put([1, 'UPDATE setting SET 알림소리 = 1'])
                self.windowQ.put([2, '알림소리 ON'])
                if self.dict_bool['알림소리']:
                    self.soundQ.put('알림소리 설정이 ON으로 변경되었습니다.')

    def GetChart(self, gubun, code, name, tradeday=None):
        prec = self.GetMasterLastPrice(code)
        if gubun in [ui_num['차트P1'], ui_num['차트P3']]:
            df = self.Block_Request('opt10081', 종목코드=code, 기준일자=self.dict_strg['당일날짜'], 수정주가구분=1,
                                    output='주식일봉차트조회', next=0)
            df2 = self.Block_Request('opt10080', 종목코드=code, 틱범위=3, 수정주가구분=1, output='주식분봉차트조회', next=0)
            if gubun == ui_num['차트P1']:
                self.chart1Q.put([name, prec, df, ''])
                self.chart2Q.put([name, prec, df2, ''])
            elif gubun == ui_num['차트P3']:
                self.chart3Q.put([name, prec, df, ''])
                self.chart4Q.put([name, prec, df2, ''])
        elif gubun == ui_num['차트P5'] and tradeday is not None:
            df2 = self.Block_Request('opt10080', 종목코드=code, 틱범위=3, 수정주가구분=1, output='주식분봉차트조회', next=0)
            self.chart5Q.put([name, prec, df2, tradeday])
        self.dict_chat[gubun] = code

    def GetTujajaChegeolH(self, code):
        df1 = self.Block_Request('opt10059', 일자=self.dict_strg['당일날짜'], 종목코드=code, 금액수량구분=1, 매매구분=0,
                                 단위구분=1, output='종목별투자자', next=0)
        df2 = self.Block_Request('opt10046', 종목코드=code, 틱구분=1, 체결강도구분=1, output='체결강도추이', next=0)
        self.chart1Q.put([code, df1, df2])

    def GetAccountjanGo(self):
        self.dict_bool['계좌잔고조회'] = True
        self.windowQ.put([2, '계좌평가 및 잔고'])

        while True:
            df = self.Block_Request('opw00004', 계좌번호=self.dict_strg['계좌번호'], 비밀번호='', 상장폐지조회구분=0,
                                    비밀번호입력매체구분='00', output='계좌평가현황', next=0)
            if df['D+2추정예수금'][0] != '':
                if self.dict_bool['모의투자']:
                    con = sqlite3.connect(db_stg)
                    df = pd.read_sql('SELECT * FROM tradelist', con)
                    con.close()
                    self.dict_intg['예수금'] = \
                        100000000 - self.dict_df['잔고목록']['매입금액'].sum() + df['수익금'].sum()
                else:
                    self.dict_intg['예수금'] = int(df['D+2추정예수금'][0])
                self.dict_intg['추정예수금'] = self.dict_intg['예수금']
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 오류가 발생하여 계좌평가현황을 재조회합니다.'])
                time.sleep(3.35)

        while True:
            df = self.Block_Request('opw00018', 계좌번호=self.dict_strg['계좌번호'], 비밀번호='', 비밀번호입력매체구분='00',
                                    조회구분=2, output='계좌평가결과', next=0)
            if df['추정예탁자산'][0] != '':
                if self.dict_bool['모의투자']:
                    self.dict_intg['추정예탁자산'] = self.dict_intg['예수금'] + self.dict_df['잔고목록']['평가금액'].sum()
                else:
                    self.dict_intg['추정예탁자산'] = int(df['추정예탁자산'][0])

                self.dict_intg['종목당투자금'] = int(self.dict_intg['추정예탁자산'] * 0.99 / TUJAGMDIVIDE)
                self.stgQ.put(self.dict_intg['종목당투자금'])

                if self.dict_bool['모의투자']:
                    self.dict_df['잔고평가'].at[self.dict_strg['당일날짜']] = \
                        self.dict_intg['추정예탁자산'], self.dict_intg['예수금'], 0, 0, 0, 0, 0
                else:
                    tsp = float(int(df['총수익률(%)'][0]) / 100)
                    tsg = int(df['총평가손익금액'][0])
                    tbg = int(df['총매입금액'][0])
                    tpg = int(df['총평가금액'][0])
                    self.dict_df['잔고평가'].at[self.dict_strg['당일날짜']] = \
                        self.dict_intg['추정예탁자산'], self.dict_intg['예수금'], 0, tsp, tsg, tbg, tpg
                self.windowQ.put([ui_num['잔고평가'], self.dict_df['잔고평가']])
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 오류가 발생하여 계좌평가결과를 재조회합니다.'])
                time.sleep(3.35)

        if len(self.dict_df['거래목록']) > 0:
            self.UpdateTotaltradelist(first=True)

    def GetKospiKosdaqChart(self):
        self.dict_bool['업종차트조회'] = True
        self.windowQ.put([2, '코스피 코스닥 차트'])
        while True:
            df = self.Block_Request('opt20006', 업종코드='001', 기준일자=self.dict_strg['당일날짜'],
                                    output='업종일봉조회', next=0)
            if df['현재가'][0] != '':
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 오류가 발생하여 코스피 일봉차트를 재조회합니다.'])
                time.sleep(3.35)

        while True:
            df2 = self.Block_Request('opt20005', 업종코드='001', 틱범위='3', output='업종분봉조회', next=0)
            if df2['현재가'][0] != '':
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 오류가 발생하여 코스피 분봉차트를 재조회합니다.'])
                time.sleep(3.35)

        prec = abs(round(float(df['현재가'][1]) / 100, 2))
        self.chart6Q.put(['코스피종합', prec, df, ''])
        self.chart7Q.put(['코스피종합', prec, df2, ''])

        while True:
            df = self.Block_Request('opt20006', 업종코드='101', 기준일자=self.dict_strg['당일날짜'],
                                    output='업종일봉조회', next=0)
            if df['현재가'][0] != '':
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 오류가 발생하여 코스닥 일봉차트를 재조회합니다.'])
                time.sleep(3.35)

        while True:
            df2 = self.Block_Request('opt20005', 업종코드='101', 틱범위='3', output='업종분봉조회', next=0)
            if df2['현재가'][0] != '':
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 오류가 발생하여 코스닥 분봉차트를 재조회합니다.'])
                time.sleep(3.35)

        prec = abs(round(float(df['현재가'][1]) / 100, 2))
        self.chart8Q.put(['코스닥종합', prec, df, ''])
        self.chart9Q.put(['코스닥종합', prec, df2, ''])
        time.sleep(1)

    def OperationRealreg(self):
        self.dict_bool['장운영시간등록'] = True
        self.windowQ.put([2, '장운영시간 알림 등록'])
        self.traderQ.put([sn_oper, ' ', '215;20;214', 0])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 장운영시간 알림 등록 완료'])

    def UpjongjisuRealreg(self):
        self.dict_bool['업종지수등록'] = True
        self.windowQ.put([2, '업종지수 주식체결 등록'])
        self.traderQ.put([sn_oper, '001', '10;15;20', 1])
        self.traderQ.put([sn_oper, '101', '10;15;20', 1])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 업종지수 주식체결 등록 완료'])

    def ViRealreg(self):
        self.dict_bool['VI발동해제등록'] = True
        self.windowQ.put([2, 'VI발동해제 등록'])
        self.Block_Request('opt10054', 시장구분='000', 장전구분='1', 종목코드='', 발동구분='1', 제외종목='111111011',
                           거래량구분='0', 거래대금구분='0', 발동방향='0', output='발동종목', next=0)
        if self.dict_bool['알림소리']:
            self.soundQ.put('자동매매 시스템을 시작하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 시스템 시작 완료'])
        self.teleQ.put('시스템을 시작하였습니다.')

    def JangoChungsan(self):
        self.dict_bool['잔고청산'] = True
        self.windowQ.put([2, '잔고청산'])
        if len(self.dict_df['잔고목록']) > 0:
            for code in self.dict_df['잔고목록'].index:
                if code in self.list_sell:
                    continue
                c = self.dict_df['잔고목록']['현재가'][code]
                oc = self.dict_df['잔고목록']['보유수량'][code]
                name = self.dict_name[code]
                if self.dict_bool['모의투자']:
                    self.list_sell.append(code)
                    self.UpdateChejanData(code, name, '체결', '매도', c, c, oc, 0,
                                          strf_time('%Y%m%d%H%M%S%f'), strf_time('%Y%m%d%H%M%S%f'))
                else:
                    self.Order('매도', code, name, c, oc)
        if self.dict_bool['알림소리']:
            self.soundQ.put('잔고청산 주문을 전송하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 잔고청산 주문 완료'])

    def AllRemoveRealreg(self):
        self.windowQ.put([2, '실시간 데이터 수신 중단'])
        self.traderQ.put(['ALL', 'ALL'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('실시간 데이터의 수신을 중단하였습니다.')

    def SaveDatabase(self):
        if len(self.dict_df['거래목록']) > 0:
            df = self.dict_df['실현손익'][['총매수금액', '총매도금액', '총수익금액', '총손실금액', '수익률', '수익금합계']].copy()
            self.queryQ.put([1, df, 'totaltradelist', 'append'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('일별실현손익를 저장하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 일별실현손익 저장 완료'])

    @thread_decorator
    def PutHogaJanngo(self):
        if 0 in self.dict_hoga.keys() and self.dict_hoga[0][1]:
            self.windowQ.put([ui_num['호가잔고0'], self.dict_hoga[0][2]])
            self.dict_hoga[0][1] = False
        if 1 in self.dict_hoga.keys() and self.dict_hoga[1][1]:
            self.windowQ.put([ui_num['호가잔고1'], self.dict_hoga[1][2]])
            self.dict_hoga[1][1] = False

    @thread_decorator
    def UpdateTotaljango(self):
        if len(self.dict_df['잔고목록']) > 0:
            tsg = self.dict_df['잔고목록']['평가손익'].sum()
            tbg = self.dict_df['잔고목록']['매입금액'].sum()
            tpg = self.dict_df['잔고목록']['평가금액'].sum()
            bct = len(self.dict_df['잔고목록'])
            tsp = round(tsg / tbg * 100, 2)
            ttg = self.dict_intg['예수금'] + tpg
            self.dict_df['잔고평가'].at[self.dict_strg['당일날짜']] = \
                ttg, self.dict_intg['예수금'], bct, tsp, tsg, tbg, tpg
        else:
            self.dict_df['잔고평가'].at[self.dict_strg['당일날짜']] = \
                self.dict_intg['예수금'], self.dict_intg['예수금'], 0, 0.0, 0, 0, 0
        self.windowQ.put([ui_num['잔고목록'], self.dict_df['잔고목록']])
        self.windowQ.put([ui_num['잔고평가'], self.dict_df['잔고평가']])

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

    def OnEventConnect(self, err_code):
        if err_code == 0:
            self.dict_bool['로그인'] = True

    def OnReceiveTrData(self, screen, rqname, trcode, record, nnext):
        if screen == '' and record == '':
            return
        if 'ORD' in trcode:
            return

        items = None
        self.dict_bool['TR다음'] = True if nnext == '2' else False
        for output in self.dict_item['output']:
            record = list(output.keys())[0]
            items = list(output.values())[0]
            if record == self.dict_strg['TR명']:
                break
        rows = self.ocx.dynamicCall('GetRepeatCnt(QString, QString)', trcode, rqname)
        if rows == 0:
            rows = 1
        df2 = []
        for row in range(rows):
            row_data = []
            for item in items:
                data = self.ocx.dynamicCall('GetCommData(QString, QString, int, QString)', trcode, rqname, row, item)
                row_data.append(data.strip())
            df2.append(row_data)
        df = pd.DataFrame(data=df2, columns=items)
        self.dict_df['TRDF'] = df
        self.dict_bool['TR수신'] = True
        self.windowQ.put([1, f"조회 데이터 수신 완료 - {rqname} [{trcode}] {self.dict_strg['TR종목명']}"])

    def OnReceiveRealData(self, code, realtype, realdata):
        if realdata == '':
            return

        if realtype == '장시작시간':
            if self.dict_intg['장운영상태'] == 8:
                return
            try:
                self.dict_intg['장운영상태'] = int(self.GetCommRealData(code, 215))
                current = self.GetCommRealData(code, 20)
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData 장시작시간 {e}'])
            else:
                self.OperationAlert(current)
        elif realtype == '업종지수':
            if self.dict_bool['실시간데이터수신중단']:
                return
            try:
                c = abs(float(self.GetCommRealData(code, 10)))
                v = int(self.GetCommRealData(code, 15))
                d = self.GetCommRealData(code, 20)
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData 업종지수 {e}'])
            else:
                if code == '001':
                    self.chart6Q.put([d, c, v])
                    self.chart7Q.put([d, c, v])
                elif code == '101':
                    self.chart8Q.put([d, c, v])
                    self.chart9Q.put([d, c, v])
        elif realtype == 'VI발동/해제':
            if self.dict_bool['실시간데이터수신중단']:
                return
            try:
                code = self.GetCommRealData(code, 9001).strip('A').strip('Q')
                gubun = self.GetCommRealData(code, 9068)
                name = self.dict_name[code]
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData VI발동/해제 {e}'])
            else:
                if gubun == '1' and \
                        (code not in self.dict_vipr.keys() or
                         (self.dict_vipr[code][0] and now() > self.dict_vipr[code][1])):
                    self.UpdateViPrice(code, name)
        elif realtype == '주식체결':
            if self.dict_bool['실시간데이터수신중단']:
                return
            try:
                c = abs(int(self.GetCommRealData(code, 10)))
                o = abs(int(self.GetCommRealData(code, 16)))
                h = abs(int(self.GetCommRealData(code, 17)))
                low = abs(int(self.GetCommRealData(code, 18)))
                per = float(self.GetCommRealData(code, 12))
                v = int(self.GetCommRealData(code, 15))
                ch = float(self.GetCommRealData(code, 228))
                t = self.GetCommRealData(code, 20)
                name = self.dict_name[code]
                prec = self.GetMasterLastPrice(code)
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData 주식체결 {e}'])
            else:
                if self.dict_intg['장운영상태'] == 3:
                    if code not in self.dict_vipr.keys():
                        self.InsertViPrice(code, o)
                    elif not self.dict_vipr[code][0] and now() > self.dict_vipr[code][1]:
                        self.UpdateViPrice(code, c)
                    if code in self.dict_df['잔고목록'].index:
                        self.UpdateJango(code, name, c, o, h, low, per, ch)
                self.UpdateChartHoga(code, name, c, o, h, low, per, ch, v, t, prec)
        elif realtype == '주식호가잔량':
            if self.dict_bool['실시간데이터수신중단']:
                return
            if (0 in self.dict_hoga.keys() and code == self.dict_hoga[0][0]) or \
                    (1 in self.dict_hoga.keys() and code == self.dict_hoga[1][0]):
                try:
                    if code not in self.dict_sghg.keys():
                        Sanghanga, Hahanga = self.GetSangHahanga(code)
                        self.dict_sghg[code] = [Sanghanga, Hahanga]
                    else:
                        Sanghanga = self.dict_sghg[code][0]
                        Hahanga = self.dict_sghg[code][1]
                    prec = self.GetMasterLastPrice(code)
                    vp = [int(float(self.GetCommRealData(code, 139))),
                          int(self.GetCommRealData(code, 90)), int(self.GetCommRealData(code, 89)),
                          int(self.GetCommRealData(code, 88)), int(self.GetCommRealData(code, 87)),
                          int(self.GetCommRealData(code, 86)), int(self.GetCommRealData(code, 85)),
                          int(self.GetCommRealData(code, 84)), int(self.GetCommRealData(code, 83)),
                          int(self.GetCommRealData(code, 82)), int(self.GetCommRealData(code, 81)),
                          int(self.GetCommRealData(code, 91)), int(self.GetCommRealData(code, 92)),
                          int(self.GetCommRealData(code, 93)), int(self.GetCommRealData(code, 94)),
                          int(self.GetCommRealData(code, 95)), int(self.GetCommRealData(code, 96)),
                          int(self.GetCommRealData(code, 97)), int(self.GetCommRealData(code, 98)),
                          int(self.GetCommRealData(code, 99)), int(self.GetCommRealData(code, 100)),
                          int(float(self.GetCommRealData(code, 129)))]
                    jc = [int(self.GetCommRealData(code, 121)),
                          int(self.GetCommRealData(code, 70)), int(self.GetCommRealData(code, 69)),
                          int(self.GetCommRealData(code, 68)), int(self.GetCommRealData(code, 67)),
                          int(self.GetCommRealData(code, 66)), int(self.GetCommRealData(code, 65)),
                          int(self.GetCommRealData(code, 64)), int(self.GetCommRealData(code, 63)),
                          int(self.GetCommRealData(code, 62)), int(self.GetCommRealData(code, 61)),
                          int(self.GetCommRealData(code, 71)), int(self.GetCommRealData(code, 72)),
                          int(self.GetCommRealData(code, 73)), int(self.GetCommRealData(code, 74)),
                          int(self.GetCommRealData(code, 75)), int(self.GetCommRealData(code, 76)),
                          int(self.GetCommRealData(code, 77)), int(self.GetCommRealData(code, 78)),
                          int(self.GetCommRealData(code, 79)), int(self.GetCommRealData(code, 80)),
                          int(self.GetCommRealData(code, 125))]
                    hg = [Sanghanga,
                          abs(int(self.GetCommRealData(code, 50))), abs(int(self.GetCommRealData(code, 49))),
                          abs(int(self.GetCommRealData(code, 48))), abs(int(self.GetCommRealData(code, 47))),
                          abs(int(self.GetCommRealData(code, 46))), abs(int(self.GetCommRealData(code, 45))),
                          abs(int(self.GetCommRealData(code, 44))), abs(int(self.GetCommRealData(code, 43))),
                          abs(int(self.GetCommRealData(code, 42))), abs(int(self.GetCommRealData(code, 41))),
                          abs(int(self.GetCommRealData(code, 51))), abs(int(self.GetCommRealData(code, 52))),
                          abs(int(self.GetCommRealData(code, 53))), abs(int(self.GetCommRealData(code, 54))),
                          abs(int(self.GetCommRealData(code, 55))), abs(int(self.GetCommRealData(code, 56))),
                          abs(int(self.GetCommRealData(code, 57))), abs(int(self.GetCommRealData(code, 58))),
                          abs(int(self.GetCommRealData(code, 59))), abs(int(self.GetCommRealData(code, 60))),
                          Hahanga]
                    per = [round((hg[0] / prec - 1) * 100, 2), round((hg[1] / prec - 1) * 100, 2),
                           round((hg[2] / prec - 1) * 100, 2), round((hg[3] / prec - 1) * 100, 2),
                           round((hg[4] / prec - 1) * 100, 2), round((hg[5] / prec - 1) * 100, 2),
                           round((hg[6] / prec - 1) * 100, 2), round((hg[7] / prec - 1) * 100, 2),
                           round((hg[8] / prec - 1) * 100, 2), round((hg[9] / prec - 1) * 100, 2),
                           round((hg[10] / prec - 1) * 100, 2), round((hg[11] / prec - 1) * 100, 2),
                           round((hg[12] / prec - 1) * 100, 2), round((hg[13] / prec - 1) * 100, 2),
                           round((hg[14] / prec - 1) * 100, 2), round((hg[15] / prec - 1) * 100, 2),
                           round((hg[16] / prec - 1) * 100, 2), round((hg[17] / prec - 1) * 100, 2),
                           round((hg[18] / prec - 1) * 100, 2), round((hg[19] / prec - 1) * 100, 2),
                           round((hg[20] / prec - 1) * 100, 2), round((hg[21] / prec - 1) * 100, 2)]
                except Exception as e:
                    self.windowQ.put([1, f'OnReceiveRealData 주식호가잔량 {e}'])
                else:
                    self.UpdateHogajanryang(code, vp, jc, hg, per)

    @thread_decorator
    def OperationAlert(self, current):
        if self.dict_intg['장운영상태'] == 3:
            self.windowQ.put([2, '장운영상태'])
        if self.dict_bool['알림소리']:
            if current == '084000':
                self.soundQ.put('장시작 20분 전입니다.')
            elif current == '085000':
                self.soundQ.put('장시작 10분 전입니다.')
            elif current == '085500':
                self.soundQ.put('장시작 5분 전입니다.')
            elif current == '085900':
                self.soundQ.put('장시작 1분 전입니다.')
            elif current == '085930':
                self.soundQ.put('장시작 30초 전입니다.')
            elif current == '085940':
                self.soundQ.put('장시작 20초 전입니다.')
            elif current == '085950':
                self.soundQ.put('장시작 10초 전입니다.')
            elif current == '090000':
                self.soundQ.put(f"{self.dict_strg['당일날짜'][:4]}년 {self.dict_strg['당일날짜'][4:6]}월 "
                                f"{self.dict_strg['당일날짜'][6:]}일 장이 시작되었습니다.")
            elif current == '152000':
                self.soundQ.put('장마감 10분 전입니다.')
            elif current == '152500':
                self.soundQ.put('장마감 5분 전입니다.')
            elif current == '152900':
                self.soundQ.put('장마감 1분 전입니다.')
            elif current == '152930':
                self.soundQ.put('장마감 30초 전입니다.')
            elif current == '152940':
                self.soundQ.put('장마감 20초 전입니다.')
            elif current == '152950':
                self.soundQ.put('장마감 10초 전입니다.')
            elif current == '153000':
                self.soundQ.put(f"{self.dict_strg['당일날짜'][:4]}년 {self.dict_strg['당일날짜'][4:6]}월 "
                                f"{self.dict_strg['당일날짜'][6:]}일 장이 종료되었습니다.")

    def InsertViPrice(self, code, o):
        uvi, dvi, uvid5 = self.GetVIPrice(code, o)
        self.dict_vipr[code] = [True, timedelta_sec(-180), timedelta_sec(-180), uvi, dvi, uvid5]

    def GetVIPrice(self, code, std_price):
        uvi = std_price * 1.1
        x = self.GetHogaunit(code, uvi)
        if uvi % x != 0:
            uvi = uvi + (x - uvi % x)
        uvid5 = uvi - x * 5
        dvi = std_price * 0.9
        x = self.GetHogaunit(code, dvi)
        if dvi % x != 0:
            dvi = dvi - dvi % x
        return int(uvi), int(dvi), int(uvid5)

    def GetHogaunit(self, code, price):
        if price < 1000:
            x = 1
        elif 1000 <= price < 5000:
            x = 5
        elif 5000 <= price < 10000:
            x = 10
        elif 10000 <= price < 50000:
            x = 50
        elif code in self.list_kosd:
            x = 100
        elif 50000 <= price < 100000:
            x = 100
        elif 100000 <= price < 500000:
            x = 500
        else:
            x = 1000
        return x

    def UpdateViPrice(self, code, key):
        if type(key) == str:
            try:
                self.dict_vipr[code][:3] = False, timedelta_sec(5), timedelta_sec(180)
            except KeyError:
                self.dict_vipr[code] = [False, timedelta_sec(5), timedelta_sec(180), 0, 0, 0]
            self.traderQ.put([sn_vijc, code, '10;12;14;30;228', 1])
        elif type(key) == int:
            uvi, dvi, uvid5 = self.GetVIPrice(code, key)
            self.dict_vipr[code] = [True, now(), timedelta_sec(180), uvi, dvi, uvid5]
            self.traderQ.put([sn_vijc, code])

    def UpdateJango(self, code, name, c, o, h, low, per, ch):
        self.lock.acquire()
        prec = self.dict_df['잔고목록']['현재가'][code]
        if prec != c:
            bg = self.dict_df['잔고목록']['매입금액'][code]
            jc = int(self.dict_df['잔고목록']['보유수량'][code])
            pg, sg, sp = self.GetPgSgSp(bg, jc * c)
            columns = ['현재가', '수익률', '평가손익', '평가금액', '시가', '고가', '저가']
            self.dict_df['잔고목록'].at[code, columns] = c, sp, sg, pg, o, h, low
            self.stgQ.put([code, name, per, sp, jc, ch, c])
        self.lock.release()

    # noinspection PyMethodMayBeStatic
    def GetPgSgSp(self, bg, cg):
        gtexs = cg * 0.0023
        gsfee = cg * 0.00015
        gbfee = bg * 0.00015
        texs = gtexs - (gtexs % 1)
        sfee = gsfee - (gsfee % 10)
        bfee = gbfee - (gbfee % 10)
        pg = int(cg - texs - sfee - bfee)
        sg = pg - bg
        sp = round(sg / bg * 100, 2)
        return pg, sg, sp

    @thread_decorator
    def UpdateChartHoga(self, code, name, c, o, h, low, per, ch, v, t, prec):
        if ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]:
            self.chart1Q.put([code, t, c, per, ch])
            self.chart1Q.put([t, c, v])
            self.chart2Q.put([t, c, v])
        elif ui_num['차트P3'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P3']]:
            self.chart3Q.put([t, c, v])
            self.chart4Q.put([t, c, v])

        if 0 in self.dict_hoga.keys() and code == self.dict_hoga[0][0]:
            self.hoga1Q.put([v, ch])
            self.UpdateHogajango(0, code, name, c, o, h, low, prec)
        elif 1 in self.dict_hoga.keys() and code == self.dict_hoga[1][0]:
            self.hoga2Q.put([v, ch])
            self.UpdateHogajango(1, code, name, c, o, h, low, prec)

    def UpdateHogajango(self, gubun, code, name, c, o, h, low, prec):
        try:
            uvi, dvi = self.dict_vipr[code][3:5]
        except KeyError:
            uvi, dvi = 0, 0
        if code in self.dict_df['잔고목록'].index:
            df = self.dict_df['잔고목록'][self.dict_df['잔고목록'].index == code].copy()
            df['UVI'] = uvi
            df['DVI'] = dvi
            self.dict_hoga[gubun] = [code, True, df.rename(columns={'종목명': '호가종목명'})]
        else:
            df = pd.DataFrame([[name, 0, c, 0., 0, 0, 0, o, h, low, prec, 0, uvi, dvi]],
                              columns=columns_hj, index=[code])
            self.dict_hoga[gubun] = [code, True, df]

    def GetSangHahanga(self, code):
        predayclose = self.GetMasterLastPrice(code)
        uplimitprice = predayclose * 1.30
        x = self.GetHogaunit(code, uplimitprice)
        if uplimitprice % x != 0:
            uplimitprice -= uplimitprice % x
        downlimitprice = predayclose * 0.70
        x = self.GetHogaunit(code, downlimitprice)
        if downlimitprice % x != 0:
            downlimitprice += x - downlimitprice % x
        return int(uplimitprice), int(downlimitprice)

    @thread_decorator
    def UpdateHogajanryang(self, code, vp, jc, hg, per):
        per = [0 if p == -100 else p for p in per]
        og, op, omc = '', '', ''
        name = self.dict_name[code]
        cond = (self.dict_df['체결목록']['종목명'] == name) & (self.dict_df['체결목록']['미체결수량'] > 0)
        df = self.dict_df['체결목록'][cond]
        if len(df) > 0:
            og = df['주문구분'][0]
            op = df['주문가격'][0]
            omc = df['미체결수량'][0]
        if 0 in self.dict_hoga.keys() and code == self.dict_hoga[0][0]:
            self.hoga1Q.put([vp, jc, hg, per, og, op, omc])
        elif 1 in self.dict_hoga.keys() and code == self.dict_hoga[1][0]:
            self.hoga2Q.put([vp, jc, hg, per, og, op, omc])

    def OnReceiveChejanData(self, gubun, itemcnt, fidlist):
        if gubun != '0' and itemcnt != '' and fidlist != '':
            return
        if self.dict_bool['모의투자']:
            return
        on = self.GetChejanData(9203)
        if on == '':
            return

        try:
            code = self.GetChejanData(9001).strip('A')
            name = self.dict_name[code]
            ot = self.GetChejanData(913)
            og = self.GetChejanData(905)[1:]
            op = int(self.GetChejanData(901))
            oc = int(self.GetChejanData(900))
            omc = int(self.GetChejanData(902))
            dt = self.dict_strg['당일날짜'] + self.GetChejanData(908)
        except Exception as e:
            self.windowQ.put([1, f'OnReceiveChejanData {e}'])
        else:
            try:
                cp = int(self.GetChejanData(910))
            except ValueError:
                cp = 0
            self.UpdateChejanData(code, name, ot, og, op, cp, oc, omc, on, dt)

    @thread_decorator
    def UpdateChejanData(self, code, name, ot, og, op, cp, oc, omc, on, dt):
        self.lock.acquire()
        if ot == '체결' and omc == 0 and cp != 0:
            if og == '매수':
                self.dict_intg['예수금'] -= oc * cp
                self.dict_intg['추정예수금'] = self.dict_intg['예수금']
                self.UpdateChegeoljango(code, name, og, oc, cp)
                self.windowQ.put([1, f'매매 시스템 체결 알림 - {name} {oc}주 {og}'])
            elif og == '매도':
                bp = self.dict_df['잔고목록']['매입가'][code]
                bg = bp * oc
                pg, sg, sp = self.GetPgSgSp(bg, oc * cp)
                self.dict_intg['예수금'] += pg
                self.dict_intg['추정예수금'] = self.dict_intg['예수금']
                self.UpdateChegeoljango(code, name, og, oc, cp)
                self.UpdateTradelist(name, oc, sp, sg, bg, pg, on)
                self.windowQ.put([1, f"매매 시스템 체결 알림 - {name} {oc}주 {og}, 수익률 {sp}% 수익금{format(sg, ',')}원"])
        self.UpdateChegeollist(name, og, oc, omc, op, cp, dt, on)
        self.lock.release()

    def UpdateChegeoljango(self, code, name, og, oc, cp):
        columns = ['매입가', '현재가', '수익률', '평가손익', '매입금액', '평가금액', '보유수량']
        if og == '매수':
            if code not in self.dict_df['잔고목록'].index:
                bg = oc * cp
                pg, sg, sp = self.GetPgSgSp(bg, oc * cp)
                prec = self.GetMasterLastPrice(code)
                self.dict_df['잔고목록'].at[code] = name, cp, cp, sp, sg, bg, pg, 0, 0, 0, prec, oc
                self.receivQ.put(f'잔고편입 {code}')
                self.traderQ.put([sn_jscg, code, '10;12;14;30;228', 1])
            else:
                jc = self.dict_df['잔고목록']['보유수량'][code]
                bg = self.dict_df['잔고목록']['매입금액'][code]
                jc = jc + oc
                bg = bg + oc * cp
                bp = int(bg / jc)
                pg, sg, sp = self.GetPgSgSp(bg, jc * cp)
                self.dict_df['잔고목록'].at[code, columns] = bp, cp, sp, sg, bg, pg, jc
        elif og == '매도':
            jc = self.dict_df['잔고목록']['보유수량'][code]
            if jc - oc == 0:
                self.dict_df['잔고목록'].drop(index=code, inplace=True)
                self.receivQ.put(f'잔고청산 {code}')
                self.traderQ.put([sn_jscg, code])
            else:
                bp = self.dict_df['잔고목록']['매입가'][code]
                jc = jc - oc
                bg = jc * bp
                pg, sg, sp = self.GetPgSgSp(bg, jc * cp)
                self.dict_df['잔고목록'].at[code, columns] = bp, cp, sp, sg, bg, pg, jc

        if og == '매수':
            self.stgQ.put(['매수완료', code])
            self.list_buy.remove(code)
        elif og == '매도':
            self.stgQ.put(['매도완료', code])
            self.list_sell.remove(code)

        columns = ['매입가', '현재가', '평가손익', '매입금액']
        self.dict_df['잔고목록'][columns] = self.dict_df['잔고목록'][columns].astype(int)
        self.dict_df['잔고목록'].sort_values(by=['매입금액'], inplace=True)
        self.queryQ.put([1, self.dict_df['잔고목록'], 'jangolist', 'replace'])
        if self.dict_bool['알림소리']:
            self.soundQ.put(f'{name} {oc}주를 {og}하였습니다')

    def UpdateTradelist(self, name, oc, sp, sg, bg, pg, on):
        dt = strf_time('%Y%m%d%H%M%S%f')
        if self.dict_bool['모의투자'] and len(self.dict_df['거래목록']) > 0:
            if dt in self.dict_df['거래목록']['체결시간'].values:
                while dt in self.dict_df['거래목록']['체결시간'].values:
                    dt = str(int(dt) + 1)
                on = dt

        self.dict_df['거래목록'].at[on] = name, bg, pg, oc, sp, sg, dt
        self.dict_df['거래목록'].sort_values(by=['체결시간'], ascending=False, inplace=True)
        self.windowQ.put([ui_num['거래목록'], self.dict_df['거래목록']])

        df = pd.DataFrame([[name, bg, pg, oc, sp, sg, dt]], columns=columns_td, index=[on])
        self.queryQ.put([1, df, 'tradelist', 'append'])
        self.UpdateTotaltradelist()

    def UpdateTotaltradelist(self, first=False):
        tsg = self.dict_df['거래목록']['매도금액'].sum()
        tbg = self.dict_df['거래목록']['매수금액'].sum()
        tsig = self.dict_df['거래목록'][self.dict_df['거래목록']['수익금'] > 0]['수익금'].sum()
        tssg = self.dict_df['거래목록'][self.dict_df['거래목록']['수익금'] < 0]['수익금'].sum()
        sg = self.dict_df['거래목록']['수익금'].sum()
        sp = round(sg / self.dict_intg['추정예탁자산'] * 100, 2)
        tdct = len(self.dict_df['거래목록'])
        self.dict_df['실현손익'] = pd.DataFrame([[tdct, tbg, tsg, tsig, tssg, sp, sg]],
                                            columns=columns_tt, index=[self.dict_strg['당일날짜']])
        self.windowQ.put([ui_num['거래합계'], self.dict_df['실현손익']])

        if not first:
            self.teleQ.put(
                f"거래횟수 {len(self.dict_df['거래목록'])}회 / 총매수금액 {format(int(tbg), ',')}원 / "
                f"총매도금액 {format(int(tsg), ',')}원 / 총수익금액 {format(int(tsig), ',')}원 / "
                f"총손실금액 {format(int(tssg), ',')}원 / 수익률 {sp}% / 수익금합계 {format(int(sg), ',')}원")

    def UpdateChegeollist(self, name, og, oc, omc, op, cp, dt, on):
        if self.dict_bool['모의투자'] and len(self.dict_df['거래목록']) > 0:
            if dt in self.dict_df['거래목록']['체결시간'].values:
                while dt in self.dict_df['거래목록']['체결시간'].values:
                    dt = str(int(dt) + 1)
                on = dt

        if on in self.dict_df['체결목록'].index:
            self.dict_df['체결목록'].at[on, ['미체결수량', '체결가', '체결시간']] = omc, cp, dt
        else:
            self.dict_df['체결목록'].at[on] = name, og, oc, omc, op, cp, dt
        self.dict_df['체결목록'].sort_values(by=['체결시간'], ascending=False, inplace=True)
        self.windowQ.put([ui_num['체결목록'], self.dict_df['체결목록']])

        if omc == 0:
            df = pd.DataFrame([[name, og, oc, omc, op, cp, dt]], columns=columns_cj, index=[on])
            self.queryQ.put([1, df, 'chegeollist', 'append'])

    def OnReceiveConditionVer(self, ret, msg):
        if msg == '':
            return

        if ret == 1:
            self.dict_bool['CD수신'] = True

    def OnReceiveTrCondition(self, screen, code_list, cond_name, cond_index, nnext):
        if screen == "" and cond_name == "" and cond_index == "" and nnext == "":
            return

        codes = code_list.split(';')[:-1]
        self.list_trcd = codes
        self.dict_bool['CR수신'] = True

    def Block_Request(self, *args, **kwargs):
        if self.dict_intg['TR제한수신횟수'] == 0:
            self.dict_time['TR시작'] = now()
        if '종목코드' in kwargs.keys() and kwargs['종목코드'] != '':
            self.dict_strg['TR종목명'] = self.dict_name[kwargs['종목코드']]
        else:
            self.dict_strg['TR종목명'] = ''
        trcode = args[0].lower()
        liness = readEnc(trcode)
        self.dict_item = parseDat(trcode, liness)
        self.dict_strg['TR명'] = kwargs['output']
        nnext = kwargs['next']
        for i in kwargs:
            if i.lower() != 'output' and i.lower() != 'next':
                self.ocx.dynamicCall('SetInputValue(QString, QString)', i, kwargs[i])
        self.dict_bool['TR수신'] = False
        self.dict_bool['TR다음'] = False
        if trcode == 'optkwfid':
            code_list = args[1]
            code_count = args[2]
            self.ocx.dynamicCall('CommKwRqData(QString, bool, int, int, QString, QString)',
                                 code_list, 0, code_count, '0', self.dict_strg['TR명'], sn_brrq)
        elif trcode == 'opt10054':
            self.ocx.dynamicCall('CommRqData(QString, QString, int, QString)',
                                 self.dict_strg['TR명'], trcode, nnext, sn_brrd)
        else:
            self.ocx.dynamicCall('CommRqData(QString, QString, int, QString)',
                                 self.dict_strg['TR명'], trcode, nnext, sn_brrq)
        sleeptime = timedelta_sec(0.25)
        while not self.dict_bool['TR수신'] or now() < sleeptime:
            pythoncom.PumpWaitingMessages()
        if trcode != 'opt10054':
            self.DisconnectRealData(sn_brrq)
        self.UpdateTrtime()
        return self.dict_df['TRDF']

    @thread_decorator
    def UpdateTrtime(self):
        if self.dict_intg['TR제한수신횟수'] > 95:
            self.dict_time['TR재개'] = timedelta_sec(self.dict_intg['TR제한수신횟수'] * 3.35, self.dict_time['TR시작'])
            remaintime = (self.dict_time['TR재개'] - now()).total_seconds()
            if remaintime > 0:
                self.windowQ.put([1, f'시스템 명령 실행 알림 - TR 조회 재요청까지 남은 시간은 {round(remaintime, 2)}초입니다.'])
            self.dict_intg['TR제한수신횟수'] = 0

    @property
    def TrtimeCondition(self):
        return now() > self.dict_time['TR재개']

    @property
    def RemainedTrtime(self):
        return round((self.dict_time['TR재개'] - now()).total_seconds(), 2)

    def SendCondition(self, screen, cond_name, cond_index, search):
        self.dict_bool['CR수신'] = False
        self.ocx.dynamicCall('SendCondition(QString, QString, int, int)', screen, cond_name, cond_index, search)
        while not self.dict_bool['CR수신']:
            pythoncom.PumpWaitingMessages()
        return self.list_trcd

    def DisconnectRealData(self, screen):
        self.ocx.dynamicCall('DisconnectRealData(QString)', screen)

    def GetMasterCodeName(self, code):
        return self.ocx.dynamicCall('GetMasterCodeName(QString)', code)

    def GetCodeListByMarket(self, market):
        data = self.ocx.dynamicCall('GetCodeListByMarket(QString)', market)
        tokens = data.split(';')[:-1]
        return tokens

    def GetMasterLastPrice(self, code):
        return int(self.ocx.dynamicCall('GetMasterLastPrice(QString)', code))

    def GetCommRealData(self, code, fid):
        return self.ocx.dynamicCall('GetCommRealData(QString, int)', code, fid)

    def GetChejanData(self, fid):
        return self.ocx.dynamicCall('GetChejanData(int)', fid)

    def SysExit(self):
        self.windowQ.put([2, '시스템 종료'])
        self.teleQ.put('60초 후 시스템을 종료합니다.')
        if self.dict_bool['알림소리']:
            self.soundQ.put('60초 후 시스템을 종료합니다.')
        else:
            self.windowQ.put([1, '시스템 명령 실행 알림 - 60초 후 시스템을 종료합니다.'])
        i = 60
        while i > 0:
            if i <= 10:
                self.windowQ.put([1, f'시스템 명령 실행 알림 - 시스템 종료 카운터 {i}'])
            i -= 1
            time.sleep(1)
        self.windowQ.put([1, '시스템 명령 실행 알림 - 시스템 종료'])
