import os
import sys
import time
import psutil
import pythoncom
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from threading import Timer
from PyQt5.QAxContainer import QAxWidget
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import *
from utility.setting import *

TUJAGMDIVIDE = 10    # 종목당 투자금 분할계수


class Updater(QtCore.QThread):
    data1 = QtCore.pyqtSignal(list)
    data2 = QtCore.pyqtSignal(str)

    def __init__(self, traderQ):
        super().__init__()
        self.traderQ = traderQ

    def run(self):
        while True:
            data = self.traderQ.get()
            if type(data) == list:
                self.data1.emit(data)
            elif type(data) == str:
                self.data2.emit(data)


class Trader:
    def __init__(self, qlist):
        app = QtWidgets.QApplication(sys.argv)
        """
           0        1        2      3      4       5       6      7       8        9       10       11
        windowQ, traderQ, receivQ, stgQ, soundQ, queryQ, teleQ, hoga1Q, hoga2Q, chart1Q, chart2Q, chart3Q,
        chart4Q, chart5Q, chart6Q, chart7Q, chart8Q, chart9Q, chart10Q, tick1Q, tick2Q, tick3Q, tick4Q
          12       13       14       15       16       17       18        19      20      21      22
        """
        self.windowQ = qlist[0]
        self.traderQ = qlist[1]
        self.receivQ = qlist[2]
        self.stgQ = qlist[3]
        self.soundQ = qlist[4]
        self.queryQ = qlist[5]
        self.teleQ = qlist[6]
        self.hoga1Q = qlist[7]
        self.hoga2Q = qlist[8]
        self.chart1Q = qlist[9]
        self.chart2Q = qlist[10]
        self.chart3Q = qlist[11]
        self.chart4Q = qlist[12]
        self.chart5Q = qlist[13]
        self.chart6Q = qlist[14]
        self.chart7Q = qlist[15]
        self.chart8Q = qlist[16]
        self.chart9Q = qlist[17]

        self.df_cj = pd.DataFrame(columns=columns_cj)   # 체결목록
        self.df_jg = pd.DataFrame(columns=columns_jg)   # 잔고목록
        self.df_tj = pd.DataFrame(columns=columns_tj)   # 잔고평가
        self.df_td = pd.DataFrame(columns=columns_td)   # 거래목록
        self.df_tt = pd.DataFrame(columns=columns_tt)   # 실현손익
        self.df_tr = None

        self.dict_sghg = {}     # key: 종목코드, value: [상한가, 하한가]
        self.dict_hoga = {}     # key: 호가창번호, value: [종목코드, 갱신여부, 호가잔고(DataFrame)]
        self.dict_chat = {}     # key: UI번호, value: 종목코드

        self.dict_name = {}     # key: 종목코드, value: 종목명
        self.dict_vipr = {}     # key: 종목코드, value: [갱신여부, 발동시간+5초, UVI, DVI, UVID5]
        self.dict_buyt = {}     # key: 종목코드, value: 매수시간
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
            'TR명': ''
        }
        self.dict_bool = {
            '데이터베이스로딩': False,
            '계좌잔고조회': False,
            '장운영시간등록': False,
            '업종차트조회': False,
            '업종지수등록': False,
            '장초전략잔고청산': False,
            '장중전략잔고청산': False,
            '실시간데이터수신중단': False,
            '당일거래목록저장': False,

            '테스트': False,
            '모의투자': False,
            '알림소리': False,

            '로그인': False,
            'TR수신': False,
            'TR다음': False
        }
        remaintime = (strp_time('%Y%m%d%H%M%S', self.dict_strg['당일날짜'] + '090100') - now()).total_seconds()
        self.dict_time = {
            '휴무종료': timedelta_sec(remaintime) if remaintime > 0 else timedelta_sec(600),
            '호가정보': now(),
            '거래정보': now(),
            '부가정보': now(),
            'TR시작': now(),
            'TR재개': now()
        }
        self.dict_item = None
        self.list_kosd = None
        self.list_buy = []
        self.list_sell = []

        self.ocx = QAxWidget('KHOPENAPI.KHOpenAPICtrl.1')
        self.ocx.OnEventConnect.connect(self.OnEventConnect)
        self.ocx.OnReceiveTrData.connect(self.OnReceiveTrData)
        self.ocx.OnReceiveRealData.connect(self.OnReceiveRealData)
        self.ocx.OnReceiveChejanData.connect(self.OnReceiveChejanData)

        self.LoadDatabase()
        self.CommConnect()

        self.updater = Updater(self.traderQ)
        self.updater.data1.connect(self.UpdateList)
        self.updater.data2.connect(self.UpdateStr)
        self.updater.start()

        self.qtimer = QtCore.QTimer()
        self.qtimer.setInterval(1000)
        self.qtimer.timeout.connect(self.Scheduler)
        self.qtimer.start()

        app.exec_()

    def LoadDatabase(self):
        self.dict_bool['데이터베이스로딩'] = True
        self.windowQ.put([2, '데이터베이스 로딩'])
        con = sqlite3.connect(DB_STG)
        df = pd.read_sql('SELECT * FROM setting', con)
        df = df.set_index('index')
        self.dict_bool['테스트'] = df['테스트'][0]
        self.dict_bool['모의투자'] = df['모의투자'][0]
        self.dict_bool['알림소리'] = df['알림소리'][0]
        self.windowQ.put([2, f"테스트모드 {self.dict_bool['테스트']}"])
        self.windowQ.put([2, f"모의투자 {self.dict_bool['모의투자']}"])
        self.windowQ.put([2, f"알림소리 {self.dict_bool['알림소리']}"])

        df = pd.read_sql(f"SELECT * FROM chegeollist WHERE 체결시간 LIKE '{self.dict_strg['당일날짜']}%'", con)
        self.df_cj = df.set_index('index').sort_values(by=['체결시간'], ascending=False)

        df = pd.read_sql(f"SELECT * FROM tradelist WHERE 체결시간 LIKE '{self.dict_strg['당일날짜']}%'", con)
        self.df_td = df.set_index('index').sort_values(by=['체결시간'], ascending=False)

        df = pd.read_sql(f'SELECT * FROM jangolist', con)
        self.df_jg = df.set_index('index').sort_values(by=['매입금액'], ascending=False)

        if len(self.df_cj) > 0:
            self.windowQ.put([ui_num['체결목록'], self.df_cj])
        if len(self.df_td) > 0:
            self.windowQ.put([ui_num['거래목록'], self.df_td])

        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 DB 정보 로딩 완료'])

    def CommConnect(self):
        self.windowQ.put([2, '트레이더 OPENAPI 로그인'])
        self.ocx.dynamicCall('CommConnect()')
        while not self.dict_bool['로그인']:
            pythoncom.PumpWaitingMessages()

        self.dict_strg['계좌번호'] = self.ocx.dynamicCall('GetLoginInfo(QString)', 'ACCNO').split(';')[0]

        self.list_kosd = self.GetCodeListByMarket('10')
        list_code = self.GetCodeListByMarket('0') + self.list_kosd
        for code in list_code:
            name = self.GetMasterCodeName(code)
            self.dict_name[code] = name

        self.chart9Q.put(self.dict_name)
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 OpenAPI 로그인 완료'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('키움증권 오픈에이피아이에 로그인하였습니다.')

        if len(self.df_jg) > 0:
            for code in self.df_jg.index:
                cond = (self.df_cj['주문구분'] == '매수') & (self.df_cj['종목명'] == self.dict_name[code])
                df = self.df_cj[cond]
                if len(df) > 0:
                    self.dict_buyt[code] = strp_time('%Y%m%d%H%M%S%f', df['체결시간'].iloc[0])
                else:
                    self.dict_buyt[code] = now()
                self.receivQ.put(f'잔고편입 {code}')

        if int(strf_time('%H%M%S')) > 90000:
            self.windowQ.put([2, '장운영상태'])
            self.dict_intg['장운영상태'] = 3

    def UpdateList(self, data):
        if len(data) == 5:
            self.BuySell(data[0], data[1], data[2], data[3], data[4])
        elif len(data) == 6:
            self.UpdateJango(data[0], data[1], data[2], data[3], data[4], data[5])
        elif len(data) == 2:
            self.dict_vipr = data[1]

    def BuySell(self, gubun, code, name, c, oc):
        if gubun == '매수':
            if code in self.df_jg.index:
                self.stgQ.put(['매수취소', code])
                return
            if code in self.list_buy:
                self.stgQ.put(['매수취소', code])
                self.windowQ.put([1, '매매 시스템 오류 알림 - 현재 매도 주문중인 종목입니다.'])
                return
            if self.dict_intg['추정예수금'] < oc * c:
                cond = (self.df_cj['주문구분'] == '시드부족') & (self.df_cj['종목명'] == name)
                df = self.df_cj[cond]
                if len(df) == 0 or \
                        (len(df) > 0 and now() > timedelta_sec(180, strp_time('%Y%m%d%H%M%S%f', df['체결시간'][0]))):
                    self.Order('시드부족', code, name, c, oc)
                self.stgQ.put(['매수취소', code])
                return
        elif gubun == '매도':
            if code not in self.df_jg.index:
                self.stgQ.put(['매도취소', code])
                return
            if code in self.list_sell:
                self.stgQ.put(['매도취소', code])
                self.windowQ.put([1, '매매 시스템 오류 알림 - 현재 매수 주문중인 종목입니다.'])
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
            self.UpdateChejanData(code, name, '체결', gubun, c, c, oc, 0, strf_time('%Y%m%d%H%M%S%f'))
        else:
            self.SendOrder([gubun, '4989', self.dict_strg['계좌번호'], on, code, oc, 0, '03', '', name])

    def SendOrder(self, order):
        name = order[-1]
        del order[-1]
        ret = self.ocx.dynamicCall(
            'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)', order)
        if ret != 0:
            self.windowQ.put([1, f'매매 시스템 오류 알림 - {name} {order[5]}주 {order[0]} 주문 실패'])
        sleeptime = timedelta_sec(0.21)
        while now() < sleeptime:
            pythoncom.PumpWaitingMessages()

    def UpdateJango(self, code, name, c, o, h, low):
        try:
            prec = self.df_jg['현재가'][code]
        except KeyError:
            return
        if prec != c:
            bg = self.df_jg['매입금액'][code]
            oc = int(self.df_jg['보유수량'][code])
            pg, sg, sp = self.GetPgSgSp(bg, oc * c)
            columns = ['현재가', '수익률', '평가손익', '평가금액', '시가', '고가', '저가']
            self.df_jg.at[code, columns] = c, sp, sg, pg, o, h, low
            if code in self.dict_buyt.keys():
                self.stgQ.put([code, name, sp, oc, c, self.dict_buyt[code]])

    def UpdateStr(self, data):
        if '현재가' in data:
            gubun = int(data.split(' ')[0][3:5])
            code = data.split(' ')[-1]
            name = self.dict_name[code]
            if gubun == ui_num['차트P0']:
                gubun = ui_num['차트P1']
                if ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]:
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 트레이더 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[data]).start()
                    return
                self.chart6Q.put('기업개요 ' + code)
                self.chart7Q.put('기업공시 ' + code)
                self.chart8Q.put('종목뉴스 ' + code)
                self.chart9Q.put('재무제표 ' + code)
                self.hoga1Q.put('초기화')
                self.SetRealReg([sn_cthg, code, '10;12;14;30;228;41;61;71;81', 1])
                if 0 in self.dict_hoga.keys():
                    self.SetRealRemove([sn_cthg, self.dict_hoga[0][0]])
                self.dict_hoga[0] = [code, True, pd.DataFrame(columns=columns_hj)]
                self.GetChart(gubun, code, name)
                self.GetTujajaChegeolH(code)
            elif gubun == ui_num['차트P1']:
                if (ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]) or \
                        (ui_num['차트P3'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P3']]):
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 트레이더 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[data]).start()
                    return
                self.hoga1Q.put('초기화')
                self.SetRealReg([sn_cthg, code, '10;12;14;30;228;41;61;71;81', 1])
                if 0 in self.dict_hoga.keys():
                    self.SetRealRemove([sn_cthg, self.dict_hoga[0][0]])
                self.dict_hoga[0] = [code, True, pd.DataFrame(columns=columns_hj)]
                self.GetChart(gubun, code, name)
            elif gubun == ui_num['차트P3']:
                if (ui_num['차트P1'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P1']]) or \
                        (ui_num['차트P3'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P3']]):
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 트레이더 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[data]).start()
                    return
                self.hoga2Q.put('초기화')
                self.SetRealReg([sn_cthg, code, '10;12;14;30;228;41;61;71;81', 1])
                if 1 in self.dict_hoga.keys():
                    self.SetRealRemove([sn_cthg, self.dict_hoga[1][0]])
                self.dict_hoga[1] = [code, True, pd.DataFrame(columns=columns_hj)]
                self.GetChart(gubun, code, name)
            elif gubun == ui_num['차트P5']:
                tradeday = data.split(' ')[-2]
                if ui_num['차트P5'] in self.dict_chat.keys() and code == self.dict_chat[ui_num['차트P5']]:
                    return
                if int(tradeday) < int(strf_time('%Y%m%d', timedelta_day(-5))):
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 트레이더 5일 이전의 체결정보는 조회할 수 없습니다.'])
                    return
                if not self.TrtimeCondition:
                    self.windowQ.put([1, f'시스템 명령 오류 알림 - 트레이더 해당 명령은 {self.RemainedTrtime}초 후에 실행됩니다.'])
                    Timer(self.RemainedTrtime, self.traderQ.put, args=[data]).start()
                    return
                self.GetChart(gubun, code, name, tradeday)
        elif '매수취소' in data:
            code = data.split(' ')[1]
            name = self.dict_name[code]
            term = (self.df_cj['종목명'] == name) & (self.df_cj['미체결수량'] > 0) & (self.df_cj['주문구분'] == '매수')
            df = self.df_cj[term]
            if len(df) == 1:
                on = df.index[0]
                omc = df['미체결수량'][on]
                order = ['매수취소', '4989', self.dict_strg['계좌번호'], 3, code, omc, 0, '00', on, name]
                self.SendOrder(order)
        elif '매도취소' in data:
            code = data.split(' ')[1]
            name = self.dict_name[code]
            term = (self.df_cj['종목명'] == name) & (self.df_cj['미체결수량'] > 0) & (self.df_cj['주문구분'] == '매도')
            df = self.df_cj[term]
            if len(df) == 1:
                on = df.index[0]
                omc = df['미체결수량'][on]
                order = ['매도취소', '4989', self.dict_strg['계좌번호'], 4, code, omc, 0, '00', on, name]
                self.SendOrder(order)
        elif data == '데이터베이스 로딩':
            if not self.dict_bool['데이터베이스로딩']:
                self.LoadDatabase()
        elif data == 'OPENAPI 로그인':
            if self.ocx.dynamicCall('GetConnectState()') == 0:
                self.CommConnect()
        elif data == '계좌평가 및 잔고':
            if not self.dict_bool['계좌잔고조회']:
                self.GetAccountjanGo()
        elif data == '코스피 코스닥 차트':
            if not self.dict_bool['업종차트조회']:
                self.GetKospiKosdaqChart()
        elif data == '장운영시간 알림 등록':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 리시버에서 자동실행됩니다.'])
        elif data == '업종지수 주식체결 등록':
            if not self.dict_bool['업종지수등록']:
                self.UpjongjisuRealreg()
        elif data == 'VI발동해제 등록':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 리시버에서 자동실행됩니다.'])
        elif data == '장운영상태':
            if self.dict_intg['장운영상태'] != 3:
                self.windowQ.put([2, '장운영상태'])
                self.dict_intg['장운영상태'] = 3
        elif data == '실시간 조건검색식 등록':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 리시버에서 자동실행됩니다.'])
        elif data == '장초전략 잔고청산':
            if not self.dict_bool['장초전략잔고청산']:
                self.JangoChungsan1()
        elif data == '실시간 조건검색식 중단':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 리시버에서 자동실행됩니다.'])
        elif data == '장중전략 시작':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 리시버에서 자동실행됩니다.'])
        elif data == '장중전략 잔고청산':
            if not self.dict_bool['장중전략잔고청산']:
                self.JangoChungsan2()
        elif data == '실시간 데이터 수신 중단':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 리시버에서 자동실행됩니다.'])
        elif data == '당일거래목록 저장':
            if not self.dict_bool['당일거래목록저장']:
                self.SaveDayData()
        elif data == '틱데이터 저장':
            self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 해당 명령은 콜렉터에서 자동실행됩니다.'])
        elif data == '시스템 종료':
            self.SysExit()
        elif data == '/당일체결목록':
            if len(self.df_cj) > 0:
                self.teleQ.put(self.df_cj)
            else:
                self.teleQ.put('현재는 거래목록이 없습니다.')
        elif data == '/당일거래목록':
            if len(self.df_td) > 0:
                self.teleQ.put(self.df_td)
            else:
                self.teleQ.put('현재는 거래목록이 없습니다.')
        elif data == '/계좌잔고평가':
            if len(self.df_jg) > 0:
                self.teleQ.put(self.df_jg)
            else:
                self.teleQ.put('현재는 잔고목록이 없습니다.')
        elif data == '/잔고청산주문':
            if not self.dict_bool['장초전략잔고청산']:
                self.JangoChungsan1()
            elif not self.dict_bool['장중전략잔고청산']:
                self.JangoChungsan2()
        elif '설정' in data:
            bot_number = data.split(' ')[1]
            chat_id = int(data.split(' ')[2])
            self.queryQ.put([1, f"UPDATE telegram SET str_bot = '{bot_number}', int_id = '{chat_id}'"])
            if self.dict_bool['알림소리']:
                self.soundQ.put('텔레그램 봇넘버 및 아이디가 변경되었습니다.')
            else:
                self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 텔레그램 봇넘버 및 아이디 설정 완료'])
        elif data == '테스트모드 ON/OFF':
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
        elif data == '모의투자 ON/OFF':
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
        elif data == '알림소리 ON/OFF':
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
        elif data == '틱데이터저장완료':
            self.SysExit()

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

    def Scheduler(self):
        if not self.dict_bool['계좌잔고조회']:
            self.GetAccountjanGo()
        if not self.dict_bool['업종차트조회']:
            self.GetKospiKosdaqChart()
        if not self.dict_bool['장운영시간등록']:
            self.OperationRealreg()
        if not self.dict_bool['업종지수등록']:
            self.UpjongjisuRealreg()
        if self.dict_intg['장운영상태'] == 1 and now() > self.dict_time['휴무종료']:
            self.SysExit()
        if int(strf_time('%H%M%S')) >= 100000 and not self.dict_bool['장초전략잔고청산']:
            self.JangoChungsan1()
        if int(strf_time('%H%M%S')) >= 152900 and not self.dict_bool['장중전략잔고청산']:
            self.JangoChungsan2()
        if self.dict_intg['장운영상태'] == 8 and not self.dict_bool['실시간데이터수신중단']:
            self.RemoveAllRealreg()
        if self.dict_intg['장운영상태'] == 8 and not self.dict_bool['당일거래목록저장']:
            self.SaveDayData()

        if now() > self.dict_time['호가정보']:
            self.PutHogaJanngo()
            self.dict_time['호가정보'] = timedelta_sec(0.25)
        if now() > self.dict_time['거래정보']:
            self.UpdateTotaljango()
            self.dict_time['거래정보'] = timedelta_sec(1)
        if now() > self.dict_time['부가정보']:
            self.UpdateInfo()
            self.dict_time['부가정보'] = timedelta_sec(2)

    def GetAccountjanGo(self):
        self.dict_bool['계좌잔고조회'] = True
        self.windowQ.put([2, '계좌평가 및 잔고'])

        while True:
            df = self.Block_Request('opw00004', 계좌번호=self.dict_strg['계좌번호'], 비밀번호='', 상장폐지조회구분=0,
                                    비밀번호입력매체구분='00', output='계좌평가현황', next=0)
            if df['D+2추정예수금'][0] != '':
                if self.dict_bool['모의투자']:
                    con = sqlite3.connect(DB_STG)
                    df = pd.read_sql('SELECT * FROM tradelist', con)
                    con.close()
                    self.dict_intg['예수금'] = 100000000 - self.df_jg['매입금액'].sum() + df['수익금'].sum()
                else:
                    self.dict_intg['예수금'] = int(df['D+2추정예수금'][0])
                self.dict_intg['추정예수금'] = self.dict_intg['예수금']
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 오류가 발생하여 계좌평가현황을 재조회합니다.'])
                time.sleep(3.35)

        while True:
            df = self.Block_Request('opw00018', 계좌번호=self.dict_strg['계좌번호'], 비밀번호='', 비밀번호입력매체구분='00',
                                    조회구분=2, output='계좌평가결과', next=0)
            if df['추정예탁자산'][0] != '':
                if self.dict_bool['모의투자']:
                    self.dict_intg['추정예탁자산'] = self.dict_intg['예수금'] + self.df_jg['평가금액'].sum()
                else:
                    self.dict_intg['추정예탁자산'] = int(df['추정예탁자산'][0])

                self.dict_intg['종목당투자금'] = int(self.dict_intg['추정예탁자산'] * 0.99 / TUJAGMDIVIDE)
                self.stgQ.put(self.dict_intg['종목당투자금'])

                if self.dict_bool['모의투자']:
                    self.df_tj.at[self.dict_strg['당일날짜']] = \
                        self.dict_intg['추정예탁자산'], self.dict_intg['예수금'], 0, 0, 0, 0, 0
                else:
                    tsp = float(int(df['총수익률(%)'][0]) / 100)
                    tsg = int(df['총평가손익금액'][0])
                    tbg = int(df['총매입금액'][0])
                    tpg = int(df['총평가금액'][0])
                    self.df_tj.at[self.dict_strg['당일날짜']] = \
                        self.dict_intg['추정예탁자산'], self.dict_intg['예수금'], 0, tsp, tsg, tbg, tpg
                self.windowQ.put([ui_num['잔고평가'], self.df_tj])
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 오류가 발생하여 계좌평가결과를 재조회합니다.'])
                time.sleep(3.35)

        if len(self.df_td) > 0:
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
                self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 오류가 발생하여 코스피 일봉차트를 재조회합니다.'])
                time.sleep(3.35)

        while True:
            df2 = self.Block_Request('opt20005', 업종코드='001', 틱범위='3', output='업종분봉조회', next=0)
            if df2['현재가'][0] != '':
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 오류가 발생하여 코스피 분봉차트를 재조회합니다.'])
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
                self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 오류가 발생하여 코스닥 일봉차트를 재조회합니다.'])
                time.sleep(3.35)

        while True:
            df2 = self.Block_Request('opt20005', 업종코드='101', 틱범위='3', output='업종분봉조회', next=0)
            if df2['현재가'][0] != '':
                break
            else:
                self.windowQ.put([1, '시스템 명령 오류 알림 - 트레이더 오류가 발생하여 코스닥 분봉차트를 재조회합니다.'])
                time.sleep(3.35)

        prec = abs(round(float(df['현재가'][1]) / 100, 2))
        self.chart8Q.put(['코스닥종합', prec, df, ''])
        self.chart9Q.put(['코스닥종합', prec, df2, ''])
        time.sleep(1)

    def OperationRealreg(self):
        self.dict_bool['장운영시간등록'] = True
        self.SetRealReg([sn_oper, ' ', '215;20;214', 0])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 장운영시간 등록 완료'])

    def UpjongjisuRealreg(self):
        self.dict_bool['업종지수등록'] = True
        self.windowQ.put([2, '업종지수 주식체결 등록'])
        self.SetRealReg([sn_oper, '001', '10;15;20', 1])
        self.SetRealReg([sn_oper, '101', '10;15;20', 1])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 업종지수 주식체결 등록 완료'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('자동매매 시스템을 시작하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 시작 완료'])
        self.teleQ.put('시스템을 시작하였습니다.')

    def JangoChungsan1(self):
        self.dict_bool['장초전략잔고청산'] = True
        self.windowQ.put([2, '장초전략 잔고청산'])
        if len(self.df_jg) > 0:
            for code in self.df_jg.index:
                if code in self.list_sell:
                    continue
                c = self.df_jg['현재가'][code]
                oc = self.df_jg['보유수량'][code]
                name = self.dict_name[code]
                if self.dict_bool['모의투자']:
                    self.list_sell.append(code)
                    self.UpdateChejanData(code, name, '체결', '매도', c, c, oc, 0, strf_time('%Y%m%d%H%M%S%f'))
                else:
                    self.Order('매도', code, name, c, oc)
        if self.dict_bool['알림소리']:
            self.soundQ.put('장초전략 잔고청산 주문을 전송하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 장초전략 잔고청산 주문 완료'])

    def JangoChungsan2(self):
        self.dict_bool['장중전략잔고청산'] = True
        self.windowQ.put([2, '장중전략 잔고청산'])
        if len(self.df_jg) > 0:
            for code in self.df_jg.index:
                if code in self.list_sell:
                    continue
                c = self.df_jg['현재가'][code]
                oc = self.df_jg['보유수량'][code]
                name = self.dict_name[code]
                if self.dict_bool['모의투자']:
                    self.list_sell.append(code)
                    self.UpdateChejanData(code, name, '체결', '매도', c, c, oc, 0, strf_time('%Y%m%d%H%M%S%f'))
                else:
                    self.Order('매도', code, name, c, oc)
        if self.dict_bool['알림소리']:
            self.soundQ.put('장중전략 잔고청산 주문을 전송하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 장중전략 잔고청산 주문 완료'])

    def RemoveAllRealreg(self):
        self.dict_bool['실시간데이터수신중단'] = True
        self.SetRealRemove(['ALL', 'ALL'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('실시간 데이터의 수신을 중단하였습니다.')

    def SaveDayData(self):
        self.dict_bool['당일거래목록저장'] = True
        self.windowQ.put([2, '당일거래목록 저장'])
        if len(self.df_td) > 0:
            df = self.df_tt[['총매수금액', '총매도금액', '총수익금액', '총손실금액', '수익률', '수익금합계']].copy()
            self.queryQ.put([1, df, 'totaltradelist', 'append'])
        if self.dict_bool['알림소리']:
            self.soundQ.put('일별실현손익를 저장하였습니다.')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 일별실현손익 저장 완료'])

    @thread_decorator
    def PutHogaJanngo(self):
        if 0 in self.dict_hoga.keys() and self.dict_hoga[0][1]:
            self.windowQ.put([ui_num['호가잔고0'], self.dict_hoga[0][2]])
            self.dict_hoga[0][1] = False
        if 1 in self.dict_hoga.keys() and self.dict_hoga[1][1]:
            self.windowQ.put([ui_num['호가잔고1'], self.dict_hoga[1][2]])
            self.dict_hoga[1][1] = False

    def UpdateTotaljango(self):
        if len(self.df_jg) > 0:
            tsg = self.df_jg['평가손익'].sum()
            tbg = self.df_jg['매입금액'].sum()
            tpg = self.df_jg['평가금액'].sum()
            bct = len(self.df_jg)
            tsp = round(tsg / tbg * 100, 2)
            ttg = self.dict_intg['예수금'] + tpg
            self.df_tj.at[self.dict_strg['당일날짜']] = \
                ttg, self.dict_intg['예수금'], bct, tsp, tsg, tbg, tpg
        else:
            self.df_tj.at[self.dict_strg['당일날짜']] = \
                self.dict_intg['예수금'], self.dict_intg['예수금'], 0, 0.0, 0, 0, 0
        self.windowQ.put([ui_num['잔고목록'], self.df_jg])
        self.windowQ.put([ui_num['잔고평가'], self.df_tj])

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
        self.df_tr = df
        self.dict_bool['TR수신'] = True

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
            uvi, dvi = self.dict_vipr[code][2:4]
        except KeyError:
            uvi, dvi = 0, 0
        if code in self.df_jg.index:
            df = self.df_jg[self.df_jg.index == code].copy()
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

    @thread_decorator
    def UpdateHogajanryang(self, code, vp, jc, hg, per):
        per = [0 if p == -100 else p for p in per]
        og, op, omc = '', '', ''
        name = self.dict_name[code]
        cond = (self.df_cj['종목명'] == name) & (self.df_cj['미체결수량'] > 0)
        df = self.df_cj[cond]
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
        except Exception as e:
            self.windowQ.put([1, f'OnReceiveChejanData {e}'])
        else:
            try:
                cp = int(self.GetChejanData(910))
            except ValueError:
                cp = 0
            self.UpdateChejanData(code, name, ot, og, op, cp, oc, omc, on)

    def UpdateChejanData(self, code, name, ot, og, op, cp, oc, omc, on):
        if ot == '체결' and omc == 0 and cp != 0:
            if og == '매수':
                self.UpdateChegeoljango(code, name, og, oc, cp)
                self.dict_buyt[code] = now()
                self.list_buy.remove(code)
                self.receivQ.put(f'잔고편입 {code}')
                self.stgQ.put(['매수완료', code])
                self.dict_intg['예수금'] -= oc * cp
                self.dict_intg['추정예수금'] = self.dict_intg['예수금']
                self.windowQ.put([1, f'매매 시스템 체결 알림 - {name} {oc}주 {og}'])
            elif og == '매도':
                bp = self.df_jg['매입가'][code]
                bg = bp * oc
                pg, sg, sp = self.GetPgSgSp(bg, oc * cp)
                self.UpdateChegeoljango(code, name, og, oc, cp)
                self.UpdateTradelist(name, oc, sp, sg, bg, pg, on)
                self.list_sell.remove(code)
                self.receivQ.put(f'잔고청산 {code}')
                self.dict_intg['종목당투자금'] = \
                    int(self.df_tj['추정예탁자산'][self.dict_strg['당일날짜']] * 0.99 / TUJAGMDIVIDE)
                self.stgQ.put(self.dict_intg['종목당투자금'])
                self.stgQ.put(['매도완료', code])
                self.dict_intg['예수금'] += pg
                self.dict_intg['추정예수금'] = self.dict_intg['예수금']
                self.windowQ.put([1, f"매매 시스템 체결 알림 - {name} {oc}주 {og}, 수익률 {sp}% 수익금{format(sg, ',')}원"])
        self.UpdateChegeollist(name, og, oc, omc, op, cp, on)

    def UpdateChegeoljango(self, code, name, og, oc, cp):
        columns = ['매입가', '현재가', '수익률', '평가손익', '매입금액', '평가금액', '보유수량']
        if og == '매수':
            if code not in self.df_jg.index:
                bg = oc * cp
                pg, sg, sp = self.GetPgSgSp(bg, oc * cp)
                prec = self.GetMasterLastPrice(code)
                self.df_jg.at[code] = name, cp, cp, sp, sg, bg, pg, 0, 0, 0, prec, oc
            else:
                jc = self.df_jg['보유수량'][code]
                bg = self.df_jg['매입금액'][code]
                jc = jc + oc
                bg = bg + oc * cp
                bp = int(bg / jc)
                pg, sg, sp = self.GetPgSgSp(bg, jc * cp)
                self.df_jg.at[code, columns] = bp, cp, sp, sg, bg, pg, jc
        elif og == '매도':
            jc = self.df_jg['보유수량'][code]
            if jc - oc == 0:
                self.df_jg.drop(index=code, inplace=True)
            else:
                bp = self.df_jg['매입가'][code]
                jc = jc - oc
                bg = jc * bp
                pg, sg, sp = self.GetPgSgSp(bg, jc * cp)
                self.df_jg.at[code, columns] = bp, cp, sp, sg, bg, pg, jc

        columns = ['매입가', '현재가', '평가손익', '매입금액']
        self.df_jg[columns] = self.df_jg[columns].astype(int)
        self.df_jg.sort_values(by=['매입금액'], inplace=True)
        self.queryQ.put([1, self.df_jg, 'jangolist', 'replace'])
        if self.dict_bool['알림소리']:
            self.soundQ.put(f'{name} {oc}주를 {og}하였습니다')

    def UpdateTradelist(self, name, oc, sp, sg, bg, pg, on):
        dt = strf_time('%Y%m%d%H%M%S%f')
        if self.dict_bool['모의투자'] and len(self.df_td) > 0:
            if on in self.df_td.index:
                while on in self.df_td.index:
                    on = str(int(on) + 1)
                dt = on

        self.df_td.at[on] = name, bg, pg, oc, sp, sg, dt
        self.df_td.sort_values(by=['체결시간'], ascending=False, inplace=True)
        self.windowQ.put([ui_num['거래목록'], self.df_td])

        df = pd.DataFrame([[name, bg, pg, oc, sp, sg, dt]], columns=columns_td, index=[on])
        self.queryQ.put([1, df, 'tradelist', 'append'])
        self.UpdateTotaltradelist()

    def UpdateTotaltradelist(self, first=False):
        tsg = self.df_td['매도금액'].sum()
        tbg = self.df_td['매수금액'].sum()
        tsig = self.df_td[self.df_td['수익금'] > 0]['수익금'].sum()
        tssg = self.df_td[self.df_td['수익금'] < 0]['수익금'].sum()
        sg = self.df_td['수익금'].sum()
        sp = round(sg / self.dict_intg['추정예탁자산'] * 100, 2)
        tdct = len(self.df_td)
        self.df_tt = pd.DataFrame(
            [[tdct, tbg, tsg, tsig, tssg, sp, sg]], columns=columns_tt, index=[self.dict_strg['당일날짜']]
        )
        self.windowQ.put([ui_num['거래합계'], self.df_tt])

        if not first:
            self.teleQ.put(
                f"거래횟수 {len(self.df_td)}회 / 총매수금액 {format(int(tbg), ',')}원 / "
                f"총매도금액 {format(int(tsg), ',')}원 / 총수익금액 {format(int(tsig), ',')}원 / "
                f"총손실금액 {format(int(tssg), ',')}원 / 수익률 {sp}% / 수익금합계 {format(int(sg), ',')}원")

    def UpdateChegeollist(self, name, og, oc, omc, op, cp, on):
        dt = strf_time('%Y%m%d%H%M%S%f')
        if self.dict_bool['모의투자'] and len(self.df_cj) > 0:
            if on in self.df_cj.index:
                while on in self.df_cj.index:
                    on = str(int(on) + 1)
                dt = on

        if on in self.df_cj.index:
            self.df_cj.at[on, ['미체결수량', '체결가', '체결시간']] = omc, cp, dt
        else:
            self.df_cj.at[on] = name, og, oc, omc, op, cp, dt
        self.df_cj.sort_values(by=['체결시간'], ascending=False, inplace=True)
        self.windowQ.put([ui_num['체결목록'], self.df_cj])

        if omc == 0:
            df = pd.DataFrame([[name, og, oc, omc, op, cp, dt]], columns=columns_cj, index=[on])
            self.queryQ.put([1, df, 'chegeollist', 'append'])

    def Block_Request(self, *args, **kwargs):
        if self.dict_intg['TR제한수신횟수'] == 0:
            self.dict_time['TR시작'] = now()
        trcode = args[0].lower()
        lines = readEnc(trcode)
        self.dict_item = parseDat(trcode, lines)
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
        return self.df_tr

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

    def SetRealReg(self, rreg):
        self.ocx.dynamicCall('SetRealReg(QString, QString, QString, QString)', rreg)

    def SetRealRemove(self, rreg):
        self.ocx.dynamicCall('SetRealRemove(QString, QString)', rreg)

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
        self.stgQ.put('전략프로세스종료')
        self.windowQ.put([1, '시스템 명령 실행 알림 - 트레이더 종료'])
        self.teleQ.put('10초 후 시스템을 종료합니다.')
        if self.dict_bool['알림소리']:
            self.soundQ.put('10초 후 시스템을 종료합니다.')
        else:
            self.windowQ.put([1, '시스템 명령 실행 알림 - 60초 후 시스템을 종료합니다.'])
        i = 10
        while i > 0:
            self.windowQ.put([1, f'시스템 명령 실행 알림 - 시스템 종료 카운터 {i}'])
            i -= 1
            time.sleep(1)
        self.windowQ.put([1, '시스템 명령 실행 알림 - 시스템 종료'])
        sys.exit()
