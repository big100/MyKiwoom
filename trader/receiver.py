import os
import sys
import psutil
import pythoncom
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5.QAxContainer import QAxWidget
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import *
from utility.setting import *

MONEYTOP_MINUTE = 5         # 최근거래대금순위을 집계할 시간
MONEYTOP_RANK = 20          # 최근거래대금순위중 관심종목으로 선정할 순위


class Updater(QtCore.QThread):
    data1 = QtCore.pyqtSignal(str)

    def __init__(self, receivQ):
        super().__init__()
        self.receivQ = receivQ

    def run(self):
        while True:
            data = self.receivQ.get()
            self.data1.emit(data)


class Receiver:
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
        self.queryQ = qlist[5]
        self.tick1Q = qlist[19]
        self.tick2Q = qlist[20]
        self.tick3Q = qlist[21]
        self.tick4Q = qlist[22]

        self.dict_bool = {
            '리시버시작': False,
            '실시간조건검색시작': False,
            '실시간조건검색중단': False,
            '장중단타전략시작': False,

            '로그인': False,
            'TR수신': False,
            'TR다음': False,
            'CD수신': False,
            'CR수신': False
        }
        self.dict_intg = {
            '스레드': 0,
            '시피유': 0.,
            '메모리': 0.
        }
        self.dict_cdjm = {}
        self.dict_vipr = {}
        self.dict_tick = {}
        self.dict_hoga = {}
        self.dict_cond = {}
        self.dict_name = {}
        self.dict_code = {}

        self.list_gsjm1 = []
        self.list_gsjm2 = []
        self.list_trcd = []
        self.list_jang = []
        self.list_prmt = []
        self.list_kosd = None
        self.list_code = None
        self.list_code1 = None
        self.list_code2 = None
        self.list_code3 = None
        self.list_code4 = None

        self.df_tr = None
        self.dict_item = None
        self.str_tname = None

        self.operation = 1
        self.df_mt = pd.DataFrame(columns=['거래대금순위'])
        self.df_mc = pd.DataFrame(columns=['최근거래대금'])
        self.str_tday = strf_time('%Y%m%d')
        self.str_jcct = self.str_tday + '090000'
        self.dt_mtct = None

        remaintime = (strp_time('%Y%m%d%H%M%S', self.str_tday + '090100') - now()).total_seconds()
        self.dict_time = {
            '휴무종료': timedelta_sec(remaintime) if remaintime > 0 else timedelta_sec(600),
            '거래대금순위기록': now(),
            '거래대금순위저장': now(),
            '부가정보': now()
        }

        self.ocx = QAxWidget('KHOPENAPI.KHOpenAPICtrl.1')
        self.ocx.OnEventConnect.connect(self.OnEventConnect)
        self.ocx.OnReceiveTrData.connect(self.OnReceiveTrData)
        self.ocx.OnReceiveRealData.connect(self.OnReceiveRealData)
        self.ocx.OnReceiveTrCondition.connect(self.OnReceiveTrCondition)
        self.ocx.OnReceiveConditionVer.connect(self.OnReceiveConditionVer)
        self.ocx.OnReceiveRealCondition.connect(self.OnReceiveRealCondition)

        self.CommConnect()

        self.updater = Updater(self.receivQ)
        self.updater.data1.connect(self.UpdateJangolist)
        self.updater.start()

        self.qtimer1 = QtCore.QTimer()
        self.qtimer1.setInterval(1000)
        self.qtimer1.timeout.connect(self.Scheduler)
        self.qtimer1.start()

        self.qtimer2 = QtCore.QTimer()
        self.qtimer2.setInterval(10000)
        self.qtimer2.timeout.connect(self.MoneyTopSearch)

        app.exec_()

    def CommConnect(self):
        self.ocx.dynamicCall('CommConnect()')
        while not self.dict_bool['로그인']:
            pythoncom.PumpWaitingMessages()

        self.dict_bool['CD수신'] = False
        self.ocx.dynamicCall('GetConditionLoad()')
        while not self.dict_bool['CD수신']:
            pythoncom.PumpWaitingMessages()

        con = sqlite3.connect(DB_TICK)
        df = pd.read_sql("SELECT name FROM sqlite_master WHERE TYPE = 'table'", con)
        con.close()
        table_list = list(df['name'].values)

        self.list_kosd = self.GetCodeListByMarket('10')
        list_code = self.GetCodeListByMarket('0') + self.list_kosd

        df = pd.DataFrame(columns=['종목명'])
        for code in list_code:
            name = self.GetMasterCodeName(code)
            df.at[code] = name
            self.dict_name[code] = name
            self.dict_code[name] = code
            if code not in table_list:
                query = f'CREATE TABLE "{code}" ("index" TEXT, "현재가" REAL, "시가" REAL, "고가" REAL,' \
                         '"저가" REAL, "등락율" REAL, "당일거래대금" REAL, "체결강도" REAL, "초당매수수량" REAL,' \
                         '"초당매도수량" REAL, "VI해제시간" TEXT, "VI아래5호가" REAL, "매도총잔량" REAL, "매수총잔량" REAL,' \
                         '"매도호가5" REAL, "매도호가4" REAL, "매도호가3" REAL, "매도호가2" REAL, "매도호가1" REAL,' \
                         '"매수호가1" REAL, "매수호가2" REAL, "매수호가3" REAL, "매수호가4" REAL, "매수호가5" REAL,' \
                         '"매도잔량5" REAL, "매도잔량4" REAL, "매도잔량3" REAL, "매도잔량2" REAL, "매도잔량1" REAL,' \
                         '"매수잔량1" REAL, "매수잔량2" REAL, "매수잔량3" REAL, "매수잔량4" REAL, "매수잔량5" REAL);'
                self.queryQ.put([2, query])
                query = f'CREATE INDEX "ix_{code}_index" ON "{code}"("index")'
                self.queryQ.put([2, query])
        self.queryQ.put('디비트리거시작')
        self.queryQ.put([2, df, 'codename', 'replace'])
        self.windowQ.put([3, self.dict_code])
        self.windowQ.put([4, self.dict_name])

        data = self.ocx.dynamicCall('GetConditionNameList()')
        conditions = data.split(';')[:-1]
        for condition in conditions:
            cond_index, cond_name = condition.split('^')
            self.dict_cond[int(cond_index)] = cond_name

        self.windowQ.put([1, self.dict_cond])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 OpenAPI 로그인 완료'])

    def UpdateJangolist(self, data):
        code = data.split(' ')[1]
        if '잔고편입' in data and code not in self.list_jang:
            self.list_jang.append(code)
            if code not in self.list_gsjm2:
                self.stgQ.put(['조건진입', code])
                self.list_gsjm2.append(code)
        elif '잔고청산' in data and code in self.list_jang:
            self.list_jang.remove(code)
            if code not in self.list_gsjm1 and code in self.list_gsjm2:
                self.stgQ.put(['조건이탈', code])
                self.list_gsjm2.remove(code)

    def Scheduler(self):
        if not self.dict_bool['리시버시작']:
            self.OperationRealreg()
        if self.operation == 1 and now() > self.dict_time['휴무종료']:
            self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 종료'])
            sys.exit()
        if self.operation == 3:
            if int(strf_time('%H%M%S')) < 100000:
                if not self.dict_bool['실시간조건검색시작']:
                    self.ConditionSearchStart()
            else:
                if self.dict_bool['실시간조건검색시작'] and not self.dict_bool['실시간조건검색중단']:
                    self.ConditionSearchStop()
                if not self.dict_bool['장중단타전략시작']:
                    self.StartJangjungStrategy()
        if self.operation == 8 and int(strf_time('%H%M%S')) >= 153500:
            self.RemoveAllRealreg()
            self.SaveTickData()
            self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 종료'])
            sys.exit()

        if now() > self.dict_time['거래대금순위기록']:
            if len(self.list_gsjm1) > 0:
                self.UpdateMoneyTop()
            self.dict_time['거래대금순위기록'] = timedelta_sec(1)
        if now() > self.dict_time['부가정보']:
            self.UpdateInfo()
            self.dict_time['부가정보'] = timedelta_sec(2)

    def OperationRealreg(self):
        self.dict_bool['리시버시작'] = True
        self.windowQ.put([2, '장운영시간 알림 등록'])
        self.SetRealReg([sn_oper, ' ', '215;20;214', 0])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 장운영시간 등록 완료'])

        self.windowQ.put([2, 'VI발동해제 등록'])
        self.Block_Request('opt10054', 시장구분='000', 장전구분='1', 종목코드='', 발동구분='1', 제외종목='111111011',
                           거래량구분='0', 거래대금구분='0', 발동방향='0', output='발동종목', next=0)
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 VI발동해제 등록 완료'])

        self.list_code = self.SendCondition([sn_oper, self.dict_cond[1], 1, 0])
        self.list_code1 = [code for i, code in enumerate(self.list_code) if i % 4 == 0]
        self.list_code2 = [code for i, code in enumerate(self.list_code) if i % 4 == 1]
        self.list_code3 = [code for i, code in enumerate(self.list_code) if i % 4 == 2]
        self.list_code4 = [code for i, code in enumerate(self.list_code) if i % 4 == 3]
        k = 0
        for i in range(0, len(self.list_code), 100):
            rreg = [sn_jchj + k, ';'.join(self.list_code[i:i + 100]), '10;12;14;30;228;41;61;71;81', 1]
            self.SetRealReg(rreg)
            text = f"실시간 알림 등록 완료 - [{sn_jchj + k}] 종목갯수 {len(rreg[1].split(';'))}"
            self.windowQ.put([1, text])
            k += 1
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 실시간주식체결 등록 완료'])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 시작'])

    def ConditionSearchStart(self):
        self.dict_bool['실시간조건검색시작'] = True
        self.windowQ.put([2, '실시간 조건검색식 등록'])
        codes = self.SendCondition([sn_cond, self.dict_cond[0], 0, 1])
        if len(codes) > 0:
            for code in codes:
                self.InsertGsjmlist(code)
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 실시간조건검색 등록 완료'])

    def ConditionSearchStop(self):
        self.dict_bool['실시간조건검색중단'] = True
        self.windowQ.put([2, '실시간 조건검색식 중단'])
        self.SendConditionStop([sn_cond, self.dict_cond[0], 0])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 실시간조건검색 중단 완료'])

    def StartJangjungStrategy(self):
        self.dict_bool['장중단타전략시작'] = True
        self.windowQ.put([2, '장중전략 시작'])
        self.df_mc.sort_values(by=['최근거래대금'], ascending=False, inplace=True)
        list_top = list(self.df_mc.index[:MONEYTOP_RANK])
        insert_list = set(list_top) - set(self.list_gsjm1)
        if len(insert_list) > 0:
            for code in list(insert_list):
                self.InsertGsjmlist(code)
        delete_list = set(self.list_gsjm1) - set(list_top)
        if len(delete_list) > 0:
            for code in list(delete_list):
                self.DeleteGsjmlist(code)
        self.list_prmt = list_top
        self.qtimer2.start()

    def RemoveAllRealreg(self):
        self.dict_bool['실시간데이터수신중단'] = True
        self.windowQ.put([2, '실시간 데이터 수신 중단'])
        self.SetRealRemove(['ALL', 'ALL'])
        self.windowQ.put([1, '시스템 명령 실행 알림 - 리시버 실시간 데이터 중단 완료'])

    def SaveTickData(self):
        self.windowQ.put([2, '틱데이터 저장'])
        con = sqlite3.connect(DB_STG)
        df = pd.read_sql(f"SELECT * FROM chegeollist WHERE 체결시간 LIKE '{self.str_tday}%'", con).set_index('index')
        con.close()
        codes = []
        for name in list(df['종목명'].values):
            code = self.dict_code[name]
            if code not in codes:
                codes.append(code)
        self.tick1Q.put(['콜렉터종료', codes])
        self.tick2Q.put(['콜렉터종료', codes])
        self.tick3Q.put(['콜렉터종료', codes])
        self.tick4Q.put(['콜렉터종료', codes])

    def UpdateMoneyTop(self):
        timetype = '%Y%m%d%H%M%S'
        list_text = ';'.join(self.list_gsjm1)
        curr_strtime = self.str_jcct
        curr_datetime = strp_time(timetype, curr_strtime)
        if self.dt_mtct is not None:
            gap_seconds = (curr_datetime - self.dt_mtct).total_seconds()
            while gap_seconds > 1:
                gap_seconds -= 1
                pre_time = strf_time(timetype, timedelta_sec(-gap_seconds, curr_datetime))
                self.df_mt.at[pre_time] = list_text
        if curr_datetime != self.dt_mtct:
            self.df_mt.at[curr_strtime] = list_text
            self.dt_mtct = curr_datetime

        if now() > self.dict_time['거래대금순위저장']:
            self.queryQ.put([2, self.df_mt, 'moneytop', 'append'])
            self.df_mt = pd.DataFrame(columns=['거래대금순위'])
            self.dict_time['거래대금순위저장'] = timedelta_sec(10)

    def MoneyTopSearch(self):
        self.df_mc.sort_values(by=['최근거래대금'], ascending=False, inplace=True)
        list_top = list(self.df_mc.index[:MONEYTOP_RANK])
        insert_list = set(list_top) - set(self.list_prmt)
        if len(insert_list) > 0:
            for code in list(insert_list):
                self.InsertGsjmlist(code)
        delete_list = set(self.list_prmt) - set(list_top)
        if len(delete_list) > 0:
            for code in list(delete_list):
                self.DeleteGsjmlist(code)
        self.list_prmt = list_top

    def InsertGsjmlist(self, code):
        if code not in self.list_gsjm1:
            self.list_gsjm1.append(code)
        if code not in self.list_jang and code not in self.list_gsjm2:
            self.stgQ.put(['조건진입', code])
            self.list_gsjm2.append(code)

    def DeleteGsjmlist(self, code):
        if code in self.list_gsjm1:
            self.list_gsjm1.remove(code)
        if code not in self.list_jang and code in self.list_gsjm2:
            self.stgQ.put(['조건이탈', code])
            self.list_gsjm2.remove(code)

    @thread_decorator
    def UpdateInfo(self):
        info = [7, self.dict_intg['메모리'], self.dict_intg['스레드'], self.dict_intg['시피유']]
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

    def OnReceiveRealCondition(self, code, IorD, cname, cindex):
        if cname == '' and cindex == '':
            return
        if int(strf_time('%H%M%S')) > 100000:
            return

        if IorD == 'I':
            self.InsertGsjmlist(code)
        elif IorD == 'D':
            self.DeleteGsjmlist(code)

    def OnReceiveRealData(self, code, realtype, realdata):
        if realdata == '':
            return

        if realtype == '장시작시간':
            try:
                self.operation = int(self.GetCommRealData(code, 215))
                current = self.GetCommRealData(code, 20)
                remain = self.GetCommRealData(code, 214)
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData 장시작시간 {e}'])
            else:
                self.windowQ.put([1, f'장운영 시간 수신 알림 - {self.operation} {current[:2]}:{current[2:4]}:{current[4:]}'
                                     f' 남은시간 {remain[:2]}:{remain[2:4]}:{remain[4:]}'])
        elif realtype == 'VI발동/해제':
            try:
                code = self.GetCommRealData(code, 9001).strip('A').strip('Q')
                gubun = self.GetCommRealData(code, 9068)
                name = self.GetMasterCodeName(code)
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData VI발동/해제 {e}'])
            else:
                if gubun == '1' and code in self.list_code and \
                        (code not in self.dict_vipr.keys() or (self.dict_vipr[code][0] and now() > self.dict_vipr[code][1])):
                    self.UpdateViPrice(code, name)
        elif realtype == '주식체결':
            try:
                c = abs(int(self.GetCommRealData(code, 10)))
                o = abs(int(self.GetCommRealData(code, 16)))
                v = self.GetCommRealData(code, 15)
                dt = self.str_tday + self.GetCommRealData(code, 20)
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData 주식체결 {e}'])
            else:
                if self.operation == 1:
                    self.operation = 3
                if dt != self.str_jcct and int(dt) > int(self.str_jcct):
                    self.str_jcct = dt
                if code not in self.dict_vipr.keys():
                    self.InsertViPrice(code, o)
                elif not self.dict_vipr[code][0] and now() > self.dict_vipr[code][1]:
                    self.UpdateViPrice(code, c)
                try:
                    predt, bid_volumns, ask_volumns = self.dict_tick[code]
                except KeyError:
                    predt, bid_volumns, ask_volumns = None, 0, 0
                if '+' in v:
                    self.dict_tick[code] = [dt, bid_volumns + abs(int(v)), ask_volumns]
                elif '-' in v:
                    self.dict_tick[code] = [dt, bid_volumns, ask_volumns + abs(int(v))]
                if dt != predt:
                    try:
                        bids, asks = self.dict_tick[code][1:]
                    except KeyError:
                        bids, asks = 0, 0
                    self.dict_tick[code] = [dt, 0, 0]
                    try:
                        h = abs(int(self.GetCommRealData(code, 17)))
                        low = abs(int(self.GetCommRealData(code, 18)))
                        per = float(self.GetCommRealData(code, 12))
                        dm = int(self.GetCommRealData(code, 14))
                        ch = float(self.GetCommRealData(code, 228))
                        name = self.GetMasterCodeName(code)
                    except Exception as e:
                        self.windowQ.put([1, f'OnReceiveRealData 주식체결 {e}'])
                    else:
                        if code in self.dict_hoga.keys():
                            self.UpdateTickData(code, name, c, o, h, low, per, dm, ch, bids, asks, dt, now())
        elif realtype == '주식호가잔량':
            try:
                tsjr = int(self.GetCommRealData(code, 121))
                tbjr = int(self.GetCommRealData(code, 125))
                s5hg, b5hg = abs(int(self.GetCommRealData(code, 45))), abs(int(self.GetCommRealData(code, 55)))
                s4hg, b4hg = abs(int(self.GetCommRealData(code, 44))), abs(int(self.GetCommRealData(code, 54)))
                s3hg, b3hg = abs(int(self.GetCommRealData(code, 43))), abs(int(self.GetCommRealData(code, 53)))
                s2hg, b2hg = abs(int(self.GetCommRealData(code, 42))), abs(int(self.GetCommRealData(code, 52)))
                s1hg, b1hg = abs(int(self.GetCommRealData(code, 41))), abs(int(self.GetCommRealData(code, 51)))
                s5jr, b5jr = int(self.GetCommRealData(code, 65)), int(self.GetCommRealData(code, 75))
                s4jr, b4jr = int(self.GetCommRealData(code, 64)), int(self.GetCommRealData(code, 74))
                s3jr, b3jr = int(self.GetCommRealData(code, 63)), int(self.GetCommRealData(code, 73))
                s2jr, b2jr = int(self.GetCommRealData(code, 62)), int(self.GetCommRealData(code, 72))
                s1jr, b1jr = int(self.GetCommRealData(code, 61)), int(self.GetCommRealData(code, 71))
            except Exception as e:
                self.windowQ.put([1, f'OnReceiveRealData 주식호가잔량 {e}'])
            else:
                self.dict_hoga[code] = [tsjr, tbjr,
                                        s5hg, s4hg, s3hg, s2hg, s1hg, b1hg, b2hg, b3hg, b4hg, b5hg,
                                        s5jr, s4jr, s3jr, s2jr, s1jr, b1jr, b2jr, b3jr, b4jr, b5jr]

    def InsertViPrice(self, code, o):
        uvi, dvi, vid5price = self.GetVIPrice(code, o)
        self.dict_vipr[code] = [True, timedelta_sec(-180), uvi, dvi, vid5price]

    def UpdateViPrice(self, code, key):
        if type(key) == str:
            try:
                self.dict_vipr[code][:2] = False, timedelta_sec(5)
            except KeyError:
                self.dict_vipr[code] = [False, timedelta_sec(5), 0, 0, 0]
            self.windowQ.put([1, f'변동성 완화 장치 발동 - [{code}] {key}'])
        elif type(key) == int:
            uvi, dvi, vid5price = self.GetVIPrice(code, key)
            self.dict_vipr[code] = [True, timedelta_sec(5), uvi, dvi, vid5price]

    def GetVIPrice(self, code, std_price):
        uvi = std_price * 1.1
        x = self.GetHogaunit(code, uvi)
        if uvi % x != 0:
            uvi = uvi + (x - uvi % x)
        vid5price = uvi - x * 5
        dvi = std_price * 0.9
        x = self.GetHogaunit(code, dvi)
        if dvi % x != 0:
            dvi = dvi - dvi % x
        return int(uvi), int(dvi), int(vid5price)

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

    def UpdateTickData(self, code, name, c, o, h, low, per, dm, ch, bids, asks, dt, receivetime):
        dt_ = dt[:13]
        if code not in self.dict_cdjm.keys():
            columns = ['10초누적거래대금', '10초전당일거래대금']
            self.dict_cdjm[code] = pd.DataFrame([[0, dm]], columns=columns, index=[dt_])
        elif dt_ != self.dict_cdjm[code].index[-1]:
            predm = self.dict_cdjm[code]['10초전당일거래대금'][-1]
            self.dict_cdjm[code].at[dt_] = dm - predm, dm
            if len(self.dict_cdjm[code]) == MONEYTOP_MINUTE * 6:
                if per > 0:
                    self.df_mc.at[code] = self.dict_cdjm[code]['10초누적거래대금'].sum()
                elif code in self.df_mc.index:
                    self.df_mc.drop(index=code, inplace=True)
                self.dict_cdjm[code].drop(index=self.dict_cdjm[code].index[0], inplace=True)

        vitime = self.dict_vipr[code][1]
        vid5price = self.dict_vipr[code][4]
        data = [c, o, h, low, per, dm, ch, bids, asks, vitime, vid5price]
        data += self.dict_hoga[code] + [code, dt, receivetime]

        if code in self.list_gsjm2:
            injango = code in self.list_jang
            self.stgQ.put(data + [name, injango])
            if injango:
                self.traderQ.put([code, name, c, o, h, low])

        data[9] = strf_time('%Y%m%d%H%M%S', vitime)
        if code in self.list_code1:
            self.tick1Q.put(data)
        elif code in self.list_code2:
            self.tick2Q.put(data)
        elif code in self.list_code3:
            self.tick3Q.put(data)
        elif code in self.list_code4:
            self.tick4Q.put(data)

    def OnReceiveTrData(self, screen, rqname, trcode, record, nnext):
        if screen == '' and record == '':
            return
        items = None
        self.dict_bool['TR다음'] = True if nnext == '2' else False
        for output in self.dict_item['output']:
            record = list(output.keys())[0]
            items = list(output.values())[0]
            if record == self.str_tname:
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

    def Block_Request(self, *args, **kwargs):
        trcode = args[0].lower()
        lines = readEnc(trcode)
        self.dict_item = parseDat(trcode, lines)
        self.str_tname = kwargs['output']
        nnext = kwargs['next']
        for i in kwargs:
            if i.lower() != 'output' and i.lower() != 'next':
                self.ocx.dynamicCall('SetInputValue(QString, QString)', i, kwargs[i])
        self.dict_bool['TR수신'] = False
        self.dict_bool['TR다음'] = False
        self.ocx.dynamicCall('CommRqData(QString, QString, int, QString)', self.str_tname, trcode, nnext, sn_brrq)
        sleeptime = timedelta_sec(0.25)
        while not self.dict_bool['TR수신'] or now() < sleeptime:
            pythoncom.PumpWaitingMessages()
        return self.df_tr

    def SendCondition(self, cond):
        self.dict_bool['CR수신'] = False
        self.ocx.dynamicCall('SendCondition(QString, QString, int, int)', cond)
        sleeptime = timedelta_sec(0.25)
        while not self.dict_bool['CR수신'] or now() < sleeptime:
            pythoncom.PumpWaitingMessages()
        return self.list_trcd

    def SendConditionStop(self, cond):
        self.ocx.dynamicCall("SendConditionStop(QString, QString, int)", cond)

    def SetRealReg(self, rreg):
        self.ocx.dynamicCall('SetRealReg(QString, QString, QString, QString)', rreg)

    def SetRealRemove(self, rreg):
        self.ocx.dynamicCall('SetRealRemove(QString, QString)', rreg)

    def GetMasterCodeName(self, code):
        return self.ocx.dynamicCall('GetMasterCodeName(QString)', code)

    def GetCodeListByMarket(self, market):
        data = self.ocx.dynamicCall('GetCodeListByMarket(QString)', market)
        tokens = data.split(';')[:-1]
        return tokens

    def GetCommRealData(self, code, fid):
        return self.ocx.dynamicCall('GetCommRealData(QString, int)', code, fid)
