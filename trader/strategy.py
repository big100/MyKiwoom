import os
import sys
import psutil
import numpy as np
import pandas as pd
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.setting import columns_gj1, ui_num, DICT_SET
from utility.static import now, timedelta_sec, thread_decorator


class Strategy:
    def __init__(self, windowQ, traderQ, stgQ):
        self.windowQ = windowQ
        self.traderQ = traderQ
        self.stgQ = stgQ

        self.list_buy = []
        self.list_sell = []

        self.dict_gsjm = {}     # key: 종목코드, value: 10시이전 DataFrame, 10시이후 list
        self.dict_hgjr = {}     # key: 종목코드, value: list
        self.dict_time = {
            '관심종목': now(),
            '부가정보': now(),
            '연산시간': now()
        }
        self.dict_intg = {
            '종목당투자금': 0,
            '스레드': 0,
            '시피유': 0.,
            '메모리': 0.
        }

        self.Start()

    def Start(self):
        while True:
            data = self.stgQ.get()
            if type(data) == int:
                self.UpdateTotaljasan(data)
            elif type(data) == list:
                if len(data) == 2:
                    self.UpdateList(data[0], data[1])
                elif len(data) == 38:
                    self.BuyStrategy(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8],
                                     data[9], data[10], data[11], data[12], data[13], data[14], data[15], data[16],
                                     data[17], data[18], data[19], data[20], data[21], data[22], data[23], data[24],
                                     data[25], data[26], data[27], data[28], data[29], data[30], data[31], data[32],
                                     data[33], data[34], data[35], data[36], data[37])
                elif len(data) == 6:
                    self.SellStrategy(data[0], data[1], data[2], data[3], data[4], data[5])
            elif data == '전략연산프로세스종료':
                break

            if now() > self.dict_time['관심종목']:
                self.windowQ.put([ui_num['관심종목'], self.dict_gsjm])
                self.dict_time['관심종목'] = timedelta_sec(1)
            if now() > self.dict_time['부가정보']:
                self.UpdateInfo()
                self.dict_time['부가정보'] = timedelta_sec(2)

        self.windowQ.put([1, '시스템 명령 실행 알림 - 전략 연산 프로세스 종료'])

    def UpdateTotaljasan(self, data):
        self.dict_intg['종목당투자금'] = data

    def UpdateList(self, gubun, code):
        if '조건진입' in gubun:
            if code not in self.dict_gsjm.keys():
                data = np.zeros((DICT_SET['평균시간'] + 2, len(columns_gj1))).tolist()
                df = pd.DataFrame(data, columns=columns_gj1)
                df['체결시간'] = '090000'
                self.dict_gsjm[code] = df.copy()
        elif gubun == '조건이탈':
            if code in self.dict_gsjm.keys():
                del self.dict_gsjm[code]
        elif gubun in ['매수완료', '매수취소']:
            if code in self.list_buy:
                self.list_buy.remove(code)
        elif gubun in ['매도완료', '매도취소']:
            if code in self.list_sell:
                self.list_sell.remove(code)

    def BuyStrategy(self, 현재가, 시가, 고가, 저가, 등락율, 당일거래대금, 체결강도,
                    초당매수수량, 초당매도수량, VI해제시간, VI아래5호가, 매도총잔량, 매수총잔량,
                    매도호가5, 매도호가4, 매도호가3, 매도호가2, 매도호가1, 매수호가1, 매수호가2, 매수호가3, 매수호가4, 매수호가5,
                    매도잔량5, 매도잔량4, 매도잔량3, 매도잔량2, 매도잔량1, 매수잔량1, 매수잔량2, 매수잔량3, 매수잔량4, 매수잔량5,
                    종목코드, 체결시간, 틱수신시간, 종목명, 잔고종목):
        if 종목코드 not in self.dict_gsjm.keys():
            return

        고저평균 = round((고가 + 저가) / 2)
        고저평균대비등락율 = round((현재가 / 고저평균 - 1) * 100, 2)
        직전당일거래대금 = self.dict_gsjm[종목코드]['당일거래대금'][1]
        초당거래대금 = 0 if 직전당일거래대금 == 0 else int(당일거래대금 - 직전당일거래대금)

        self.dict_gsjm[종목코드] = self.dict_gsjm[종목코드].shift(1)
        self.dict_gsjm[종목코드].at[0] = 등락율, 고저평균대비등락율, 초당거래대금, 당일거래대금, 체결강도, 0.
        self.dict_hgjr[종목코드] = \
            [매도총잔량, 매수총잔량,
             매도호가5, 매도호가4, 매도호가3, 매도호가2, 매도호가1, 매수호가1, 매수호가2, 매수호가3, 매수호가4, 매수호가5,
             매도잔량5, 매도잔량4, 매도잔량3, 매도잔량2, 매도잔량1, 매수잔량1, 매수잔량2, 매수잔량3, 매수잔량4, 매수잔량5]
        if self.dict_gsjm[종목코드]['체결강도'][DICT_SET[f'평균시간']] != 0.:
            평균값인덱스 = DICT_SET['평균시간'] + 1
            초당거래대금평균 = int(self.dict_gsjm[종목코드]['초당거래대금'][1:평균값인덱스].mean())
            체결강도평균 = round(self.dict_gsjm[종목코드]['체결강도'][1:평균값인덱스].mean(), 2)
            최고체결강도 = round(self.dict_gsjm[종목코드]['체결강도'][1:평균값인덱스].max(), 2)
            self.dict_gsjm[종목코드].at[평균값인덱스] = 0., 0., 초당거래대금평균, 0, 체결강도평균, 최고체결강도

            if 잔고종목:
                return
            if 종목코드 in self.list_buy:
                return

            매수 = True

            # 전략 비공개

            if 매수:
                매수수량 = int(self.dict_intg['종목당투자금'] / 현재가)
                if 매수수량 > 0:
                    if 매수수량 <= 매도잔량1:
                        예상체결가 = 매도호가1
                    else:
                        매수수량1 = 매도잔량1
                        남은매수수량 = 매수수량 - 매수수량1
                        예상체결가 = round((매도호가1 * 매수수량1 + 매도호가2 * 남은매수수량) / 매수수량, 2)
                    self.list_buy.append(종목코드)
                    self.traderQ.put(['매수', 종목코드, 종목명, 예상체결가, 매수수량])

        if now() > self.dict_time['연산시간']:
            gap = (now() - 틱수신시간).total_seconds()
            self.windowQ.put([1, f'전략스 연산 시간 알림 - 수신시간과 연산시간의 차이는 [{gap}]초입니다.'])
            self.dict_time['연산시간'] = timedelta_sec(60)

    def SellStrategy(self, 종목코드, 종목명, 수익률, 보유수량, 현재가, 매수시간):
        if 종목코드 not in self.dict_gsjm.keys() or 종목코드 not in self.dict_hgjr.keys():
            return
        if 종목코드 in self.list_sell:
            return
        if self.dict_gsjm[종목코드]['체결강도'][DICT_SET[f'평균시간']] == 0.:
            return

        매도 = False
        평균값인덱스 = DICT_SET['평균시간'] + 1
        등락율 = self.dict_gsjm[종목코드]['등락율'][0]
        체결강도 = self.dict_gsjm[종목코드]['체결강도'][0]
        고저평균대비등락율 = self.dict_gsjm[종목코드]['고저평균대비등락율'][0]
        초당거래대금평균 = self.dict_gsjm[종목코드]['초당거래대금'][평균값인덱스]
        체결강도평균 = self.dict_gsjm[종목코드]['체결강도'][평균값인덱스]
        최고체결강도 = self.dict_gsjm[종목코드]['체결강도'][평균값인덱스]
        매도총잔량, 매수총잔량, \
            매도호가5, 매도호가4, 매도호가3, 매도호가2, 매도호가1, 매수호가1, 매수호가2, 매수호가3, 매수호가4, 매수호가5, \
            매도잔량5, 매도잔량4, 매도잔량3, 매도잔량2, 매도잔량1, 매수잔량1, 매수잔량2, 매수잔량3, 매수잔량4, 매수잔량5 = \
            self.dict_hgjr[종목코드]

        """ 매도 조건 예시 """
        if 수익률 <= -2 or 수익률 >= 3:
            매도 = True

        # 전략 비공개

        if 매도:
            if 보유수량 <= 매수잔량1:
                예상체결가 = 매수호가1
            else:
                보유수량1 = 매수잔량1
                남은보유수량 = 보유수량 - 보유수량1
                예상체결가 = round((매수호가1 * 보유수량1 + 매수호가2 * 남은보유수량) / 보유수량, 2)
            self.list_sell.append(종목코드)
            self.traderQ.put(['매도', 종목코드, 종목명, 예상체결가, 보유수량])

    @thread_decorator
    def UpdateInfo(self):
        info = [6, self.dict_intg['메모리'], self.dict_intg['스레드'], self.dict_intg['시피유']]
        self.windowQ.put(info)
        self.UpdateSysinfo()

    def UpdateSysinfo(self):
        p = psutil.Process(os.getpid())
        self.dict_intg['메모리'] = round(p.memory_info()[0] / 2 ** 20.86, 2)
        self.dict_intg['스레드'] = p.num_threads()
        self.dict_intg['시피유'] = round(p.cpu_percent(interval=2) / 2, 2)
