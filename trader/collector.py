import os
import sys
import psutil
import warnings
import numpy as np
import pandas as pd
warnings.filterwarnings("ignore", category=np.VisibleDeprecationWarning)
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import now, strf_time, timedelta_sec, thread_decorator


class Collector:
    def __init__(self, windowQ, queryQ, tickQ):
        self.windowQ = windowQ
        self.queryQ = queryQ
        self.tickQ = tickQ

        self.dict_df = {}
        self.dict_dm = {}
        self.dict_time = {
            '기록시간': now(),
            '부가정보': now()
        }
        self.dict_intg = {
            '스레드': 0,
            '시피유': 0.,
            '메모리': 0.
        }
        self.str_tday = strf_time('%Y%m%d')
        self.Start()

    def Start(self):
        while True:
            tick = self.tickQ.get()
            if type(tick) == list:
                self.UpdateTickData(tick[0], tick[1], tick[2], tick[3], tick[4], tick[5], tick[6], tick[7],
                                    tick[8], tick[9], tick[10], tick[11], tick[12], tick[13], tick[14],
                                    tick[15], tick[16], tick[17], tick[18], tick[19], tick[20], tick[21], tick[22])
            elif tick == '틱데이터저장':
                self.queryQ.put([2, self.dict_df])
                break

            if now() > self.dict_time['부가정보']:
                self.UpdateInfo()
                self.dict_time['부가정보'] = timedelta_sec(2)

        self.windowQ.put([1, '시스템 명령 실행 알림 - 콜렉터 종료'])

    def UpdateTickData(self, code, c, o, h, low, per, dm, ch, vp, bids, asks, vitime, vid5,
                       s2hg, s1hg, b1hg, b2hg, s2jr, s1jr, b1jr, b2jr, d, receivetime):
        try:
            hlm = int(round((h + low) / 2))
            hlmp = round((c / hlm - 1) * 100, 2)
        except ZeroDivisionError:
            return

        try:
            predm = self.dict_dm[code]
        except KeyError:
            predm = dm

        self.dict_dm[code] = dm
        sm = dm - predm
        d = self.str_tday + d

        if code not in self.dict_df.keys():
            self.dict_df[code] = pd.DataFrame(
                [[c, o, h, per, hlmp, sm, dm, ch, vp, bids, asks, vitime, vid5,
                  s2hg, s1hg, b1hg, b2hg, s2jr, s1jr, b1jr, b2jr]],
                columns=['현재가', '시가', '고가', '등락율', '고저평균대비등락율', '거래대금', '누적거래대금', '체결강도',
                         '전일거래량대비', '매수수량', '매도수량', 'VI발동시간', '상승VID5가격',
                         '매도호가2', '매도호가1', '매수호가1', '매수호가2',
                         '매도잔량2', '매도잔량1', '매수잔량1', '매수잔량2'],
                index=[d])
        else:
            self.dict_df[code].at[d] = \
                c, o, h, per, hlmp, sm, dm, ch, vp, bids, asks, vitime, vid5,\
                s2hg, s1hg, b1hg, b2hg, s2jr, s1jr, b1jr, b2jr

        if now() > self.dict_time['기록시간']:
            gap = (now() - receivetime).total_seconds()
            self.windowQ.put([1, f'콜렉터 수신 기록 알림 - 수신시간과 기록시간의 차이는 [{gap}]초입니다.'])
            self.queryQ.put([2, self.dict_df])
            self.dict_df = {}
            self.dict_time['기록시간'] = timedelta_sec(60)

    def SaveTickData(self, codes):
        self.queryQ.put([2, self.dict_df])

    @thread_decorator
    def UpdateInfo(self):
        info = [8, self.dict_intg['메모리'], self.dict_intg['스레드'], self.dict_intg['시피유']]
        self.windowQ.put(info)
        self.UpdateSysinfo()

    def UpdateSysinfo(self):
        p = psutil.Process(os.getpid())
        self.dict_intg['메모리'] = round(p.memory_info()[0] / 2 ** 20.86, 2)
        self.dict_intg['스레드'] = p.num_threads()
        self.dict_intg['시피유'] = round(p.cpu_percent(interval=2) / 2, 2)
