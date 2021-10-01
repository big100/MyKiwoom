import os
import sys
import psutil
import pandas as pd
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import now, timedelta_sec, thread_decorator


class UpdaterHoga:
    def __init__(self, windowQ, hogaQ, gubun):
        self.windowQ = windowQ
        self.hogaQ = hogaQ
        self.gubun = gubun

        self.df_hc = None
        self.df_hg = None
        self.df_so = None
        self.df_bo = None
        self.bool_hcup = False
        self.bool_hgup = False
        self.dict_time = {
            '호가갱신': now(),
            '부가정보': now()
        }
        self.dict_intg = {
            '스레드': 0,
            '시피유': 0.,
            '메모리': 0.
        }
        self.UpdateHoga('초기화')
        self.Start()

    def Start(self):
        while True:
            hoga = self.hogaQ.get()
            if type(hoga) == str:
                self.UpdateHoga(hoga)
            elif type(hoga) == list:
                if len(hoga) == 2:
                    self.UpdateChegeolcount(hoga[0], hoga[1])
                elif len(hoga) == 3:
                    self.UpdateMichejeolcount(hoga[0], hoga[1], hoga[2])
                elif len(hoga) == 7:
                    self.UpdateHogajalryang(hoga[0], hoga[1], hoga[2], hoga[3], hoga[4], hoga[5], hoga[6])

            if now() > self.dict_time['호가갱신']:
                if self.bool_hcup and self.df_hc is not None:
                    self.windowQ.put([self.gubun + 3, self.df_hc])
                    self.bool_hcup = False
                if self.bool_hgup and self.df_hg is not None:
                    self.windowQ.put([self.gubun + 4, self.df_hg])
                    self.bool_hgup = False
                self.dict_time['호가갱신'] = timedelta_sec(0.25)
            if now() > self.dict_time['부가정보']:
                self.UpdateInfo()
                self.dict_time['부가정보'] = timedelta_sec(2)

    def UpdateHoga(self, text):
        if text == '초기화':
            mc = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
            cc = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            ch = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
                  0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
            self.df_so = pd.DataFrame({'매도미체결수량': mc})
            self.df_hc = pd.DataFrame({'체결수량': cc, '체결강도': ch})
            self.df_hg = pd.DataFrame({'증감': cc, '잔량': cc, '호가': cc, '등락율': ch})
            self.df_bo = pd.DataFrame({'매수미체결수량': mc})
            self.windowQ.put([self.gubun + 2, self.df_so])
            self.windowQ.put([self.gubun + 3, self.df_hc])
            self.windowQ.put([self.gubun + 4, self.df_hg])
            self.windowQ.put([self.gubun + 5, self.df_bo])

    def UpdateChegeolcount(self, v, ch):
        if v > 0:
            tbc = self.df_hc['체결수량'][0] + v
            tsc = self.df_hc['체결수량'][21]
        else:
            tbc = self.df_hc['체결수량'][0]
            tsc = self.df_hc['체결수량'][21] + abs(v)
        hch = self.df_hc['체결강도'][0]
        if hch < ch:
            hch = ch
        lch = self.df_hc['체결강도'][21]
        if lch == 0 or lch > ch:
            lch = ch
        self.df_hc = self.df_hc.shift(1)
        self.df_hc.at[0, ['체결수량', '체결강도']] = tbc, hch
        self.df_hc.at[1, ['체결수량', '체결강도']] = v, ch
        self.df_hc.at[21, ['체결수량', '체결강도']] = tsc, lch
        self.df_hc[['체결수량']] = self.df_hc[['체결수량']].astype(int)
        self.bool_hcup = True

    def UpdateMichejeolcount(self, og, op, omc):
        mc = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
        if omc != 0:
            if op == 0:
                mc[21] = omc
            else:
                df = self.df_hg[self.df_hg['호가'] == op]
                if len(df) > 0:
                    mc[df.index[0]] = omc
        if og == '매도':
            self.df_so = pd.DataFrame({'매도미체결수량': mc})
            self.windowQ.put([self.gubun + 2, self.df_so])
        elif og == '매수':
            self.df_bo = pd.DataFrame({'매수미체결수량': mc})
            self.windowQ.put([self.gubun + 5, self.df_bo])

    def UpdateHogajalryang(self, vp, jc, hg, per, og, op, omc):
        if og in ['매수', '매도']:
            mindex = ''
            if og == '매수':
                df = self.df_bo[self.df_bo['매수미체결수량'] != '']
                if len(df) > 0:
                    mindex = df.index[0]
            elif og == '매도':
                df = self.df_so[self.df_so['매도미체결수량'] != '']
                if len(df) > 0:
                    mindex = df.index[0]
            self.df_hg = pd.DataFrame({'증감': vp, '잔량': jc, '호가': hg, '등락율': per})
            if mindex != '' and op != self.df_hg['호가'][mindex]:
                self.UpdateMichejeolcount(og, op, omc)
        else:
            self.df_hg = pd.DataFrame({'증감': vp, '잔량': jc, '호가': hg, '등락율': per})
        self.bool_hgup = True

    @thread_decorator
    def UpdateInfo(self):
        info = [self.gubun, self.dict_intg['메모리'], self.dict_intg['스레드'], self.dict_intg['시피유']]
        self.windowQ.put(info)
        self.UpdateSysinfo()

    def UpdateSysinfo(self):
        p = psutil.Process(os.getpid())
        self.dict_intg['메모리'] = round(p.memory_info()[0] / 2 ** 20.86, 2)
        self.dict_intg['스레드'] = p.num_threads()
        self.dict_intg['시피유'] = round(p.cpu_percent(interval=2) / 2, 2)
