import os
import sys
import time
import sqlite3
import zipfile
import pythoncom
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QAxContainer import QAxWidget
from multiprocessing import Process, Queue
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import strf_time, now, telegram_msg
from utility.setting import openapi_path, sn_brrq, db_day, db_stg, sn_cond
app = QtWidgets.QApplication(sys.argv)

PREMONEY = 300000       # 단위 백만
AVGPERIOD = 7           # 돌파계수 평균기간

"""
조건검색식 3번을 이용하여 장마감 후 실행해야한다.
조건검색식 조건은 시가총액 2천억 이상, 관리종목 제외, ETF, ETN, 스팩 등 제외, 거래대금 3천억 이상
"""


class UpdaterShort:
    def __init__(self, gubun, queryQQ, lockk):
        self.gubun = gubun
        self.queryQ = queryQQ
        self.lock = lockk
        self.str_trname = None
        self.str_tday = strf_time('%Y%m%d')
        self.df_tr = None
        self.list_trcd = []
        self.dict_tritems = None
        self.dict_cond = {}
        self.dict_bool = {
            '로그인': False,
            'TR수신': False,
            'CD수신': False,
            'CR수신': False
        }

        self.ocx = QAxWidget('KHOPENAPI.KHOpenAPICtrl.1')
        self.ocx.OnEventConnect.connect(self.OnEventConnect)
        self.ocx.OnReceiveTrData.connect(self.OnReceiveTrData)
        self.ocx.OnReceiveTrCondition.connect(self.OnReceiveTrCondition)
        self.ocx.OnReceiveConditionVer.connect(self.OnReceiveConditionVer)
        self.Start()

    def Start(self):
        self.CommConnect()
        self.Updater()

    def CommConnect(self):
        self.ocx.dynamicCall('CommConnect()')
        while not self.dict_bool['로그인']:
            pythoncom.PumpWaitingMessages()

        self.dict_bool['CD수신'] = False
        self.ocx.dynamicCall('GetConditionLoad()')
        while not self.dict_bool['CD수신']:
            pythoncom.PumpWaitingMessages()

        data = self.ocx.dynamicCall('GetConditionNameList()')
        conditions = data.split(';')[:-1]
        for condition in conditions:
            cond_index, cond_name = condition.split('^')
            self.dict_cond[int(cond_index)] = cond_name

    def Updater(self):
        codes = self.SendCondition(sn_cond, self.dict_cond[2], 2, 0)
        conn = sqlite3.connect(db_stg)
        df = pd.read_sql(f'SELECT * FROM jangolist', conn)
        conn.close()
        df = df.set_index('index')
        df = df[df['전략구분'] == '단기']
        jango_list = list(df.index)
        count = len(codes)
        for i, code in enumerate(codes):
            time.sleep(3.6)
            df = self.Block_Request('opt10081', 종목코드=code, 기준일자=self.str_tday, 수정주가구분=1,
                                    output='주식일봉차트조회', next=0)
            if len(df) < AVGPERIOD:
                continue
            df = df.set_index('일자')
            df = df[::-1]
            columns = ['현재가', '시가', '고가', '저가', '거래대금']
            df[columns] = df[columns].astype(int).abs()
            df['고가저가폭'] = df['고가'] - df['저가']
            df['종가시가폭'] = df['현재가'] - df['시가']
            df[['종가시가폭']] = df[['종가시가폭']].abs()
            df['돌파계수'] = df['종가시가폭'] / df['고가저가폭']
            df[['돌파계수']] = df[['돌파계수']].astype(float).round(2)
            df['평균돌파계수'] = df['돌파계수'].rolling(window=AVGPERIOD).mean().round(2)
            if df['거래대금'][-1] >= PREMONEY or code in jango_list:
                k = int(df['고가저가폭'][-1] * df['평균돌파계수'][-1])
                self.queryQ.put([code, k])
            print(f'[{now()}] {self.gubun} 데이터 업데이트 중 ... [{i + 1}/{count}]')
        self.queryQ.put('업데이트완료')
        sys.exit()

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

    def OnReceiveTrData(self, screen, rqname, trcode, record, nnext):
        if screen == '' and record == '':
            return
        items = None
        self.dict_bool['TR다음'] = True if nnext == '2' else False
        for output in self.dict_tritems['output']:
            record = list(output.keys())[0]
            items = list(output.values())[0]
            if record == self.str_trname:
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

    def SendCondition(self, screen, cond_name, cond_index, search):
        self.dict_bool['CR수신'] = False
        self.ocx.dynamicCall('SendCondition(QString, QString, int, int)', screen, cond_name, cond_index, search)
        while not self.dict_bool['CR수신']:
            pythoncom.PumpWaitingMessages()
        return self.list_trcd

    def Block_Request(self, *args, **kwargs):
        trcode = args[0].lower()
        liness = self.ReadEnc(trcode)
        self.dict_tritems = self.ParseDat(trcode, liness)
        self.str_trname = kwargs['output']
        nnext = kwargs['next']
        for i in kwargs:
            if i.lower() != 'output' and i.lower() != 'next':
                self.ocx.dynamicCall('SetInputValue(QString, QString)', i, kwargs[i])
        self.dict_bool['TR수신'] = False
        self.ocx.dynamicCall('CommRqData(QString, QString, int, QString)', self.str_trname, trcode, nnext, sn_brrq)
        while not self.dict_bool['TR수신']:
            pythoncom.PumpWaitingMessages()
        return self.df_tr

    # noinspection PyMethodMayBeStatic
    def ReadEnc(self, trcode):
        enc = zipfile.ZipFile(f'{openapi_path}/data/{trcode}.enc')
        liness = enc.read(trcode.upper() + '.dat').decode('cp949')
        return liness

    # noinspection PyMethodMayBeStatic
    def ParseDat(self, trcode, liness):
        liness = liness.split('\n')
        start = [i for i, x in enumerate(liness) if x.startswith('@START')]
        end = [i for i, x in enumerate(liness) if x.startswith('@END')]
        block = zip(start, end)
        enc_data = {'trcode': trcode, 'input': [], 'output': []}
        for start, end in block:
            block_data = liness[start - 1:end + 1]
            block_info = block_data[0]
            block_type = 'input' if 'INPUT' in block_info else 'output'
            record_line = block_data[1]
            tokens = record_line.split('_')[1].strip()
            record = tokens.split('=')[0]
            fields = block_data[2:-1]
            field_name = []
            for line in fields:
                field = line.split('=')[0].strip()
                field_name.append(field)
            fields = {record: field_name}
            enc_data['input'].append(fields) if block_type == 'input' else enc_data['output'].append(fields)
        return enc_data


class Query:
    def __init__(self, queryQQ):
        self.queryQ = queryQQ
        self.con = sqlite3.connect(db_day)
        self.Start()

    def __del__(self):
        self.con.close()

    def Start(self):
        columns = ['변동성']
        df_mid = pd.DataFrame(columns=columns)
        while True:
            data = self.queryQ.get()
            if data != '업데이트완료':
                df_mid.at[data[0]] = data[1]
            else:
                break

        df_mid[columns] = df_mid[columns].astype(int)
        con = sqlite3.connect(db_stg)
        df_mid.to_sql('short', con, if_exists='replace', chunksize=1000)
        con.close()
        telegram_msg('short DB를 업데이트하였습니다.')
        sys.exit()


if __name__ == '__main__':
    queryQ = Queue()
    Process(target=Query, args=(queryQ,)).start()
    Process(target=UpdaterShort, args=(queryQ,)).start()
