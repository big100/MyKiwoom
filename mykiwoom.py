import time
import psutil
import logging
from PyQt5.QtGui import QPalette
from multiprocessing import Process, Queue
from utility.setui import *
from utility.static import *
from utility.query import Query
from utility.sound import Sound
from utility.telegrammsg import TelegramMsg
from trader.trader import Trader
from trader.strategy import Strategy
from trader.receiver import Receiver
from trader.collector import Collector
from trader.updater_hoga import UpdaterHoga
from trader.updater_chart import UpdaterChart


class Window(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.log = logging.getLogger('Window')
        self.log.setLevel(logging.INFO)
        filehandler = logging.FileHandler(filename=f"{system_path}/Log/S{strf_time('%Y%m%d')}.txt", encoding='utf-8')
        self.log.addHandler(filehandler)

        SetUI(self)

        self.dict_code = {}
        self.dict_name = {}
        self.dict_mcpg_lastindex = {}
        self.dict_mcpg_lastchuse = {}
        self.dict_mcpg_lastmoveavg = {}
        self.dict_mcpg_lastcandlestick = {}
        self.dict_mcpg_lastmoneybar = {}
        self.dict_mcpg_infiniteline = {}
        self.dict_mcpg_legend1 = {}
        self.dict_mcpg_legend2 = {}
        self.dict_mcpg_name = {}
        self.dict_mcpg_close = {}

        self.mode0 = 0
        self.mode1 = 0
        self.mode2 = 0

        self.list_info = [
            [0., 0, 0.],    # 0 ui
            [0., 0, 0.],    # 1 strategy
            [0., 0, 0.],    # 2 receiver
            [0., 0, 0.],    # 3 collector1
            [0., 0, 0.],    # 4 collector2
            [0., 0, 0.],    # 5 collector3
            [0., 0, 0.],    # 6 collector4
            [0., 0, 0.],    # 7 hoga1
            [0., 0, 0.],    # 8 hoga2
            [0., 0, 0.],    # 9 chart1
            [0., 0, 0.],    # 10 chart2
            [0., 0, 0.],    # 11 chart3
            [0., 0, 0.],    # 12 chart4
            [0., 0, 0.],    # 13 chart5
            [0., 0, 0.],    # 14 chart6
            [0., 0, 0.],    # 15 chart7
            [0., 0, 0.],    # 16 chart8
            [0., 0, 0.]     # 17 chart9
        ]
        self.rowcol = [     # Qt.Key_Return, Qt.Key_Enter
            [-1, -1],       # self.td_tableWidget [row, col]
            [-1, -1],       # self.jg_tableWidget [row, col]
            [-1, -1],       # self.cj_tableWidget [row, col]
            [-1, -1],       # self.gj_tableWidget [row, col]
            [-1, -1]        # self.dd_tableWidget [row, col]
        ]

        self.writer = Writer()
        self.writer.data0.connect(self.UpdateTexedit)
        self.writer.data1.connect(self.UpdateChart)
        self.writer.data2.connect(self.UpdateGaonsimJongmok)
        self.writer.data3.connect(self.UpdateTablewidget)
        self.writer.start()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            item = None
            col_number = 0
            row = self.td_tableWidget.currentIndex().row()
            col = self.td_tableWidget.currentIndex().column()
            if col <= 1 and (row != self.rowcol[0][0] or col != self.rowcol[0][1]):
                item, col_number = self.td_tableWidget.item(row, 0), col
                self.rowcol[0] = [row, col]
            row = self.jg_tableWidget.currentIndex().row()
            col = self.jg_tableWidget.currentIndex().column()
            if col <= 1 and (row != self.rowcol[1][0] or col != self.rowcol[1][1]):
                item, col_number = self.jg_tableWidget.item(row, 0), col
                self.rowcol[1] = [row, col]
            row = self.cj_tableWidget.currentIndex().row()
            col = self.cj_tableWidget.currentIndex().column()
            if col <= 1 and (row != self.rowcol[2][0] or col != self.rowcol[2][1]):
                item, col_number = self.cj_tableWidget.item(row, 0), col
                self.rowcol[2] = [row, col]
            row = self.gj_tableWidget.currentIndex().row()
            col = self.gj_tableWidget.currentIndex().column()
            if col <= 1 and (row != self.rowcol[3][0] or col != self.rowcol[3][1]):
                item, col_number = self.gj_tableWidget.item(row, 0), col
                self.rowcol[3] = [row, col]
            row = self.dd_tableWidget.currentIndex().row()
            col = self.dd_tableWidget.currentIndex().column()
            if col <= 1 and (row != self.rowcol[4][0] or col != self.rowcol[4][1]):
                item, col_number = self.dd_tableWidget.item(row, 1), col
                self.rowcol[4] = [row, col]
            if item is None:
                return
            code = self.dict_code[item.text()]
            self.PutTraderQ(code, col_number)

    def ReturnPressed_1(self):
        codeorname = self.ct_lineEdit_01.text()
        try:
            code = self.dict_code[codeorname]
        except KeyError:
            if codeorname in self.dict_code.values():
                code = codeorname
            else:
                windowQ.put([1, '시스템 명령 오류 알림 - 종목명 또는 종목코드를 잘못입력하였습니다.'])
                return
        if self.mode1 == 0:
            traderQ.put(f"현재가{ui_num['차트P1']} {code}")
        elif self.mode1 == 1:
            traderQ.put(f"현재가{ui_num['차트P0']} {code}")

    def ReturnPressed_2(self):
        codeorname = self.ct_lineEdit_02.text()
        try:
            code = self.dict_code[codeorname]
        except KeyError:
            if codeorname in self.dict_code.values():
                code = codeorname
            else:
                windowQ.put([1, '시스템 명령 오류 알림 - 종목명 또는 종목코드를 잘못입력하였습니다.'])
                return
        traderQ.put(f"현재가{ui_num['차트P3']} {code}")

    def UpdateTexedit(self, msg):
        if msg[0] == 0:
            self.gg_textEdit.clear()
            self.gg_textEdit.append(msg[1])
        elif msg[0] == 1:
            if '오류' in msg[1] or '실패' in msg[1]:
                self.lg_textEdit.setTextColor(color_fg_bc)
            elif '매매 시스템 체결 알림' in msg[1]:
                self.lg_textEdit.setTextColor(color_fg_bt)
            else:
                self.lg_textEdit.setTextColor(color_fg_dk)
            self.lg_textEdit.append(f'[{now()}] {msg[1]}')
            self.log.info(f'[{now()}] {msg[1]}')
            if msg[1] == '시스템 명령 실행 알림 - 시스템 종료':
                app.quit()
        elif msg[0] == 2:
            pushbutton = None
            if msg[1] == '데이터베이스 불러오기':
                pushbutton = self.sj_pushButton_02
            elif msg[1] == 'OPENAPI 로그인':
                pushbutton = self.sj_pushButton_03
                self.ButtonClicked_4(0)
            elif msg[1] == '계좌평가 및 잔고':
                pushbutton = self.sj_pushButton_04
            elif msg[1] == '코스피 코스닥 차트':
                pushbutton = self.sj_pushButton_05
            elif msg[1] == '장운영시간 알림 등록':
                pushbutton = self.sj_pushButton_06
            elif msg[1] == '업종지수 주식체결 등록':
                pushbutton = self.sj_pushButton_07
            elif msg[1] == 'VI발동해제 등록':
                self.ButtonClicked_4(2)
                pushbutton = self.sj_pushButton_08
            elif msg[1] == '장운영상태':
                pushbutton = self.sj_pushButton_09
            elif msg[1] == '실시간 조건검색식 등록':
                pushbutton = self.sj_pushButton_10
            elif msg[1] == '잔고청산':
                pushbutton = self.sj_pushButton_11
            elif msg[1] == '실시간 데이터 수신 중단':
                pushbutton = self.sj_pushButton_12
            elif msg[1] == '틱데이터 저장':
                pushbutton = self.sj_pushButton_13
            elif msg[1] == '시스템 종료':
                pushbutton = self.sj_pushButton_17
            if pushbutton is not None:
                pushbutton.setStyleSheet(style_bc_dk)
                pushbutton.setFont(qfont12)

            pushbutton = None
            text = None
            if '테스트모드' in msg[1]:
                pushbutton = self.sj_pushButton_14
                text = '테스트모드 ON' if msg[1].split(' ')[-1] in ['ON', '1'] else '테스트모드 OFF'
            elif '모의투자' in msg[1]:
                pushbutton = self.sj_pushButton_15
                text = '모의투자 ON' if msg[1].split(' ')[-1] in ['ON', '1'] else '모의투자 OFF'
            elif '알림소리' in msg[1]:
                pushbutton = self.sj_pushButton_16
                text = '알림소리 ON' if msg[1].split(' ')[-1] in ['ON', '1'] else '알림소리 OFF'

            if pushbutton is not None and text is not None:
                pushbutton.setText(text)
                if msg[1].split(' ')[-1] in ['ON', '1']:
                    pushbutton.setStyleSheet(style_bc_bt)
                else:
                    pushbutton.setStyleSheet(style_bc_dk)
                pushbutton.setFont(qfont12)

            if '텔레그램봇넘버' in msg[1]:
                text = msg[1].split(' ')[-1]
                self.sj_lineEdit_01.setText(text)
                self.sj_lineEdit_01.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            elif '사용자아이디' in msg[1]:
                text = msg[1].split(' ')[-1]
                self.sj_lineEdit_02.setText(text)
                self.sj_lineEdit_02.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        elif msg[0] == 3:
            self.dict_code = msg[1]
        elif msg[0] == 4:
            self.dict_name = msg[1]
            completer = QtWidgets.QCompleter(list(self.dict_code.keys()) + list(self.dict_name.keys()))
            self.ct_lineEdit_01.setCompleter(completer)
            self.ct_lineEdit_02.setCompleter(completer)
        elif msg[0] == 5:
            float_memory = float2str3p2(msg[1])
            int_thread = msg[2]
            float_cpuper = float2str2p2(msg[3])
            label01text = f'Trader Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%'

            float_memory = float2str3p2(self.list_info[1][0])
            int_thread = self.list_info[1][1]
            float_cpuper = float2str2p2(self.list_info[1][2])
            label06text = f"Stretegy Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%"

            float_memory = float2str3p2(self.list_info[2][0])
            int_thread = self.list_info[2][1]
            float_cpuper = float2str2p2(self.list_info[2][2])
            label02text = f"Receiver Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%"

            float_memory = float2str3p2(
                round(self.list_info[3][0] + self.list_info[4][0] + self.list_info[5][0] + self.list_info[6][0], 2)
            )
            int_thread = self.list_info[3][1] + self.list_info[4][1] + self.list_info[5][1] + self.list_info[6][1]
            float_cpuper = float2str2p2(
                round(self.list_info[3][2] + self.list_info[4][2] + self.list_info[5][2] + self.list_info[6][2], 2)
            )
            label07text = f"Collector Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%"

            float_memory = float2str3p2(self.list_info[7][0] + self.list_info[8][0])
            int_thread = self.list_info[7][1] + self.list_info[8][1]
            float_cpuper = float2str2p2(self.list_info[7][2] + self.list_info[8][2])
            label03text = f"Hoga Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%"

            float_memory = float2str3p2(
                round(self.list_info[9][0] + self.list_info[10][0] + self.list_info[11][0] +
                      self.list_info[12][0] + self.list_info[13][0] + self.list_info[14][0] +
                      self.list_info[15][0] + self.list_info[16][0] + self.list_info[17][0], 2)
            )
            int_thread = \
                self.list_info[9][1] + self.list_info[10][1] + self.list_info[11][1] + \
                self.list_info[12][1] + self.list_info[13][1] + self.list_info[14][1] + \
                self.list_info[15][1] + self.list_info[16][1] + self.list_info[17][1]
            float_cpuper = float2str2p2(
                round(self.list_info[9][2] + self.list_info[10][2] + self.list_info[11][2] +
                      self.list_info[12][2] + self.list_info[13][2] + self.list_info[14][2] +
                      self.list_info[15][2] + self.list_info[16][2] + self.list_info[17][2], 2)
            )
            label08text = f"Chart Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%"

            float_memory = float2str3p2(self.list_info[0][0])
            int_thread = self.list_info[0][1]
            float_cpuper = float2str2p2(self.list_info[0][2])
            label04text = f"UI Process - Memory {float_memory}MB | Threads {int_thread}EA | CPU {float_cpuper}%"

            total_memory = msg[1]
            for info in self.list_info:
                total_memory += info[0]
            total_memory = float2str3p2(round(total_memory, 2))
            total_threads = msg[2]
            for info in self.list_info:
                total_threads += info[1]
            total_cpuper = msg[3]
            for info in self.list_info:
                total_cpuper += info[2]
            total_cpuper = float2str2p2(round(total_cpuper, 2))
            label05text = f'Total Process - Memory {total_memory}MB | Threads {total_threads}EA | CPU {total_cpuper}%'

            chartq_size = chart1Q.qsize() + chart2Q.qsize() + chart3Q.qsize() + chart4Q.qsize() + chart5Q.qsize()
            chartq_size += chart6Q.qsize() + chart7Q.qsize() + chart8Q.qsize() + chart9Q.qsize()
            hogaq_size = hoga1Q.qsize() + hoga2Q.qsize()
            label09text = f'Queue - windowQ {windowQ.qsize()} | workerQ {traderQ.qsize()} | stgQ {stgQ.qsize()} | '\
                          f'chartQ {chartq_size} | hogaQ {hogaq_size} | queryQ {queryQ.qsize()} | '\
                          f'soundQ {soundQ.qsize()} | teleQ {teleQ.qsize()}'

            if self.mode2 == 0:
                self.info_label_01.setText(label01text)
                self.info_label_02.setText(label02text)
                self.info_label_03.setText(label03text)
                self.info_label_04.setText(label04text)
                self.info_label_06.setText(label06text)
                self.info_label_07.setText(label07text)
                self.info_label_08.setText(label08text)
                self.info_label_09.setText(label09text)
            self.info_label_05.setText(label05text)
            self.UpdateSysinfo()
        elif msg[0] == 6:
            self.list_info[1] = [msg[1], msg[2], msg[3]]
        elif msg[0] == 7:
            self.list_info[2] = [msg[1], msg[2], msg[3]]
        elif msg[0] == 8:
            self.list_info[3] = [msg[1], msg[2], msg[3]]
        elif msg[0] == 9:
            self.list_info[4] = [msg[1], msg[2], msg[3]]
        elif msg[0] == 10:
            self.list_info[5] = [msg[1], msg[2], msg[3]]
        elif msg[0] == 11:
            self.list_info[6] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['호가P0']:
            self.list_info[7] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['호가P1']:
            self.list_info[8] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P1']:
            self.list_info[9] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P2']:
            self.list_info[10] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P3']:
            self.list_info[11] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P4']:
            self.list_info[12] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P5']:
            self.list_info[13] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P6']:
            self.list_info[14] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P7']:
            self.list_info[15] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P8']:
            self.list_info[16] = [msg[1], msg[2], msg[3]]
        elif msg[0] == ui_num['차트P9']:
            self.list_info[17] = [msg[1], msg[2], msg[3]]

    def UpdateChart(self, data):
        gubun = data[0]
        df = data[1]

        if self.mode2 != 0:
            return
        if gubun in [ui_num['차트P1'], ui_num['차트P2']] and (self.mode0 != 0 or self.mode1 not in [0, 1]):
            return
        if gubun in [ui_num['차트P3'], ui_num['차트P4']] and (self.mode0 != 0 or self.mode1 != 0):
            return
        if gubun in [ui_num['차트P6'], ui_num['차트P7'], ui_num['차트P8'], ui_num['차트P9']] and \
                (self.mode0 != 1 or self.mode1 != 0):
            return
        if gubun == ui_num['차트P5'] and self.mode1 != 2:
            return

        def crosshair(yminn, pc, main_pg=None, sub_pg=None):
            if main_pg is not None:
                vLine1 = pg.InfiniteLine()
                vLine1.setPen(pg.mkPen(color_fg_bk, width=1))
                hLine = pg.InfiniteLine(angle=0)
                hLine.setPen(pg.mkPen(color_fg_bk, width=1))
                main_pg.addItem(vLine1, ignoreBounds=True)
                main_pg.addItem(hLine, ignoreBounds=True)
                main_vb = main_pg.getViewBox()
                label = pg.TextItem(anchor=(0, 1), color=color_fg_bt, border=color_bg_bt, fill=color_bg_ld)
                label.setFont(qfont12)
                label.setPos(-0.25, yminn)
                main_pg.addItem(label)
            if sub_pg is not None:
                vLine2 = pg.InfiniteLine()
                vLine2.setPen(pg.mkPen(color_fg_bk, width=1))
                sub_pg.addItem(vLine2, ignoreBounds=True)
                sub_vb = sub_pg.getViewBox()

            def mouseMoved(evt):
                pos = evt[0]
                if main_pg is not None and main_pg.sceneBoundingRect().contains(pos):
                    mousePoint = main_vb.mapSceneToView(pos)
                    per = round((mousePoint.y() / pc - 1) * 100, 2)
                    label.setText(f"십자선 {format(int(mousePoint.y()), ',')}\n등락율 {per}%")
                    vLine1.setPos(mousePoint.x())
                    hLine.setPos(mousePoint.y())
                    if sub_pg is not None:
                        vLine2.setPos(mousePoint.x())
                if sub_pg is not None and sub_pg.sceneBoundingRect().contains(pos):
                    mousePoint = sub_vb.mapSceneToView(pos)
                    vLine1.setPos(mousePoint.x())
                    vLine2.setPos(mousePoint.x())
            if main_pg is not None:
                main_pg.proxy = pg.SignalProxy(main_pg.scene().sigMouseMoved, rateLimit=20, slot=mouseMoved)
            if sub_pg is not None:
                sub_pg.proxy = pg.SignalProxy(main_pg.scene().sigMouseMoved, rateLimit=20, slot=mouseMoved)

        def getMainLegendText():
            cc = df['현재가'][-1]
            per = round((c / df['전일종가'][0] - 1) * 100, 2)
            nlist = [ui_num['차트P2'], ui_num['차트P4'], ui_num['차트P5'], ui_num['차트P7'], ui_num['차트P9']]
            if gubun in nlist:
                ema05 = df['지수이평05'][-1]
                ema10 = df['지수이평10'][-1]
                ema20 = df['지수이평20'][-1]
                textt = f"05이평 {format(ema05, ',')}\n10이평 {format(ema10, ',')}\n" \
                        f"20이평 {format(ema20, ',')}\n현재가  {format(cc, ',')}\n등락율  {per}%"
            else:
                ema05 = df['지수이평05'][-1]
                ema20 = df['지수이평20'][-1]
                ema60 = df['지수이평60'][-1]
                textt = f"05이평 {format(ema05, ',')}\n20이평 {format(ema20, ',')}\n" \
                        f"60이평 {format(ema60, ',')}\n현재가  {format(cc, ',')}\n등락율  {per}%"
            return textt

        def getSubLegendText():
            money = int(df['거래량'][-1])
            per = round(df['거래량'][-1] / df['거래량'][-2] * 100, 2)
            textt = f"거래량 {format(money, ',')}\n증감비 {per}%"
            return textt

        x = len(df) - 1
        c = df['현재가'][-1]
        o = df['시가'][-1]
        prec = df['전일종가'][0]
        v = df['거래량'][-1]
        vmax = df['거래량'].max()
        name = df['종목명'][0]

        if gubun in [ui_num['차트P1'], ui_num['차트P3']]:
            ymin = min(df['저가'].min(), df['지수이평05'].min(), df['지수이평20'].min(), df['지수이평60'].min(),
                       df['지수이평120'].min(), df['지수이평240'].min(), df['지수이평480'].min())
            ymax = max(df['고가'].max(), df['지수이평05'].max(), df['지수이평20'].max(), df['지수이평60'].max(),
                       df['지수이평120'].max(), df['지수이평240'].max(), df['지수이평480'].max())
        elif gubun in [ui_num['차트P2'], ui_num['차트P4'], ui_num['차트P5'], ui_num['차트P7'], ui_num['차트P9']]:
            ymin = min(df['저가'].min(), df['지수이평05'].min(), df['지수이평10'].min(), df['지수이평20'].min(),
                       df['지수이평40'].min(), df['지수이평60'].min(), df['지수이평120'].min())
            ymax = max(df['고가'].max(), df['지수이평05'].max(), df['지수이평10'].max(), df['지수이평20'].max(),
                       df['지수이평40'].max(), df['지수이평60'].max(), df['지수이평120'].max())
        else:
            ymin = min(df['저가'].min(), df['지수이평05'].min(), df['지수이평20'].min(), df['지수이평60'].min())
            ymax = max(df['고가'].max(), df['지수이평05'].max(), df['지수이평20'].max(), df['지수이평60'].max())

        if gubun not in self.dict_mcpg_lastindex.keys() or self.dict_mcpg_lastindex[gubun] != df.index[-1] or \
                gubun not in self.dict_mcpg_name.keys() or self.dict_mcpg_name[gubun] != name:
            self.dict_ctpg[gubun][0].clear()
            self.dict_ctpg[gubun][1].clear()
            self.dict_mcpg_lastindex[gubun] = df.index[-1]
            self.dict_ctpg[gubun][0].addItem(ChuseItem(df, ymin, ymax))
            self.dict_ctpg[gubun][0].addItem(MoveavgItem(df, gubun))
            self.dict_ctpg[gubun][0].addItem(CandlestickItem(df))
            self.dict_mcpg_lastchuse[gubun] = LastChuseItem(df, ymin, ymax)
            self.dict_mcpg_lastmoveavg[gubun] = LastMoveavgItem(df, gubun)
            self.dict_mcpg_lastcandlestick[gubun] = LastCandlestickItem(df)
            self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_lastchuse[gubun])
            self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_lastmoveavg[gubun])
            self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_lastcandlestick[gubun])
            if gubun == ui_num['차트P5'] and self.mode1 == 2:
                for i, index2 in enumerate(df.index):
                    if df['매수체결가'][index2] != '':
                        for price in df['매수체결가'][index2].split(';'):
                            arrow = pg.ArrowItem(angle=-180, tipAngle=30, baseAngle=20, headLen=20, tailLen=10,
                                                 tailWidth=2, pen=None, brush='r')
                            arrow.setPos(i, int(price))
                            self.dict_ctpg[gubun][0].addItem(arrow)
                            text = pg.TextItem(anchor=(1, 0.5), color=color_fg_bt, border=color_bg_bt, fill=color_bg_ld)
                            text.setFont(qfont12)
                            text.setPos(i - 1, int(price))
                            text.setText(price)
                            self.dict_ctpg[gubun][0].addItem(text)
                    if df['매도체결가'][index2] != '':
                        for price in df['매도체결가'][index2].split(';'):
                            arrow = pg.ArrowItem(angle=-180, tipAngle=30, baseAngle=20, headLen=20, tailLen=10,
                                                 tailWidth=2, pen=None, brush='b')
                            arrow.setPos(i, int(price))
                            self.dict_ctpg[gubun][0].addItem(arrow)
                            text = pg.TextItem(anchor=(1, 0.5), color=color_fg_bt, border=color_bg_bt, fill=color_bg_ld)
                            text.setFont(qfont12)
                            text.setPos(i - 1, int(price))
                            text.setText(price)
                            self.dict_ctpg[gubun][0].addItem(text)
            self.dict_mcpg_infiniteline[gubun] = pg.InfiniteLine(angle=0)
            self.dict_mcpg_infiniteline[gubun].setPen(pg.mkPen(color_cifl))
            self.dict_mcpg_infiniteline[gubun].setPos(c)
            self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_infiniteline[gubun])
            xticks = [list(zip(range(len(df.index))[::12], df.index[::12]))]
            self.dict_ctpg[gubun][0].getAxis('bottom').setTicks(xticks)
            self.dict_ctpg[gubun][1].addItem(VolumeBarsItem(df))
            self.dict_mcpg_lastmoneybar[gubun] = LastVolumeBarItem(x, c, o, v)
            self.dict_ctpg[gubun][1].addItem(self.dict_mcpg_lastmoneybar[gubun])
            self.dict_ctpg[gubun][1].getAxis('bottom').setLabel(text=name)
            self.dict_ctpg[gubun][1].getAxis('bottom').setTicks(xticks)
            crosshair(ymin, prec, main_pg=self.dict_ctpg[gubun][0], sub_pg=self.dict_ctpg[gubun][1])
            self.dict_mcpg_legend1[gubun] = pg.TextItem(color=color_fg_bt, border=color_bg_bt, fill=color_bg_ld)
            self.dict_mcpg_legend1[gubun].setFont(qfont12)
            self.dict_mcpg_legend1[gubun].setPos(-0.25, ymax)
            self.dict_mcpg_legend1[gubun].setText(getMainLegendText())
            self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_legend1[gubun])
            self.dict_mcpg_legend2[gubun] = pg.TextItem(color=color_fg_bt, border=color_bg_bt, fill=color_bg_ld)
            self.dict_mcpg_legend2[gubun].setFont(qfont12)
            self.dict_mcpg_legend2[gubun].setPos(-0.25, vmax)
            self.dict_mcpg_legend2[gubun].setText(getSubLegendText())
            self.dict_ctpg[gubun][1].addItem(self.dict_mcpg_legend2[gubun])
            if gubun not in self.dict_mcpg_name.keys() or self.dict_mcpg_name[gubun] != name:
                self.dict_ctpg[gubun][0].enableAutoRange(enable=True)
                self.dict_ctpg[gubun][1].enableAutoRange(enable=True)
                self.dict_mcpg_name[gubun] = name
        else:
            if gubun not in self.dict_mcpg_close.keys() or self.dict_mcpg_close[gubun] != c:
                self.dict_ctpg[gubun][0].removeItem(self.dict_mcpg_lastchuse[gubun])
                self.dict_ctpg[gubun][0].removeItem(self.dict_mcpg_lastmoveavg[gubun])
                self.dict_ctpg[gubun][0].removeItem(self.dict_mcpg_lastcandlestick[gubun])
                self.dict_mcpg_lastchuse[gubun] = LastChuseItem(df, ymin, ymax)
                self.dict_mcpg_lastmoveavg[gubun] = LastMoveavgItem(df, gubun)
                self.dict_mcpg_lastcandlestick[gubun] = LastCandlestickItem(df)
                self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_lastchuse[gubun])
                self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_lastmoveavg[gubun])
                self.dict_ctpg[gubun][0].addItem(self.dict_mcpg_lastcandlestick[gubun])
                self.dict_mcpg_infiniteline[gubun].setPos(c)
                self.dict_mcpg_legend1[gubun].setText(getMainLegendText())
                self.dict_mcpg_close[gubun] = c
            self.dict_ctpg[gubun][1].removeItem(self.dict_mcpg_lastmoneybar[gubun])
            self.dict_mcpg_lastmoneybar[gubun] = LastVolumeBarItem(x, c, o, v)
            self.dict_ctpg[gubun][1].addItem(self.dict_mcpg_lastmoneybar[gubun])
            self.dict_mcpg_legend2[gubun].setText(getSubLegendText())

    def UpdateGaonsimJongmok(self, data):
        gubun = data[0]
        dict_df = data[1]

        if gubun == ui_num['관심종목'] and self.table_tabWidget.currentWidget() != self.gj_tab:
            return

        if len(dict_df) == 0:
            if self.gj_tableWidget.item(0, 0) is not None:
                self.gj_tableWidget.clearContents()
            return

        def changeFormat(text):
            text = str(text)
            try:
                format_data = format(int(text), ',')
            except ValueError:
                format_data = format(float(text), ',')
                if len(format_data.split('.')) >= 2 and len(format_data.split('.')[1]) == 1:
                    format_data += '0'
            return format_data

        self.gj_tableWidget.setRowCount(len(dict_df))
        for j, code in enumerate(list(dict_df.keys())):
            item = QtWidgets.QTableWidgetItem(self.dict_name[code])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self.gj_tableWidget.setItem(j, 0, item)

            smavg = dict_df[code]['거래대금'][DICT_SET['평균시간'] + 1]
            item = QtWidgets.QTableWidgetItem(changeFormat(smavg).split('.')[0])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.gj_tableWidget.setItem(j, columns_gj3.index('smavg'), item)

            chavg = dict_df[code]['체결강도'][DICT_SET['평균시간'] + 1]
            item = QtWidgets.QTableWidgetItem(changeFormat(chavg))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.gj_tableWidget.setItem(j, columns_gj3.index('chavg'), item)

            chhigh = dict_df[code]['최고체결강도'][DICT_SET['평균시간'] + 1]
            item = QtWidgets.QTableWidgetItem(changeFormat(chhigh))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.gj_tableWidget.setItem(j, columns_gj3.index('chhigh'), item)

            for i, column in enumerate(columns_gj2):
                if column in ['거래대금', '누적거래대금']:
                    item = QtWidgets.QTableWidgetItem(changeFormat(dict_df[code][column][0]).split('.')[0])
                else:
                    item = QtWidgets.QTableWidgetItem(changeFormat(dict_df[code][column][0]))
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                if column == '등락율':
                    if DICT_SET['등락율하한'] <= dict_df[code][column][0] <= DICT_SET['등락율상한']:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif column == '고저평균대비등락율':
                    if dict_df[code][column][0] >= 0:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif column == '거래대금':
                    if dict_df[code][column][0] >= smavg + DICT_SET['거래대금차이']:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif column == '누적거래대금':
                    if dict_df[code][column][0] >= DICT_SET['누적거래대금하한']:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif column == '체결강도':
                    if dict_df[code][column][0] >= DICT_SET['체결강도하한'] and \
                            dict_df[code][column][0] >= chavg + DICT_SET['체결강도차이']:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                self.gj_tableWidget.setItem(j, i + 1, item)

        if len(dict_df) < 46:
            self.gj_tableWidget.setRowCount(46)

    def UpdateTablewidget(self, data):
        gubun = data[0]
        df = data[1]

        if gubun == ui_num['체결강도'] and (self.mode2 != 0 or self.mode1 != 1):
            return
        nlist = [ui_num['호가잔고0'], ui_num['매도주문0'], ui_num['체결수량0'], ui_num['호가0'], ui_num['매수주문0']]
        if gubun in nlist and (self.mode2 != 0 or self.mode1 not in [0, 1]):
            return
        nlist = [ui_num['호가잔고1'], ui_num['매도주문1'], ui_num['체결수량1'], ui_num['호가1'], ui_num['매수주문1']]
        if gubun in nlist and (self.mode2 != 0 or self.mode1 != 0):
            return

        tableWidget = None
        if gubun == ui_num['거래합계']:
            tableWidget = self.tt_tableWidget
        elif gubun == ui_num['거래목록']:
            tableWidget = self.td_tableWidget
        elif gubun == ui_num['잔고평가']:
            tableWidget = self.tj_tableWidget
        elif gubun == ui_num['잔고목록']:
            tableWidget = self.jg_tableWidget
        elif gubun == ui_num['체결목록']:
            tableWidget = self.cj_tableWidget
        elif gubun == ui_num['기업공시']:
            tableWidget = self.gs_tableWidget
        elif gubun == ui_num['기업뉴스']:
            tableWidget = self.ns_tableWidget
        elif gubun == ui_num['투자자']:
            tableWidget = self.jj_tableWidget
        elif gubun == ui_num['재무년도']:
            tableWidget = self.jm1_tableWidget
            tableWidget.setHorizontalHeaderLabels(df.columns)
        elif gubun == ui_num['재무분기']:
            tableWidget = self.jm2_tableWidget
            tableWidget.setHorizontalHeaderLabels(df.columns)
        elif gubun == ui_num['동업종비교']:
            tableWidget = self.jb_tableWidget
            tableWidget.setHorizontalHeaderLabels(df.columns)
        elif gubun == ui_num['체결강도']:
            tableWidget = self.ch_tableWidget
        elif gubun == ui_num['당일합계']:
            tableWidget = self.dt_tableWidget
        elif gubun == ui_num['당일상세']:
            tableWidget = self.dd_tableWidget
        elif gubun == ui_num['누적합계']:
            tableWidget = self.nt_tableWidget
        elif gubun == ui_num['누적상세']:
            tableWidget = self.nd_tableWidget
        elif gubun == ui_num['호가잔고0']:
            tableWidget = self.hoga_00_hj_tableWidget
        elif gubun == ui_num['매도주문0']:
            tableWidget = self.hoga_00_hs_tableWidget
        elif gubun == ui_num['체결수량0']:
            tableWidget = self.hoga_00_hc_tableWidget
        elif gubun == ui_num['호가0']:
            tableWidget = self.hoga_00_hg_tableWidget
        elif gubun == ui_num['매수주문0']:
            tableWidget = self.hoga_00_hb_tableWidget
        elif gubun == ui_num['호가잔고1']:
            tableWidget = self.hoga_01_hj_tableWidget
        elif gubun == ui_num['매도주문1']:
            tableWidget = self.hoga_01_hs_tableWidget
        elif gubun == ui_num['체결수량1']:
            tableWidget = self.hoga_01_hc_tableWidget
        elif gubun == ui_num['호가1']:
            tableWidget = self.hoga_01_hg_tableWidget
        elif gubun == ui_num['매수주문1']:
            tableWidget = self.hoga_01_hb_tableWidget
        if tableWidget is None:
            return

        if len(df) == 0:
            if tableWidget.item(0, 0) is not None:
                tableWidget.clearContents()
            return

        def changeFormat(text):
            text = str(text)
            try:
                format_data = format(int(text), ',')
            except ValueError:
                format_data = format(float(text), ',')
                if len(format_data.split('.')) >= 2 and len(format_data.split('.')[1]) == 1:
                    format_data += '0'
            nnlist = [ui_num['체결수량0'], ui_num['호가0'], ui_num['체결수량1'], ui_num['호가1']]
            if gubun in nnlist and format_data in ['0', '0.00']:
                format_data = ''
            return format_data

        tableWidget.setRowCount(len(df))
        for j, index in enumerate(df.index):
            for i, column in enumerate(df.columns):
                if column == '체결시간':
                    cgtime = str(df[column][index])
                    if gubun == ui_num['체결강도']:
                        cgtime = f'{cgtime[:2]}:{cgtime[2:4]}:{cgtime[4:6]}'
                    else:
                        cgtime = f'{cgtime[8:10]}:{cgtime[10:12]}:{cgtime[12:14]}'
                    item = QtWidgets.QTableWidgetItem(cgtime)
                elif column in ['거래일자', '일자']:
                    day = df[column][index]
                    if '.' not in day:
                        day = day[:4] + '.' + day[4:6] + '.' + day[6:]
                    item = QtWidgets.QTableWidgetItem(day)
                elif column in ['종목명', '주문구분', '호가종목명', '기간', '매도미체결수량', '매수미체결수량',
                                '공시', '정보제공', '언론사', '제목']:
                    item = QtWidgets.QTableWidgetItem(str(df[column][index]))
                elif gubun in [ui_num['재무년도'], ui_num['재무분기'], ui_num['동업종비교']]:
                    try:
                        item = QtWidgets.QTableWidgetItem(str(df[column][index]))
                    except KeyError:
                        continue
                elif column not in ['수익률', '등락율', '고저평균대비등락율', '체결강도',
                                    '체결강도5분', '체결강도20분', '체결강도60분', '최고체결강도']:
                    item = QtWidgets.QTableWidgetItem(changeFormat(df[column][index]).split('.')[0])
                else:
                    item = QtWidgets.QTableWidgetItem(changeFormat(df[column][index]))

                if column in ['종목명', '호가종목명', '공시', '제목', '구분']:
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                elif column in ['거래횟수', '추정예탁자산', '추정예수금', '보유종목수', '주문구분', '체결시간', '거래일자', '기간',
                                '일자', '매도미체결수량', '매도미체결수량', '정보제공', '언론사']:
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)

                if column == '체결수량':
                    if j == 0:
                        item.setIcon(self.icon_totalb)
                    elif j == 21:
                        item.setIcon(self.icon_totals)
                elif column == '체결강도' and gubun in [ui_num['체결수량0'], ui_num['체결수량1']]:
                    if j == 0:
                        item.setIcon(self.icon_up)
                    elif j == 21:
                        item.setIcon(self.icon_down)
                elif gubun in [ui_num['호가0'], ui_num['호가1']]:
                    if column == '증감':
                        if j == 0:
                            item.setIcon(self.icon_perb)
                        elif j == 21:
                            item.setIcon(self.icon_pers)
                    elif column == '잔량':
                        if j == 0:
                            item.setIcon(self.icon_totalb)
                        elif j == 21:
                            item.setIcon(self.icon_totals)
                    elif column == '호가':
                        if j == 0:
                            item.setIcon(self.icon_up)
                        elif j == 21:
                            item.setIcon(self.icon_down)
                        else:
                            if gubun == ui_num['호가0']:
                                hj_tableWidget = self.hoga_00_hj_tableWidget
                            else:
                                hj_tableWidget = self.hoga_01_hj_tableWidget
                            if hj_tableWidget.item(0, 0) is not None:
                                o = comma2int(hj_tableWidget.item(0, columns_hj.index('시가')).text())
                                h = comma2int(hj_tableWidget.item(0, columns_hj.index('고가')).text())
                                low = comma2int(hj_tableWidget.item(0, columns_hj.index('저가')).text())
                                if o != 0:
                                    if df[column][index] == o:
                                        item.setIcon(self.icon_open)
                                    elif df[column][index] == h:
                                        item.setIcon(self.icon_high)
                                    elif df[column][index] == low:
                                        item.setIcon(self.icon_low)
                    elif column == '등락율':
                        if j == 0:
                            item.setIcon(self.icon_up)
                        elif j == 21:
                            item.setIcon(self.icon_down)
                        else:
                            if gubun == ui_num['호가0']:
                                hj_tableWidget = self.hoga_00_hj_tableWidget
                            else:
                                hj_tableWidget = self.hoga_01_hj_tableWidget
                            if hj_tableWidget.item(0, 0) is not None:
                                uvi = comma2int(hj_tableWidget.item(0, columns_hj.index('UVI')).text())
                                dvi = comma2int(hj_tableWidget.item(0, columns_hj.index('DVI')).text())
                                if df[column][index] != 0:
                                    if j < 11:
                                        if df['호가'][index] == uvi:
                                            item.setIcon(self.icon_vi)
                                    else:
                                        if df['호가'][index] == dvi:
                                            item.setIcon(self.icon_vi)

                if '수익률' in df.columns:
                    if df['수익률'][index] >= 0:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif gubun == ui_num['체결목록']:
                    if df['주문구분'][index] == '매수':
                        item.setForeground(color_fg_bt)
                    elif df['주문구분'][index] == '매도':
                        item.setForeground(color_fg_dk)
                    elif df['주문구분'][index] in ['매도취소', '매수취소']:
                        item.setForeground(color_fg_bc)
                elif gubun in [ui_num['기업공시'], ui_num['기업뉴스']]:
                    cname = '공시' if gubun == ui_num['기업공시'] else '제목'
                    if '단기과열' in df[cname][index] or '투자주의' in df[cname][index] or \
                            '투자경고' in df[cname][index] or '투자위험' in df[cname][index] or \
                            '거래정지' in df[cname][index] or '환기종목' in df[cname][index] or \
                            '불성실공시' in df[cname][index] or '관리종목' in df[cname][index] or \
                            '정리매매' in df[cname][index] or '유상증자' in df[cname][index] or \
                            '무상증자' in df[cname][index]:
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif gubun == ui_num['투자자']:
                    if column in ['등락율', '개인투자자', '외국인투자자', '기관계']:
                        if df[column][index] >= 0:
                            item.setForeground(color_fg_bt)
                        else:
                            item.setForeground(color_fg_dk)
                elif gubun in [ui_num['재무년도'],  ui_num['재무분기'],  ui_num['동업종비교']]:
                    if '-' not in df[column][index] and column != '구분':
                        item.setForeground(color_fg_bt)
                    else:
                        item.setForeground(color_fg_dk)
                elif gubun == ui_num['체결강도']:
                    if column == '등락율':
                        if df[column][index] >= 0:
                            item.setForeground(color_fg_bt)
                        else:
                            item.setForeground(color_fg_dk)
                    elif '체결강도' in column:
                        if df[column][index] >= 100:
                            item.setForeground(color_fg_bt)
                        else:
                            item.setForeground(color_fg_dk)
                elif gubun in [ui_num['체결수량0'], ui_num['체결수량1']]:
                    if column == '체결수량':
                        if j == 0:
                            if df[column][index] > df[column][21]:
                                item.setForeground(color_fg_bt)
                            else:
                                item.setForeground(color_fg_dk)
                        elif j == 21:
                            if df[column][index] > df[column][0]:
                                item.setForeground(color_fg_bt)
                            else:
                                item.setForeground(color_fg_dk)
                        else:
                            if gubun == ui_num['체결수량0']:
                                hg_tableWidget = self.hoga_00_hg_tableWidget
                            else:
                                hg_tableWidget = self.hoga_01_hg_tableWidget
                            if hg_tableWidget.item(0, 0) is not None and \
                                    hg_tableWidget.item(10, columns_hg.index('호가')).text() != '':
                                c = comma2int(hg_tableWidget.item(10, columns_hg.index('호가')).text())
                                if df[column][index] > 0:
                                    item.setForeground(color_fg_bt)
                                    if df[column][index] * c > 90000000:
                                        item.setBackground(color_bf_bt)
                                elif df[column][index] < 0:
                                    item.setForeground(color_fg_dk)
                                    if df[column][index] * c < -90000000:
                                        item.setBackground(color_bf_dk)
                    elif column == '체결강도':
                        if df[column][index] >= 100:
                            item.setForeground(color_fg_bt)
                        else:
                            item.setForeground(color_fg_dk)
                elif gubun in [ui_num['호가0'], ui_num['호가1']]:
                    if '증감' in column:
                        if j == 0:
                            if df[column][index] > 100:
                                item.setForeground(color_fg_bt)
                            else:
                                item.setForeground(color_fg_dk)
                        elif j == 21:
                            if df[column][index] > 100:
                                item.setForeground(color_fg_bt)
                            else:
                                item.setForeground(color_fg_dk)
                        elif df[column][index] > 0:
                            item.setForeground(color_fg_bt)
                            if df[column][index] * df['호가'][10] > 90000000:
                                item.setBackground(color_bf_bt)
                        elif df[column][index] < 0:
                            item.setForeground(color_fg_dk)
                            if df[column][index] * df['호가'][11] < -90000000:
                                item.setBackground(color_bf_dk)
                    elif column == '잔량':
                        if j == 0:
                            if df[column][index] > df[column][21]:
                                item.setForeground(color_fg_bt)
                            else:
                                item.setForeground(color_fg_dk)
                        elif j == 21:
                            if df[column][index] > df[column][0]:
                                item.setForeground(color_fg_bt)
                            else:
                                item.setForeground(color_fg_dk)
                        elif j < 11:
                            item.setForeground(color_fg_bt)
                        else:
                            item.setForeground(color_fg_dk)
                    elif column in ['호가', '등락율']:
                        if df['등락율'][index] > 0:
                            item.setForeground(color_fg_bt)
                        elif df['등락율'][index] < 0:
                            item.setForeground(color_fg_dk)
                        if column == '호가' and df[column][index] != 0:
                            if gubun == ui_num['호가0']:
                                hj_tableWidget = self.hoga_00_hj_tableWidget
                            else:
                                hj_tableWidget = self.hoga_01_hj_tableWidget
                            if hj_tableWidget.item(0, 0) is not None:
                                c = comma2int(hj_tableWidget.item(0, columns_hj.index('현재가')).text())
                                if j not in [0, 21] and df[column][index] == c:
                                    item.setBackground(color_bf_bt)
                                if hj_tableWidget.item(0, columns_hj.index('매입가')).text() != '0':
                                    gap = df[column][19] - df[column][20]
                                    buyprice = comma2int(hj_tableWidget.item(0, columns_hj.index('매입가')).text())
                                    if df[column][index] <= buyprice < df[column][index] + gap:
                                        item.setBackground(color_bf_dk)
                elif gubun in [ui_num['매도주문0'], ui_num['매도주문1'], ui_num['매수주문0'], ui_num['매수주문1']]:
                    item.setForeground(color_fg_bt)
                    item.setBackground(color_bg_bt)
                tableWidget.setItem(j, i, item)

        if len(df) < 13 and gubun in [ui_num['거래목록'], ui_num['잔고목록'], ui_num['체결목록']]:
            tableWidget.setRowCount(13)
        elif len(df) < 22 and gubun == ui_num['기업공시']:
            tableWidget.setRowCount(22)
        elif len(df) < 12 and gubun == ui_num['기업뉴스']:
            tableWidget.setRowCount(12)
        elif len(df) < 28 and gubun == ui_num['체결강도']:
            tableWidget.setRowCount(28)
        elif len(df) < 31 and gubun == ui_num['당일상세']:
            tableWidget.setRowCount(31)
        elif len(df) < 41 and gubun == ui_num['누적상세']:
            tableWidget.setRowCount(41)

    @QtCore.pyqtSlot(int)
    def CellClicked_1(self, row):
        item = self.hoga_00_hj_tableWidget.item(0, 0)
        if item is None:
            return
        name = item.text()
        if self.hoga_00_hj_tableWidget.item(0, columns_jg.index('보유수량')).text() == '':
            return
        jc = comma2int(self.hoga_00_hj_tableWidget.item(0, columns_jg.index('보유수량')).text())
        if self.hoga_00_hg_tableWidget.item(row, columns_hg.index('호가')).text() == '':
            return
        hg = comma2int(self.hoga_00_hg_tableWidget.item(row, columns_hg.index('호가')).text())

        bper = 0
        if self.hoga_00_sell_radioButton_01.isChecked():
            bper = 10
        elif self.hoga_00_sell_radioButton_02.isChecked():
            bper = 25
        elif self.hoga_00_sell_radioButton_03.isChecked():
            bper = 33
        elif self.hoga_00_sell_radioButton_04.isChecked():
            bper = 50
        elif self.hoga_00_sell_radioButton_05.isChecked():
            bper = 75
        elif self.hoga_00_sell_radioButton_06.isChecked():
            bper = 100
        if bper == 0:
            windowQ.put([2, '시스템 명령 오류 알림 - 매도비율을 선택하십시오.'])
            return

        oc = int(jc * (bper / 100))
        if oc == 0:
            oc = 1
        order = ['매도', self.dict_code[name], name, hg, oc]
        traderQ.put(order)

    @QtCore.pyqtSlot(int)
    def CellClicked_2(self, row):
        item = self.hoga_00_hj_tableWidget.item(0, 0)
        if item is None:
            return
        name = item.text()
        if self.hoga_00_hg_tableWidget.item(row, columns_hg.index('호가')).text() == '':
            return
        hg = comma2int(self.hoga_00_hg_tableWidget.item(row, columns_hg.index('호가')).text())

        og = 0
        if self.hoga_00_buy_radioButton_01.isChecked():
            og = 100000
        elif self.hoga_00_buy_radioButton_02.isChecked():
            og = 500000
        elif self.hoga_00_buy_radioButton_03.isChecked():
            og = 1000000
        elif self.hoga_00_buy_radioButton_04.isChecked():
            og = 5000000
        elif self.hoga_00_buy_radioButton_05.isChecked():
            og = 10000000
        elif self.hoga_00_buy_radioButton_06.isChecked():
            og = 50000000
        if og == 0:
            windowQ.put([2, '시스템 명령 오류 알림 - 매수금액을 선택하십시오.'])
            return

        oc = int(og / hg)
        order = ['매수', self.dict_code[name], name, hg, oc]
        traderQ.put(order)

    @QtCore.pyqtSlot(int)
    def CellClicked_3(self, row):
        item = self.hoga_01_hj_tableWidget.item(0, 0)
        if item is None:
            return
        name = item.text()
        if self.hoga_01_hj_tableWidget.item(0, columns_jg.index('보유수량')).text() == '':
            return
        jc = comma2int(self.hoga_01_hj_tableWidget.item(0, columns_jg.index('보유수량')).text())
        if self.hoga_01_hg_tableWidget.item(row, columns_hg.index('호가')).text() == '':
            return
        hg = comma2int(self.hoga_01_hg_tableWidget.item(row, columns_hg.index('호가')).text())

        bper = 0
        if self.hoga_01_sell_radioButton_01.isChecked():
            bper = 10
        elif self.hoga_01_sell_radioButton_02.isChecked():
            bper = 25
        elif self.hoga_01_sell_radioButton_03.isChecked():
            bper = 33
        elif self.hoga_01_sell_radioButton_04.isChecked():
            bper = 50
        elif self.hoga_01_sell_radioButton_05.isChecked():
            bper = 75
        elif self.hoga_01_sell_radioButton_06.isChecked():
            bper = 100
        if bper == 0:
            windowQ.put([2, '시스템 명령 오류 알림 - 매도비율을 선택하십시오.'])
            return

        oc = int(jc * (bper / 100))
        if oc == 0:
            oc = 1
        order = ['매도', self.dict_code[name], name, hg, oc]
        traderQ.put(order)

    @QtCore.pyqtSlot(int)
    def CellClicked_4(self, row):
        item = self.hoga_01_hj_tableWidget.item(0, 0)
        if item is None:
            return
        name = item.text()
        if self.hoga_01_hg_tableWidget.item(row, columns_hg.index('호가')).text() == '':
            return
        hg = comma2int(self.hoga_01_hg_tableWidget.item(row, columns_hg.index('호가')).text())

        og = 0
        if self.hoga_01_buy_radioButton_01.isChecked():
            og = 100000
        elif self.hoga_01_buy_radioButton_02.isChecked():
            og = 500000
        elif self.hoga_01_buy_radioButton_03.isChecked():
            og = 1000000
        elif self.hoga_01_buy_radioButton_04.isChecked():
            og = 5000000
        elif self.hoga_01_buy_radioButton_05.isChecked():
            og = 10000000
        elif self.hoga_01_buy_radioButton_06.isChecked():
            og = 50000000
        if og == 0:
            windowQ.put([2, '시스템 명령 오류 알림 - 매수금액을 선택하십시오.'])
            return

        oc = int(og / hg)
        order = ['매수', self.dict_code[name], name, hg, oc]
        traderQ.put(order)

    @QtCore.pyqtSlot(int, int)
    def CellClicked_5(self, row, col):
        if col > 1:
            return
        item = self.td_tableWidget.item(row, 0)
        if item is None:
            return
        code = self.dict_code[item.text()]
        self.PutTraderQ(code, col)

    @QtCore.pyqtSlot(int, int)
    def CellClicked_6(self, row, col):
        if col > 1:
            return
        item = self.jg_tableWidget.item(row, 0)
        if item is None:
            return
        code = self.dict_code[item.text()]
        self.PutTraderQ(code, col)

    @QtCore.pyqtSlot(int, int)
    def CellClicked_7(self, row, col):
        if col > 1:
            return
        item = self.cj_tableWidget.item(row, 0)
        if item is None:
            return
        code = self.dict_code[item.text()]
        self.PutTraderQ(code, col)

    @QtCore.pyqtSlot(int, int)
    def CellClicked_8(self, row, col):
        if col > 1:
            return
        item = self.gj_tableWidget.item(row, 0)
        if item is None:
            return
        code = self.dict_code[item.text()]
        self.PutTraderQ(code, col)

    @QtCore.pyqtSlot(int, int)
    def CellClicked_9(self, row, col):
        if col > 1:
            return
        item = self.dd_tableWidget.item(row, 1)
        if item is None:
            return
        code = self.dict_code[item.text()]
        self.PutTraderQ(code, col)

    def PutTraderQ(self, code, col):
        if self.mode2 != 0:
            return
        if self.mode1 == 0:
            if self.mode0 == 1:
                self.ButtonClicked_4(0)
            if col == 0:
                traderQ.put(f"현재가{ui_num['차트P1']} {code}")
            elif col == 1:
                traderQ.put(f"현재가{ui_num['차트P3']} {code}")
        elif self.mode1 == 1:
            traderQ.put(f"현재가{ui_num['차트P0']} {code}")
        elif self.mode1 == 2:
            if self.table_tabWidget.currentWidget() == self.st_tab:
                date = self.calendarWidget.selectedDate()
                tradeday = date.toString('yyyyMMdd')
            elif 0 < int(strf_time('%H%M%S')) < 90000:
                tradeday = strf_time('%Y%m%d', timedelta_day(-1))
            else:
                tradeday = strf_time('%Y%m%d')
            traderQ.put(f"현재가{ui_num['차트P5']} {tradeday} {code}")

    def CalendarClicked(self):
        date = self.calendarWidget.selectedDate()
        searchday = date.toString('yyyyMMdd')
        con = sqlite3.connect(db_stg)
        df = pd.read_sql(f"SELECT * FROM tradelist WHERE 체결시간 LIKE '{searchday}%'", con)
        con.close()
        if len(df) > 0:
            df = df.set_index('index')
            df.sort_values(by=['체결시간'], ascending=True, inplace=True)
            df = df[['체결시간', '종목명', '매수금액', '매도금액', '주문수량', '수익률', '수익금']].copy()
            nbg, nsg = df['매수금액'].sum(), df['매도금액'].sum()
            sp = round((nsg / nbg - 1) * 100, 2)
            npg, nmg, nsig = df[df['수익금'] > 0]['수익금'].sum(), df[df['수익금'] < 0]['수익금'].sum(), df['수익금'].sum()
            df2 = pd.DataFrame(columns=columns_dt)
            df2.at[0] = searchday, nbg, nsg, npg, nmg, sp, nsig
        else:
            df = pd.DataFrame(columns=columns_dd)
            df2 = pd.DataFrame(columns=columns_dt)
        windowQ.put([ui_num['당일합계'], df2])
        windowQ.put([ui_num['당일상세'], df])

    def ButtonClicked_1(self, gubun):
        if gubun in ['시장가매도0', '매도취소0']:
            hj_tableWidget = self.hoga_00_hj_tableWidget
            hg_tableWidget = self.hoga_00_hg_tableWidget
            sell_radioButton_01 = self.hoga_00_sell_radioButton_01
            sell_radioButton_02 = self.hoga_00_sell_radioButton_02
            sell_radioButton_03 = self.hoga_00_sell_radioButton_03
            sell_radioButton_04 = self.hoga_00_sell_radioButton_04
            sell_radioButton_05 = self.hoga_00_sell_radioButton_05
            sell_radioButton_06 = self.hoga_00_sell_radioButton_06
        else:
            hj_tableWidget = self.hoga_01_hj_tableWidget
            hg_tableWidget = self.hoga_01_hg_tableWidget
            sell_radioButton_01 = self.hoga_01_sell_radioButton_01
            sell_radioButton_02 = self.hoga_01_sell_radioButton_02
            sell_radioButton_03 = self.hoga_01_sell_radioButton_03
            sell_radioButton_04 = self.hoga_01_sell_radioButton_04
            sell_radioButton_05 = self.hoga_01_sell_radioButton_05
            sell_radioButton_06 = self.hoga_01_sell_radioButton_06
        item = hj_tableWidget.item(0, 0)
        if item is None:
            return
        code = self.dict_code[item.text()]
        if '시장가매도' in gubun:
            bper = 0
            if sell_radioButton_01.isChecked():
                bper = 1
            elif sell_radioButton_02.isChecked():
                bper = 2
            elif sell_radioButton_03.isChecked():
                bper = 3
            elif sell_radioButton_04.isChecked():
                bper = 5
            elif sell_radioButton_05.isChecked():
                bper = 7.5
            elif sell_radioButton_06.isChecked():
                bper = 10
            if bper == 0:
                windowQ.put([1, '시스템 명령 오류 알림 - 매도비율을 선택하십시오.'])
                return

            c = comma2int(hg_tableWidget.item(11, columns_hg.index('호가')).text())
            if hj_tableWidget.item(0, columns_jg.index('보유수량')).text() == '':
                return
            jc = comma2int(hj_tableWidget.item(0, columns_jg.index('보유수량')).text())
            oc = int(jc * (bper / 10))
            if oc == 0:
                oc = 1
            name = hj_tableWidget.item(0, 0).text()
            order = ['매도', self.dict_code[name], name, c, oc]
            traderQ.put(order)
        elif '매도취소' in gubun:
            order = f'매도취소 {code}'
            traderQ.put(order)

    def ButtonClicked_2(self, gubun):
        if gubun in ['시장가매수0', '매수취소0']:
            hj_tableWidget = self.hoga_00_hj_tableWidget
            hg_tableWidget = self.hoga_00_hg_tableWidget
            buy_radioButton_01 = self.hoga_00_buy_radioButton_01
            buy_radioButton_02 = self.hoga_00_buy_radioButton_02
            buy_radioButton_03 = self.hoga_00_buy_radioButton_03
            buy_radioButton_04 = self.hoga_00_buy_radioButton_04
            buy_radioButton_05 = self.hoga_00_buy_radioButton_05
            buy_radioButton_06 = self.hoga_00_buy_radioButton_06
        else:
            hj_tableWidget = self.hoga_01_hj_tableWidget
            hg_tableWidget = self.hoga_01_hg_tableWidget
            buy_radioButton_01 = self.hoga_01_buy_radioButton_01
            buy_radioButton_02 = self.hoga_01_buy_radioButton_02
            buy_radioButton_03 = self.hoga_01_buy_radioButton_03
            buy_radioButton_04 = self.hoga_01_buy_radioButton_04
            buy_radioButton_05 = self.hoga_01_buy_radioButton_05
            buy_radioButton_06 = self.hoga_01_buy_radioButton_06
        item = hj_tableWidget.item(0, 0)
        if item is None:
            return
        code = self.dict_code[item.text()]
        if '시장가매수' in gubun:
            og = 0
            if buy_radioButton_01.isChecked():
                og = 100000
            elif buy_radioButton_02.isChecked():
                og = 500000
            elif buy_radioButton_03.isChecked():
                og = 1000000
            elif buy_radioButton_04.isChecked():
                og = 5000000
            elif buy_radioButton_05.isChecked():
                og = 10000000
            elif buy_radioButton_06.isChecked():
                og = 50000000
            if og == 0:
                windowQ.put([1, '시스템 명령 오류 알림 - 매수금액을 선택하십시오.'])
                return

            c = comma2int(hg_tableWidget.item(10, columns_hg.index('호가')).text())
            oc = int(og / c)
            name = hj_tableWidget.item(0, 0).text()
            order = ['매수', self.dict_code[name], name, c, oc]
            traderQ.put(order)
        elif '매수취소' in gubun:
            order = f'매수취소 {code}'
            traderQ.put(order)

    def ButtonClicked_3(self, cmd):
        if cmd == '설정':
            text1 = self.sj_lineEdit_01.text()
            text2 = self.sj_lineEdit_02.text()
            if text1 == '' or text2 == '':
                windowQ.put([1, '시스템 명령 오류 알림 - 텔레그램 봇넘버 및 사용자 아이디를 입력하십시오.'])
                return
            traderQ.put(f'{cmd} {text1} {text2}')
        elif '집계' in cmd:
            con = sqlite3.connect(db_stg)
            df = pd.read_sql('SELECT * FROM totaltradelist', con)
            con.close()
            df = df[::-1]
            if len(df) > 0:
                sd = strp_time('%Y%m%d', df['index'][df.index[0]])
                ld = strp_time('%Y%m%d', df['index'][df.index[-1]])
                pr = str((sd - ld).days + 1) + '일'
                nbg, nsg = df['총매수금액'].sum(), df['총매도금액'].sum()
                sp = round((nsg / nbg - 1) * 100, 2)
                npg, nmg = df['총수익금액'].sum(), df['총손실금액'].sum()
                nsig = df['수익금합계'].sum()
                df2 = pd.DataFrame(columns=columns_nt)
                df2.at[0] = pr, nbg, nsg, npg, nmg, sp, nsig
                windowQ.put([ui_num['누적합계'], df2])
            else:
                return
            if cmd == '일별집계':
                df = df.rename(columns={'index': '일자'})
                windowQ.put([ui_num['누적상세'], df])
            elif cmd == '월별집계':
                df['일자'] = df['index'].apply(lambda x: x[:6])
                df2 = pd.DataFrame(columns=columns_nd)
                lastmonth = df['일자'][df.index[-1]]
                month = strf_time('%Y%m')
                while int(month) >= int(lastmonth):
                    df3 = df[df['일자'] == month]
                    if len(df3) > 0:
                        tbg, tsg = df3['총매수금액'].sum(), df3['총매도금액'].sum()
                        sp = round((tsg / tbg - 1) * 100, 2)
                        tpg, tmg = df3['총수익금액'].sum(), df3['총손실금액'].sum()
                        ttsg = df3['수익금합계'].sum()
                        df2.at[month] = month, tbg, tsg, tpg, tmg, sp, ttsg
                    month = str(int(month) - 89) if int(month[4:]) == 1 else str(int(month) - 1)
                windowQ.put([ui_num['누적상세'], df2])
            elif cmd == '연도별집계':
                df['일자'] = df['index'].apply(lambda x: x[:4])
                df2 = pd.DataFrame(columns=columns_nd)
                lastyear = df['일자'][df.index[-1]]
                year = strf_time('%Y')
                while int(year) >= int(lastyear):
                    df3 = df[df['일자'] == year]
                    if len(df3) > 0:
                        tbg, tsg = df3['총매수금액'].sum(), df3['총매도금액'].sum()
                        sp = round((tsg / tbg - 1) * 100, 2)
                        tpg, tmg = df3['총수익금액'].sum(), df3['총손실금액'].sum()
                        ttsg = df3['수익금합계'].sum()
                        df2.at[year] = year, tbg, tsg, tpg, tmg, sp, ttsg
                    year = str(int(year) - 1)
                windowQ.put([ui_num['누적상세'], df2])
        else:
            traderQ.put(f'{cmd}')

    def ButtonClicked_4(self, gubun):
        if (gubun in [0, 1] and self.mode2 == 1) or (gubun == 0 and self.mode1 in [1, 2]):
            windowQ.put([1, '시스템 명령 오류 알림 - 현재 모드에서는 작동하지 않습니다.'])
            return
        if gubun == 0 and self.mode0 == 0 and self.mode1 == 0 and self.mode2 == 0:
            self.chart_00_tabWidget.setCurrentWidget(self.chart_05_tab)
            self.chart_01_tabWidget.setCurrentWidget(self.chart_06_tab)
            self.chart_02_tabWidget.setCurrentWidget(self.chart_07_tab)
            self.chart_03_tabWidget.setCurrentWidget(self.chart_08_tab)
            self.etc_pushButton_00.setStyleSheet(style_bc_dk)
            self.etc_pushButton_00.setFont(qfont12)
            self.ct_label_01.setGeometry(3500, 5, 140, 20)
            self.ct_lineEdit_01.setGeometry(3500, 5, 100, 20)
            self.ct_label_02.setGeometry(3500, 702, 140, 20)
            self.ct_lineEdit_02.setGeometry(3500, 702, 100, 20)
            self.mode0 = 1
        elif gubun == 0 and self.mode0 == 1 and self.mode1 == 0 and self.mode2 == 0:
            self.chart_00_tabWidget.setCurrentWidget(self.chart_00_tab)
            self.chart_01_tabWidget.setCurrentWidget(self.chart_01_tab)
            self.chart_02_tabWidget.setCurrentWidget(self.chart_02_tab)
            self.chart_03_tabWidget.setCurrentWidget(self.chart_03_tab)
            self.etc_pushButton_00.setStyleSheet(style_bc_bt)
            self.etc_pushButton_00.setFont(qfont12)
            self.ct_label_01.setGeometry(1820, 5, 140, 20)
            self.ct_lineEdit_01.setGeometry(1960, 5, 100, 20)
            self.ct_label_02.setGeometry(1820, 702, 140, 20)
            self.ct_lineEdit_02.setGeometry(1960, 702, 100, 20)
            self.mode0 = 0
        elif gubun == 1 and self.mode1 == 0 and self.mode2 == 0:
            self.chart_02_tabWidget.setGeometry(3500, 702, 1025, 692)
            self.chart_03_tabWidget.setGeometry(3500, 702, 1026, 692)
            self.info_label_06.setGeometry(3500, 699, 400, 30)
            self.info_label_07.setGeometry(3500, 699, 400, 30)
            self.info_label_08.setGeometry(3500, 699, 400, 30)
            self.info_label_09.setGeometry(3500, 699, 600, 30)
            self.gg_tabWidget.setGeometry(5, 702, 682, 120)
            self.gs_tabWidget.setGeometry(5, 827, 682, 567)
            self.ns_tabWidget.setGeometry(692, 702, 682, 332)
            self.jj_tabWidget.setGeometry(692, 1039, 682, 355)
            self.jm_tabWidget.setGeometry(1379, 702, 682, 332)
            self.jb_tabWidget.setGeometry(1379, 1039, 682, 355)
            self.ch_tabWidget.setGeometry(2066, 702, 682, 692)
            self.etc_pushButton_01.setStyleSheet(style_bc_dk)
            self.etc_pushButton_01.setFont(qfont12)
            self.chart_00_tabWidget.setCurrentWidget(self.chart_00_tab)
            self.chart_01_tabWidget.setCurrentWidget(self.chart_01_tab)
            self.chart_02_tabWidget.setCurrentWidget(self.chart_02_tab)
            self.chart_03_tabWidget.setCurrentWidget(self.chart_03_tab)
            self.etc_pushButton_00.setStyleSheet(style_bc_bt)
            self.etc_pushButton_00.setFont(qfont12)
            self.mode0 = 0
            self.ct_label_01.setGeometry(1820, 5, 140, 20)
            self.ct_lineEdit_01.setGeometry(1960, 5, 100, 20)
            self.ct_label_02.setGeometry(3500, 702, 65, 20)
            self.ct_lineEdit_02.setGeometry(3500, 702, 100, 20)
            self.mode1 = 1
        elif gubun == 1 and self.mode1 == 1 and self.mode2 == 0:
            self.chart_00_tabWidget.setGeometry(3500, 5, 1025, 692)
            self.chart_01_tabWidget.setGeometry(3500, 5, 1026, 692)
            self.chart_04_tabWidget.setGeometry(5, 5, 2743, 1390)
            self.hoga_00_tabWidget.setGeometry(3500, 5, 682, 692)
            self.hoga_01_tabWidget.setGeometry(3500, 702, 682, 692)
            self.gg_tabWidget.setGeometry(3500, 702, 682, 120)
            self.gs_tabWidget.setGeometry(3500, 827, 682, 567)
            self.ns_tabWidget.setGeometry(3500, 702, 682, 332)
            self.jj_tabWidget.setGeometry(3500, 1039, 682, 355)
            self.jm_tabWidget.setGeometry(3500, 702, 682, 332)
            self.jb_tabWidget.setGeometry(3500, 1039, 682, 355)
            self.ch_tabWidget.setGeometry(3500, 702, 682, 692)
            self.ct_label_01.setGeometry(3500, 5, 140, 20)
            self.ct_lineEdit_01.setGeometry(3500, 5, 100, 20)
            self.mode1 = 2
        elif gubun == 1 and self.mode1 == 2 and self.mode2 == 0:
            self.chart_00_tabWidget.setGeometry(5, 5, 1025, 692)
            self.chart_01_tabWidget.setGeometry(1035, 5, 1026, 692)
            self.chart_02_tabWidget.setGeometry(5, 702, 1025, 692)
            self.chart_03_tabWidget.setGeometry(1035, 702, 1026, 692)
            self.chart_04_tabWidget.setGeometry(3500, 5, 2743, 1390)
            self.chart_00_tabWidget.setCurrentWidget(self.chart_00_tab)
            self.chart_01_tabWidget.setCurrentWidget(self.chart_01_tab)
            self.chart_02_tabWidget.setCurrentWidget(self.chart_02_tab)
            self.chart_03_tabWidget.setCurrentWidget(self.chart_03_tab)
            self.info_label_06.setGeometry(155, 699, 400, 30)
            self.info_label_07.setGeometry(600, 699, 600, 30)
            self.info_label_08.setGeometry(1185, 699, 400, 30)
            self.info_label_09.setGeometry(2145, 699, 600, 30)
            self.etc_pushButton_00.setStyleSheet(style_bc_bt)
            self.etc_pushButton_00.setFont(qfont12)
            self.mode0 = 0
            self.hoga_00_tabWidget.setGeometry(2066, 5, 682, 692)
            self.hoga_01_tabWidget.setGeometry(2066, 702, 682, 692)
            self.etc_pushButton_01.setStyleSheet(style_bc_bt)
            self.etc_pushButton_01.setFont(qfont12)
            self.ct_label_01.setGeometry(1820, 5, 140, 20)
            self.ct_lineEdit_01.setGeometry(1960, 5, 100, 20)
            self.ct_label_02.setGeometry(1820, 702, 140, 20)
            self.ct_lineEdit_02.setGeometry(1960, 702, 100, 20)
            self.mode1 = 0
        elif gubun == 2 and self.mode2 == 0:
            self.setGeometry(2748, 0, 692, 1400)
            self.chart_00_tabWidget.setGeometry(3500, 5, 1025, 692)
            self.chart_01_tabWidget.setGeometry(3500, 702, 1026, 692)
            self.chart_02_tabWidget.setGeometry(3500, 5, 1025, 692)
            self.chart_03_tabWidget.setGeometry(3500, 702, 1026, 692)
            self.chart_04_tabWidget.setGeometry(3500, 5, 2743, 1390)
            self.hoga_00_tabWidget.setGeometry(3500, 5, 682, 692)
            self.hoga_01_tabWidget.setGeometry(3500, 702, 682, 692)
            self.lgsj_tabWidget.setGeometry(5, 5, 682, 282)
            self.table_tabWidget.setGeometry(5, 292, 682, 1103)
            self.info_label_01.setGeometry(3500, 1, 400, 30)
            self.info_label_02.setGeometry(3500, 1, 400, 30)
            self.info_label_03.setGeometry(3500, 1, 400, 30)
            self.info_label_04.setGeometry(3500, 1, 400, 30)
            self.info_label_05.setGeometry(140, 1, 400, 30)
            self.info_label_06.setGeometry(3500, 699, 400, 30)
            self.info_label_07.setGeometry(3500, 699, 400, 30)
            self.info_label_08.setGeometry(3500, 699, 400, 30)
            self.info_label_09.setGeometry(3500, 699, 600, 30)
            self.etc_pushButton_00.setGeometry(437, 291, 80, 20)
            self.etc_pushButton_01.setGeometry(522, 291, 80, 20)
            self.etc_pushButton_02.setGeometry(607, 291, 80, 20)
            self.etc_pushButton_02.setStyleSheet(style_bc_dk)
            self.etc_pushButton_02.setFont(qfont12)
            self.ct_label_01.setGeometry(3500, 5, 140, 20)
            self.ct_lineEdit_01.setGeometry(3500, 5, 100, 20)
            self.ct_label_02.setGeometry(3500, 702, 140, 20)
            self.ct_lineEdit_02.setGeometry(3500, 702, 100, 20)
            self.mode2 = 1
        elif gubun == 2 and self.mode2 == 1:
            self.setGeometry(0, 0, 3440, 1400)
            self.info_label_01.setGeometry(155, 1, 400, 30)
            self.info_label_02.setGeometry(600, 1, 400, 30)
            self.info_label_03.setGeometry(1185, 1, 400, 30)
            self.info_label_04.setGeometry(2145, 1, 400, 30)
            self.info_label_05.setGeometry(2888, 1, 400, 30)
            if self.mode1 == 0:
                self.chart_00_tabWidget.setGeometry(5, 5, 1025, 692)
                self.chart_01_tabWidget.setGeometry(1035, 5, 1026, 692)
                self.chart_02_tabWidget.setGeometry(5, 702, 1025, 692)
                self.chart_03_tabWidget.setGeometry(1035, 702, 1026, 692)
                self.hoga_00_tabWidget.setGeometry(2066, 5, 682, 692)
                self.hoga_01_tabWidget.setGeometry(2066, 702, 682, 692)
                self.info_label_06.setGeometry(155, 699, 400, 30)
                self.info_label_07.setGeometry(600, 699, 600, 30)
                self.info_label_08.setGeometry(1185, 699, 400, 30)
                self.info_label_09.setGeometry(2145, 699, 600, 30)
            elif self.mode1 == 1:
                self.chart_00_tabWidget.setGeometry(5, 5, 1025, 692)
                self.chart_01_tabWidget.setGeometry(1035, 5, 1026, 692)
                self.chart_00_tabWidget.setCurrentWidget(self.chart_00_tab)
                self.chart_01_tabWidget.setCurrentWidget(self.chart_01_tab)
                self.hoga_00_tabWidget.setGeometry(2066, 5, 682, 692)
            else:
                self.chart_04_tabWidget.setGeometry(5, 5, 2743, 1390)
            self.lgsj_tabWidget.setGeometry(2753, 5, 682, 282)
            self.table_tabWidget.setGeometry(2753, 292, 682, 1103)
            self.etc_pushButton_00.setGeometry(3185, 291, 80, 20)
            self.etc_pushButton_01.setGeometry(3270, 291, 80, 20)
            self.etc_pushButton_02.setGeometry(3355, 291, 80, 20)
            self.etc_pushButton_02.setStyleSheet(style_bc_bt)
            self.etc_pushButton_02.setFont(qfont12)
            if self.mode0 == 0 and self.mode1 == 0:
                self.ct_label_01.setGeometry(1820, 5, 140, 20)
                self.ct_lineEdit_01.setGeometry(1960, 5, 100, 20)
                self.ct_label_02.setGeometry(1820, 702, 140, 20)
                self.ct_lineEdit_02.setGeometry(1960, 702, 100, 20)
            elif self.mode1 == 1:
                self.ct_label_01.setGeometry(1820, 5, 120, 20)
                self.ct_lineEdit_01.setGeometry(1960, 5, 100, 20)
            self.mode2 = 0

    @thread_decorator
    def UpdateSysinfo(self):
        p = psutil.Process(os.getpid())
        self.list_info[0] = [
            round(p.memory_info()[0] / 2 ** 20.86, 2), p.num_threads(), round(p.cpu_percent(interval=2) / 2, 2)
        ]


class Writer(QtCore.QThread):
    data0 = QtCore.pyqtSignal(list)
    data1 = QtCore.pyqtSignal(list)
    data2 = QtCore.pyqtSignal(list)
    data3 = QtCore.pyqtSignal(list)

    def __init__(self):
        super().__init__()

    def run(self):
        tlist = [ui_num['단타설정'], ui_num['관심종목'], ui_num['관심종목'] + 100]
        clist = [ui_num['차트P1'], ui_num['차트P2'], ui_num['차트P3'], ui_num['차트P4'], ui_num['차트P5'],
                 ui_num['차트P6'], ui_num['차트P7'], ui_num['차트P8'], ui_num['차트P9']]
        while True:
            data = windowQ.get()
            if data[0] not in tlist and type(data[1]) != pd.DataFrame:
                self.data0.emit(data)
            elif data[0] in clist:
                self.data1.emit(data)
            elif data[0] in tlist:
                self.data2.emit(data)
            else:
                self.data3.emit(data)


if __name__ == '__main__':
    windowQ, traderQ, receivQ, stgQ, soundQ, queryQ, teleQ, hoga1Q, hoga2Q, chart1Q, chart2Q, chart3Q, chart4Q, \
        chart5Q, chart6Q, chart7Q, chart8Q, chart9Q, tick1Q, tick2Q, tick3Q, tick4Q = \
        Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue(),  Queue(), Queue(), Queue(), \
        Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue(), Queue()

    Process(target=Sound, args=(soundQ,), daemon=True).start()
    Process(target=Query, args=(windowQ, queryQ), daemon=True).start()
    Process(target=TelegramMsg, args=(windowQ, traderQ, queryQ, teleQ), daemon=True).start()

    os.system(f'python {system_path}/login/versionupdater.py')
    os.system(f'python {system_path}/login/autologin2.py')

    Process(target=Collector, args=(1, windowQ, queryQ, tick1Q), daemon=True).start()
    Process(target=Collector, args=(2, windowQ, queryQ, tick2Q), daemon=True).start()
    Process(target=Collector, args=(3, windowQ, queryQ, tick3Q), daemon=True).start()
    Process(target=Collector, args=(4, windowQ, queryQ, tick4Q), daemon=True).start()
    Process(target=Receiver,
            args=(windowQ, receivQ, traderQ, stgQ, queryQ, tick1Q, tick2Q, tick3Q, tick4Q), daemon=True).start()
    time.sleep(15)

    os.system(f'python {system_path}/login/autologin1.py')

    Process(target=UpdaterHoga, args=(windowQ, hoga1Q, ui_num['호가P0']), daemon=True).start()
    Process(target=UpdaterHoga, args=(windowQ, hoga2Q, ui_num['호가P1']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart1Q, ui_num['차트P1']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart2Q, ui_num['차트P2']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart3Q, ui_num['차트P3']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart4Q, ui_num['차트P4']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart5Q, ui_num['차트P5']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart6Q, ui_num['차트P6']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart7Q, ui_num['차트P7']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart8Q, ui_num['차트P8']), daemon=True).start()
    Process(target=UpdaterChart, args=(windowQ, traderQ, chart9Q, ui_num['차트P9']), daemon=True).start()
    Process(target=Strategy, args=(windowQ, traderQ, stgQ), daemon=True).start()
    Process(target=Trader,
            args=(windowQ, traderQ, stgQ, receivQ, soundQ, queryQ, teleQ, hoga1Q, hoga2Q, chart1Q,
                  chart2Q, chart3Q, chart4Q, chart5Q, chart6Q, chart7Q, chart8Q, chart9Q), daemon=True).start()

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('fusion')
    palette = QPalette()
    palette.setColor(QPalette.Window, color_bg_bc)
    palette.setColor(QPalette.Background, color_bg_bc)
    palette.setColor(QPalette.WindowText, color_fg_bc)
    palette.setColor(QPalette.Base, color_bg_bc)
    palette.setColor(QPalette.AlternateBase, color_bg_dk)
    palette.setColor(QPalette.Text, color_fg_bc)
    palette.setColor(QPalette.Button, color_bg_bc)
    palette.setColor(QPalette.ButtonText, color_fg_bc)
    palette.setColor(QPalette.Link, color_fg_bk)
    palette.setColor(QPalette.Highlight, color_fg_bk)
    palette.setColor(QPalette.HighlightedText, color_bg_bk)
    app.setPalette(palette)
    window = Window()
    window.show()
    app.exec_()

    if int(strf_time('%H%M%S')) < 100000:
        os.system('shutdown /s /t 60')
        sys.exit()

    os.system(f'python {system_path}/backtester/backtester_vc.py')
    os.system('shutdown /s /t 60')
    sys.exit()
