from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QIcon, QPixmap
from trader.chartItem import *


def SetUI(self):
    def setIcon(path):
        icon = QIcon()
        icon.addPixmap(QPixmap(path))
        return icon

    def setLine(tab, width):
        line = QtWidgets.QFrame(tab)
        line.setLineWidth(width)
        line.setStyleSheet(style_fc_dk)
        line.setFrameShape(QtWidgets.QFrame.HLine)
        return line

    def setLineedit(groupbox, returnPressed=None):
        lineedit = QtWidgets.QLineEdit(groupbox)
        if returnPressed is not None:
            lineedit.setAlignment(Qt.AlignCenter)
        else:
            lineedit.setAlignment(Qt.AlignRight)
        if returnPressed is not None:
            lineedit.setStyleSheet(style_bc_bt)
        else:
            lineedit.setStyleSheet(style_fc_bt)
        lineedit.setFont(qfont12)
        if returnPressed is not None:
            lineedit.returnPressed.connect(returnPressed)
        return lineedit

    def setPushbutton(name, groupbox, buttonclicked, cmd=None):
        pushbutton = QtWidgets.QPushButton(name, groupbox)
        pushbutton.setStyleSheet(style_bc_bt)
        pushbutton.setFont(qfont12)
        if cmd is not None:
            pushbutton.clicked.connect(lambda: buttonclicked(cmd))
        else:
            pushbutton.clicked.connect(lambda: buttonclicked(name))
        return pushbutton

    def setTextEdit(tab, qfont=None):
        textedit = QtWidgets.QTextEdit(tab)
        textedit.setReadOnly(True)
        textedit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        textedit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        textedit.setStyleSheet(style_bc_dk)
        if qfont is not None:
            textedit.setFont(qfont)
        else:
            textedit.setFont(qfont12)
        return textedit

    def setPg(tname, tabwidget, tab):
        ctpg = pg.GraphicsLayoutWidget()
        ctpg_01 = ctpg.addPlot(row=0, col=0, viewBox=CustomViewBox1())
        ctpg_02 = ctpg.addPlot(row=1, col=0, viewBox=CustomViewBox2())
        ctpg_01.showAxis('left', False)
        ctpg_01.showAxis('right', True)
        ctpg_01.getAxis('right').setStyle(tickTextWidth=45, autoExpandTextSpace=False)
        ctpg_01.getAxis('right').setTickFont(qfont12)
        ctpg_01.getAxis('bottom').setTickFont(qfont12)
        ctpg_02.showAxis('left', False)
        ctpg_02.showAxis('right', True)
        ctpg_02.getAxis('right').setStyle(tickTextWidth=45, autoExpandTextSpace=False)
        ctpg_02.getAxis('right').setTickFont(qfont12)
        ctpg_02.getAxis('bottom').setTickFont(qfont12)
        ctpg_02.setXLink(ctpg_01)
        qGraphicsGridLayout = ctpg.ci.layout
        qGraphicsGridLayout.setRowStretchFactor(0, 2)
        qGraphicsGridLayout.setRowStretchFactor(1, 1)
        ctpg_vboxLayout = QtWidgets.QVBoxLayout(tab)
        ctpg_vboxLayout.setContentsMargins(5, 5, 5, 5)
        ctpg_vboxLayout.addWidget(ctpg)
        tabwidget.addTab(tab, tname)
        return [ctpg_01, ctpg_02]

    def setTablewidget(tab, columns, colcount, rowcount, sectionsize=None, clicked=None, color=False, qfont=None):
        tableWidget = QtWidgets.QTableWidget(tab)
        if sectionsize is not None:
            tableWidget.verticalHeader().setDefaultSectionSize(sectionsize)
        else:
            tableWidget.verticalHeader().setDefaultSectionSize(23)
        tableWidget.verticalHeader().setVisible(False)
        tableWidget.setAlternatingRowColors(True)
        tableWidget.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        tableWidget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tableWidget.setColumnCount(len(columns))
        tableWidget.setRowCount(rowcount)
        tableWidget.setHorizontalHeaderLabels(columns)
        if qfont is not None:
            tableWidget.setFont(qfont)
        if colcount == 1:
            tableWidget.setColumnWidth(0, 84)
        elif colcount == 2:
            tableWidget.setColumnWidth(0, 84)
            tableWidget.setColumnWidth(1, 84)
        elif colcount == 3:
            tableWidget.setColumnWidth(0, 126)
            tableWidget.setColumnWidth(1, 90)
            tableWidget.setColumnWidth(2, 450)
        elif colcount == 4:
            tableWidget.setColumnWidth(0, 84)
            tableWidget.setColumnWidth(1, 84)
            tableWidget.setColumnWidth(2, 84)
            tableWidget.setColumnWidth(3, 84)
        elif colcount == 5:
            tableWidget.setColumnWidth(0, 66)
            tableWidget.setColumnWidth(1, 60)
            tableWidget.setColumnWidth(2, 60)
            tableWidget.setColumnWidth(3, 60)
            tableWidget.setColumnWidth(4, 60)
        elif colcount == 6:
            if rowcount == 13:
                tableWidget.setColumnWidth(0, 60)
                tableWidget.setColumnWidth(1, 60)
                tableWidget.setColumnWidth(2, 60)
                tableWidget.setColumnWidth(3, 60)
                tableWidget.setColumnWidth(4, 60)
                tableWidget.setColumnWidth(5, 60)
            else:
                tableWidget.setColumnWidth(0, 111)
                tableWidget.setColumnWidth(1, 111)
                tableWidget.setColumnWidth(2, 111)
                tableWidget.setColumnWidth(3, 111)
                tableWidget.setColumnWidth(4, 111)
                tableWidget.setColumnWidth(5, 111)
        elif columns[-1] == 'chhigh':
            tableWidget.setColumnWidth(0, 122)
            tableWidget.setColumnWidth(1, 68)
            tableWidget.setColumnWidth(2, 68)
            tableWidget.setColumnWidth(3, 68)
            tableWidget.setColumnWidth(4, 68)
            tableWidget.setColumnWidth(5, 68)
            tableWidget.setColumnWidth(6, 68)
            tableWidget.setColumnWidth(7, 68)
            tableWidget.setColumnWidth(8, 68)
        else:
            if columns[0] in ['기간', '일자']:
                tableWidget.setColumnWidth(0, 100)
                tableWidget.setColumnWidth(1, 100)
                tableWidget.setColumnWidth(2, 100)
                tableWidget.setColumnWidth(3, 100)
                tableWidget.setColumnWidth(4, 100)
                tableWidget.setColumnWidth(5, 66)
                tableWidget.setColumnWidth(6, 100)
            else:
                tableWidget.setColumnWidth(0, 126)
                tableWidget.setColumnWidth(1, 90)
                tableWidget.setColumnWidth(2, 90)
                tableWidget.setColumnWidth(3, 90)
                tableWidget.setColumnWidth(4, 90)
                tableWidget.setColumnWidth(5, 90)
                tableWidget.setColumnWidth(6, 90)
            if colcount >= 8:
                tableWidget.setColumnWidth(7, 90)
            if colcount >= 9:
                tableWidget.setColumnWidth(8, 90)
            if colcount >= 11:
                tableWidget.setColumnWidth(9, 90)
                tableWidget.setColumnWidth(10, 90)
            if colcount >= 12:
                tableWidget.setColumnWidth(11, 90)
            if colcount >= 13:
                tableWidget.setColumnWidth(12, 90)
            if colcount >= 14:
                tableWidget.setColumnWidth(13, 90)
        if clicked is not None:
            tableWidget.cellClicked.connect(clicked)
        if color:
            for i in range(22):
                tableitem = QtWidgets.QTableWidgetItem()
                tableitem.setBackground(color_bg_bt)
                tableWidget.setItem(i, 0, tableitem)
        return tableWidget

    self.setFont(qfont12)
    self.setWindowTitle('MyKiwoom')
    self.setWindowFlags(Qt.FramelessWindowHint)
    self.setWindowIcon(QIcon(f'{system_path}/Icon/python.png'))
    self.icon_open = setIcon(f'{system_path}/Icon/open.bmp')
    self.icon_high = setIcon(f'{system_path}/Icon/high.bmp')
    self.icon_low = setIcon(f'{system_path}/Icon/low.bmp')
    self.icon_up = setIcon(f'{system_path}/Icon/up.bmp')
    self.icon_down = setIcon(f'{system_path}/Icon/down.bmp')
    self.icon_vi = setIcon(f'{system_path}/Icon/vi.bmp')
    self.icon_totals = setIcon(f'{system_path}/Icon/totals.bmp')
    self.icon_totalb = setIcon(f'{system_path}/Icon/totalb.bmp')
    self.icon_pers = setIcon(f'{system_path}/Icon/pers.bmp')
    self.icon_perb = setIcon(f'{system_path}/Icon/perb.bmp')

    pg.setConfigOptions(background=color_bg_dk, foreground=color_fg_dk, leftButtonPan=False)
    self.dict_ctpg = {}

    self.chart_00_tabWidget = QtWidgets.QTabWidget(self)
    self.chart_00_tab = QtWidgets.QWidget()
    self.chart_05_tab = QtWidgets.QWidget()
    self.dict_ctpg[ui_num['차트P1']] = setPg('일봉 차트', self.chart_00_tabWidget, self.chart_00_tab)
    self.dict_ctpg[ui_num['차트P6']] = setPg('일봉 차트', self.chart_00_tabWidget, self.chart_05_tab)

    self.chart_01_tabWidget = QtWidgets.QTabWidget(self)
    self.chart_01_tab = QtWidgets.QWidget()
    self.chart_06_tab = QtWidgets.QWidget()
    self.dict_ctpg[ui_num['차트P2']] = setPg('분봉 차트', self.chart_01_tabWidget, self.chart_01_tab)
    self.dict_ctpg[ui_num['차트P7']] = setPg('분봉 차트', self.chart_01_tabWidget, self.chart_06_tab)

    self.chart_02_tabWidget = QtWidgets.QTabWidget(self)
    self.chart_02_tab = QtWidgets.QWidget()
    self.chart_07_tab = QtWidgets.QWidget()
    self.dict_ctpg[ui_num['차트P3']] = setPg('일봉 차트', self.chart_02_tabWidget, self.chart_02_tab)
    self.dict_ctpg[ui_num['차트P8']] = setPg('일봉 차트', self.chart_02_tabWidget, self.chart_07_tab)

    self.chart_03_tabWidget = QtWidgets.QTabWidget(self)
    self.chart_03_tab = QtWidgets.QWidget()
    self.chart_08_tab = QtWidgets.QWidget()
    self.dict_ctpg[ui_num['차트P4']] = setPg('분봉 차트', self.chart_03_tabWidget, self.chart_03_tab)
    self.dict_ctpg[ui_num['차트P9']] = setPg('분봉 차트', self.chart_03_tabWidget, self.chart_08_tab)

    self.chart_04_tabWidget = QtWidgets.QTabWidget(self)
    self.chart_04_tab = QtWidgets.QWidget()
    self.dict_ctpg[ui_num['차트P5']] = setPg('복기 차트', self.chart_04_tabWidget, self.chart_04_tab)

    self.hoga_00_tabWidget = QtWidgets.QTabWidget(self)
    self.hoga_00_tab = QtWidgets.QWidget()
    self.hoga_00_sellper_groupBox = QtWidgets.QGroupBox(' ', self.hoga_00_tab)
    self.hoga_00_sell_radioButton_01 = QtWidgets.QRadioButton('10%', self.hoga_00_sellper_groupBox)
    self.hoga_00_sell_radioButton_02 = QtWidgets.QRadioButton('25%', self.hoga_00_sellper_groupBox)
    self.hoga_00_sell_radioButton_03 = QtWidgets.QRadioButton('33%', self.hoga_00_sellper_groupBox)
    self.hoga_00_sell_radioButton_04 = QtWidgets.QRadioButton('50%', self.hoga_00_sellper_groupBox)
    self.hoga_00_sell_radioButton_05 = QtWidgets.QRadioButton('75%', self.hoga_00_sellper_groupBox)
    self.hoga_00_sell_radioButton_06 = QtWidgets.QRadioButton('100%', self.hoga_00_sellper_groupBox)
    self.hoga_00_buywon_groupBox = QtWidgets.QGroupBox(' ', self.hoga_00_tab)
    self.hoga_00_buy_radioButton_01 = QtWidgets.QRadioButton('100,000', self.hoga_00_buywon_groupBox)
    self.hoga_00_buy_radioButton_02 = QtWidgets.QRadioButton('500,000', self.hoga_00_buywon_groupBox)
    self.hoga_00_buy_radioButton_03 = QtWidgets.QRadioButton('1,000,000', self.hoga_00_buywon_groupBox)
    self.hoga_00_buy_radioButton_04 = QtWidgets.QRadioButton('5,000,000', self.hoga_00_buywon_groupBox)
    self.hoga_00_buy_radioButton_05 = QtWidgets.QRadioButton('10,000,000', self.hoga_00_buywon_groupBox)
    self.hoga_00_buy_radioButton_06 = QtWidgets.QRadioButton('50,000,000', self.hoga_00_buywon_groupBox)
    self.hoga_00_sell_pushButton_01 = setPushbutton('시장가 매도', self.hoga_00_tab, self.ButtonClicked_1,
                                                    '시장가매도0')
    self.hoga_00_sell_pushButton_02 = setPushbutton('매도 취소', self.hoga_00_tab, self.ButtonClicked_1,
                                                    '매도취소0')
    self.hoga_00_buy_pushButton_01 = setPushbutton('매수 취소', self.hoga_00_tab, self.ButtonClicked_2,
                                                   '매수취소0')
    self.hoga_00_buy_pushButton_02 = setPushbutton('시장가 매수', self.hoga_00_tab, self.ButtonClicked_2,
                                                   '시장가매수0')
    self.hoga_00_hj_tableWidget = setTablewidget(self.hoga_00_tab, columns_hj, len(columns_hj), 1)
    self.hoga_00_hs_tableWidget = setTablewidget(self.hoga_00_tab, columns_hs, len(columns_hs), 22,
                                                 clicked=self.CellClicked_1, color=True)
    self.hoga_00_hc_tableWidget = setTablewidget(self.hoga_00_tab, columns_hc, len(columns_hc), 22)
    self.hoga_00_hg_tableWidget = setTablewidget(self.hoga_00_tab, columns_hg, len(columns_hg), 22)
    self.hoga_00_hb_tableWidget = setTablewidget(self.hoga_00_tab, columns_hb, len(columns_hb), 22,
                                                 clicked=self.CellClicked_2, color=True)
    self.hoga_00_line = setLine(self.hoga_00_tab, 1)
    self.hoga_00_tabWidget.addTab(self.hoga_00_tab, '호가 주문')

    self.hoga_01_tabWidget = QtWidgets.QTabWidget(self)
    self.hoga_01_tab = QtWidgets.QWidget()
    self.hoga_01_sellper_groupBox = QtWidgets.QGroupBox(' ', self.hoga_01_tab)
    self.hoga_01_sell_radioButton_01 = QtWidgets.QRadioButton('10%', self.hoga_01_sellper_groupBox)
    self.hoga_01_sell_radioButton_02 = QtWidgets.QRadioButton('25%', self.hoga_01_sellper_groupBox)
    self.hoga_01_sell_radioButton_03 = QtWidgets.QRadioButton('33%', self.hoga_01_sellper_groupBox)
    self.hoga_01_sell_radioButton_04 = QtWidgets.QRadioButton('50%', self.hoga_01_sellper_groupBox)
    self.hoga_01_sell_radioButton_05 = QtWidgets.QRadioButton('75%', self.hoga_01_sellper_groupBox)
    self.hoga_01_sell_radioButton_06 = QtWidgets.QRadioButton('100%', self.hoga_01_sellper_groupBox)
    self.hoga_01_buywon_groupBox = QtWidgets.QGroupBox(' ', self.hoga_01_tab)
    self.hoga_01_buy_radioButton_01 = QtWidgets.QRadioButton('100,000', self.hoga_01_buywon_groupBox)
    self.hoga_01_buy_radioButton_02 = QtWidgets.QRadioButton('500,000', self.hoga_01_buywon_groupBox)
    self.hoga_01_buy_radioButton_03 = QtWidgets.QRadioButton('1,000,000', self.hoga_01_buywon_groupBox)
    self.hoga_01_buy_radioButton_04 = QtWidgets.QRadioButton('5,000,000', self.hoga_01_buywon_groupBox)
    self.hoga_01_buy_radioButton_05 = QtWidgets.QRadioButton('10,000,000', self.hoga_01_buywon_groupBox)
    self.hoga_01_buy_radioButton_06 = QtWidgets.QRadioButton('50,000,000', self.hoga_01_buywon_groupBox)
    self.hoga_01_sell_pushButton_01 = setPushbutton('시장가 매도', self.hoga_01_tab, self.ButtonClicked_1,
                                                    '시장가매도1')
    self.hoga_01_sell_pushButton_02 = setPushbutton('매도 취소', self.hoga_01_tab, self.ButtonClicked_1,
                                                    '매도취소1')
    self.hoga_01_buy_pushButton_01 = setPushbutton('매수 취소', self.hoga_01_tab, self.ButtonClicked_2,
                                                   '매수취소1')
    self.hoga_01_buy_pushButton_02 = setPushbutton('시장가 매수', self.hoga_01_tab, self.ButtonClicked_2,
                                                   '시장가매수1')
    self.hoga_01_hj_tableWidget = setTablewidget(self.hoga_01_tab, columns_hj, len(columns_hj), 1)
    self.hoga_01_hs_tableWidget = setTablewidget(self.hoga_01_tab, columns_hs, len(columns_hs), 22,
                                                 clicked=self.CellClicked_3, color=True)
    self.hoga_01_hc_tableWidget = setTablewidget(self.hoga_01_tab, columns_hc, len(columns_hc), 22)
    self.hoga_01_hg_tableWidget = setTablewidget(self.hoga_01_tab, columns_hg, len(columns_hg), 22)
    self.hoga_01_hb_tableWidget = setTablewidget(self.hoga_01_tab, columns_hb, len(columns_hb), 22,
                                                 clicked=self.CellClicked_4, color=True)
    self.hoga_01_line = setLine(self.hoga_01_tab, 1)
    self.hoga_01_tabWidget.addTab(self.hoga_01_tab, '호가 주문')

    self.gg_tabWidget = QtWidgets.QTabWidget(self)
    self.gg_tab = QtWidgets.QWidget()
    self.gg_textEdit = setTextEdit(self.gg_tab, qfont14)
    self.gg_tabWidget.addTab(self.gg_tab, '기업 개요')

    self.gs_tabWidget = QtWidgets.QTabWidget(self)
    self.gs_tab = QtWidgets.QWidget()
    self.gs_tableWidget = setTablewidget(self.gs_tab, columns_gc, len(columns_gc), 22, qfont=qfont13)
    self.gs_tabWidget.addTab(self.gs_tab, '기업 공시')

    self.ns_tabWidget = QtWidgets.QTabWidget(self)
    self.ns_tab = QtWidgets.QWidget()
    self.ns_tableWidget = setTablewidget(self.ns_tab, columns_ns, len(columns_ns), 12, qfont=qfont13)
    self.ns_tabWidget.addTab(self.ns_tab, '종목 뉴스')

    self.jj_tabWidget = QtWidgets.QTabWidget(self)
    self.jj_tab = QtWidgets.QWidget()
    self.jj_tableWidget = setTablewidget(self.jj_tab, columns_jj, len(columns_jj), 28)
    self.jj_tabWidget.addTab(self.jj_tab, '투자자별 매매동향')

    self.jm_tabWidget = QtWidgets.QTabWidget(self)
    self.jm_tab = QtWidgets.QWidget()
    self.jm1_tableWidget = setTablewidget(self.jm_tab, columns_jm1, len(columns_jm1), 13, sectionsize=21)
    self.jm2_tableWidget = setTablewidget(self.jm_tab, columns_jm2, len(columns_jm2), 13, sectionsize=21)
    self.jm_tabWidget.addTab(self.jm_tab, '재무제표')

    self.jb_tabWidget = QtWidgets.QTabWidget(self)
    self.jb_tab = QtWidgets.QWidget()
    self.jb_tableWidget = setTablewidget(self.jb_tab, columns_jb, len(columns_jb), 14, sectionsize=21)
    self.jb_tabWidget.addTab(self.jb_tab, '동일업종비교')

    self.ch_tabWidget = QtWidgets.QTabWidget(self)
    self.ch_tab = QtWidgets.QWidget()
    self.ch_tableWidget = setTablewidget(self.ch_tab, columns_ch, len(columns_ch), 28)
    self.ch_tabWidget.addTab(self.ch_tab, '체결강도')

    self.lgsj_tabWidget = QtWidgets.QTabWidget(self)
    self.lg_tab = QtWidgets.QWidget()
    self.sj_tab = QtWidgets.QWidget()

    self.lg_textEdit = setTextEdit(self.lg_tab)

    self.sj_groupBox_01 = QtWidgets.QGroupBox(self.sj_tab)
    self.sj_label_01 = QtWidgets.QLabel('텔레그램봇 넘버', self.sj_groupBox_01)
    self.sj_label_02 = QtWidgets.QLabel('사용자 아이디', self.sj_groupBox_01)
    self.sj_lineEdit_01 = setLineedit(self.sj_groupBox_01)
    self.sj_lineEdit_02 = setLineedit(self.sj_groupBox_01)
    self.sj_pushButton_01 = setPushbutton('설정', self.sj_groupBox_01, self.ButtonClicked_3)
    self.sj_groupBox_02 = QtWidgets.QGroupBox(self.sj_tab)
    self.sj_pushButton_02 = setPushbutton('데이터베이스 로딩', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_03 = setPushbutton('OPENAPI 로그인', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_04 = setPushbutton('계좌평가 및 잔고', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_05 = setPushbutton('코스피 코스닥 차트', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_06 = setPushbutton('장운영시간 알림 등록', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_07 = setPushbutton('업종지수 주식체결 등록', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_08 = setPushbutton('VI발동해제 등록', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_09 = setPushbutton('장운영상태', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_10 = setPushbutton('실시간 조건검색식 등록', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_11 = setPushbutton('잔고청산', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_12 = setPushbutton('실시간 데이터 수신 중단', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_13 = setPushbutton('틱데이터 저장', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_14 = setPushbutton('테스트모드 ON/OFF', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_15 = setPushbutton('모의투자 ON/OFF', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_16 = setPushbutton('알림소리 ON/OFF', self.sj_groupBox_02, self.ButtonClicked_3)
    self.sj_pushButton_17 = setPushbutton('시스템 종료', self.sj_groupBox_02, self.ButtonClicked_3)

    self.lgsj_tabWidget.addTab(self.lg_tab, '로그')
    self.lgsj_tabWidget.addTab(self.sj_tab, '시스템설정')

    self.table_tabWidget = QtWidgets.QTabWidget(self)
    self.td_tab = QtWidgets.QWidget()
    self.gj_tab = QtWidgets.QWidget()
    self.st_tab = QtWidgets.QWidget()
    self.sg_tab = QtWidgets.QWidget()

    self.tt_tableWidget = setTablewidget(self.td_tab, columns_tt, len(columns_tt), 1)
    self.td_tableWidget = setTablewidget(self.td_tab, columns_td, len(columns_td), 13, clicked=self.CellClicked_5)
    self.tj_tableWidget = setTablewidget(self.td_tab, columns_tj, len(columns_tj), 1)
    self.jg_tableWidget = setTablewidget(self.td_tab, columns_jg, len(columns_jg), 13, clicked=self.CellClicked_6)
    self.cj_tableWidget = setTablewidget(self.td_tab, columns_cj, len(columns_cj), 12, clicked=self.CellClicked_7)
    self.gj_tableWidget = setTablewidget(self.gj_tab, columns_gj3, len(columns_gj3), 46, clicked=self.CellClicked_8)

    self.st_groupBox = QtWidgets.QGroupBox(self.st_tab)
    self.calendarWidget = QtWidgets.QCalendarWidget(self.st_groupBox)
    todayDate = QtCore.QDate.currentDate()
    self.calendarWidget.setCurrentPage(todayDate.year(), todayDate.month())
    self.calendarWidget.clicked.connect(self.CalendarClicked)
    self.dt_tableWidget = setTablewidget(self.st_tab, columns_dt, len(columns_dt), 1)
    self.dd_tableWidget = setTablewidget(self.st_tab, columns_dd, len(columns_dd), 31, clicked=self.CellClicked_9)

    self.sg_groupBox = QtWidgets.QGroupBox(self.sg_tab)
    self.sg_pushButton_01 = setPushbutton('일별집계', self.sg_groupBox, self.ButtonClicked_3)
    self.sg_pushButton_02 = setPushbutton('월별집계', self.sg_groupBox, self.ButtonClicked_3)
    self.sg_pushButton_03 = setPushbutton('연도별집계', self.sg_groupBox, self.ButtonClicked_3)
    self.nt_tableWidget = setTablewidget(self.sg_tab, columns_nt, len(columns_nt), 1)
    self.nd_tableWidget = setTablewidget(self.sg_tab, columns_nd, len(columns_nd), 41)

    self.table_tabWidget.addTab(self.td_tab, '계좌평가')
    self.table_tabWidget.addTab(self.gj_tab, '관심종목')
    self.table_tabWidget.addTab(self.st_tab, '거래목록')
    self.table_tabWidget.addTab(self.sg_tab, '수익현황')

    self.info_label_01 = QtWidgets.QLabel(self)
    self.info_label_02 = QtWidgets.QLabel(self)
    self.info_label_03 = QtWidgets.QLabel(self)
    self.info_label_04 = QtWidgets.QLabel(self)
    self.info_label_05 = QtWidgets.QLabel(self)
    self.info_label_06 = QtWidgets.QLabel(self)
    self.info_label_07 = QtWidgets.QLabel(self)
    self.info_label_08 = QtWidgets.QLabel(self)
    self.info_label_09 = QtWidgets.QLabel(self)

    self.etc_pushButton_00 = setPushbutton('차트탭변경', self, self.ButtonClicked_4, 0)
    self.etc_pushButton_01 = setPushbutton('차트유형변경', self, self.ButtonClicked_4, 1)
    self.etc_pushButton_02 = setPushbutton('창크기변경', self, self.ButtonClicked_4, 2)

    self.ct_label_01 = QtWidgets.QLabel('종목명 또는 종목코드 조회', self)
    self.ct_label_02 = QtWidgets.QLabel('종목명 또는 종목코드 조회', self)
    self.ct_lineEdit_01 = setLineedit(self, self.ReturnPressed_1)
    self.ct_lineEdit_02 = setLineedit(self, self.ReturnPressed_2)

    self.setGeometry(0, 0, 3440, 1400)
    self.chart_00_tabWidget.setGeometry(5, 5, 1025, 692)
    self.chart_01_tabWidget.setGeometry(1035, 5, 1026, 692)
    self.chart_02_tabWidget.setGeometry(5, 702, 1025, 692)
    self.chart_03_tabWidget.setGeometry(1035, 702, 1026, 692)
    self.chart_04_tabWidget.setGeometry(3500, 5, 2743, 1390)
    self.hoga_00_tabWidget.setGeometry(2066, 5, 682, 692)
    self.hoga_01_tabWidget.setGeometry(2066, 702, 682, 692)
    self.lgsj_tabWidget.setGeometry(2753, 5, 682, 282)
    self.table_tabWidget.setGeometry(2753, 292, 682, 1103)

    self.info_label_01.setGeometry(155, 1, 400, 30)
    self.info_label_02.setGeometry(600, 1, 400, 30)
    self.info_label_03.setGeometry(1185, 1, 400, 30)
    self.info_label_04.setGeometry(2145, 1, 400, 30)
    self.info_label_05.setGeometry(2888, 1, 400, 30)
    self.info_label_06.setGeometry(155, 699, 400, 30)
    self.info_label_07.setGeometry(600, 699, 600, 30)
    self.info_label_08.setGeometry(1185, 699, 400, 30)
    self.info_label_09.setGeometry(2145, 699, 600, 30)

    self.etc_pushButton_00.setGeometry(3185, 291, 80, 20)
    self.etc_pushButton_01.setGeometry(3270, 291, 80, 20)
    self.etc_pushButton_02.setGeometry(3355, 291, 80, 20)
    self.ct_label_01.setGeometry(3500, 5, 140, 20)
    self.ct_label_02.setGeometry(3500, 702, 140, 20)
    self.ct_lineEdit_01.setGeometry(3500, 5, 100, 20)
    self.ct_lineEdit_02.setGeometry(3500, 702, 100, 20)

    self.hoga_00_sellper_groupBox.setGeometry(5, -10, 331, 65)
    self.hoga_00_sell_radioButton_01.setGeometry(10, 22, 100, 20)
    self.hoga_00_sell_radioButton_02.setGeometry(110, 22, 100, 20)
    self.hoga_00_sell_radioButton_03.setGeometry(220, 22, 100, 20)
    self.hoga_00_sell_radioButton_04.setGeometry(10, 42, 100, 20)
    self.hoga_00_sell_radioButton_05.setGeometry(110, 42, 100, 20)
    self.hoga_00_sell_radioButton_06.setGeometry(220, 42, 100, 20)
    self.hoga_00_buywon_groupBox.setGeometry(341, -10, 331, 65)
    self.hoga_00_buy_radioButton_01.setGeometry(10, 22, 100, 20)
    self.hoga_00_buy_radioButton_02.setGeometry(110, 22, 100, 20)
    self.hoga_00_buy_radioButton_03.setGeometry(220, 22, 100, 20)
    self.hoga_00_buy_radioButton_04.setGeometry(10, 42, 100, 20)
    self.hoga_00_buy_radioButton_05.setGeometry(110, 42, 100, 20)
    self.hoga_00_buy_radioButton_06.setGeometry(220, 42, 100, 20)
    self.hoga_00_sell_pushButton_01.setGeometry(5, 60, 163, 20)
    self.hoga_00_sell_pushButton_02.setGeometry(173, 60, 163, 20)
    self.hoga_00_buy_pushButton_01.setGeometry(341, 60, 163, 20)
    self.hoga_00_buy_pushButton_02.setGeometry(509, 60, 163, 20)
    self.hoga_00_hj_tableWidget.setGeometry(5, 85, 668, 42)
    self.hoga_00_hs_tableWidget.setGeometry(5, 132, 84, 525)
    self.hoga_00_hc_tableWidget.setGeometry(88, 132, 168, 525)
    self.hoga_00_hg_tableWidget.setGeometry(255, 132, 336, 525)
    self.hoga_00_hb_tableWidget.setGeometry(590, 132, 84, 525)
    self.hoga_00_line.setGeometry(6, 402, 667, 1)

    self.hoga_01_sellper_groupBox.setGeometry(5, -10, 331, 65)
    self.hoga_01_sell_radioButton_01.setGeometry(10, 22, 100, 20)
    self.hoga_01_sell_radioButton_02.setGeometry(110, 22, 100, 20)
    self.hoga_01_sell_radioButton_03.setGeometry(220, 22, 100, 20)
    self.hoga_01_sell_radioButton_04.setGeometry(10, 42, 100, 20)
    self.hoga_01_sell_radioButton_05.setGeometry(110, 42, 100, 20)
    self.hoga_01_sell_radioButton_06.setGeometry(220, 42, 100, 20)
    self.hoga_01_buywon_groupBox.setGeometry(341, -10, 331, 65)
    self.hoga_01_buy_radioButton_01.setGeometry(10, 22, 100, 20)
    self.hoga_01_buy_radioButton_02.setGeometry(110, 22, 100, 20)
    self.hoga_01_buy_radioButton_03.setGeometry(220, 22, 100, 20)
    self.hoga_01_buy_radioButton_04.setGeometry(10, 42, 100, 20)
    self.hoga_01_buy_radioButton_05.setGeometry(110, 42, 100, 20)
    self.hoga_01_buy_radioButton_06.setGeometry(220, 42, 100, 20)
    self.hoga_01_sell_pushButton_01.setGeometry(5, 60, 163, 20)
    self.hoga_01_sell_pushButton_02.setGeometry(173, 60, 163, 20)
    self.hoga_01_buy_pushButton_01.setGeometry(341, 60, 163, 20)
    self.hoga_01_buy_pushButton_02.setGeometry(509, 60, 163, 20)
    self.hoga_01_hj_tableWidget.setGeometry(5, 85, 668, 42)
    self.hoga_01_hs_tableWidget.setGeometry(5, 132, 84, 525)
    self.hoga_01_hc_tableWidget.setGeometry(88, 132, 168, 525)
    self.hoga_01_hg_tableWidget.setGeometry(255, 132, 336, 525)
    self.hoga_01_hb_tableWidget.setGeometry(590, 132, 84, 525)
    self.hoga_01_line.setGeometry(6, 402, 667, 1)

    self.gg_tabWidget.setGeometry(3500, 702, 682, 120)
    self.gs_tabWidget.setGeometry(3500, 807, 682, 567)
    self.ns_tabWidget.setGeometry(3500, 702, 682, 330)
    self.jj_tabWidget.setGeometry(3500, 1039, 682, 360)
    self.jm_tabWidget.setGeometry(3500, 702, 682, 330)
    self.jb_tabWidget.setGeometry(3500, 1039, 682, 360)
    self.ch_tabWidget.setGeometry(3500, 702, 682, 682)

    self.gg_textEdit.setGeometry(5, 5, 668, 80)
    self.gs_tableWidget.setGeometry(5, 5, 668, 527)
    self.ns_tableWidget.setGeometry(5, 5, 668, 293)
    self.jj_tableWidget.setGeometry(5, 5, 668, 315)
    self.jm1_tableWidget.setGeometry(5, 5, 310, 293)
    self.jm2_tableWidget.setGeometry(310, 5, 363, 293)
    self.jb_tableWidget.setGeometry(5, 5, 668, 315)
    self.ch_tableWidget.setGeometry(5, 5, 668, 653)

    self.lg_textEdit.setGeometry(5, 5, 668, 242)

    self.sj_groupBox_01.setGeometry(5, 3, 668, 65)
    self.sj_label_01.setGeometry(10, 13, 180, 20)
    self.sj_label_02.setGeometry(10, 38, 180, 20)
    self.sj_lineEdit_01.setGeometry(95, 12, 486, 20)
    self.sj_lineEdit_02.setGeometry(95, 37, 486, 20)
    self.sj_pushButton_01.setGeometry(586, 12, 75, 45)

    self.sj_groupBox_02.setGeometry(5, 70, 668, 178)
    self.sj_pushButton_02.setGeometry(5, 10, 161, 37)
    self.sj_pushButton_03.setGeometry(171, 10, 161, 37)
    self.sj_pushButton_04.setGeometry(337, 10, 161, 37)
    self.sj_pushButton_05.setGeometry(503, 10, 160, 37)
    self.sj_pushButton_06.setGeometry(5, 52, 161, 37)
    self.sj_pushButton_07.setGeometry(171, 52, 161, 37)
    self.sj_pushButton_08.setGeometry(337, 52, 161, 37)
    self.sj_pushButton_09.setGeometry(503, 52, 161, 37)
    self.sj_pushButton_10.setGeometry(5, 94, 160, 37)
    self.sj_pushButton_11.setGeometry(171, 94, 161, 37)
    self.sj_pushButton_12.setGeometry(337, 94, 161, 37)
    self.sj_pushButton_13.setGeometry(503, 94, 161, 37)
    self.sj_pushButton_14.setGeometry(5, 136, 161, 37)
    self.sj_pushButton_15.setGeometry(171, 136, 161, 37)
    self.sj_pushButton_16.setGeometry(337, 136, 161, 37)
    self.sj_pushButton_17.setGeometry(503, 136, 161, 37)

    self.tt_tableWidget.setGeometry(5, 5, 668, 42)
    self.td_tableWidget.setGeometry(5, 52, 668, 320)
    self.tj_tableWidget.setGeometry(5, 377, 668, 42)
    self.jg_tableWidget.setGeometry(5, 424, 668, 320)
    self.cj_tableWidget.setGeometry(5, 749, 668, 320)
    self.gj_tableWidget.setGeometry(5, 5, 668, 1063)

    self.st_groupBox.setGeometry(5, 3, 668, 278)
    self.calendarWidget.setGeometry(5, 11, 658, 258)
    self.dt_tableWidget.setGeometry(5, 287, 668, 42)
    self.dd_tableWidget.setGeometry(5, 334, 668, 735)

    self.sg_groupBox.setGeometry(5, 3, 668, 48)
    self.sg_pushButton_01.setGeometry(5, 11, 216, 30)
    self.sg_pushButton_02.setGeometry(226, 12, 216, 30)
    self.sg_pushButton_03.setGeometry(447, 12, 216, 30)
    self.nt_tableWidget.setGeometry(5, 57, 668, 42)
    self.nd_tableWidget.setGeometry(5, 104, 668, 965)
