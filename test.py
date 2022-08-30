# -*- coding: utf-8 -*-
import os
import xml.etree.ElementTree as ET
import sys
import logging
from PyQt5.QtWidgets import *
from PyQt5 import uic

# UI파일 연결
# 단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("test.ui")[0]


# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QDialog, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.chk1_sync.stateChanged.connect(self.chk_btn_func)
        self.btn1_mosaic.clicked.connect(self.btn1_mosaic_func)
        self.btn2_clear.clicked.connect(self.btn2_clear_func)
        self.btn3_current.clicked.connect(self.bnt3_current_func)
        self.btn4_sync.clicked.connect(self.btn4_sync_func)
        self.btn5_disableMosaic.clicked.connect(self.btn5_disableMosaic_func)
        self.combo_hz.currentIndexChanged.connect(self.combo_hz_func)
        self.combo_sync.currentIndexChanged.connect(self.combo_sync_func)

    ###############################################
    # FUNTIONS ####################################
    ###############################################

    def disableMosaic(self):
        hz = self.combo_hz_func()
        f = "c:\configureMosaic.exe test rows=1 cols=1 res=3840,2160,{0} out=0,0 nextgrid rows=1 cols=1 " \
            "res=3840,2160,{1} out=1,0".format(hz,hz)

        xmlString = os.popen(f).read()
        xml = ET.fromstring(xmlString)
        return xml, xmlString


    # 모자이크 활성화 RETURN[xml, str]
    def enableMosaic(self):
        hz = self.combo_hz_func()
        f = "C:\configureMosaic.exe set cols=1 rows=2 res=3840,2160,{0} out=1,0 out=0,0 maxperf".format(hz)

        xmlString = os.popen(f).read()
        xml = ET.fromstring(xmlString)
        return xml, xmlString

    # 모자이크 체크 RETURN[bool, list, list]
    @staticmethod
    def isMosaiced():
        cmd = os.popen("c:\\configureMosaic.exe listconfigcmd").read().split(" ")

        # nextgrid가 cmdlist에 있으면 True, 없으면 False 반환
        isMosaiced = False if 'nextgrid' in cmd else True

        firstgrid = cmd[2:5]
        firstgrid.insert(0, '1st_grid')
        # nextgrid
        if isMosaiced:
            nextgrid = None
        else:
            nextgridIndex = cmd.index('nextgrid')
            nextgrid = cmd[nextgridIndex:nextgridIndex + 4]

        return isMosaiced, firstgrid, nextgrid


    ###############################################
    # COMBOBOX FUNCTION ###########################
    ###############################################
    def combo_hz_func(self):
        hz = float(self.combo_hz.currentText()[:-2])
        return hz

    def combo_sync_func(self):
        current = self.combo_sync.currentText()
        self.label1.setText("Sync parameter set " + current)
        return current

    ###############################################
    # CHECKBOX FUNCTION ###########################
    ###############################################
    def chk_btn_func(self):
        if self.chk1_sync.isChecked():
            self.label1.setText("모자이크 활성화에 이어서 싱크작업을 진행합니다.")
            self.btn4_sync.setEnabled(False)
        if not self.chk1_sync.isChecked():
            self.label1.setText("모자이크 활성화만 진행합니다.")
            self.btn4_sync.setEnabled(True)

    ###############################################
    # BUTTON FUNCTION #############################
    ###############################################

    # SYNC_BUTTON
    def btn4_sync_func(self):
        current = self.combo_sync_func()
        self.label1.setText(current + " sync 활성화를 진행합니다.")

    # CURRENT_BUTTON
    def bnt3_current_func(self):
        self.textLog.clear()
        self.label1.setText("현재 상태입니다.")
        isMosaiced, firstgrid, nextgrid = self.isMosaiced()
        if isMosaiced:
            self.textLog.appendPlainText("현재 모자이크 상태입니다.")
            self.textLog.appendPlainText(str(firstgrid))
            self.textLog.appendPlainText(str(nextgrid))
        else:
            self.textLog.appendPlainText("현재 모자이크 상태가 아닙니다.")
            self.textLog.appendPlainText(str(firstgrid))
            self.textLog.appendPlainText(str(nextgrid))

    # CLEAR_BUTTON
    def btn2_clear_func(self):
        self.textLog.clear()
        self.label1.setText("log is cleared!")

    # MOSAIC_BUTTON
    def btn1_mosaic_func(self):

        xml, xmlString = self.enableMosaic()
        self.textLog.appendPlainText(xmlString)
        hz = self.combo_hz_func()

        if xml.attrib['valid'] == "1":
            self.label1.setText("{0}hz 모자이크가 활성화 되었습니다.".format(hz))
        else:
            self.bnt3_current_func()
            self.label1.setText("모자이크가 활성화에 실패했습니다.")

    # DISABLE_MOSAIC_BUTTON
    def btn5_disableMosaic_func(self):
        self.textLog.clear()
        xml, xmlString = self.disableMosaic()
        self.textLog.appendPlainText(xmlString)

    # DISABLE_MOSAIC_BUTTON


###############################################
# main #
###############################################

if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec()
