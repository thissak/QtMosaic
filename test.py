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


# hz = 30.00


# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QDialog, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.chk1_sync.stateChanged.connect(self.chk_btn_fuc)
        self.btn1_mosaic.clicked.connect(self.btn1_mosaic_fnc)
        self.btn2_clear.clicked.connect(self.btn2_clear_fnc)
        self.btn3_current.clicked.connect(self.bnt3_current_fnc)
        self.list1_hz.itemClicked.connect(self.chkItemClicked)
        # 현재 LISTWIDGET을 0번째 줄로 설정
        self.list1_hz.setCurrentRow(0)

    # FUNTIONS #

    # 모자이크 활성화 RETURN XML(xml), XMLSTRING(str)
    def enableMosaic(self):
        hz = self.chkItemClicked()
        f = "C:\\configureMosaic.exe test rows=1 cols=1 res=2560,1600,{0} gridPos=0,0 out=0,0 rotate=0  nextgrid " \
            "rows=1 cols=1 res=3840,2160,{0} gridPos=2560,0 out=0,1 rotate=0".format(hz, hz)
        xmlString = os.popen(f).read()
        xml = ET.fromstring(xmlString)
        return xml, xmlString

    # listconfigcmd 모자이크 체크 RETURN bool. firstgrid(list), nextgrid(list)
    @staticmethod
    def isMosaiced():
        cmd = os.popen("c:\\configureMosaic.exe listconfigcmd").read().split(" ")
        rows = cmd[2].split("=")[-1]
        isMosaiced = (int(rows) == 2)

        firstgrid = cmd[2:5]
        firstgrid.insert(0, '1st_grid')
        try:
            nextgridIndex = cmd.index('nextgrid')
            nextgrid = cmd[nextgridIndex:nextgridIndex + 4]
        except ValueError:
            nextgrid = None

        return isMosaiced, firstgrid, nextgrid

    # BUTTONS #

    # LISTWIDGET FUNTION RETURN hz(float)

    def chkItemClicked(self):
        hz = self.list1_hz.currentItem().text()
        return float(hz[:-2])

    # CHECKBOX FUNCTION
    def chk_btn_fuc(self):
        if self.chk_1.isChecked():
            self.label1.setText("모자이크 활성화에 이어서 싱크작업을 진행합니다.")
        if not self.chk_1.isChecked():
            self.label1.setText("모자이크 활성화만 진행합니다.")

    # CURRENT BUTTON FUNCTION
    def bnt3_current_fnc(self):
        self.textLog.clear()
        self.label1.setText("현재 상태입니다.")
        isMosaiced, firstgrid, nextgrid = WindowClass.isMosaiced()
        if isMosaiced:
            self.textLog.appendPlainText("현재 모자이크 상태입니다.")
            self.textLog.appendPlainText(str(firstgrid))
            self.textLog.appendPlainText(str(nextgrid))
        else:
            self.textLog.appendPlainText("현재 모자이크 상태가 아닙니다.")
            self.textLog.appendPlainText(str(firstgrid))
            self.textLog.appendPlainText(str(nextgrid))

    # CLEAR BUTTON FUNCTION
    def btn2_clear_fnc(self):
        self.textLog.clear()
        self.label1.setText("")

    # MOSAIC BUTTON FUNCTION
    def btn1_mosaic_fnc(self):
        self.textLog.clear()
        xml, xmlString = self.enableMosaic()
        # isMosaiced, firstgrid, nextgrid = WindowClass.isMosaiced()
        # self.textLog.appendPlainText("현재 모자이크 상태는 " + str(isMosaiced) + " 입니다.")
        # self.textLog.appendPlainText(str(firstgrid))
        # self.textLog.appendPlainText(str(nextgrid))
        self.textLog.appendPlainText(xmlString)

        if xml.attrib['valid'] == "1":
            self.label1.setText("모자이크가 활성화 되었습니다.")
        else:
            self.bnt3_current_fnc()
            self.label1.setText("모자이크가 활성화에 실패했습니다.")


if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec()
