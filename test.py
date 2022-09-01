# -*- coding: utf-8 -*-

#########################
# thissa@naver.com ######
# 2022.09.01       ######
#########################


import os
import xml.etree.ElementTree as ET
import sys
import subprocess
import configparser
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

        self.btn1_mosaic.clicked.connect(self.btn1_mosaic_func)
        self.btn2_clear.clicked.connect(self.btn2_clear_func)
        self.btn3_current.clicked.connect(self.bnt3_current_func)
        self.btn4_sync.clicked.connect(self.btn4_sync_func)
        self.btn5_disableMosaic.clicked.connect(self.btn5_disableMosaic_func)
        self.dialog_ok.accepted.connect(self.dialog_ok_func)
        self.chk1_sync.stateChanged.connect(self.chk_btn_func)
        self.combo_hz.currentIndexChanged.connect(self.combo_hz_func)
        self.combo_master.currentIndexChanged.connect(self.combo_master_func)

    ###############################################
    # FUNTIONS ####################################
    ###############################################

    # 현재 PARMETER를 RETURN STR(CHECKBOX_SYNC), STR(COMBOBOX_HZ), STR(COMBOBOX_MASTER)
    def getParam(self):
        cSync = self.chk1_sync.isChecked()
        hz = self.combo_hz.currentIndex()
        master = self.combo_master.currentIndex()
        return cSync, hz, master

    def setParam(self, cSync, hz, master):
        self.chk1_sync.setChecked(cSync)
        self.combo_hz.setCurrentIndex(hz)
        self.combo_master.setCurrentIndex(master)

    @staticmethod
    def generateConfig(cSync, hz, master):
        config = configparser.ConfigParser()

        # 설정파일 오브잭트 만들기
        config['sicmo'] = {}
        config['sicmo']['cSync'] = str(cSync)
        config['sicmo']['hz'] = str(hz)
        config['sicmo']['master'] = str(master)

        with open('config.ini', 'w', encoding='utf-8') as configFile:
            config.write(configFile)

    @staticmethod
    def readConfig():
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        cSync = config['sicmo']['cSync']
        if cSync == 'True':
            cSync = True
        else:
            cSync = False
        hz = config['sicmo']['hz']
        master = config['sicmo']['master']

        return cSync, hz, master

    def enableSync(self):
        current = self.combo_master_func()
        self.label1.setText("%s sync 활성화를 진행합니다." % current)
        if current == 'Master':
            result = os.popen(
                'C:\\configureQSync.exe +qsync master display 0,1 source HOUSESYNC vmode NTSCPALSECAM').read()
        elif current == 'Slave':
            result = os.popen('C:\\configureQSync.exe +qsync slave display 0,1').read()

    # 현재 싱크상태와 MASTER_SLAVE 상태 RETURN BOOL, STR, STR
    @staticmethod
    def isSynced():
        f = os.popen('C:\\configureQSync.exe ?queryTopo').read()
        Is_synced = f.split('Is synced')[1].split('\n')[0].replace(" ", "").split(":")[-1] == 'YES'
        syncState = f.split('Sync State')[1].split('\n')[0].replace(" ", "").split(":")[-1]
        return Is_synced, syncState, f

    # DISABLE_MOSAIC RETURN XML, STR
    def disableMosaic(self):
        hz = self.combo_hz_func()
        f = "c:\\configureMosaic.exe set rows=1 cols=1 res=3840,2160,{0} out=0,0 nextgrid rows=1 cols=1 " \
            "res=3840,2160,{1} out=1,0".format(hz, hz)
        xmlString = os.popen(f).read()
        xml = ET.fromstring(xmlString)
        return xml, xmlString

    # 현재 상태 출력
    def printCurrentState(self):
        isMosaiced, firstgrid, nextgrid = self.isMosaiced()
        self.textLog.clear()
        self.textLog.appendPlainText(str(firstgrid))
        self.textLog.appendPlainText(str(nextgrid))

    # 프로그램 강제 종료
    def killProcess(self, xml):
        preventApps = self.findPreventApp(xml)
        for app in preventApps:
            subprocess.call('taskkill /IM %s /F' % app)

    # FIND PREVENTAPP RETURN LIST
    @staticmethod
    def findPreventApp(xml):
        preventApp = []
        prevent = xml[0].find('appspreventingmosaic')
        for app in prevent.iter('app'):
            preventApp.append(app.get('name'))
        return preventApp

    # 모자이크 활성화 RETURN[xml, str]
    def enableMosaic(self):
        hz = self.combo_hz_func()
        f = "C:\\configureMosaic.exe set cols=1 rows=2 res=3840,2160,{0} out=1,0 out=0,0 maxperf".format(hz)

        xmlString = os.popen(f).read()
        xml = ET.fromstring(xmlString)
        return xml, xmlString

    # 모자이크 체크 RETURN[bool, list(firstgrid), list(nextgrid)]
    @staticmethod
    def isMosaiced():
        cmd = os.popen("c:\\configureMosaic.exe listconfigcmd").read().split(" ")

        # nextgrid가 cmdlist에 있으면 True, 없으면 False 반환
        isMosaiced = False if 'nextgrid' in cmd else True

        firstgrid = cmd[2:5]
        firstgrid.insert(0, '1st_grid')
        # nextgrid
        if isMosaiced:
            nextgrid = ''
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

    def combo_master_func(self):
        current = self.combo_master.currentText()
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

    # DIALOG_OK_BUTTON
    def dialog_ok_func(self):
        cSync, hz, master = self.getParam()
        self.generateConfig(str(cSync), hz, master)
        print('ok')

    # ENABLE_SYNC_BUTTON
    def btn4_sync_func(self):
        # MASTER, SLAVE COMBO_BOX
        current = self.combo_master_func()
        self.textLog.clear()
        IsSynced, syncState, f = self.isSynced()
        isMosaic, _, _ = self.isMosaiced()
        if IsSynced:
            self.label1.setText("현재 %s 싱크 상태입니다." % syncState)
            self.textLog.appendPlainText(f)
        elif not isMosaic:
            self.label1.setText("먼저 모자이크 활성화가 필요합니다.")
            self.printCurrentState()
        elif isMosaic and not IsSynced:
            self.enableSync()
            IsSynced, syncState, f = self.isSynced()
            if IsSynced:
                self.label1.setText("%s sync 활성화에 성공했습니다." % current)
            else:
                self.label1.setText("%s sync 활성화에 실패했습니다." % current)
            # self.textLog.appendPlainText(result)

    # CHECK_CURRENT_BUTTON
    # try except 본사 컴퓨터 테스트용
    def bnt3_current_func(self):
        try:
            self.textLog.clear()
            self.label1.setText("현재 상태입니다.")
            IsSynced, syncState, f = self.isSynced()
            isMosaiced, firstgrid, nextgrid = self.isMosaiced()
            if isMosaiced:
                self.textLog.appendPlainText("현재 모자이크 상태입니다.")
                self.textLog.appendPlainText(str(firstgrid))
                self.textLog.appendPlainText(str(nextgrid))
            else:
                self.textLog.appendPlainText("현재 모자이크 상태가 아닙니다.")
                self.textLog.appendPlainText(str(firstgrid))
                self.textLog.appendPlainText(str(nextgrid))
            if IsSynced:
                self.textLog.appendPlainText("현재 %s 싱크 상태입니다." % syncState)
            else:
                self.textLog.appendPlainText("현재 싱크 상태가 아닙니다.")
        except Exception as e:
            self.textLog.appendPlainText(str(e))

    # CLEAR_BUTTON
    def btn2_clear_func(self):
        self.textLog.clear()
        self.label1.setText("log is cleared!")

    # ENABLE_MOSAIC_BUTTON
    def btn1_mosaic_func(self):
        # if self.chk1_sync.isChecked():
        #     self.label1.setText("싱크작업을 진행합니다")
        isMosaiced, firstgrid, nextgrid = self.isMosaiced()
        if isMosaiced:
            self.textLog.clear()
            self.label1.setText("현재 모자이크 상태입니다.")
            self.printCurrentState()
            return
        # 모자이크 활성화
        else:
            self.textLog.clear()
            xml, xmlString = self.enableMosaic()
            hz = self.combo_hz_func()

            # 모자이크 활성화에 성공했을때
            if xml.attrib['valid'] == "1":
                if self.chk1_sync.isChecked():
                    self.enableSync()
                self.label1.setText("{0}hz 모자이크 활성화에 성공했습니다.".format(hz))
                self.printCurrentState()

            else:
                preventApp = self.findPreventApp(xml)
                reply = self.killProcess_info_event(preventApp, True)
                if reply == QMessageBox.Yes:
                    self.textLog.appendPlainText('프로그램을 강제 종료하고 모자이크를 활성화 합니다')
                    self.killProcess(xml)
                    # RECURSION_FUNCTION
                    self.btn1_mosaic_func()
                    return
                else:
                    self.textLog.appendPlainText('응용프로그램 종료 후 다시 시도하세요.')
                    return

    # QMESSAGE_BOX
    def killProcess_info_event(self, preventApp, isEnableMosaic):
        if isEnableMosaic:
            message = '\n프로그램이 실행 중 입니다. \n강제 종료 후 모자이크 활성화를 할까요?'
        else:
            message = '\n프로그램이 실행되고 있어 모자이크를 비활성화 할 수 없습니다. \n강제 종료 후 모자이크 비활성화를 할까요?'
        reply = QMessageBox.information(self, 'waring', str(preventApp) + message, QMessageBox.Yes | QMessageBox.No)
        return reply

    # DISABLE_MOSAIC_BUTTON
    def btn5_disableMosaic_func(self):
        # 만약 싱크 상태라면 해재한다.
        IsSynced, syncState, f = self.isSynced()
        if IsSynced:
            os.popen('C:\\configureQSync.exe -qsync')
        # 모자이크 체크
        isMosaiced, firstgrid, nextgrid = self.isMosaiced()
        hz = self.combo_hz_func()
        if not isMosaiced:
            self.label1.setText("현재 모자이크 비활성화 상태입니다.")
            self.printCurrentState()
            return
        # 모자이크 비활성화
        self.textLog.clear()
        xml, xmlString = self.disableMosaic()
        self.textLog.appendPlainText(xmlString)
        valid = xml.get('valid')
        if valid == "0":
            preventApp = self.findPreventApp(xml)
            reply = self.killProcess_info_event(preventApp, False)
            if reply == QMessageBox.Yes:
                self.textLog.appendPlainText('프로그램을 강제 종료하고 모자이크를 비활성화 합니다')
                self.killProcess(xml)
                # RECURSION_FUNCTION
                self.btn5_disableMosaic_func()
                return
            else:
                self.label1.setText("현재 모자이크 활성화 상태입니다.")
                self.textLog.appendPlainText('응용프로그램 종료 후 다시 시도하세요.')
                return
        elif xml.attrib['valid'] == "1":
            self.label1.setText("{0}hz 모자이크 비활성화에 성공했습니다.".format(hz))
            self.printCurrentState()


###############################################
# main #
###############################################

if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()
    # ini파일 읽어서 적용하기
    try:
        cSync, hz, master = myWindow.readConfig()
        myWindow.setParam(cSync, int(hz), int(master))
    except:
        myWindow.generateConfig(False, 0, 0)

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec()
