# -*- coding: utf-8 -*-

#########################
# thissa@naver.com ######
# 2022.09.01       ######
#########################

import time
import os
import xml.etree.ElementTree as ET
import sys
import subprocess
import configparser
from PyQt5.QtWidgets import *
from PyQt5 import uic

# pyinstaller -w --uac-admin -F test.py admin 권한으로 실행되는 exe

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
        self.btn5_disableMosaic.clicked.connect(self.btn5_disable_mosaic_func)
        self.btn6_openNvcpl.clicked.connect(self.btn6_open_nvcpl_func)
        self.btn7_setDefault.clicked.connect(self.btn7_set_default_func)
        self.btn8_setAdvanced.clicked.connect(self.btn8_set_advanced_func)
        self.dialog_ok.accepted.connect(self.dialog_ok_func)
        self.chk1_sync.stateChanged.connect(self.chk_btn_func)
        self.combo_hz.currentIndexChanged.connect(self.combo_hz_func)
        self.combo_master.currentIndexChanged.connect(self.combo_master_func)

    ###############################################
    # FUNCTIONS ####################################
    ###############################################

    # 파워쉘의 권한을 Unrestricted로 설정합니다.
    def set_powershell_policy(self):
        f = os.popen("powershell.exe Get-ExecutionPolicy").read()
        if not f == "Unrestricted\n":
            os.popen("powershell.exe Set-ExecutionPolicy Unrestricted")
            time.sleep(2)
        else:
            self.textLog.appendPlainText("PowerShell ExecutionPolicy : Unrestricted")

    # 현재 PARMETER를 RETURN STR(CHECKBOX_SYNC), STR(COMBOBOX_HZ), STR(COMBOBOX_MASTER)
    def get_params(self):
        c_sync_ = self.chk1_sync.isChecked()
        hz_ = self.combo_hz.currentIndex()
        master_ = self.combo_master.currentIndex()
        return c_sync_, hz_, master_

    def set_params(self, c_sync_, hz_, master_):
        self.chk1_sync.setChecked(c_sync_)
        self.combo_hz.setCurrentIndex(hz_)
        self.combo_master.setCurrentIndex(master_)

    @staticmethod
    def generate_config(c_sync_, hz_, master_):
        config = configparser.ConfigParser()

        # 설정파일 오브잭트 만들기
        config['sicmo'] = {}
        config['sicmo']['cSync'] = str(c_sync_)
        config['sicmo']['hz'] = str(hz_)
        config['sicmo']['master'] = str(master_)

        with open('config.ini', 'w', encoding='utf-8') as configFile:
            config.write(configFile)

    @staticmethod
    def read_config():
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        c_sync_ = config['sicmo']['cSync']
        if c_sync_ == 'True':
            c_sync_ = True
        else:
            c_sync_ = False
        hz_ = config['sicmo']['hz']
        master_ = config['sicmo']['master']

        return c_sync_, hz_, master_

    # result 설정필요
    def enable_sync(self):
        current = self.combo_master_func()
        self.label1.setText("%s sync 활성화를 진행합니다." % current)
        if current == 'Master':
            result = os.popen(
                'C:\\configureQSync.exe +qsync master display 0,1 source HOUSESYNC vmode NTSCPALSECAM').read()
        elif current == 'Slave':
            result = os.popen('C:\\configureQSync.exe +qsync slave display 0,1').read()

    # 현재 싱크상태와 MASTER_SLAVE 상태 RETURN BOOL, STR, STR
    @staticmethod
    def is_synced():
        f = os.popen('C:\\configureQSync.exe ?queryTopo').read()
        is_synced = f.split('Is synced')[1].split('\n')[0].replace(" ", "").split(":")[-1] == 'YES'
        sync_state = f.split('Sync State')[1].split('\n')[0].replace(" ", "").split(":")[-1]
        return is_synced, sync_state, f

    # DISABLE_MOSAIC RETURN XML, STR
    def disable_mosaic(self):
        hz_ = self.combo_hz_func()
        f = "c:\\configureMosaic.exe set rows=1 cols=1 res=3840,2160,{0} out=0,0 nextgrid rows=1 cols=1 " \
            "res=3840,2160,{1} out=1,0".format(hz_, hz_)
        xml_string = os.popen(f).read()
        xml = ET.fromstring(xml_string)
        return xml, xml_string

    # 현재 상태 출력
    def print_current_state(self):
        _, first_grid, next_grid = self.is_mosaic()
        self.textLog.clear()
        self.textLog.appendPlainText(str(first_grid))
        self.textLog.appendPlainText(str(next_grid))

    # 프로그램 강제 종료
    def kill_process(self, xml):
        prevent_apps = self.find_prevent_apps(xml)
        for app_ in prevent_apps:
            subprocess.call('taskkill /IM %s /F' % app_)

    # FIND PREVENTAPP RETURN LIST
    @staticmethod
    def find_prevent_apps(xml):
        prevent_app = []
        prevent = xml[0].find('appspreventingmosaic')
        for app_ in prevent.iter('app'):
            prevent_app.append(app_.get('name'))
        return prevent_app

    # 모자이크 활성화 RETURN[xml, str]
    def enable_mosaic(self):
        hz_ = self.combo_hz_func()
        f = "C:\\configureMosaic.exe set cols=1 rows=2 res=3840,2160,{0} out=1,0 out=0,0 maxperf".format(hz_)

        xml_string = os.popen(f).read()
        xml = ET.fromstring(xml_string)
        return xml, xml_string

    # 모자이크 체크 RETURN[bool, list(firstgrid), list(nextgrid)]
    @staticmethod
    def is_mosaic():
        cmd = os.popen("c:\\configureMosaic.exe listconfigcmd").read().split(" ")

        # nextgrid가 cmdlist에 있으면 True, 없으면 False 반환
        is_mosaic = False if 'nextgrid' in cmd else True

        firstgrid = cmd[2:5]
        firstgrid.insert(0, '1st_grid')
        # nextgrid
        if is_mosaic:
            nextgrid = ''
        else:
            nextgrid_index = cmd.index('nextgrid')
            nextgrid = cmd[nextgrid_index:nextgrid_index + 4]

        return is_mosaic, firstgrid, nextgrid

    ###############################################
    # COMBOBOX FUNCTION ###########################
    ###############################################
    def combo_hz_func(self):
        hz__ = float(self.combo_hz.currentText()[:-2])
        return hz__

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
        c_sync_, hz_, master_ = self.get_params()
        self.generate_config(str(c_sync_), hz_, master_)
        print('ok')

    # ENABLE_SYNC_BUTTON
    def btn4_sync_func(self):
        # MASTER, SLAVE COMBO_BOX
        current = self.combo_master_func()
        self.textLog.clear()
        is_synced, sync_state, f = self.is_synced()
        is_mosaic, _, _ = self.is_mosaic()
        if is_synced:
            self.label1.setText("현재 %s 싱크 상태입니다." % sync_state)
            self.textLog.appendPlainText(f)
        elif not is_mosaic:
            self.label1.setText("먼저 모자이크 활성화가 필요합니다.")
            self.print_current_state()
        elif is_mosaic and not is_synced:
            self.enable_sync()
            is_synced, sync_state, f = self.is_synced()
            if is_synced:
                self.label1.setText("%s sync 활성화에 성공했습니다." % current)
            else:
                self.label1.setText("%s sync 활성화에 실패했습니다." % current)
            # self.textLog.appendPlainText(result)

    # CHECK_CURRENT_BUTTON
    # try except 본사 컴퓨터 테스트용
    def bnt3_current_func(self):
        self.textLog.clear()
        self.label1.setText("현재 상태입니다.")
        is_synced, sync_state, f = self.is_synced()
        is_mosaic, first_grid, next_grid = self.is_mosaic()
        if is_mosaic:
            self.textLog.appendPlainText("현재 모자이크 상태입니다.")
            self.textLog.appendPlainText(str(first_grid))
            self.textLog.appendPlainText(str(next_grid))
        else:
            self.textLog.appendPlainText("현재 모자이크 상태가 아닙니다.")
            self.textLog.appendPlainText(str(first_grid))
            self.textLog.appendPlainText(str(next_grid))
        if is_synced:
            self.textLog.appendPlainText("현재 %s 싱크 상태입니다." % sync_state)
        else:
            self.textLog.appendPlainText("현재 싱크 상태가 아닙니다.")

    # CLEAR_BUTTON
    def btn2_clear_func(self):
        self.textLog.clear()
        self.label1.setText("log is cleared!")

    # ENABLE_MOSAIC_BUTTON
    def btn1_mosaic_func(self):
        # if self.chk1_sync.isChecked():
        #     self.label1.setText("싱크작업을 진행합니다")
        is_mosaic, firstgrid, nextgrid = self.is_mosaic()
        if is_mosaic:
            self.textLog.clear()
            self.label1.setText("현재 모자이크 상태입니다.")
            self.print_current_state()
            return
        # 모자이크 활성화
        else:
            self.textLog.clear()
            xml, xml_string = self.enable_mosaic()
            hz_ = self.combo_hz_func()

            # 모자이크 활성화에 성공했을때
            if xml.attrib['valid'] == "1":
                if self.chk1_sync.isChecked():
                    self.enable_sync()
                self.label1.setText("{0}hz 모자이크 활성화에 성공했습니다.".format(hz_))
                self.print_current_state()

            else:
                prevent_app = self.find_prevent_apps(xml)
                reply = self.kill_process_info_event(prevent_app, True)
                if reply == QMessageBox.Yes:
                    self.textLog.appendPlainText('프로그램을 강제 종료하고 모자이크를 활성화 합니다')
                    self.kill_process(xml)
                    # RECURSION_FUNCTION
                    self.btn1_mosaic_func()
                    return
                else:
                    self.textLog.appendPlainText('응용프로그램 종료 후 다시 시도하세요.')
                    return

    # QMESSAGE_BOX
    def kill_process_info_event(self, prevent_app, is_enable_mosaic):
        if is_enable_mosaic:
            message = '\n프로그램이 실행 중 입니다. \n강제 종료 후 모자이크 활성화를 할까요?'
        else:
            message = '\n프로그램이 실행되고 있어 모자이크를 비활성화 할 수 없습니다. \n강제 종료 후 모자이크 비활성화를 할까요?'
        reply = QMessageBox.information(self, 'waring', str(prevent_app) + message, QMessageBox.Yes | QMessageBox.No)
        return reply

    # DISABLE_MOSAIC_BUTTON
    def btn5_disable_mosaic_func(self):
        # 만약 싱크 상태라면 해재한다.
        is_synced, sync_state, f = self.is_synced()
        if is_synced:
            os.popen('C:\\configureQSync.exe -qsync')
        # 모자이크 체크
        is_mosaic, firstgrid, nextgrid = self.is_mosaic()
        hz_ = self.combo_hz_func()
        if not is_mosaic:
            self.label1.setText("현재 모자이크 비활성화 상태입니다.")
            self.print_current_state()
            return
        # 모자이크 비활성화
        self.textLog.clear()
        xml, xml_string = self.disable_mosaic()
        self.textLog.appendPlainText(xml_string)
        valid = xml.get('valid')
        if valid == "0":
            prevent_app = self.find_prevent_apps(xml)
            reply = self.kill_process_info_event(prevent_app, False)
            if reply == QMessageBox.Yes:
                self.textLog.appendPlainText('프로그램을 강제 종료하고 모자이크를 비활성화 합니다')
                self.kill_process(xml)
                # RECURSION_FUNCTION
                self.btn5_disable_mosaic_func()
                return
            else:
                self.label1.setText("현재 모자이크 활성화 상태입니다.")
                self.textLog.appendPlainText('응용프로그램 종료 후 다시 시도하세요.')
                return
        elif xml.attrib['valid'] == "1":
            self.label1.setText("{0}hz 모자이크 비활성화에 성공했습니다.".format(hz_))
            self.print_current_state()

    # OPEN NVCPL
    def btn6_open_nvcpl_func(self):
        os.popen("C:\\Program Files\\WindowsApps\\NVIDIACorp.NVIDIAControlPanel_8.1.962.0_x64__56jybvy8sckqj\\nvcplui"
                 ".exe")
        self.label1.setText("nvidia 제어판을 엽니다")

    # Gobal3DPreset SET DEFAULT
    def btn7_set_default_func(self):
        self.set_powershell_policy()
        f = os.popen("powershell.exe .\\setDefault.ps1").read()
        if f == "Profile manager instance unavailable\n":
            self.textLog.clear()
            QMessageBox.information(self, "information", '<a href="https://www.nvidia.com/ko-kr/drivers/nvwmi/">NVWMI 설치가 '
                                                   '필요합니다.</a>')
            self.label1.setText("Profile3D 설정에 실패했습니다.")
            self.textLog.appendPlainText(f)
            self.textLog.appendPlainText('이 기능을 실행하려면 NVWMI가 필요합니다.')
        else:
            self.textLog.clear()
            self.label1.setText("Profile3D: Default 설정에 성공했습니다.")
            self.textLog.appendPlainText(f)

    # Gobal3DPreset SET Workstation App - Advanced Streaming
    def btn8_set_advanced_func(self):
        self.set_powershell_policy()
        f = os.popen("powershell.exe .\\setAdvanced.ps1").read()
        if f == "Profile manager instance unavailable\n":
            self.textLog.clear()
            QMessageBox.information(self, "information", '<a href="https://www.nvidia.com/ko-kr/drivers/nvwmi/">NVWMI 설치가 '
                                                   '필요합니다.</a>')
            self.label1.setText("Profile3D 설정에 실패했습니다.")
            self.textLog.appendPlainText(f)
            self.textLog.appendPlainText('이 기능을 실행하려면 NVWMI가 필요합니다.')
        else:
            self.textLog.clear()
            self.label1.setText("Profile3D: Workstation App - Advanced Streaming 설정에 성공했습니다.")
            self.textLog.appendPlainText(f)


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
        c_sync, hz, master = myWindow.read_config()
        myWindow.set_params(c_sync, int(hz), int(master))
    except KeyError:
        myWindow.generate_config(False, 0, 0)

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec()
