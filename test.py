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
from PyQt5.QtTest import *

# pyinstaller -w --uac-admin -F --icon=./mosync.ico test.py /////admin 권한으로 실행되는 exe

# UI파일 연결
# 단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("test.ui")[0]


# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QDialog, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.btn1_enable_mosaic.clicked.connect(self.btn1_enable_mosaic_func)
        self.btn2_enable_sync.clicked.connect(self.btn2_enable_sync_func)
        self.btn3_current.clicked.connect(self.bnt3_check_current_func)
        self.btn4_clear.clicked.connect(self.btn4_clear_func)
        self.btn5_disableMosaic.clicked.connect(self.btn5_disable_mosaic_func)
        self.btn6_openNvcpl.clicked.connect(self.btn6_open_nvcpl_func)
        self.btn7_setDefault.clicked.connect(self.btn7_set_default_func)
        self.btn8_setDynamic.clicked.connect(self.btn8_set_dynamic_func)
        self.btn9_openListener.clicked.connect(self.btn9_open_switchboard_listener_func)
        self.chk1_sync.stateChanged.connect(self.chk1_set_sync_func)
        self.chk2_set_default.stateChanged.connect(self.chk2_set_default_func)
        self.chk3_set_dynamic.stateChanged.connect(self.chk3_set_dynamic_func)
        self.combo_hz.currentIndexChanged.connect(self.get_combo_hz_func)
        self.combo_master.currentIndexChanged.connect(self.get_combo_master_func)

    ###############################################
    # FUNCTIONS ###################################
    ###############################################

    # QMESSAGE_BOX RETURN QMessageBox.Yex or QMessageBox.No
    def query_do_kill_processes(self, prevent_app, is_enable_mosaic):
        if is_enable_mosaic:
            message = '\n프로그램이 실행 중 입니다. \n강제 종료 후 모자이크 활성화를 할까요?'
        else:
            message = '\n프로그램이 실행되고 있어 모자이크를 비활성화 할 수 없습니다. \n강제 종료 후 모자이크 비활성화를 할까요?'
        reply = QMessageBox.information(self, 'waring', str(prevent_app) + message, QMessageBox.Yes | QMessageBox.No)
        return reply

    # 버튼 클릭시 기존의 label, log를 지우고 information에 message를 출력
    def set_info_message(self, message):
        self.btn4_clear_func()
        self.label1.setText(message)
        QTest.qWait(self, 100)

    # 파워쉘의 권한을 Unrestricted로 설정합니다.
    def set_powershell_policy(self):
        f = os.popen("powershell.exe Get-ExecutionPolicy").read()
        if not f == "Unrestricted\n":
            os.popen("powershell.exe Set-ExecutionPolicy Unrestricted")
            QTest.qWait(self, 1000)
        else:
            self.textLog.appendPlainText("PowerShell ExecutionPolicy : Unrestricted")

    # 현재 PARMETER를 RETURN STR(CHECKBOX_SYNC), STR(COMBOBOX_HZ), STR(COMBOBOX_MASTER)
    def get_params(self):
        c_sync_ = self.chk1_sync.isChecked()
        hz_ = self.combo_hz.currentIndex()
        master_ = self.combo_master.currentIndex()
        profile_ = self.chk2_set_default.isChecked()
        profile_d = self.chk3_set_dynamic.isChecked()
        return c_sync_, hz_, master_, profile_, profile_d

    def set_params(self, c_sync_, hz_, master_, profile_, profile_d):
        self.chk1_sync.setChecked(c_sync_)
        self.combo_hz.setCurrentIndex(hz_)
        self.combo_master.setCurrentIndex(master_)
        self.chk2_set_default.setChecked(profile_)
        self.chk3_set_dynamic.setChecked(profile_d)

    @staticmethod
    def generate_config(c_sync_, hz_, master_, profile_, profile_d):
        config = configparser.ConfigParser()

        # 설정파일 오브잭트 만들기
        config['sicmo'] = {}
        config['sicmo']['cSync'] = str(c_sync_)
        config['sicmo']['hz'] = str(hz_)
        config['sicmo']['master'] = str(master_)
        config['sicmo']['profile3d'] = str(profile_)
        config['sicmo']['profile_d'] = str(profile_d)

        with open('config.ini', 'w', encoding='utf-8') as configFile:
            config.write(configFile)

    # QWidget 창을 닫을때 config생성
    def closeEvent(self, QCloseEvent):
        c_sync_, hz_, master_, profile_, profile_d = self.get_params()
        self.generate_config(str(c_sync_), hz_, master_, str(profile_), str(profile_d))

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
        # Profile3d Default
        profile_ = config['sicmo']['profile3d']
        if profile_ == 'True':
            profile_ = True
        else:
            profile_ = False
        # Profile 3d Dynamic
        profile_d = config['sicmo']['profile_d']
        if profile_d == 'True':
            profile_d = True
        else:
            profile_d = False

        return c_sync_, hz_, master_, profile_, profile_d

    # result 설정필요
    def set_enable_sync(self):
        self.set_info_message("동기화를 실행합니다.")
        f = ""
        current = self.get_combo_master_func()
        if current == 'Master':
            f = os.popen(
                'C:\\configureQSync.exe +qsync master display 0,1 source HOUSESYNC vmode NTSCPALSECAM').read()
        elif current == 'Slave':
            f = os.popen('C:\\configureQSync.exe +qsync slave display 0,1').read()
        self.textLog.appendPlainText(f)
        QTest.qWait(self, 3000)

    # 현재 동기화상태와 MASTER_SLAVE 상태 체크: RETURN: BOOL, STR, STR
    def is_synced(self):
        f = os.popen('C:\\configureQSync.exe ?queryTopo').read()
        self.textLog.appendPlainText(f)
        if "UNSYNCED" in f:
            is_synced = False
        else:
            is_synced = True
        sync_state = f.split('Sync State')[1].split('\n')[0].replace(" ", "").split(":")[-1]
        return is_synced, sync_state, f

    # SET DISABLE_MOSAIC RETURN XML, STR
    def set_disable_mosaic(self):
        self.set_info_message("동기화를 비활성화 합니다.")
        hz_ = self.get_combo_hz_func()
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
    def set_enable_mosaic(self):
        hz_ = self.get_combo_hz_func()
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
    def get_combo_hz_func(self):
        hz__ = float(self.combo_hz.currentText()[:-2])
        return hz__

    def get_combo_master_func(self):
        current = self.combo_master.currentText()
        self.label1.setText("Sync parameter set " + current)
        return current

    ###############################################
    # CHECKBOX FUNCTION ###########################
    ###############################################

    # ENABLE MOSAIC에 연관된 옵션
    def chk1_set_sync_func(self):
        if self.chk1_sync.isChecked():
            self.label1.setText("모자이크 활성화에 이어서 동기화작업을 실행합니다.")
            self.btn4_sync.setEnabled(False)
        elif not self.chk1_sync.isChecked() and not self.chk3_set_dynamic.isChecked():
            self.label1.setText("모자이크 활성화만 실행합니다.")
            self.btn4_sync.setEnabled(True)
        elif not self.chk1_sync.isChecked() and self.chk3_set_dynamic.isChecked():
            self.label1.setText("모자이크 활성화에 이어서 Profile 3d setting을 Dynamic으로 설정합니다.")
            self.btn4_sync.setEnabled(True)

    # DISABLE MOSAIC에 연관된 옵션
    def chk2_set_default_func(self):
        if self.chk2_set_default.isChecked():
            self.label1.setText("모자이크 비활성화에 이어서 Profile 3d setting을 default로 설정합니다.")
        if not self.chk2_set_default.isChecked():
            self.btn7_setDefault.setStyleSheet('background-color: rgb()')
            self.label1.setText("모자이크 비활성화만 실행합니다.")

    def chk3_set_dynamic_func(self):
        if self.chk3_set_dynamic.isChecked():
            self.label1.setText("모자이크 활성화에 이어서 Profile 3d setting을 Dynamic으로 설정합니다.")
        elif not self.chk3_set_dynamic.isChecked() and not self.chk1_sync.isChecked():
            self.label1.setText("모자이크 활성화만 실행합니다.")
        elif not self.chk3_set_dynamic.isChecked() and self.chk1_sync.isChecked():
            self.label1.setText("모자이크 활성화와 동기화를 실행합니다.")

    ###############################################
    # BUTTON FUNCTION #############################
    ###############################################

    # ENABLE_SYNC_BUTTON
    def btn2_enable_sync_func(self):
        self.set_info_message("동기화를 실행합니다.")
        # MASTER, SLAVE COMBO_BOX
        current = self.get_combo_master_func()
        self.textLog.clear()
        is_synced, sync_state, f = self.is_synced()
        is_mosaic, _, _ = self.is_mosaic()
        if is_synced:
            self.label1.setText("현재 동기화(%s) 상태입니다." % sync_state)
            self.textLog.appendPlainText(f)
        elif not is_mosaic:
            self.label1.setText("먼저 모자이크 활성화가 필요합니다.")
            self.print_current_state()
        elif is_mosaic and not is_synced:
            self.set_enable_sync()
            # 10초 wait
            QTest.qWait(self, 8000)
            is_synced, sync_state, f = self.is_synced()
            if is_synced:
                self.set_info_message("동기화(%s) 작업에 성공했습니다." % sync_state)

    # CHECK_CURRENT_BUTTON
    def bnt3_check_current_func(self):
        self.textLog.clear()
        is_synced, sync_state, f = self.is_synced()
        is_mosaic, first_grid, next_grid = self.is_mosaic()

        # 모자이크상태라면
        if is_mosaic:
            self.textLog.appendPlainText(str(first_grid))
            self.textLog.appendPlainText(str(next_grid))
            mosaic_message = "활성화"
            # 모자이크 버튼세팅
            self.btn1_enable_mosaic.setStyleSheet('color: grey; font: Italic')
        else:
            self.textLog.appendPlainText(str(first_grid))
            self.textLog.appendPlainText(str(next_grid))
            mosaic_message = "비활성화"
            # 모자이크 버튼세팅
            self.btn1_enable_mosaic.setStyleSheet("")

        # 싱크상태라면
        if is_synced:
            sync_message = "활성화 ({0})".format(sync_state)
            # 싱크 버튼세팅
            self.btn2_enable_sync.setStyleSheet('color: grey; font: Italic')
        else:
            sync_message = "비활성화"
            # 싱크 버튼세팅
            self.btn2_enable_sync.setStyleSheet("")
        t = "모자이크-{0}, 동기화-{1} 상태입니다.".format(mosaic_message, sync_message)
        self.label1.setText(t)

    # CLEAR_BUTTON
    def btn4_clear_func(self):
        # 버튼 텍스트 클리어
        self.btn2_enable_sync.setStyleSheet("")
        self.btn1_enable_mosaic.setStyleSheet("")

        # 라벨, 로그 클리어
        self.textLog.clear()
        self.label1.setText("Information window ")

    # ENABLE_MOSAIC_BUTTON
    def btn1_enable_mosaic_func(self):
        self.set_info_message("모자이크 활성화를 실행합니다.")
        is_mosaic, firstgrid, nextgrid = self.is_mosaic()
        if is_mosaic:
            self.textLog.clear()
            self.label1.setText("현재 모자이크 상태입니다.")
            self.print_current_state()
            return

        # 모자이크 상태가 아니면 모자이크 활성화
        if not is_mosaic:
            self.textLog.clear()
            xml, xml_string = self.set_enable_mosaic()
            hz_ = self.get_combo_hz_func()

            # 모자이크 활성화에 성공했을때
            if xml.attrib['valid'] == "1":
                self.label1.setText("{0}hz 모자이크 활성화에 성공했습니다.".format(hz_))
                self.print_current_state()

                # CHECKBOX3 DYNAMIC 세팅
                if self.chk3_set_dynamic.isChecked():
                    QTest.qWait(self, 3000)
                    self.btn8_set_dynamic_func()

                # CHECKBOX2 동기화작업
                if self.chk1_sync.isChecked():
                    QTest.qWait(self, 3000)
                    self.btn2_enable_sync_func()
            else:
                prevent_app = self.find_prevent_apps(xml)
                reply = self.query_do_kill_processes(prevent_app, True)
                if reply == QMessageBox.Yes:
                    self.textLog.appendPlainText('프로그램을 강제 종료하고 모자이크를 활성화 합니다')
                    self.kill_process(xml)
                    # RECURSION_FUNCTION
                    self.btn1_enable_mosaic_func()
                    return
                else:
                    self.btn4_clear_func()
                    self.textLog.appendPlainText('응용프로그램 종료 후 다시 시도하세요.')
                    return

    # DISABLE_MOSAIC_BUTTON
    def btn5_disable_mosaic_func(self):
        # 만약 동기화 상태라면 해재한다.
        is_synced, sync_state, f = self.is_synced()
        if is_synced:
            self.set_info_message("동기화 비활성화를 실행합니다.")
            os.popen('C:\\configureQSync.exe -qsync')
        hz_ = self.get_combo_hz_func()

        # 만약 모자이크 상태가 아니라면
        is_mosaic, firstgrid, nextgrid = self.is_mosaic()
        if not is_mosaic:
            self.label1.setText("현재 모자이크 비활성화 상태입니다.")
            self.print_current_state()
            return
        # 만약 모자이크 상태라면 모자이크 비활성화
        if is_mosaic:
            # profile3D 체크
            if self.chk2_set_default.isChecked():
                self.btn7_set_default_func()
                QTest.qWait(self, 3000)
            xml, xml_string = self.set_disable_mosaic()
            self.textLog.appendPlainText(xml_string)
            valid = xml.get('valid')
            # 모자이크 비활성화에 실패했다면
            if valid == "0":
                prevent_app = self.find_prevent_apps(xml)
                reply = self.query_do_kill_processes(prevent_app, False)
                if reply == QMessageBox.Yes:
                    self.textLog.appendPlainText('프로그램을 강제 종료하고 모자이크 비활성화를 실행합니다')
                    self.kill_process(xml)
                    # RECURSION_FUNCTION
                    self.btn5_disable_mosaic_func()
                    return
                else:
                    self.label1.setText("현재 모자이크 활성화 상태입니다.")
                    self.textLog.appendPlainText('응용프로그램 종료 후 다시 시도하세요.')
                    return
            # 모자이크 비활성화에 성공했다면
            elif xml.attrib['valid'] == "1":
                self.label1.setText("{0}hz 모자이크 비활성화에 성공했습니다.".format(hz_))
                self.print_current_state()

    # OPEN NVCPL
    def btn6_open_nvcpl_func(self):
        self.set_info_message("nvidia 제어판을 실행합니다.")
        os.popen("C:\\Program Files\\WindowsApps\\NVIDIACorp.NVIDIAControlPanel_8.1.962.0_x64__56jybvy8sckqj\\nvcplui"
                 ".exe")

    # Gobal3DPreset SET DEFAULT
    def btn7_set_default_func(self):
        is_mosaic, _, _ = self.is_mosaic()
        if is_mosaic:
            self.set_powershell_policy()
            self.set_info_message("Profile3D: default로 설정합니다.")
            QTest.qWait(self, 1000)
            f = os.popen("powershell.exe .\\.ddi\\setDefault.ps1").read()
            if "succeed" in f:
                self.label1.setText("Profile3D: Default 설정에 성공했습니다.")
                self.textLog.appendPlainText(f)
            else:
                QMessageBox.information(self, "information",
                                        '<a href="https://www.nvidia.com/ko-kr/drivers/nvwmi/">NVWMI 설치가 '
                                        '필요합니다.</a>')
                self.label1.setText("Profile3D 설정에 실패했습니다.")
                self.textLog.appendPlainText(f)
                self.textLog.appendPlainText('이 기능을 실행하려면 NVWMI가 필요합니다.')
                return
        else:
            self.set_info_message("모자이크 상태에서만 실행 가능합니다.")

    # Global 3DPreset SET Workstation App - Advanced Streaming
    def btn8_set_dynamic_func(self):
        is_mosaic, _, _ = self.is_mosaic()
        if is_mosaic:
            self.set_powershell_policy()
            self.set_info_message("Profile3D: Workstation App - Dynamic Streaming으로 설정합니다.")
            QTest.qWait(self, 1000)
            f = os.popen("powershell.exe .\\.ddi\\setDynamic.ps1").read()
            if "succeed" in f:
                self.label1.setText("Profile3D: Workstation App - Dynamic Streaming 설정에 성공했습니다.")
                self.textLog.appendPlainText(f)
            else:
                QMessageBox.information(self, "information",
                                        '<a href="https://www.nvidia.com/ko-kr/drivers/nvwmi/">NVWMI 설치가 '
                                        '필요합니다.</a>')
                self.label1.setText("Profile3D 설정에 실패했습니다.")
                self.textLog.appendPlainText(f)
                self.textLog.appendPlainText('이 기능을 실행하려면 NVWMI가 필요합니다.')
                return
        else:
            self.set_info_message("모자이크 상태에서만 실행 가능합니다.")

    def btn9_open_switchboard_listener_func(self):
        f = os.popen("tasklist").read()
        if "SwitchboardListener.exe" not in f:
            os.popen("D:\\UE_4.27\\Engine\\Binaries\\Win64\\SwitchboardListener.exe")
            f = os.popen("tasklist").read()
            if "SwitchboardListener.exe" in f:
                self.label1.setText("Switchboard Listener 실행에 성공했습니다.")
        else:
            self.set_info_message("Switchboard Listener가 이미 실행중입니다.")


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
        c_sync, hz, master, profile_default, profile_dynamic = myWindow.read_config()
        myWindow.set_params(c_sync, int(hz), int(master), profile_default, profile_dynamic)
    except KeyError:
        myWindow.generate_config(False, 1, 0, False, False)

    # 프로그램 화면을 보여주는 코드
    myWindow.show()
    myWindow.btn4_clear_func()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec()
