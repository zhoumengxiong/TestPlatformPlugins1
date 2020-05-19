# -*- coding: utf-8 -*-
"""
开发者：周梦雄
最后更新日期：2019/7/30
"""
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel,
                             QRadioButton, QButtonGroup, QInputDialog,
                             QCheckBox, QTextBrowser)
import sys
import os
from PyQt5.QtGui import QIcon
import configparser


class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(450, 200, 660, 450)
        self.setWindowTitle('STA测试工装参数配置')
        self.setWindowIcon(QIcon('tittle.ico'))

        self.lb1 = QLabel('厂商代码：', self)
        self.lb1.move(20, 20)

        self.lb2 = QLabel('芯片代码：', self)
        self.lb2.move(20, 80)

        self.lb3 = QLabel('版本日期：', self)
        self.lb3.move(20, 140)

        self.lb4 = QLabel('版本号：', self)
        self.lb4.move(20, 200)

        self.lb5 = QLabel('硬件版本ID：', self)
        self.lb5.move(20, 260)

        self.lb6 = QLabel('详细版本信息：', self)
        self.lb6.move(20, 320)

        self.lb7 = QLabel('GD', self)
        self.lb7.move(150, 20)

        self.lb8 = QLabel('03', self)
        self.lb8.move(150, 80)

        self.lb9 = QLabel('2019年3月25日', self)
        self.lb9.move(150, 140)

        self.lb10 = QLabel('00 01', self)
        self.lb10.move(150, 200)

        self.lb11 = QLabel('135', self)
        self.lb11.move(150, 260)

        self.lb12 = QLabel('V02S03LGWT0100', self)
        self.lb12.move(150, 320)

        self.lb13 = QLabel('配置文件尚未保存！', self)
        self.lb13.move(20, 380)
        self.lb13.setStyleSheet(
            "* { color: #FF6666;font-size:20px;font-weight:bold;}")

        self.bt1 = QPushButton('修改厂商代码', self)
        self.bt1.move(320, 20)

        self.bt2 = QPushButton('修改芯片代码', self)
        self.bt2.move(320, 80)

        self.bt3 = QPushButton('修改版本日期', self)
        self.bt3.move(320, 140)

        self.bt4 = QPushButton('修改版本号', self)
        self.bt4.move(320, 200)

        self.bt5 = QPushButton('修改硬件版本ID', self)
        self.bt5.move(320, 260)

        self.bt6 = QPushButton('修改详细版本信息', self)
        self.bt6.move(320, 320)

        self.bt7 = QPushButton('保存配置文件', self)
        self.bt7.move(320, 380)
        self.bt7.setStyleSheet(
            "QPushButton{color:rgb(255,255,255);background-color:#424242}"
            "QPushButton:hover{background-color: #64dd17;border-color: #64dd17;}"
        )

        self.setStyleSheet(
            # "*{background:lightsteelblue;}"  #背景色设为 白色
            # 字体颜色 大小为15 常规 微软雅黑字体
            "*{color:rgb(0,0,0);font-size:15px;font-weight:normal;font-family:微软雅黑;}"
            "*:hover{color:rgb(100,100,100);}"
            "QPushButton{border:1px solid #333333;padding:4px;min-width: 65px;min-height: 12px;}"
            # "QPushButton:hover{background-color: #FFFFCC;border-color: #444444;}"
            "QPushButton{border-radius:3px}"
            "QPushButton:disabled{color: #333333;}")

        self.cb1 = QCheckBox('过零检测', self)
        self.cb1.move(530, 20)
        self.cb1.setChecked(True)

        self.cb2 = QCheckBox('三相过零检测', self)
        self.cb2.move(530, 80)
        self.cb2.setChecked(True)

        self.cb3 = QCheckBox('一级厂商代码', self)
        self.cb3.move(530, 140)

        self.cb4 = QCheckBox('二级厂商代码', self)
        self.cb4.move(530, 200)
        self.cb4.setChecked(True)

        self.cb5 = QCheckBox('芯片ID检测', self)
        self.cb5.move(530, 260)
        self.cb5.setChecked(True)

        self.cb6 = QCheckBox('模块ID写入', self)
        self.cb6.move(530, 320)
        self.cb6.setChecked(True)

        # bg_code = QButtonGroup(self)
        # bg_code.addButton(self.cb3)
        # bg_code.addButton(self.cb4)

        self.show()

        self.bt1.clicked.connect(self.showDialog)
        self.bt2.clicked.connect(self.showDialog)
        self.bt3.clicked.connect(self.showDialog)
        self.bt4.clicked.connect(self.showDialog)
        self.bt5.clicked.connect(self.showDialog)
        self.bt6.clicked.connect(self.showDialog)
        self.bt7.clicked.connect(self.write_ini)
        self.cb1.stateChanged.connect(self.modify_test_option)
        self.cb2.stateChanged.connect(self.modify_test_option)
        self.cb3.toggled.connect(self.modify_test_option)
        self.cb4.toggled.connect(self.modify_test_option)
        self.cb5.stateChanged.connect(self.modify_test_option)
        self.cb6.stateChanged.connect(self.modify_test_option)

    def write_ini(self):
        write_variables = [
            self.lb1.text(),
            self.lb2.text(),
            self.lb3.text(),
            self.lb4.text(),
            self.lb5.text(),
            self.lb6.text()
        ]
        write_value = [
            self.lb7.text(),
            self.lb8.text(),
            self.lb9.text(),
            self.lb10.text(),
            self.lb11.text(),
            self.lb12.text()
        ]
        version_file = r"VerInfoDef_two.txt"
        path_name = os.path.join(path, version_file)
        f = open(path_name, 'w')
        for line in zip(write_variables, write_value):
            f.write(line[0] + line[1] + '\n')
        f.close()
        self.lb13.setText('配置文件保存成功！')
        self.lb13.setStyleSheet(
            "* { color: #00c853;font-size:20px;font-weight:bold;}")

    def showDialog(self):
        sender = self.sender()
        chip_code = ['00', '01', '03']
        version_cust = [
            '00 01', '00 90', '01 21', '01 51', '01 81', '02 11', '02 70',
            '03 01', '03 31', '03 61', '04 21', '04 51', '04 81', '05 41',
            '05 70', '06 00', '06 30', '10 90', '20 90', '21 81', '25 41'
        ]
        vendor_code = [
            '65', 'BD', 'BG', 'BW', 'CL', 'CY', 'DW', 'ES', 'FC', 'GD', 'GY',
            'HB', 'HL', 'HR', 'HT', 'HY', 'JC', 'JF', 'JH', 'JN', 'JR', 'KD',
            'LD', 'LN', 'NK', 'NL', 'NZ', 'PH', 'RS', 'SD', 'SS', 'SX', 'TC',
            'TX', 'WE', 'WN', 'WS', 'XJ', 'XL', 'YD', 'YH', 'ZN'
        ]
        id_hw = ['105', '107', '135']
        if sender == self.bt1:
            text, ok = QInputDialog.getItem(self, '修改厂商代码', '选择厂商代码：',
                                            vendor_code)
            if ok:
                self.lb7.setText(text)
        elif sender == self.bt2:
            text, ok = QInputDialog.getItem(self, '修改芯片代码', '请选择芯片代码：',
                                            chip_code)
            if ok:
                self.lb8.setText(text)
        elif sender == self.bt3:
            text, ok = QInputDialog.getText(self, '修改版本日期',
                                            '请输入版本日期(格式：20190325)：')
            if ok and text != '':
                text1 = '%s年%s月%s日' % (text[:4], text[5:6], text[6:])
                self.lb9.setText(text1)
        elif sender == self.bt4:
            text, ok = QInputDialog.getItem(self, '修改版本号', '请选择版本号：',
                                            version_cust)
            if ok:
                self.lb10.setText(text)
        elif sender == self.bt5:
            text, ok = QInputDialog.getItem(self, '修改硬件版本ID', '请选择硬件版本ID：',
                                            id_hw)
            if ok:
                self.lb11.setText(text)
        elif sender == self.bt6:
            text, ok = QInputDialog.getText(self, '修改详细版本信息', '输入详细版本信息：')
            if ok and text != '':
                self.lb12.setText(text)

    def modify_test_option(self):
        sender = self.sender()
        if sender == self.cb1:
            if self.cb1.isChecked():
                prod_check['withoutzerocross'] = '0'
            else:
                prod_check['withoutzerocross'] = '1'
        elif sender == self.cb2:
            if self.cb2.isChecked():
                prod_check['withoutzerocrossThree'] = '0'
            else:
                prod_check['withoutzerocrossThree'] = '1'
        elif sender == self.cb3:
            if self.cb3.isChecked():
                prod_check['SetCustomInfo'] = '0'
            else:
                prod_check['SetCustomInfo'] = '1'
        elif sender == self.cb4:
            if self.cb4.isChecked():
                prod_check['SetCustomInfo_two'] = '0'
            else:
                prod_check['SetCustomInfo_two'] = '1'
        elif sender == self.cb5:
            if self.cb5.isChecked():
                prod_check['ChipIdCheck'] = '0'
            else:
                prod_check['ChipIdCheck'] = '1'
        elif sender == self.cb6:
            if self.cb6.isChecked():
                prod_check['ModuleIdWrite'] = '0'
            else:
                prod_check['ModuleIdWrite'] = '1'
        with open(path_name, 'w') as configfile:
            config.write(configfile)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    path = r"E:\Python\GUI_project"
    version_file = r"STAprodCheckCfg.ini"
    path_name = os.path.join(path, version_file)
    config = configparser.ConfigParser()
    config.read(path_name)
    prod_check = config['ProdCheck']
    test_items = ['macmode', 'phaseset', 'withoutzerocross', 'withouthardwareid',
                  'setcustominfo', 'setcustominfo_two', 'chipidcheck', 'moduleidwrite', 'withoutzerocrossthree']
    init_value = ['2', '2', '0', '0', '1', '0', '0', '0', '0']
    for k, v in zip(test_items, init_value):
        prod_check[k] = v
    with open(path_name, 'w') as configfile:
        config.write(configfile)
    ex = Example()
    sys.exit(app.exec_())
