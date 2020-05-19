# -*- coding: utf-8 -*-
"""
开发者：周梦雄
最后更新日期：2019/8/7
"""
import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QDateEdit
from Ui_STA_config import *
import configparser


class MyMainWindow(QWidget, Ui_STA_config):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.cb1.setChecked(True)
        self.cb2.setChecked(True)
        self.cb4.setChecked(True)
        self.cb5.setChecked(True)
        self.cb6.setChecked(True)
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
        path=r'C:\pc_prodcheck\data'
        path_name = os.path.join(path, version_file)
        f = open(path_name, 'w')
        for line in zip(write_variables, write_value):
            f.write(line[0] + line[1] + '\n')
        f.close()
        self.lb13.setText("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                          "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                          "p, li { white-space: pre-wrap; }\n"
                          "</style></head><body style=\" font-family:\'黑体\'; font-size:10pt; font-weight:400; font-style:normal;\">\n"
                          "<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'微软雅黑\'; font-weight:600;\"><br /></p>\n"
                          "<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'微软雅黑\'; font-size:24pt; color:#00e500;\">配置文件保存成功！</span></p></body></html>")

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
                text1 = '%s年%s月%s日' % (text[:4], text[4:6], text[6:])
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    path = r"C:\pc_prodcheck\Config"
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
    ex = MyMainWindow()
    sys.exit(app.exec_())
