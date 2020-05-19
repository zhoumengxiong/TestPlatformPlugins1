# -*- coding: utf-8 -*-
"""
开发者：周梦雄
最后更新日期：2019/12/26
"""
import sys
import os
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QMainWindow)
from Ui_SetConfig import *
import configparser


class MyMainWindow(QMainWindow, Ui_SetConfig):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()
        self.bt_save.clicked.connect(self.write_ini)

    def write_ini(self):
        config = configparser.ConfigParser()
        path_name = r"C:\Users\Dream\Desktop\更新带屏读ID软件\最新读ID(带屏)Debug  20190621\Debug带瓶的\IniFile\FiterParam.ini"
        config.read(path_name)  # 读文件
        section = r"ErJiBiDui"
        # 新增/修改配置文件的键值
        config.set(section, 'Value1', self.le_version_sw.text())
        config.set(section, 'Value2', self.cb_chipcode.currentText())
        config.set(section, 'Value3', self.le_date_sw.text())
        config.set(section, 'Value4', (self.cb_ext_version.currentText()[2:]+self.cb_ext_version.currentText()[:2]))
        config.set(section, 'Value5', self.cb_vendor_code.currentText())
        with open(path_name, 'w') as configfile:
            config.write(configfile)
        self.statusbar.setStyleSheet(
            "* { color: #00CD00;font-size:30px;font-weight:bold;}")
        self.statusbar.showMessage("配置文件修改成功！", 3000)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MyMainWindow()
    sys.exit(app.exec_())
