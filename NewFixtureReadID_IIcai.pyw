# -*- coding: utf-8 -*-
"""
开发者：周梦雄
最后更新日期：2019/9/5
"""
import sys
import os
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QInputDialog,
    QDateEdit,
    QMainWindow,
    QTableWidget,
    QMessageBox,
    QTableWidgetItem,
    QAbstractItemView,
)
from Ui_NewFixtureReadID import *
from PyQt5.QtCore import QDateTime
import sqlite3
from openpyxl import Workbook
import configparser


class MyMainWindow(QMainWindow, Ui_STA_database_query):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.cb_prod_type.setCurrentIndex(6)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setColumnCount(3)
        self.tableWidget.setRowCount(99)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.start_datetime.setDateTime(QDateTime.currentDateTime())
        self.textBrowser.append(
            "注意：河北、四川3105 3/5；四川3911 2/28，外部版本2181；3105芯片代码03；3911集中器芯片代码00，STA 01；北京、浙江集中器白名单关闭；"
        )
        self.textBrowser.setStyleSheet("* { color: #0000FF;}")
        self.statusbar.setStyleSheet(
            "* { color: #FF6666;font-size:30px;font-weight:bold;}"
        )
        # 数据库路径
        db_file = r"C:\Users\Dream\Desktop\II采同时写模块ID和逻辑地址\Debug\MyProtocol.db"
        # 创建数据库连接对象
        self.conn = sqlite3.connect(db_file)
        # 创建游标对象
        self.cur = self.conn.cursor()
        self.start_date = self.start_datetime.dateTime().toString("yyyy-MM-dd HH:mm")
        self.sqlstring = r"SELECT ChipIDRead,AssetIDWrite,sTime FROM ResultData where ChipIDRead<>''"
        # 初始化
        self.conf = configparser.ConfigParser()
        # 配置文件的绝对路径
        self.conf_path = r"C:\Users\Dream\Desktop\II采同时写模块ID和逻辑地址\Debug\IniFile\FiterParam.ini"
        # 读取配置文件
        self.conf.read(self.conf_path)
        # 返回section中option的值
        self.NV_configure = "软件版本：%s 芯片代码：%s 版本日期：%s 外部版本：%s 厂商代码：%s" % (
            self.conf.get("ErJiBiDui", "Value1"), self.conf.get("ErJiBiDui", "Value2"), self.conf.get("ErJiBiDui", "Value3"), self.conf.get("ErJiBiDui", "Value4"), self.conf.get("ErJiBiDui", "Value5"))
        self.exit.clicked.connect(self.buttonExit)
        self.query.clicked.connect(self.click_query)
        self.start_datetime.dateTimeChanged.connect(self.on_datetime_changed)
        self.export_id.clicked.connect(self.export_id_to_excel)
        self.btn_configure.clicked.connect(self.NV_query)

        
        self.show()
        QMessageBox.warning(
                self, "警告！", "测试之前，务必点击上位机软件的 ##清除数据## 按钮！", QMessageBox.Ok)

    def buttonExit(self):
        self.conn.commit()
        self.cur.close()
        self.conn.close()
        self.close()

    def on_datetime_changed(self):
        self.start_date = self.start_datetime.dateTime().toString("yyyy-MM-dd HH:mm")

    def click_query(self):
        self.tableWidget.clearContents()  # 每一次查询时清除表格中信息
        # 执行查询（传递开始测试日期时间参数）
        self.cur.execute(self.sqlstring)
        # 自动设置ID倒数5个字符
        result_temp = self.cur.fetchall()
        try:
            self.lineEdit.setText(result_temp[0][0][-5:])
            for k, i in enumerate(result_temp):
                print("----------", i)
                for w, j in enumerate(i):
                    if type(j) != str:
                        newItem = QTableWidgetItem(str(j))
                    else:
                        newItem = QTableWidgetItem(j)
                    # 根据循环标签一次对table中的格子进行设置
                    self.tableWidget.setItem(k, w, newItem)
            self.tableWidget.resizeColumnsToContents()
            self.tableWidget.resizeRowsToContents()
            self.textBrowser.setText("")
            self.textBrowser.append(
                "SELECT ChipID,ModID,TTime FROM ResultData where ChipID<>'';"
            )
            print("find button pressed")
        except Exception:
            QMessageBox.warning(self, '提示：', '查询结果为空！', QMessageBox.Ok)

    def export_id_to_excel(self):
        wo = self.lineEdit_2.text()
        wo1 = wo + "-" + self.cb_prod_type.currentText() + ".xlsx"
        # 工作簿保存路径
        path_name = os.path.join(
            r"C:\Users\Dream\Desktop\ID清单，请手下留情，勿删！！！", wo1)
        # 新建工作簿
        wb = Workbook(path_name)
        ws = wb.create_sheet(wo, 0)
        ws.append(["芯片ID", "模块ID"])
        # 查询结果
        self.cur.execute(self.sqlstring)
        result = self.cur.fetchall()
        result_id = [(r[0], r[1]) for r in result]
        result_unique = []
        for i in result_id:
            if i not in result_unique:
                result_unique.append(i)
        for row in result_unique:
            ws.append(list(row))
        self.statusbar.showMessage(
            "本批测试 %s 个模块，请注意检查是否有漏测！" % len(result_unique))
        if result_unique[0][0][-5:] != self.lineEdit.text().upper():
            self.statusbar.clearMessage()
            QMessageBox.warning(
                self, "警告！", "你的首个ID不正确，请排查原因！", QMessageBox.Ok)
        else:
            wb.save(path_name)
            QMessageBox.information(
                self, "好消息！", "ID对应表已成功导出到excel表格！请核对左下角状态栏信息！", QMessageBox.Ok
            )

    def NV_query(self):
        self.textBrowser.setText("")
        self.textBrowser.append(self.NV_configure)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MyMainWindow()
    sys.exit(app.exec_())
