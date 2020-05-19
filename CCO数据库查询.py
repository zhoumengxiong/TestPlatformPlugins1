# -*- coding: utf-8 -*-
"""
开发者：周梦雄
最后更新日期：2019/8/9
"""
import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QDateEdit, QMainWindow, QTableWidget, QMessageBox, QTableWidgetItem
from Ui_STA_DB_query import *
from PyQt5.QtCore import QDateTime
import pyodbc
from openpyxl import Workbook


class MyMainWindow(QMainWindow, Ui_STA_database_query):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.start_datetime.setDateTime(QDateTime.currentDateTime())
        # 数据库驱动
        # sql_driver = r'DSN=芯片ID;DBQ=C:\PC_PRODCHECK\DATA\Equip_sta.mdb;DefaultDir=C:\PC_PRODCHECK\DATA;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;'
        self.sql_driver =  'DSN=芯片ID;DBQ=C:\PC_PRODCHECK\DATA\Equip_sta.mdb;DefaultDir=C:\PC_PRODCHECK\DATA;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;'
        # 创建数据库连接对象
        self.conn = pyodbc.connect(self.sql_driver)
        # 创建游标对象
        self.cur = self.conn.cursor()
        # 日期、时间拆分
        self.start_date = self.start_datetime.dateTime().toString(
            "yyyy-MM-dd HH:mm:ss").split(' ')[0]
        self.start_time = self.start_datetime.dateTime().toString(
            "yyyy-MM-dd HH:mm:ss").split(' ')[1]
        self.sqlstring = r"""SELECT ucTestRet,stVersion_two,ucChipIdStr,ucModuleIdStr,OperDate,OperTime FROM `C:\PC_PRODCHECK\DATA\Equip_sta.mdb`.`equip_sta` where ucChipIdStr<>'' and OperDate>=? and OperTime>=? order by OperDate,OperTime asc;"""
        self.exit.clicked.connect(self.buttonExit)
        self.query.clicked.connect(self.click_query)
        self.start_datetime.dateTimeChanged.connect(self.on_datetime_changed)
        self.export_id.clicked.connect(self.export_id_to_excel)

        self.show()

    def buttonExit(self):
        self.conn.commit()
        self.cur.close()
        self.conn.close()
        self.close()

    def on_datetime_changed(self):
        self.start_date = self.start_datetime.date().toString("yyyy-MM-dd")
        self.start_time = self.start_datetime.time().toString('HH:mm:ss')

    def click_query(self):
        self.tableWidget.clearContents()  # 每一次查询时清除表格中信息
        # 执行查询（传递开始测试日期时间参数）
        self.cur.execute(self.sqlstring, self.start_date, self.start_time)
        k = 0
        for i in self.cur:
            print("----------", i)
            w = 0
            for j in i:
                # 这里是将非字符串类型转成string类型，方便后面文本设置
                # if type(j) == int:
                if type(j) != str:
                    newItem = QTableWidgetItem(str(j))

                else:
                    newItem = QTableWidgetItem(j)
                # 根据循环标签一次对table中的格子进行设置
                self.tableWidget.setItem(k, w, newItem)
                w += 1
            k += 1
        self.textBrowser.setText("")
        self.textBrowser.append(self.sqlstring)
        print("query button pressed")

    def export_id_to_excel(self):
        wo = self.lineEdit_2.text()
        wo1 = wo + '.xlsx'
        # 工作簿保存路径
        path_name = os.path.join(
            r"C:\Users\Lenovo\Desktop\ID清单，请手下留情，勿删！！！", wo1)
        # 新建工作簿
        wb = Workbook(path_name)
        ws = wb.create_sheet(wo, 0)
        ws.append(['芯片ID', '模块ID'])
        # 查询结果
        self.cur.execute(self.sqlstring, self.start_date, self.start_time)
        result = self.cur.fetchall()
        result_id = [(r[2], r[3]) for r in result]
        result_unique = []
        for i in result_id:
            if i not in result_unique:
                result_unique.append(i)
        for row in result_unique:
            ws.append(list(row))
        # -------------------
        print(result_unique)
        self.statusbar.showMessage("本批测试 %s 个模块，请注意检查是否有漏测！首个ID为：%s" %
                                   (len(result_unique), result_unique[0][0]))
        if result_unique[0][0][-5:] != self.lineEdit.text().upper():
            QMessageBox.warning(
                self, '警告！', '你的首个ID不正确，请排查原因！', QMessageBox.Ok)
        else:
            wb.save(path_name)
            QMessageBox.information(
                self, '好消息！', 'ID对应表已正确导出到excel表格！', QMessageBox.Ok)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MyMainWindow()
    sys.exit(app.exec_())
