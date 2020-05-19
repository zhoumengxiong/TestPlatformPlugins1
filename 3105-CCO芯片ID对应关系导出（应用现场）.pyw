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
from Ui_CCO_DB_query import *
from PyQt5.QtCore import QDateTime
import pyodbc
from openpyxl import Workbook


class MyMainWindow(QMainWindow, Ui_CCO_database_query):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        # self.cb_prod_type_3105.setCurrentIndex(2)
        self.tableWidget_3105.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget_3105.setColumnCount(4)
        self.tableWidget_3105.setRowCount(99)
        self.tableWidget_3105.resizeColumnsToContents()
        self.tableWidget_3105.resizeRowsToContents()
        self.value_start_datetime_3105.setDateTime(QDateTime.currentDateTime())
        self.textBrowser_3105.append(
            "注意：3105芯片代码03；3911集中器芯片代码00，STA 01；北京、浙江集中器白名单关闭；"
        )
        self.textBrowser_3105.setStyleSheet("* { color: #0000FF;}")
        self.statusbar.setStyleSheet(
            "* { color: #FF6666;fontfont-size:30px;font-weight:bold;}"
        )
        # 数据库驱动
        # sql_driver_3105 = r'DSN=芯片ID;DBQ=C:\PC_PRODCHECK\DATA\Equip_sta.mdb;DefaultDir=C:\PC_PRODCHECK\DATA;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;'
        self.sql_driver_3105 = r"DRIVER=Microsoft Access Driver (*.mdb, *.accdb);DBQ=C:/HiStudio-工装/3105集中器通信模块出厂检测/data/ndm/NoStaEquip/NoStaCCOCheckDB.mdb"
        # 创建数据库连接对象
        self.conn_3105 = pyodbc.connect(self.sql_driver_3105)
        # 创建游标对象
        self.cur_3105 = self.conn_3105.cursor()
        self.start_date_3105 = self.value_start_datetime_3105.dateTime().toString("yyyy-MM-dd HH:mm")
        self.sqlstring_3105 = r"SELECT 总体测试结果,芯片ID值,模块ID值,日期 FROM NoStaTableCCOCheck where 芯片ID值<>'' and 日期>=? order by 日期 asc;"
        self.btn_exit_3105.clicked.connect(self.buttonExit_3105)
        self.btn_id_query_3105.clicked.connect(self.click_query_3105)
        self.value_start_datetime_3105.dateTimeChanged.connect(self.on_datetime_changed_3105)
        self.export_id.clicked.connect(self.export_id_to_excel)

        self.show()

    def buttonExit_3105(self):
        self.conn_3105.commit()
        self.cur_3105.close()
        self.conn_3105.close()
        self.close()

    def on_datetime_changed_3105(self):
        self.start_date_3105 = self.value_start_datetime_3105.dateTime().toString("yyyy-MM-dd HH:mm")

    def click_query_3105(self):
        self.tableWidget_3105.clearContents()  # 每一次查询时清除表格中信息
        # 执行查询（传递开始测试日期时间参数）
        self.cur_3105.execute(self.sqlstring_3105, self.start_date_3105)
        # 自动设置ID倒数5个字符
        result_temp=self.cur_3105.fetchall()
        try:
            self.value_id_3105.setText(result_temp[0][1][-5:])
            for k, i in enumerate(result_temp):
                print("----------", i)
                for w, j in enumerate(i):
                    if type(j) != str:
                        newItem = QTableWidgetItem(str(j))
                    else:
                        newItem = QTableWidgetItem(j)
                    # 根据循环标签一次对table中的格子进行设置
                    self.tableWidget_3105.setItem(k, w, newItem)
            self.tableWidget_3105.resizeColumnsToContents()
            self.tableWidget_3105.resizeRowsToContents()
            self.textBrowser_3105.setText("")
            self.textBrowser_3105.append(
                "SELECT 总体测试结果,芯片ID值,模块ID值,日期 FROM NoStaTableCCOCheck where 芯片ID值<>'' and 日期>=%r order by 日期 asc;"
                % (self.start_date_3105)
            )
            print("find button pressed")
        except Exception:
            QMessageBox.warning(self, '提示：', '查询结果为空！', QMessageBox.Ok)

    def export_id_to_excel(self):
        wo = self.value_order_3105.text()
        wo1 = wo + "-" + self.cb_prod_type_3105.currentText() + ".xlsx"
        # 工作簿保存路径
        path_name = os.path.join(
            r"C:\Users\Lenovo\Desktop\ID清单，请手下留情，勿删！！！", wo1)
        # 新建工作簿
        wb = Workbook(path_name)
        ws = wb.create_sheet(wo, 0)
        ws.append(["芯片ID", "模块ID"])
        # 查询结果
        self.cur_3105.execute(self.sqlstring_3105, self.start_date_3105)
        result = self.cur_3105.fetchall()
        result_id = [(r[1], r[2]) for r in result]
        result_unique = []
        for i in result_id:
            if i not in result_unique:
                result_unique.append(i)
        for row in result_unique:
            ws.append(list(row))
        self.statusbar.showMessage(
            "本批测试 %s 个模块，请注意检查是否有漏测！" % len(result_unique))
        if result_unique[0][0][-5:] != self.value_id_3105.text().upper():
            self.statusbar.clearMessage()
            QMessageBox.warning(
                self, "警告！", "你的首个ID不正确，请排查原因！", QMessageBox.Ok)
        else:
            wb.save(path_name)
            QMessageBox.information(
                self, "好消息！", "ID对应表已成功导出到excel表格！请核对左下角状态栏信息！", QMessageBox.Ok
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MyMainWindow()
    sys.exit(app.exec_())
