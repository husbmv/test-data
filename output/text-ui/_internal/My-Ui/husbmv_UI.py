import openpyxl
import 表三
import 表四
import 合并文件
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd

import sys
from PyQt5 import*
from PyQt5 import QtCore,  QtWidgets
from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem
import 单盘测试

from PyQt5 import QtCore, QtGui, QtWidgets





class Ui_Form(object):
    def __init__(self):
        super().__init__()
        # self.initUI()
        self.filePaths = []


        
        

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(807, 659)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(40, 40, 551, 121))
        self.groupBox.setObjectName("groupBox")
        self.groupBox_2 = QtWidgets.QGroupBox(self.groupBox)
        self.groupBox_2.setGeometry(QtCore.QRect(20, 17, 121, 91))
        self.groupBox_2.setObjectName("groupBox_2")
        self.radioButton = QtWidgets.QRadioButton(self.groupBox_2)
        self.radioButton.setGeometry(QtCore.QRect(20, 30, 89, 16))
        self.radioButton.setChecked(True)
        self.radioButton.setObjectName("radioButton")
        self.radioButton_2 = QtWidgets.QRadioButton(self.groupBox_2)
        self.radioButton_2.setGeometry(QtCore.QRect(20, 60, 89, 16))
        self.radioButton_2.setChecked(False)
        self.radioButton_2.setAutoRepeat(False)
        self.radioButton_2.setObjectName("radioButton_2")
        self.pushButton = QtWidgets.QPushButton(self.groupBox)
        self.pushButton.setGeometry(QtCore.QRect(450, 27, 75, 31))
        self.pushButton.setObjectName("pushButton")
        self.groupBox_3 = QtWidgets.QGroupBox(self.groupBox)
        self.groupBox_3.setGeometry(QtCore.QRect(160, 17, 120, 91))
        self.groupBox_3.setObjectName("groupBox_3")
        self.comboBox = QtWidgets.QComboBox(self.groupBox_3)
        self.comboBox.setGeometry(QtCore.QRect(20, 31, 81, 31))
        self.comboBox.setObjectName("comboBox")
        self.groupBox_4 = QtWidgets.QGroupBox(self.groupBox)
        self.groupBox_4.setGeometry(QtCore.QRect(300, 17, 121, 91))
        self.groupBox_4.setObjectName("groupBox_4")
        self.textEdit = QtWidgets.QTextEdit(self.groupBox_4)
        self.textEdit.setGeometry(QtCore.QRect(20, 31, 81, 31))
        self.textEdit.setObjectName("textEdit")
        self.pushButton_2 = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_2.setGeometry(QtCore.QRect(450, 75, 75, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(90, 150, 81, 41))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(230, 160, 71, 16))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(380, 160, 71, 16))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(610, 70, 54, 12))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(610, 130, 54, 12))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(40, 180, 81, 31))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.entry1_text = QtWidgets.QLineEdit(Form)
        self.entry1_text.setGeometry(QtCore.QRect(120, 180, 291, 31))
        self.entry1_text.setObjectName("engry1_text")
        self.pushButton_3 = QtWidgets.QPushButton(Form)
        self.pushButton_3.setGeometry(QtCore.QRect(420, 180, 101, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(Form)
        self.pushButton_4.setGeometry(QtCore.QRect(530, 180, 101, 31))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QtWidgets.QPushButton(Form)
        self.pushButton_5.setGeometry(QtCore.QRect(640, 180, 101, 31))
        self.pushButton_5.setObjectName("pushButton_5")
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setGeometry(QtCore.QRect(40, 250, 721, 371))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        # self.tableWidget.setColumnCount(5)
        # row1 = df_rows + 10
        # self.tableWidget.setRowCount(row1)
        # 将DataFrame的内容写入QTableWidget
        # self.tableWidget.setRowCount(len(pdf.index))
        # self.tableWidget.setColumnCount(len(pdf.columns))
        
        # # 设置列名
        # self.tableWidget.setHorizontalHeaderLabels(pdf.columns)
        
        # # 遍历DataFrame的每一行和每一列
        # for row_index, row in pdf.iterrows():
        #     for col_index, item in enumerate(row):
        #         table_item = QTableWidgetItem(str(item))
        #         self.tableWidget.setItem(row_index, col_index, table_item)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

        self.entry1_text.setReadOnly(True)

        self.pushButton.clicked.connect(self.on_pushButton_clicked)
        self.pushButton_2.clicked.connect(self.start_merge)
        self.pushButton_3.clicked.connect(self.get_path)
        self.pushButton_4.clicked.connect(self.df_table_1)
        self.pushButton_5.clicked.connect(self.glcs_b3)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "制作测试表格"))
        self.groupBox.setTitle(_translate("Form", "单盘测试"))
        self.groupBox_2.setTitle(_translate("Form", "光缆类型："))
        self.radioButton.setText(_translate("Form", "普通光缆"))
        self.radioButton_2.setText(_translate("Form", "防鼠光缆"))
        self.pushButton.setText(_translate("Form", "开始创建"))
        self.groupBox_3.setTitle(_translate("Form", "光缆型号："))
        self.groupBox_4.setTitle(_translate("Form", "打印数量"))
        self.pushButton_2.setText(_translate("Form", "合并表格"))
        self.label.setText(_translate("Form", "TextLabel1"))
        self.label_2.setText(_translate("Form", "TextLabel2"))
        self.label_3.setText(_translate("Form", "TextLabel3"))
        self.label_4.setText(_translate("Form", "TextLabel"))
        self.label_5.setText(_translate("Form", "TextLabel"))
        self.label_6.setText(_translate("Form", "选择文件："))
        self.pushButton_3.setText(_translate("Form", "选择路径"))
        self.pushButton_4.setText(_translate("Form", "开始制作表三"))
        self.pushButton_5.setText(_translate("Form", "开始制作表四"))
        self.comboBox.addItems(["4D","8D","12D","24D","48D","96D","144D"])
        self.textEdit.setPlainText("")
        


    def on_pushButton_clicked(self):
        # 运行单盘测试程序

        # 获取当前已选中的单选按钮
        if self.radioButton.isChecked():
            型号 = "GTYA"
        elif self.radioButton_2.isChecked():
            型号 = "GTYA53"


        # 获取comboBox当前选中的项
        cbv = self.comboBox.currentText()
        comboBoxValue = cbv[:-1]
        光缆芯数 = int(comboBoxValue)
        # 获取textEdit中的文本
        textEditValue = self.textEdit.toPlainText()
        页数 = int(textEditValue)

        # 执行外部程序并传递参数
        # 假设您要执行的程序是arm，并且它接受两个命令行参数
        command = 单盘测试.dpcs_zcx(型号,光缆芯数, 页数)
        
        command = self.arm(comboBoxValue, textEditValue,型号,光缆芯数, 页数)
        if command:
            subprocess.run(command, check=True)
        
        QtWidgets.QMessageBox.information(self.pushButton, "提示", "单盘测试完成！", QtWidgets.QMessageBox.Yes)

    def arm(self, comboBoxValue, textEditValue, 型号, 光缆芯数, 页数):
        print("arm", comboBoxValue, textEditValue, 型号, 光缆芯数, 页数)


    # def update_progress(self, status, current, total, message=None):
    # """更新UI进度的回调函数"""
    # if message:
    #     # 将日志添加到您的日志变量中
    #     self.log_text += message + "\n"
        
    #     # 更新UI控件（例如文本框或标签）
    #     self.log_text_widget.insert(tk.END, message + "\n")
    #     self.log_text_widget.see(tk.END)
        
    #     # 更新进度条
    #     if status == 'progress':
    #         self.progress_bar['value'] = current
    #         self.progress_bar['maximum'] = total




    def start_merge(self):
    # """开始合并的按钮点击事件"""

        合并文件.start_merge_from_ui()
        QtWidgets.QMessageBox.information(self.pushButton_2, "提示", "合并完成！", QtWidgets.QMessageBox.Yes)
            

    # def excel_start(self):
    #     # 运行excel_start.py中的excel_start函数
    #     合并文件.start_merge_from_ui(self.filePaths)
    #     QtWidgets.QMessageBox.information(self.pushButton_2, "提示", "合并完成！", QtWidgets.QMessageBox.Yes)




    def get_path(self):
        path1, _ = QtWidgets.QFileDialog.getOpenFileName(self.pushButton_3, "选择文件", "", "Excel Files (*.xlsx)")
        self.entry1_text.setText(path1)
        self.filePaths.append(path1)
        # print(self.filePaths[0])

        

    # def xzwj(self):
    #     # 读取excel文件
    #     round_name = self.filePaths[0]
    #     print(type(round_name))
    #     df = pd.read_excel(round_name, sheet_name='数据')
    #     # 去除'接头数'列为空的行
    #     df = df.iloc[:,[0,2,3,4]]
    
        
    def df_table_1(self):
        # 运行表三.py中的表三函数
        

        # 读取excel文件
        round_name = self.filePaths[0]
        if round_name == "":
            QtWidgets.QMessageBox.warning(self.pushButton_4, "警告", "请先选择文件！", QtWidgets.QMessageBox.Yes)
        
        
        print(type(round_name))
        df = pd.read_excel(round_name, sheet_name='数据')
        # 去除不必要的列
        df = df.iloc[:,[1,2,3,4]]
        # 去除'接头数'列为空的行
        df = df.dropna(subset=['接头数'])


        print(df)

        # 计算df列表的总行数
        total_rows = df.shape[0]

        # 光缆芯数 = 12
        # 页数 = 2
        # 光缆总长 = 1000
        # 接头数 = 3
        # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"


        # 创建一个新的工作簿
        wb = openpyxl.Workbook()

        # 获取默认的活动工作表
        sheet1 = wb.active


        # 设置 A1 至 J1 列的宽度为 6.75
        for col in sheet1.iter_cols(min_col=1, max_col=10):
            for cell in col:
                sheet1.column_dimensions[cell.column_letter].width = 8.4


        for i in range(total_rows):
            # 输出第一行第一列和第一行第三列的数据
            地址 = df.iloc[i, 0]
            光缆芯数 = int(df.iloc[i, 1][:-1])
            光缆总长 = df.iloc[i, 2]
            接头数 = int(df.iloc[i, 3])

            表三.表三(sheet1, 光缆芯数, 光缆总长, 接头数, 地址)


        # 保存工作簿
        名称 = "表三接头损耗表.xlsx"
        wb.save(名称)

        QtWidgets.QMessageBox.information(self.pushButton_4, "提示", "表三制作完成！", QtWidgets.QMessageBox.Yes)

    def glcs_b3(self):
        # 读取excel文件
        round_name = self.filePaths[0]

        if round_name == "":
            QtWidgets.QMessageBox.warning(self.pushButton_5, "警告", "请先选择文件！", QtWidgets.QMessageBox.Yes)
        


        print(type(round_name))
        df = pd.read_excel(round_name, sheet_name='数据')

        df = df.iloc[:,[1,2,3,4]]
        # 计算df列表的总行数
        total_rows1 = df.shape[0]

        
        

        print(df)
        print(total_rows1)

        # 创建一个新的工作簿
        wb = openpyxl.Workbook()

        # 获取默认的活动工作表
        sheet1 = wb.active


        # 设置 A1 至 J1 列的宽度为 6.63
        for col in sheet1.iter_cols(min_col=1, max_col=12):
            for cell in col:
                sheet1.column_dimensions[cell.column_letter].width = 7.28


        for i in range(total_rows1):
            # 输出第一行第一列和第一行第三列的数据
            地址 = df.iloc[i, 0]
            光缆芯数 = int(df.iloc[i, 1][:-1])

            formatted_length = df.iloc[i, 2] / 1000
            光缆总长 = f"{formatted_length:.4f}"

            
            表四.glcs_b3(sheet1, 光缆芯数, 光缆总长, 地址)


        # 保存工作簿
        名称 = "表四损耗表.xlsx"
        wb.save(名称)

        QtWidgets.QMessageBox.information(self.pushButton_5, "提示", "表四制作完成！", QtWidgets.QMessageBox.Yes)





if __name__ == "__main__":  
    app = QtWidgets.QApplication(sys.argv) 
    form = QtWidgets.QWidget()
    w = Ui_Form()
    w.setupUi(form)
    form.show()
    sys.exit(app.exec_())


