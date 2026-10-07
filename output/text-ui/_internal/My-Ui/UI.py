import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                             QFileDialog, QLabel, QMessageBox, QHeaderView)
from PyQt6.QtWidgets import QLineEdit, QComboBox, QRadioButton, QGroupBox, QLabel, QLineEdit, QTextEdit, QCheckBox, QProgressBar, QMessageBox, QHeaderView


import pandas as pd
import openpyxl
import 单盘测试
import 合并文件
import 表三
import 表四

from PyQt6 import uic





def on_pushButton_clicked():
    # 运行单盘测试程序

    # 获取当前已选中的单选按钮
    if radioButton.isChecked():
        型号 = "GTYA"
    elif radioButton_2.isChecked():
        型号 = "GTYA53"


    # 获取comboBox当前选中的项
    cbv = comboBox.currentText()
    comboBoxValue = cbv[:-1]
    光缆芯数 = int(comboBoxValue)
    # 获取lineEdit中的文本
    textEditValue = lineEdit.text()
    页数 = int(textEditValue)

    # 执行外部程序并传递参数
    # 假设您要执行的程序是arm，并且它接受两个命令行参数
    command = 单盘测试.dpcs_zcx(型号,光缆芯数, 页数)
    
    command = self.arm(comboBoxValue, textEditValue,型号,光缆芯数, 页数)
    if command:
        subprocess.run(command, check=True)

    
    
    # QtWidgets.QMessageBox.information(self.pushButton, "提示", "单盘测试完成！", QtWidgets.QMessageBox.Yes)

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

    QMessageBox.critical(self, "错误", "合并完成！", QMessageBox.Yes)
    # QtWidgets.QMessageBox.information(self.pushButton_2, "提示", "合并完成！", QtWidgets.QMessageBox.Yes)
        

# def excel_start(self):
#     # 运行excel_start.py中的excel_start函数
#     合并文件.start_merge_from_ui(self.filePaths)
#     QtWidgets.QMessageBox.information(self.pushButton_2, "提示", "合并完成！", QtWidgets.QMessageBox.Yes)




def get_path(self):
    path1, _ = QFileDialog.getOpenFileName(pushButton_3, "选择文件", "", "Excel Files (*.xlsx)")

    lineEdit_2.setText(path1)
    filePaths.append(path1)

    if filePaths:
    try:
        # 读取 Excel 文件
        df = pd.read_excel(filePaths)
        
        # 设置表格行数和列数
        tableWidget.setRowCount(self.df.shape[0])
        tableWidget.setColumnCount(self.df.shape[1])
        
        # 设置表头
        tableWidget.setHorizontalHeaderLabels(self.df.columns)
        
        # 填充数据
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iloc[row, col]))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # 设置为不可编辑
                tableWidget.setItem(row, col, item)
        
        # 调整列宽以适应内容
        tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        
        # 更新状态标签
        label4.setText(f'已加载文件: {filePaths} | 行数: {df.shape[0]}, 列数: {df.shape[1]}')
        
        # 启用清空按钮
        pushButton_6.setEnabled(True)
        
    except Exception as e:
        QMessageBox.critical(self, '错误', f'无法读取文件:\n{str(e)}')

def clear_table(self):
    tableWidget.setRowCount(0)
    tableWidget.setColumnCount(0)
    df = None
    label4.setText('表格已清空')
    pushButton_6.setEnabled(False)

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



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = uic.loadUi('./测试资料窗体.ui')
    filePaths = []

    labe1: QLabel = ui.label        # 获取控件的引用
    lebel2: QLabel = ui.label_2
    lebel3: QLabel = ui.label_3
    label4: QLabel = ui.label_4
    pushButton: QPushButton = ui.pushButton
    pushButton_2: QPushButton = ui.pushButton_2
    pushButton_3: QPushButton = ui.pushButton_3
    pushButton_4: QPushButton = ui.pushButton_4
    pushButton_5: QPushButton = ui.pushButton_5
    pushButton_6: QPushButton = ui.pushButton_6
    lineEdit: QLineEdit = ui.lineEdit
    lineEdit_2: QLineEdit = ui.lineEdit_2
    radioButton: QRadioButton = ui.radioButton
    radioButton_2: QRadioButton = ui.radioButton_2
    comboBox: QComboBox = ui.comboBox
    progressBar: QProgressBar = ui.progressBar
    tableWidget: QTableWidget = ui.tableWidget
    listWidget: QListWidget = ui.listWidget
    # listView: QListView = ui.listView


    pushButton.clicked.connect(lambda: on_pushButton_clicked)   # 开始创建
    pushButton_2.clicked.connect(get_path)              # 选择路径
    pushButton_3.clicked.connect(lambda: df_table_1)            # 表三制作 
    pushButton_4.clicked.connect(lambda: glcs_b3)           # 表四制作
    pushButton_5.clicked.connect(lambda: start_merge)       # 合并表格
    comboBox.addItems(["4D","8D","12D","24D","48D","96D","144D"])
    


    ui.show()

    sys.exit(app.exec())
