import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                             QFileDialog, QLabel, QMessageBox, QHeaderView,QListWidget)
from PyQt6.QtWidgets import (QLineEdit, QComboBox, QRadioButton, QGroupBox, QLabel, QLineEdit, 
                             QTextEdit, QCheckBox, QProgressBar, QMessageBox, QHeaderView)
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QSizePolicy
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
from openpyxl import load_workbook, Workbook,utils
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
import 单盘测试
import 段落清单
import 表三
import 表四
import os
import excel合并
from openpyxl.styles import Border, Side, Alignment, Font
import math
from 测试资料窗体_ui import Ui_Form

from PyQt6 import uic

class MainUI(QMainWindow):
        

    filePaths = ()
    row_data = []



    def log(self, message):
        self.textEdit.append(message)  # 简单追加文本到文本编辑器


    def __init__(self):
        super().__init__() 
        # 加载ui文件
        # self.ui = uic.loadUi('测试资料窗体.ui', self)

        # 打包需要将ui文件转成py文件后，再导入使用以下代码
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        # 初始化变量
        self.filePaths = []
        self.file_paths = []
        self.df = None
        
        # 初始化UI控件引用
        self.init_ui_components()
        
        # 设置信号槽连接
        self.setup_connections()
        
        # 初始化comboBox
        self.comboBox.addItems(["4D","8D","12D","24D","48D","96D","144D"])

        # self.listwidget.setWordWrap(True)        # 设置listwidget的文本换行

        self.radioButton.setChecked(True)        # 默认选中普通光缆单选按钮
        self.radioButton_2.setChecked(False)        


        self.textEdit.setReadOnly(True)         # 设置textEdit为只读
        # self.textEdit.setFont(QFont("Consolas", 9))  # 设置textEdit的字体
        
        self.progressBar.setRange(0, 100)       # 设置进度条范围为0-100
        self.progressBar.setValue(0)            # 初始化进度条为0
        # 设置表格为整行选择模式
        self.tableWidget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableWidget_3.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        # 设置选择模式为单行选择（可根据需要改为多选）
        self.tableWidget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.tableWidget_3.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        # 连接信号：当选择改变时触发槽函数
        self.tableWidget.itemSelectionChanged.connect(self.on_selection_changed)    # 连接信号：当选择改变时触发槽函数
        self.tableWidget_3.itemSelectionChanged.connect(self.on_selection_changed_2)    # 连接信号：当选择改变时触发槽函数


        # 设置窗口的最小尺寸（可选，但建议设置以避免窗口过小）
        self.setMinimumSize(False, False)
        
        # 设置窗口尺寸策略为可扩展
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # 设置表格的尺寸策略
        self.tableWidget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.tableWidget_3.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    
        
        # 设置表格列宽模式，允许手动调整和自动拉伸
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)  # 允许手动调整
        # 或者使用 Stretch 模式让列自动填充宽度
        # header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        # 设置垂直头部的调整模式
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

    
    def init_ui_components(self):
        """初始化UI控件引用"""
        # 测试资料部分
        self.label: QLabel= self.ui.label
        self.label_2: QLabel= self.ui.label_2
        # self.label_3: QLabel = self.ui.label_3
        self.label_4: QLabel = self.ui.label_4
        self.pushButton: QPushButton= self.ui.pushButton
        self.pushButton_2: QPushButton= self.ui.pushButton_2
        self.pushButton_3: QPushButton= self.ui.pushButton_3
        self.pushButton_4: QPushButton= self.ui.pushButton_4
        self.pushButton_5: QPushButton= self.ui.pushButton_5
        self.pushButton_6: QPushButton= self.ui.pushButton_6
        self.lineEdit: QLineEdit = self.ui.lineEdit
        # self.lineEdit_2: QLineEdit= self.ui.lineEdit_2
        self.radioButton: QRadioButton= self.ui.radioButton
        self.radioButton_2: QRadioButton= self.ui.radioButton_2
        self.comboBox: QComboBox= self.ui.comboBox
        self.progressBar: QProgressBar= self.ui.progressBar
        self.tableWidget: QTableWidget= self.ui.tableWidget
        # self.listwidget: QListWidget= self.ui.listWidget
        self.textEdit: QTextEdit = self.ui.textEdit

        # 竣工资料部分
        self.pushButton_13: QPushButton= self.ui.pushButton_13
        self.pushButton_14: QPushButton= self.ui.pushButton_14
        self.pushButton_15: QPushButton= self.ui.pushButton_15
        self.pushButton_16: QPushButton= self.ui.pushButton_16
        self.pushButton_17: QPushButton= self.ui.pushButton_17
        self.pushButton_18: QPushButton= self.ui.pushButton_18
        self.pushButton_19: QPushButton= self.ui.pushButton_19
        self.pushButton_20: QPushButton= self.ui.pushButton_20
        self.lineEdit_number_dd: QLineEdit= self.ui.lineEdit_3
        self.lineEdit_number_xm: QLineEdit= self.ui.lineEdit_4
        self.lineEdit_number_ht: QLineEdit= self.ui.lineEdit_5
        self.lineEdit_name_dd: QLineEdit= self.ui.lineEdit_6
        self.lineEdit_number_jsd: QLineEdit= self.ui.lineEdit_7
        self.lineEdit_name_sgdw: QLineEdit= self.ui.lineEdit_8
        self.lineEdit_csqk: QLineEdit= self.ui.lineEdit_9
        self.lineEdit_money_sd: QLineEdit= self.ui.lineEdit_10
        self.lineEdit_money_cb: QLineEdit= self.ui.lineEdit_11
        self.lineEdit_money_hs: QLineEdit= self.ui.lineEdit_12
        self.lineEdit_money_zy: QLineEdit= self.ui.lineEdit_13
        self.lineEdit_money_wk: QLineEdit= self.ui.lineEdit_14
        self.tableWidget_3: QTableWidget= self.ui.tableWidget_3

    
    def setup_connections(self):
        """设置信号槽连接"""
        # 测试资料部分
        # 先断开所有已有的连接，确保只有一次连接
        self.pushButton.clicked.disconnect()
        self.pushButton.clicked.connect(self.on_pushButton_clicked)  # 单盘测试
        # self.pushButton_2.clicked.connect(self.get_path)         # 选择文件
        self.pushButton_2.clicked.connect(self.add_files)       # 添加文件
        self.pushButton_3.clicked.connect(self.df_table_1)       # 制作表三
        self.pushButton_4.clicked.connect(self.glcs_b4)         # 制作表四
        self.pushButton_5.clicked.connect(self.start_merge)     # 合并文件
        self.pushButton_6.clicked.connect(self.clear_table)     # 清空表格

        # 竣工资料部分
        self.pushButton_13.clicked.connect(self.add_files_2)       # 添加文件
        self.pushButton_14.clicked.connect(self.search_table)       # 查找表格
        self.pushButton_17.clicked.connect(self.on_start_dlqd)       # 制作段落清单
        self.pushButton_20.clicked.connect(self.clear_nr)           # 清空内容



    def update_status(self, message):
        """更新状态信息"""
        self.status_label.setText(message)    

    def on_selection_changed(self):
        """当选择发生变化时调用，输出当前选中行的内容"""
        # 获取当前选中的所有范围
        selected_ranges = self.tableWidget.selectedRanges()
        
        # 如果没有选中的行，则退出
        if not selected_ranges:
            print("没有行被选中")
            return

        # 遍历所有选中的范围（支持多选时）
        for selection_range in selected_ranges:  # 修改变量名避免冲突
            top_row = selection_range.topRow()
            bottom_row = selection_range.bottomRow()
            # 遍历该范围内的每一行
            for row in range(top_row, bottom_row + 1):  # 现在可以正常使用range函数
                row_data = []
                # 遍历该行的每一列
                for col in range(self.tableWidget.columnCount()):
                    item = self.tableWidget.item(row, col)
                    # 如果单元格有内容，则获取文本，否则为空字符串
                    cell_text = item.text() if item else ''
                    row_data.append(cell_text)
                # 输出当前行的内容
                print(f"行 {row} 的内容: {row_data[0]}")
                self.textEdit.append(f"行 {row} 的内容: {row_data[0]}")

    # 当选择发生变化时调用，输出当前选中行的内容
    def on_selection_changed_2(self):
        """当选择发生变化时调用，输出当前选中行的内容"""
        # 获取当前选中的所有范围
        selected_ranges = self.tableWidget_3.selectedRanges()

        def calculate_price_decimal(cost, tax_rate):
            """使用Decimal进行精确的货币计算"""
            cost_dec = Decimal(str(cost))
            tax_rate_dec = Decimal(str(tax_rate))
            price_with_tax = cost_dec * (tax_rate_dec)
            return price_with_tax.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        
        # 如果没有选中的行，则退出
        if not selected_ranges:
            print("没有行被选中")
            return

        # 遍历所有选中的范围（支持多选时）
        for selection_range in selected_ranges:  # 修改变量名避免冲突
            top_row = selection_range.topRow()  # 修改变量名
            bottom_row = selection_range.bottomRow()    
            # 遍历该范围内的每一行
            for row in range(top_row, bottom_row + 1):  # 现在可以正常使用range函数
                row_data = []
                self.row_data = row_data
                # 遍历该行的每一列
                for col in range(self.tableWidget_3.columnCount()):
                    item = self.tableWidget_3.item(row, col)
                    # 如果单元格有内容，则获取文本，否则为空字符串
                    cell_text = item.text() if item else ''
                    row_data.append(cell_text)
                # 输出当前行的内容
                print(f"行 {row} 的内容: {row_data[0]}")
                self.textEdit.append(f"行 {row} 的内容: {row_data[0]}")

                成本 = row_data[14]
                含税价 = calculate_price_decimal(成本,1.09)     # 成本*1.09
                终验款 = calculate_price_decimal(含税价,0.90)   # 含税价*90%
                尾款 = 含税价 - 终验款
                self.lineEdit_number_dd.setText(row_data[1])    # 订单号
                self.lineEdit_number_xm.setText(row_data[4])    # 项目编号
                self.lineEdit_number_ht.setText(row_data[6])    # 合同号
                self.lineEdit_name_dd.setText(row_data[2])      # 订单名称
                self.lineEdit_number_jsd.setText(row_data[3])   # 结算单号
                self.lineEdit_name_sgdw.setText(row_data[7])    # 施工单位
                self.lineEdit_csqk.setText(row_data[12])        # 抽审情况
                self.lineEdit_money_sd.setText(row_data[13])    # 审定金额
                self.lineEdit_money_cb.setText(str(成本))           # 成本
                self.lineEdit_money_hs.setText(str(含税价))    # 含税金额
                self.lineEdit_money_zy.setText(str(终验款))    # 终验款
                self.lineEdit_money_wk.setText(str(尾款))    # 尾款
                if self.lineEdit_name_dd.text() != "":
                    self.pushButton_20.setEnabled(True)
        return row_data     # 返回当前选中行的内容
                
                    
    # 段落清单
    def on_start_dlqd(self):
        
        """开始段落清单"""
        # 创建一个新的工作簿
        data_list = self.on_selection_changed_2()
        filter_list = self.filePaths[0]
        段落清单.zxcx_one(filter_list,data_list)
        QMessageBox.information(self, "提示", "创建段落清单成功", QMessageBox.StandardButton.Yes)

    # 单盘测试
    def on_pushButton_clicked(self):

        if self.radioButton.isChecked():
            型号 = "GTYA"
        elif self.radioButton_2.isChecked():
            型号 = "GTYA53"
        else:
            型号 = "GTYA"  # 默认值

        # 获取comboBox当前选中的项
        cbv = self.comboBox.currentText()
        comboBoxValue = cbv[:-1]
        光缆芯数 = int(comboBoxValue)

        # 获取lineEdit中的文本
        textEditValue = self.lineEdit.text()
        
        # 检查输入是否为空
        if textEditValue == "":
            QMessageBox.critical(self, "错误", "请输入需要打印的页数", QMessageBox.StandardButton.Yes)
            return
        
        # 检查输入是否为有效数字
        try:
            页数 = int(textEditValue)
        except ValueError:
            QMessageBox.critical(self, "错误", "请输入有效的数字", QMessageBox.StandardButton.Yes)
            return
        
        # 检查输入是否大于0
        if 页数 <= 0:
            QMessageBox.critical(self, "错误", "请输入大于0的数字", QMessageBox.StandardButton.Yes)
            return
        
        # 执行单盘测试
        try:
            单盘测试.dpcs_zcx(型号, 光缆芯数, 页数)
            QMessageBox.information(self, "提示", "单盘测试完成！", QMessageBox.StandardButton.Yes)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"单盘测试执行失败：{str(e)}", QMessageBox.StandardButton.Yes)


        
        # 获取lineEdit中的文本
        # textEditValue = self.lineEdit.text()
        # if textEditValue == "":
        #     QMessageBox.critical(self, "错误", "请输入需要打印的页数", QMessageBox.StandardButton.Yes)
        # else:
        #     页数 = int(textEditValue)
        #     if 页数 > 0:
        #         # 执行单盘测试
        #         单盘测试.dpcs_zcx(型号, 光缆芯数, 页数)
        #         QMessageBox.information(self, "提示", "单盘测试完成！", QMessageBox.StandardButton.Yes)
        #     else:
        #         QMessageBox.critical(self, "错误", "请输入大于0的数字", QMessageBox.StandardButton.Yes)

    # 添加文件路径，单盘测试
    def add_files(self):
        """添加Excel文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if len(files) == 1:
            path1 = files[0]

            if path1:  # 只有当用户选择了文件才继续
                self.label_4.setText(path1)
                self.filePaths.append(path1)
                
                try:
                    # 读取 Excel 文件
                    self.df = pd.read_excel(path1)
                    
                    # 设置表格行数和列数
                    self.tableWidget.setRowCount(self.df.shape[0])
                    self.tableWidget.setColumnCount(self.df.shape[1])
                    
                    # 设置表头
                    self.tableWidget.setHorizontalHeaderLabels(self.df.columns)
                    
                    # 填充数据
                    for row in range(self.df.shape[0]):
                        for col in range(self.df.shape[1]):
                            item = QTableWidgetItem(str(self.df.iloc[row, col]))
                            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # 设置为不可编辑
                            self.tableWidget.setItem(row, col, item)
                    
                    # 调整列宽以适应内容
                    self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

                    
                    
                    # 更新状态标签
                    # self.listwidget.addItem(f'已加载文件: {path1} | 行数: {self.df.shape[0]}, 列数: {self.df.shape[1]}')
                    
                    # 启用清空按钮
                    self.pushButton_6.setEnabled(True)
                    
                except Exception as e:
                    QMessageBox.critical(self, '错误', f'无法读取文件:\n{str(e)}')


            self.log(f"已选择 {os.path.basename(files[0])} ")
            self.log(f'已加载{self.df.shape[0]}行数据')
            self.log(f"总共已选择 {len(files)} 个文件")
          

        else:
            if files:
                for file in files:
                    if file not in self.file_paths:
                        self.file_paths.append(file)
                        # self.listwidget.addItem(os.path.basename(file))
                        self.log(f"已选择 {os.path.basename(file)} ")
            self.log(f"已选择 {os.path.basename(file)} ")
            self.log(f"总共已选择 {len(files)} 个文件")

    # 添加文件路径，竣工资料
    def add_files_2(self):
        """选择文件路径"""
        path1, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel Files (*.xlsx)")
        
        if path1:  # 只有当用户选择了文件才继续
            # self.label_4.setText(path1)
            # self.filePaths.append(path1)
            
            try:
                # 读取 Excel 文件
                self.df = pd.read_excel(path1, sheet_name='数据库')
                
                # 设置表格行数和列数
                self.tableWidget_3.setRowCount(self.df.shape[0])
                self.tableWidget_3.setColumnCount(self.df.shape[1])
                
                # 设置表头
                self.tableWidget_3.setHorizontalHeaderLabels(self.df.columns)
                
                # 填充数据
                for row in range(self.df.shape[0]):
                    for col in range(self.df.shape[1]):
                        item = QTableWidgetItem(str(self.df.iloc[row, col]))
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # 设置为不可编辑
                        self.tableWidget_3.setItem(row, col, item)
                
                # 调整列宽以适应内容
                self.tableWidget_3.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
                
                # 更新状态标签
                # self.listwidget.addItem(f'已加载文件: {path1} | 行数: {self.df.shape[0]}, 列数: {self.df.shape[1]}')
                
                # 启用清空按钮
                self.pushButton_20.setEnabled(True)
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'无法读取文件:\n{str(e)}')

    # 清空表格
    def clear_table(self):
        """清空表格"""
        try:
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            excel合并.clear_logs()
            self.textEdit.clear()
            # self.log('内容已清空')
            self.df = None
            self.label_4.clear()
            self.pushButton_6.setEnabled(False)
            # QMessageBox.critical(self, "内容已清空", QMessageBox.StandardButton.Yes)
            QMessageBox.information(self, "提示","内容已清空", QMessageBox.StandardButton.Yes)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"清空表格失败：{str(e)}", QMessageBox.StandardButton.Yes)

    # 清除内容
    def clear_nr(self):
        """清空表格"""
        try:
            self.lineEdit_number_dd.clear()     # 订单号
            self.lineEdit_number_xm.clear()     # 项目号
            self.lineEdit_number_ht.clear()     # 合同号
            self.lineEdit_name_dd.clear()       # 订单名称
            self.lineEdit_number_jsd.clear()    # 结算单号 
            self.lineEdit_name_sgdw.clear()     # 施工单位
            self.lineEdit_csqk.clear()            # 抽审情况
            self.lineEdit_money_sd.clear()      # 审定金额
            self.lineEdit_money_cb.clear()      # 成本
            self.lineEdit_money_hs.clear()      # 含税金额
            self.lineEdit_money_zy.clear()      # 终验款
            self.lineEdit_money_wk.clear()      # 尾款
            self.pushButton_20.setEnabled(False)
            # QMessageBox.critical(self, "内容已清空", QMessageBox.StandardButton.Yes)
            QMessageBox.information(self, "提示","内容已清空", QMessageBox.StandardButton.Yes)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"清空表格失败：{str(e)}", QMessageBox.StandardButton.Yes)


#---------------------------合并文件---------------------------以下是合并文件
    # 在 MainUI 类中添加一个新的回调函数
    def merge_progress_callback(self, status, current, total, percentage, message):
        """合并文件进度回调函数"""
        # 更新进度条
        self.progressBar.setValue(percentage)
        
        # 更新日志显示
        self.log(message)
        
        # 处理特殊情况
        if status == 'complete':
            QMessageBox.information(self, "提示", message, QMessageBox.StandardButton.Yes)
        elif status == 'error':
            QMessageBox.critical(self, "错误", message, QMessageBox.StandardButton.Yes)

    def start_merge(self):
        """开始合并文件"""
        if not self.file_paths:
            QMessageBox.warning(self, "警告", "请先选择要合并的文件！", QMessageBox.StandardButton.Yes)
            return
        
        try:
            # 重置进度条
            self.progressBar.setValue(0)
            
            # 调用合并函数，传入进度回调函数
            result = excel合并.start_merge(
                file_list=self.file_paths,
                output_filename="测试资料.xlsx",
                progress_callback=self.merge_progress_callback
            )
            
            # 根据结果处理完成状态
            if not result['success']:
                QMessageBox.critical(self, "错误", result['message'], QMessageBox.StandardButton.Yes)
            else:
                # 最终进度设为100%
                self.progressBar.setValue(100)
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并失败：{str(e)}", QMessageBox.StandardButton.Yes)
            self.progressBar.setValue(0)

#   ---------------------------合并文件---------------------------以上
    
    # 制作表三
    def df_table_1(self):
        """制作表三"""
        if not self.filePaths:
            QMessageBox.warning(self, "警告", "请先选择文件！", QMessageBox.StandardButton.Yes)
            return
        
        try:
            round_name = self.filePaths[0]
            df = pd.read_excel(round_name, sheet_name='数据')
            # 去除不必要的列
            df = df.iloc[:,[1,2,3,4]]
            # 去除'接头数'列为空的行
            df = df.dropna(subset=['接头数'])
            
            # 计算df列表的总行数
            total_rows = df.shape[0]
            
            # 创建一个新的工作簿
            wb = openpyxl.Workbook()
            sheet1 = wb.active
            
            # 设置列宽
            for col in sheet1.iter_cols(min_col=1, max_col=10):
                for cell in col:
                    sheet1.column_dimensions[cell.column_letter].width = 8.4
            
            for i in range(total_rows):
                地址 = df.iloc[i, 0]
                光缆芯数 = int(df.iloc[i, 1][:-1])
                光缆总长 = df.iloc[i, 2]
                接头数 = int(df.iloc[i, 3])
                
                表三.表三(sheet1, 光缆芯数, 光缆总长, 接头数, 地址)
            
            # 保存工作簿
            名称 = "表三接头损耗表.xlsx"
            wb.save(名称)
            
            QMessageBox.information(self, "提示", "表三制作完成！", QMessageBox.StandardButton.Yes)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"表三制作失败：{str(e)}", QMessageBox.StandardButton.Yes)
    
    # 制作表四
    def glcs_b4(self):
        """制作表四"""
        if not self.filePaths:
            QMessageBox.warning(self, "警告", "请先选择文件！", QMessageBox.StandardButton.Yes)
            return
        
        try:
            round_name = self.filePaths[0]
            df = pd.read_excel(round_name, sheet_name='数据')
            df = df.iloc[:,[1,2,3,4]]
            
            # 计算总行数
            total_rows1 = df.shape[0]
            
            # 创建工作簿
            wb = openpyxl.Workbook()
            sheet1 = wb.active
            
            # 设置列宽
            for col in sheet1.iter_cols(min_col=1, max_col=12):
                for cell in col:
                    sheet1.column_dimensions[cell.column_letter].width = 7.28
            
            for i in range(total_rows1):
                地址 = df.iloc[i, 0]
                光缆芯数 = int(df.iloc[i, 1][:-1])
                
                formatted_length = df.iloc[i, 2] / 1000
                光缆总长 = f"{formatted_length:.4f}"
                
                表四.glcs_b3(sheet1, 光缆芯数, 光缆总长, 地址)
            
            # 保存工作簿
            名称 = "表四损耗表.xlsx"
            wb.save(名称)
            
            QMessageBox.information(self, "提示", "表四制作完成！", QMessageBox.StandardButton.Yes)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"表四制作失败：{str(e)}", QMessageBox.StandardButton.Yes)

    # 查询表格内容
    def search_table(self):
        text = self.lineEdit_number_dd.text().strip()
        if not text:
            QMessageBox.warning(self, "提示", "请输入查询内容!")
            return
        
        self.results = []
        
        # 搜索所有单元格
        for row in range(self.tableWidget_3.rowCount()):
            for col in range(self.tableWidget_3.columnCount()):
                item = self.tableWidget_3.item(row, col)
                if item and text.lower() in item.text().lower():
                    self.results.append((row, col))
        
        # 显示结果
        if self.results:
            self.current_index = 0
            self.highlight_current()
            self.statusBar().showMessage(f"找到 {len(self.results)} 个结果")
        else:
            QMessageBox.information(self, "结果", "未找到匹配内容!")
    
    # 查找下一个匹配项
    def find_next(self):
        if self.results:
            self.current_index = (self.current_index + 1) % len(self.results)
            self.highlight_current()
    
    # 查找上一个匹配项
    def find_previous(self):
        if self.results:
            self.current_index = (self.current_index - 1) % len(self.results)
            self.highlight_current()
    
    # 高亮当前匹配项
    def highlight_current(self):
        if not self.results or self.current_index < 0:
            return
        
        row, col = self.results[self.current_index]
        self.tableWidget_3.scrollToItem(self.tableWidget_3.item(row, col))
        self.tableWidget_3.setCurrentCell(row, col)
        self.statusBar().showMessage(f"结果 {self.current_index + 1}/{len(self.results)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainUI()
    window.show()
    sys.exit(app.exec())