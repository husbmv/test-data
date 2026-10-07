import sys
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                             QFileDialog, QLabel, QMessageBox, QHeaderView)
from PyQt6.QtCore import Qt

class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.df = None
        
    def initUI(self):
        self.setWindowTitle('Excel 文件查看器 - PyQt6 版本')
        self.setGeometry(100, 100, 800, 600)
        
        # 创建中央部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)    # 垂直布局
        
        # 创建顶部按钮区域
        button_layout = QHBoxLayout()
        
        self.open_button = QPushButton('打开 Excel 文件')
        self.open_button.clicked.connect(self.open_file)
        button_layout.addWidget(self.open_button)
        
        self.clear_button = QPushButton('清空表格')
        self.clear_button.clicked.connect(self.clear_table)
        self.clear_button.setEnabled(False)
        button_layout.addWidget(self.clear_button)
        
        layout.addLayout(button_layout)
        
        # 创建状态标签
        self.status_label = QLabel('请打开一个 Excel 文件')
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)     # 添加到布局中

        # 创建表格部件
        self.table = QTableWidget()
        layout.addWidget(self.table)    # 添加到布局中
        
    def open_file(self):
        # 打开文件对话框
        file_path, _ = QFileDialog.getOpenFileName(
            self, '打开 Excel 文件', '', 'Excel Files (*.xlsx *.xls)')
        
        if file_path:
            try:
                # 读取 Excel 文件
                self.df = pd.read_excel(file_path)
                
                # 设置表格行数和列数
                self.table.setRowCount(self.df.shape[0])
                self.table.setColumnCount(self.df.shape[1])
                
                # 设置表头
                self.table.setHorizontalHeaderLabels(self.df.columns)
                
                # 填充数据
                for row in range(self.df.shape[0]):
                    for col in range(self.df.shape[1]):
                        item = QTableWidgetItem(str(self.df.iloc[row, col]))
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # 设置为不可编辑
                        self.table.setItem(row, col, item)
                
                # 调整列宽以适应内容
                self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
                
                # 更新状态标签
                self.status_label.setText(f'已加载文件: {file_path} | 行数: {self.df.shape[0]}, 列数: {self.df.shape[1]}')
                
                # 启用清空按钮
                self.clear_button.setEnabled(True)
                
            except Exception as e:
                QMessageBox.critical(self, '错误', f'无法读取文件:\n{str(e)}')
    
    def clear_table(self):
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.df = None
        self.status_label.setText('表格已清空')
        self.clear_button.setEnabled(False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    viewer = ExcelViewer()
    viewer.show()
    sys.exit(app.exec())