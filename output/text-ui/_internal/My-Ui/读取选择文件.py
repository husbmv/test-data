import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTableWidget, 
                             QTableWidgetItem, QVBoxLayout, QWidget)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 设置窗口和布局
        self.setWindowTitle('QTableWidget 获取选中行示例')
        self.setGeometry(300, 300, 600, 400)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 创建表格（5行3列）
        self.table_widget = QTableWidget(5, 3)
        layout.addWidget(self.table_widget)

        # 设置表格为整行选择模式
        self.table_widget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        # 设置选择模式为单行选择（可根据需要改为多选）
        self.table_widget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        # 初始化表格数据
        self.populate_table()

        # 连接信号：当选择改变时触发槽函数
        self.table_widget.itemSelectionChanged.connect(self.on_selection_changed)

    def populate_table(self):
        """填充表格示例数据"""
        headers = ['姓名', '年龄', '城市']
        self.table_widget.setHorizontalHeaderLabels(headers)
        
        sample_data = [
            ('Alice', '25', '北京'),
            ('Bob', '30', '上海'),
            ('Carol', '28', '广州'),
            ('David', '35', '深圳'),
            ('Eve', '22', '杭州')
        ]
        
        for row, (name, age, city) in enumerate(sample_data):
            self.table_widget.setItem(row, 0, QTableWidgetItem(name))
            self.table_widget.setItem(row, 1, QTableWidgetItem(age))
            self.table_widget.setItem(row, 2, QTableWidgetItem(city))

    def on_selection_changed(self):
        """当选择发生变化时调用，输出当前选中行的内容"""
        # 获取当前选中的所有范围
        selected_ranges = self.table_widget.selectedRanges()
        
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
                for col in range(self.table_widget.columnCount()):
                    item = self.table_widget.item(row, col)
                    # 如果单元格有内容，则获取文本，否则为空字符串
                    cell_text = item.text() if item else ''
                    row_data.append(cell_text)
                # 输出当前行的内容
                print(f"行 {row} 的内容: {row_data}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

# def mousePressEvent(self, event):
#     """重写鼠标点击事件，如果点击在空白处则清除选择"""
#     if self.table_widget.itemAt(event.pos()) is None:
#         self.table_widget.clearSelection()
#     super().mousePressEvent(event)