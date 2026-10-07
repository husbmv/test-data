import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                             QSlider, QCheckBox, QSpinBox, QProgressBar)
from PyQt6.QtCore import Qt

class ResponsiveWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # 设置窗口标题和初始大小
        self.setWindowTitle('自适应窗口示例 - 拖动边缘调整大小或点击全屏按钮')
        self.setGeometry(100, 100, 800, 600)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主垂直布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建顶部水平布局
        top_layout = QHBoxLayout()
        
        # 添加一些示例组件到顶部布局
        label = QLabel('窗口大小自适应示例')
        label.setStyleSheet('font-size: 16px; font-weight: bold; padding: 10px;')
        top_layout.addWidget(label)
        
        # 全屏按钮
        fullscreen_btn = QPushButton('切换全屏')
        fullscreen_btn.clicked.connect(self.toggle_fullscreen)
        top_layout.addWidget(fullscreen_btn)
        
        # 添加到主布局
        main_layout.addLayout(top_layout)
        
        # 创建内容区域水平布局
        content_layout = QHBoxLayout()
        
        # 左侧面板
        left_panel = QVBoxLayout()
        
        # 添加一些控件到左侧面板
        left_panel.addWidget(QLabel('设置选项:'))
        
        # 复选框
        checkbox1 = QCheckBox('启用自动调整')
        checkbox1.setChecked(True)
        left_panel.addWidget(checkbox1)
        
        checkbox2 = QCheckBox('显示网格')
        left_panel.addWidget(checkbox2)
        
        # 滑块
        slider_label = QLabel('大小调整:')
        left_panel.addWidget(slider_label)
        
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(1, 100)
        slider.setValue(50)
        left_panel.addWidget(slider)
        
        # 微调框
        spin_label = QLabel('数值设置:')
        left_panel.addWidget(spin_label)
        
        spinbox = QSpinBox()
        spinbox.setRange(0, 100)
        spinbox.setValue(25)
        left_panel.addWidget(spinbox)
        
        # 添加弹性空间使控件在顶部对齐
        left_panel.addStretch()
        
        # 右侧面板
        right_panel = QVBoxLayout()
        
        # 添加一些控件到右侧面板
        right_panel.addWidget(QLabel('内容区域:'))
        
        # 文本编辑框
        text_edit = QTextEdit()
        text_edit.setPlaceholderText('在此输入文本...')
        text_edit.setMinimumHeight(200)
        right_panel.addWidget(text_edit)
        
        # 进度条
        progress_label = QLabel('进度:')
        right_panel.addWidget(progress_label)
        
        progress_bar = QProgressBar()
        progress_bar.setValue(75)
        right_panel.addWidget(progress_bar)
        
        # 按钮
        button_layout = QHBoxLayout()
        button1 = QPushButton('确定')
        button2 = QPushButton('取消')
        button3 = QPushButton('帮助')
        button_layout.addWidget(button1)
        button_layout.addWidget(button2)
        button_layout.addWidget(button3)
        right_panel.addLayout(button_layout)
        
        # 将左右面板添加到内容布局
        content_layout.addLayout(left_panel, 1)  # 比例1
        content_layout.addLayout(right_panel, 2)  # 比例2，右侧占更多空间
        
        # 将内容布局添加到主布局
        main_layout.addLayout(content_layout, 1)  # 比例1，占据剩余空间
        
        # 底部状态栏
        status_layout = QHBoxLayout()
        status_label = QLabel('状态: 就绪')
        status_layout.addWidget(status_label)
        status_layout.addStretch()  # 将状态标签推到左侧
        
        main_layout.addLayout(status_layout)
        
    def toggle_fullscreen(self):
        if self.isFullScreen():
            self.showNormal()
        else:
            self.showFullScreen()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ResponsiveWindow()
    window.show()
    sys.exit(app.exec())