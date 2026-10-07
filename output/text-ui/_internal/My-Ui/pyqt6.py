from PyQt6.QtWidgets import QApplication, QWidget
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
import sys 


app = QApplication(sys.argv)

window = QWidget()              # 创建窗口
window.setWindowTitle('PyQt6')      # 设置窗口标题
window.setGeometry(100, 100, 800, 600)      # 设置窗口大小和位置
window.show()           # 显示窗口



sys.exit(app.exec())    # 退出程序