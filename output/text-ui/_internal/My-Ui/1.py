import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                           QPushButton, QListWidget, QLabel, QWidget, QFileDialog, 
                           QMessageBox, QProgressBar, QLineEdit, QCheckBox, 
                           QTextEdit, QSplitter, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QTextCursor
from datetime import datetime

class ExcelMergerThread(QThread):
    """后台线程用于处理Excel合并操作"""
    progress_updated = pyqtSignal(int)
    log_updated = pyqtSignal(str, str)  # 信号: (消息, 类型)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, file_paths, output_path, skip_duplicates=True):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path
        self.skip_duplicates = skip_duplicates
    
    def log(self, message, msg_type="INFO"):
        """记录日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] [{msg_type}] {message}"
        self.log_updated.emit(log_message, msg_type)
    
    def run(self):
        try:
            self.log("开始Excel文件合并过程", "INFO")
            self.log(f"共需要处理 {len(self.file_paths)} 个文件", "INFO")
            self.log(f"输出文件: {self.output_path}", "INFO")
            self.log(f"删除重复行: {'是' if self.skip_duplicates else '否'}", "INFO")
            
            # 检查输出目录是否存在
            output_dir = os.path.dirname(self.output_path)
            if not os.path.exists(output_dir):
                self.log(f"创建输出目录: {output_dir}", "INFO")
                os.makedirs(output_dir)
            
            # 创建新的Excel文件
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                total_files = len(self.file_paths)
                processed_files = 0
                
                for i, file_path in enumerate(self.file_paths):
                    # 更新进度
                    progress = int((i / total_files) * 100)
                    self.progress_updated.emit(progress)
                    
                    # 获取文件名（不带扩展名）作为工作表名
                    file_name = os.path.splitext(os.path.basename(file_path))[0]
                    self.log(f"正在处理文件: {file_name}", "INFO")
                    
                    try:
                        # 检查文件是否存在
                        if not os.path.exists(file_path):
                            self.log(f"文件不存在: {file_path}", "ERROR")
                            raise FileNotFoundError(f"文件不存在: {file_path}")
                        
                        # 读取Excel文件
                        self.log(f"读取Excel文件: {file_path}", "INFO")
                        df = pd.read_excel(file_path)
                        
                        # 记录文件基本信息
                        self.log(f"文件 '{file_name}' 包含 {len(df)} 行, {len(df.columns)} 列", "INFO")
                        
                        # 如果启用去重，删除完全重复的行
                        if self.skip_duplicates:
                            original_rows = len(df)
                            df = df.drop_duplicates()
                            removed_rows = original_rows - len(df)
                            if removed_rows > 0:
                                self.log(f"删除 {removed_rows} 个重复行", "INFO")
                        
                        # 处理工作表名称（确保名称有效且不重复）
                        sheet_name = self.get_valid_sheet_name(file_name, writer)
                        
                        # 将数据写入到新的工作表
                        self.log(f"创建工作表: {sheet_name}", "INFO")
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        self.log(f"文件 '{file_name}' 处理完成", "SUCCESS")
                        processed_files += 1
                        
                    except Exception as e:
                        self.log(f"处理文件 {file_name} 时出错: {str(e)}", "ERROR")
                        # 尝试创建一个空的工作表作为占位符
                        try:
                            sheet_name = self.get_valid_sheet_name(f"ERROR_{file_name}", writer)
                            pd.DataFrame({'错误信息': [f"处理文件时出错: {str(e)}"]}).to_excel(
                                writer, sheet_name=sheet_name, index=False)
                            self.log(f"已创建错误占位符工作表: {sheet_name}", "WARNING")
                        except Exception as e2:
                            self.log(f"创建错误占位符工作表失败: {str(e2)}", "ERROR")
                
                # 完成进度
                self.progress_updated.emit(100)
                self.log(f"合并完成！成功处理 {processed_files}/{total_files} 个文件", "SUCCESS")
                self.log(f"输出文件已保存到: {self.output_path}", "INFO")
                
            self.finished_signal.emit(True, f"成功合并 {processed_files}/{total_files} 个文件到: {self.output_path}")
            
        except Exception as e:
            self.log(f"合并过程中发生严重错误: {str(e)}", "ERROR")
            self.finished_signal.emit(False, f"合并失败: {str(e)}")
    
    def get_valid_sheet_name(self, proposed_name, writer):
        """获取有效的工作表名称（处理重复和无效字符）"""
        # 替换无效字符
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        sheet_name = proposed_name
        for char in invalid_chars:
            sheet_name = sheet_name.replace(char, '_')
        
        # 限制长度（Excel工作表名最多31个字符）
        sheet_name = sheet_name[:31]
        
        # 检查是否已存在同名工作表
        if hasattr(writer, 'book'):
            existing_sheets = writer.book.sheetnames
            base_name = sheet_name
            counter = 1
            while sheet_name in existing_sheets:
                sheet_name = f"{base_name}_{counter}"
                counter += 1
                if counter > 100:  # 防止无限循环
                    sheet_name = f"{base_name}_{datetime.now().strftime('%H%M%S')}"
                    break
        
        return sheet_name

class LogTextEdit(QTextEdit):
    """自定义日志文本框"""
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setFont(QFont("Consolas", 9))
        self.setMaximumHeight(200)
        self.setMinimumHeight(150)
    
    def append_log(self, message, msg_type="INFO"):
        """添加带颜色的日志消息"""
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        
        # 根据消息类型设置颜色
        if msg_type == "ERROR":
            color = "#FF0000"  # 红色
        elif msg_type == "WARNING":
            color = "#FFA500"  # 橙色
        elif msg_type == "SUCCESS":
            color = "#008000"  # 绿色
        else:  # INFO
            color = "#0000FF"  # 蓝色
        
        # 插入带格式的文本
        cursor.insertHtml(f'<span style="color: {color};"><b>{message}</b></span><br>')
        
        # 滚动到底部
        self.setTextCursor(cursor)
        self.ensureCursorVisible()

class ExcelMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_paths = []
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Excel文件合并工具 - 带日志输出")
        self.setGeometry(100, 100, 900, 700)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel("Excel文件合并工具")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 创建分割器，使日志区域可以调整大小
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        # 上部区域：文件操作
        upper_widget = QWidget()
        upper_layout = QVBoxLayout()
        
        # 文件选择区域
        file_selection_layout = QHBoxLayout()
        
        self.add_files_btn = QPushButton("添加Excel文件")
        self.add_files_btn.clicked.connect(self.add_files)
        file_selection_layout.addWidget(self.add_files_btn)
        
        self.add_folder_btn = QPushButton("添加文件夹")
        self.add_folder_btn.clicked.connect(self.add_folder)
        file_selection_layout.addWidget(self.add_folder_btn)
        
        self.clear_list_btn = QPushButton("清空列表")
        self.clear_list_btn.clicked.connect(self.clear_file_list)
        file_selection_layout.addWidget(self.clear_list_btn)
        
        upper_layout.addLayout(file_selection_layout)
        
        # 文件列表标签
        file_list_label = QLabel("已选择的文件:")
        upper_layout.addWidget(file_list_label)
        
        # 文件列表
        self.file_list_widget = QListWidget()
        self.file_list_widget.setMaximumHeight(150)
        upper_layout.addWidget(self.file_list_widget)
        
        # 选项区域
        options_layout = QHBoxLayout()
        
        self.skip_duplicates_check = QCheckBox("删除重复行")
        self.skip_duplicates_check.setChecked(True)
        options_layout.addWidget(self.skip_duplicates_check)
        
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)
        options_layout.addWidget(self.clear_log_btn)
        
        options_layout.addStretch()
        
        upper_layout.addLayout(options_layout)
        
        # 输出路径选择
        output_layout = QHBoxLayout()
        
        output_layout.addWidget(QLabel("输出文件:"))
        
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("选择或输入输出文件路径...")
        output_layout.addWidget(self.output_path_edit)
        
        self.browse_output_btn = QPushButton("浏览")
        self.browse_output_btn.clicked.connect(self.browse_output_path)
        output_layout.addWidget(self.browse_output_btn)
        
        upper_layout.addLayout(output_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        upper_layout.addWidget(self.progress_bar)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        
        self.merge_btn = QPushButton("开始合并")
        self.merge_btn.clicked.connect(self.start_merge)
        self.merge_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; }")
        button_layout.addWidget(self.merge_btn)
        
        self.cancel_btn = QPushButton("退出")
        self.cancel_btn.clicked.connect(self.close)
        button_layout.addWidget(self.cancel_btn)
        
        upper_layout.addLayout(button_layout)
        
        upper_widget.setLayout(upper_layout)
        splitter.addWidget(upper_widget)
        
        # 下部区域：日志输出
        log_widget = QWidget()          # 日志输出
        log_layout = QVBoxLayout()      # 日志输出布局
        
        # 日志标题
        log_label = QLabel("合并日志:")
        log_label.setStyleSheet("font-weight: bold;")
        log_layout.addWidget(log_label)
        
        # 日志文本框
        self.log_text_edit = LogTextEdit()
        log_layout.addWidget(self.log_text_edit)
        
        log_widget.setLayout(log_layout)
        splitter.addWidget(log_widget)
        
        # 设置分割器比例
        splitter.setSizes([400, 300])
        
        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)
        
        # 设置默认输出路径
        self.set_default_output_path()
        
        # 输出初始日志
        self.log("应用程序已启动", "INFO")
        self.log("请添加要合并的Excel文件", "INFO")
    
    def set_default_output_path(self):
        """设置默认输出路径"""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        default_path = os.path.join(desktop, "合并结果.xlsx")
        self.output_path_edit.setText(default_path)
        self.log(f"默认输出路径设置为: {default_path}", "INFO")
    
    def log(self, message, msg_type="INFO"):
        """添加日志消息"""
        self.log_text_edit.append_log(message, msg_type)
    
    def clear_log(self):
        """清空日志"""
        self.log_text_edit.clear()
        self.log("日志已清空", "INFO")
    
    def add_files(self):
        """添加Excel文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        
        if files:
            added_count = 0
            for file in files:
                if file not in self.file_paths:
                    self.file_paths.append(file)
                    self.file_list_widget.addItem(os.path.basename(file))
                    added_count += 1
            
            self.log(f"成功添加 {added_count} 个文件", "SUCCESS")
            self.log(f"当前已选择 {len(self.file_paths)} 个文件", "INFO")
    
    def add_folder(self):
        """添加文件夹中的所有Excel文件"""
        folder = QFileDialog.getExistingDirectory(self, "选择包含Excel文件的文件夹")
        
        if folder:
            excel_files = []
            for file in os.listdir(folder):
                if file.lower().endswith(('.xlsx', '.xls')):
                    excel_files.append(os.path.join(folder, file))
            
            if excel_files:
                added_count = 0
                for file in excel_files:
                    if file not in self.file_paths:
                        self.file_paths.append(file)
                        self.file_list_widget.addItem(os.path.basename(file))
                        added_count += 1
                
                self.log(f"从文件夹 '{os.path.basename(folder)}' 成功添加 {added_count} 个Excel文件", "SUCCESS")
                self.log(f"当前已选择 {len(self.file_paths)} 个文件", "INFO")
            else:
                self.log(f"在文件夹 '{folder}' 中没有找到Excel文件", "WARNING")
                QMessageBox.warning(self, "警告", "选择的文件夹中没有找到Excel文件")
    
    def clear_file_list(self):
        """清空文件列表"""
        file_count = len(self.file_paths)
        self.file_paths.clear()
        self.file_list_widget.clear()
        self.log(f"已清空文件列表，移除了 {file_count} 个文件", "INFO")
    
    def browse_output_path(self):
        """选择输出文件路径"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存合并结果", "", "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_path_edit.setText(file_path)
            self.log(f"输出路径设置为: {file_path}", "INFO")
    
    def start_merge(self):
        """开始合并文件"""
        if not self.file_paths:
            self.log("错误: 没有选择要合并的文件", "ERROR")
            QMessageBox.warning(self, "警告", "请先选择要合并的Excel文件")
            return
        
        output_path = self.output_path_edit.text().strip()
        if not output_path:
            self.log("错误: 输出文件路径为空", "ERROR")
            QMessageBox.warning(self, "警告", "请选择输出文件路径")
            return
        
        # 检查输出文件扩展名
        if not output_path.lower().endswith('.xlsx'):
            output_path += '.xlsx'
            self.output_path_edit.setText(output_path)
        
        # 检查是否覆盖现有文件
        if os.path.exists(output_path):
            reply = QMessageBox.question(
                self, "确认覆盖", 
                f"文件 {output_path} 已存在，是否覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                self.log("用户取消覆盖操作", "INFO")
                return
        
        # 禁用界面控件
        self.set_ui_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.log("=" * 50, "INFO")
        self.log("开始合并Excel文件", "INFO")
        self.log("=" * 50, "INFO")
        
        # 创建并启动工作线程
        self.worker_thread = ExcelMergerThread(
            self.file_paths, 
            output_path,
            self.skip_duplicates_check.isChecked()
        )
        self.worker_thread.progress_updated.connect(self.progress_bar.setValue)
        self.worker_thread.log_updated.connect(self.log_text_edit.append_log)
        self.worker_thread.finished_signal.connect(self.merge_finished)
        self.worker_thread.start()
    
    def set_ui_enabled(self, enabled):
        """启用或禁用界面控件"""
        self.add_files_btn.setEnabled(enabled)
        self.add_folder_btn.setEnabled(enabled)
        self.clear_list_btn.setEnabled(enabled)
        self.merge_btn.setEnabled(enabled)
        self.browse_output_btn.setEnabled(enabled)
        self.output_path_edit.setEnabled(enabled)
        self.skip_duplicates_check.setEnabled(enabled)
        self.clear_log_btn.setEnabled(enabled)
    
    def merge_finished(self, success, message):
        """合并完成后的处理"""
        self.set_ui_enabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            self.log("=" * 50, "SUCCESS")
            self.log("合并操作完成", "SUCCESS")
            self.log("=" * 50, "SUCCESS")
            QMessageBox.information(self, "完成", message)
        else:
            self.log("=" * 50, "ERROR")
            self.log("合并操作失败", "ERROR")
            self.log("=" * 50, "ERROR")
            QMessageBox.critical(self, "错误", message)
    
    def closeEvent(self, event):
        """关闭应用程序时的处理"""
        self.log("应用程序正在关闭...", "INFO")
        # 确保工作线程安全退出
        if hasattr(self, 'worker_thread') and self.worker_thread.isRunning():
            self.log("正在停止后台任务...", "WARNING")
            self.worker_thread.terminate()
            self.worker_thread.wait()
        event.accept()

def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序样式
    app.setStyle('Fusion')
    
    window = ExcelMergerApp()
    window.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()