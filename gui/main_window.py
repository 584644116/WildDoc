from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QProgressBar, QTextEdit, 
                            QFileDialog, QMessageBox, QTabWidget, QDialog, QFrame,
                            QGridLayout, QSizePolicy, QGroupBox, QSplitter)
import json
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QFont, QPalette
import os
import sys
import win32com.client
import pythoncom
import pandas as pd
import psutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)

# 使用绝对导入
# -------- Core Imports ---------
from core.batch_processor import BatchProcessor
from utils.error_handler import ErrorHandler
from utils.logger import Logger
from core.document_generator import DocumentGenerator
from field_config_dialog import FieldConfigDialog  # 导入配置对话框

# ==================================
#   Thread Classes
# ==================================

# 1. 信息提取线程（扫描 Word -> Excel）


class ExtractWorker(QThread):
    """批量提取信息工作线程 (Word -> Excel)"""

    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, processor: BatchProcessor, directory: str):
        super().__init__()
        self.processor = processor
        self.directory = directory

    def _count_word_files(self) -> int:
        count = 0
        for root, _, files in os.walk(self.directory):
            for file in files:
                if (file.endswith('.doc') or file.endswith('.docx')) and not file.startswith('~$'):
                    count += 1
        return count

    def run(self):
        total_files = self._count_word_files()
        processed = 0

        def local_logger(message: str):
            # relay internal logger messages
            self.log.emit(message)

        # 为 BatchProcessor 设置 logger 回调（简单替换其 logger 的 info 方法）
        original_info = self.processor.logger.info
        self.processor.logger.info = lambda m: local_logger(m)

        try:
            # 自定义进度：在 BatchProcessor 内部每处理一个文件都会调用 logger.info
            # 我们统计包含 '成功处理文件' 字样的日志数量来粗略估计进度
            def progress_listener(msg: str):
                nonlocal processed
                if '成功处理文件' in msg:
                    processed += 1
                    if total_files:
                        self.progress.emit(int(processed * 100 / total_files))

            self.log.connect(progress_listener)

            self.processor.process_directory(self.directory)
        except Exception as e:
            self.log.emit(f"提取过程中发生错误: {str(e)}")
        finally:
            # 恢复 logger
            self.processor.logger.info = original_info
            self.finished.emit()

# 2. 高性能并行文档生成线程

class ParallelGenerateWorker(QThread):
    """高性能并行文档生成工作线程"""
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal()
    stats_update = pyqtSignal(dict)  # 新增：统计信息更新信号

    def __init__(self, generator, df, output_dir, mapping, naming_rule, max_workers=4):
        super().__init__()
        self.generator = generator
        self.df = df
        self.output_dir = output_dir
        self.mapping = mapping
        self.naming_rule = naming_rule
        self.max_workers = max_workers
        self.completed_count = 0
        self.failed_count = 0
        self.start_time = None

    def generate_single_document(self, row_data):
        """生成单个文档的函数（优化版）"""
        try:
            row_idx, row = row_data
            # 创建独立的生成器实例避免线程冲突
            local_generator = DocumentGenerator()
            local_generator.load_template(self.generator.template_path)
            local_generator.set_mapping(self.mapping)
            local_generator.set_naming_rule(self.naming_rule)
            
            # 使用新的单文档生成方法（更高效）
            success, result = local_generator.generate_single_document(row, self.output_dir)
            
            if success:
                return True, result  # result是文件路径
            else:
                return False, f"行 {row_idx}: {result}"  # result是错误信息
                
        except Exception as e:
            return False, f"行 {row_idx}: {str(e)}"

    def run(self):
        self.start_time = time.time()
        total_rows = len(self.df)
        self.log.emit(f"开始并行生成 {total_rows} 个文档，使用 {self.max_workers} 个线程...")

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # 提交所有任务
            future_to_row = {
                executor.submit(self.generate_single_document, (idx, row)): idx 
                for idx, row in self.df.iterrows()
            }

            # 处理完成的任务
            for future in as_completed(future_to_row):
                try:
                    success, result = future.result()
                    if success:
                        self.completed_count += 1
                        # 只显示文件名，不显示完整路径
                        filename = os.path.basename(result) if isinstance(result, str) else f"文档_{self.completed_count}"
                        self.log.emit(f"✅ [{self.completed_count:3d}/{total_rows}] {filename}")
                    else:
                        self.failed_count += 1
                        self.log.emit(f"❌ [{self.failed_count:3d}] 失败: {result}")
                    
                    # 更新进度
                    processed = self.completed_count + self.failed_count
                    progress_value = int(processed * 100 / total_rows)
                    self.progress.emit(progress_value)
                    
                    # 发送统计信息
                    elapsed_time = time.time() - self.start_time
                    avg_speed = processed / elapsed_time if elapsed_time > 0 else 0
                    remaining_time = (total_rows - processed) / avg_speed if avg_speed > 0 else 0
                    
                    # 计算预估完成时间
                    eta_text = "计算中..." if remaining_time <= 0 else f"{remaining_time:.0f}秒"
                    if remaining_time > 60:
                        eta_text = f"{remaining_time/60:.1f}分钟"
                    
                    stats = {
                        'completed': self.completed_count,
                        'failed': self.failed_count,
                        'total': total_rows,
                        'speed': f"{avg_speed:.1f} 文档/秒",
                        'remaining_time': eta_text,
                        'memory_usage': f"{psutil.Process().memory_info().rss / 1024 / 1024:.1f} MB"
                    }
                    self.stats_update.emit(stats)
                    
                except Exception as e:
                    self.failed_count += 1
                    self.log.emit(f"❌ 处理异常: {str(e)}")

        self.finished.emit()

# 保持原有的简单生成线程作为备选
class GenerateDocsWorker(QThread):
    """文档生成工作线程"""
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, generator, df, output_dir, mapping, naming_rule):
        super().__init__()
        self.generator = generator
        self.df = df
        self.output_dir = output_dir
        self.mapping = mapping
        self.naming_rule = naming_rule

    def run(self):
        try:
            self.generator.generate_documents(
                self.df, self.output_dir, self.mapping, self.naming_rule,
                progress_callback=self.progress
            )
        except Exception as e:
            self.log.emit(f"生成过程中发生严重错误: {str(e)}")
        finally:
            self.finished.emit()

# ==================================
#              MainWindow
# ==================================

class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.processor = BatchProcessor()
        self.logger = Logger('main_window')
        # 初始化状态变量
        self.template_path = None
        self.excel_path = None
        self.output_dir = None
        self.excel_df = None
        self.field_mapping = None
        self.naming_rule = None
        self.generator = DocumentGenerator()
        
        self.initUI()
    
    def get_microsoft_card_style(self) -> str:
        """获取微软风格卡片样式"""
        return """
            QFrame {
                background-color: #ffffff;
                border: 1px solid #edebe9;
                border-radius: 2px;
                padding: 24px;
                margin: 8px;
            }
        """
    
    def get_microsoft_button_style(self) -> str:
        """获取微软风格按钮样式"""
        return """
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: 1px solid #0078d4;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                border-radius: 2px;
                min-height: 32px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #106ebe;
                border-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
                border-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #f3f2f1;
                color: #a19f9d;
                border-color: #edebe9;
            }
        """
    
    def get_microsoft_progress_style(self) -> str:
        """获取微软风格进度条样式"""
        return """
            QProgressBar {
                border: 1px solid #edebe9;
                border-radius: 2px;
                text-align: center;
                font-size: 13px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f3f2f1;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 1px;
            }
        """

    def get_microsoft_log_style(self) -> str:
        """获取微软风格日志区域样式"""
        return """
            QTextEdit {
                background-color: #faf9f8;
                border: 1px solid #edebe9;
                border-radius: 2px;
                padding: 16px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                line-height: 1.4;
                color: #323130;
            }
        """
    
    def create_microsoft_section(self, title: str) -> tuple:
        """创建微软风格区域"""
        section = QWidget()
        section.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 1px solid #edebe9;
                border-radius: 2px;
            }
        """)
        
        layout = QVBoxLayout(section)
        layout.setContentsMargins(24, 20, 24, 24)
        layout.setSpacing(16)
        
        # 添加标题
        title_label = QLabel(title)
        title_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 20px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-bottom: 8px;
                border: none;
                background: transparent;
            }
        """)
        layout.addWidget(title_label)
        
        return section, layout
    
    def initUI(self):
        """初始化微软风格UI"""
        self.setWindowTitle('Word文档生成工具')
        self.setGeometry(100, 100, 900, 700)
        self.setMinimumSize(800, 600)
        
        # 设置主窗口背景色
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f3f2f1;
            }
        """)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(32, 32, 32, 32)
        main_layout.setSpacing(24)
        
        # 标题区域
        title_label = QLabel("Word文档生成工具")
        title_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 32px;
                font-weight: 300;
                font-family: 'Segoe UI Light', 'Segoe UI', Arial, sans-serif;
                margin-bottom: 8px;
            }
        """)
        main_layout.addWidget(title_label)
        
        # 副标题
        subtitle_label = QLabel("快速生成个性化Word文档")
        subtitle_label.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 16px;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-bottom: 24px;
            }
        """)
        main_layout.addWidget(subtitle_label)
        
        # 主要操作区域
        content_widget, content_layout = self.create_microsoft_section("设置文档生成参数")
        
        # 文件选择网格
        file_grid = QGridLayout()
        file_grid.setSpacing(16)
        file_grid.setContentsMargins(0, 0, 0, 0)
        
        # 模板选择行
        template_label = QLabel("Word模板:")
        template_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-width: 80px;
            }
        """)
        
        template_btn = QPushButton("选择模板文件")
        template_btn.clicked.connect(self.select_template)
        template_btn.setStyleSheet(self.get_microsoft_button_style())
        
        self.template_path_label = QLabel("请选择Word模板文件")
        self.template_path_label.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 8px 12px;
                background-color: #f3f2f1;
                border: 1px solid #edebe9;
                border-radius: 2px;
            }
        """)
        
        file_grid.addWidget(template_label, 0, 0)
        file_grid.addWidget(template_btn, 0, 1)
        file_grid.addWidget(self.template_path_label, 0, 2)
        
        # Excel选择行
        excel_label = QLabel("Excel数据:")
        excel_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-width: 80px;
            }
        """)
        
        excel_btn = QPushButton("选择Excel文件")
        excel_btn.clicked.connect(self.select_excel_file)
        excel_btn.setStyleSheet(self.get_microsoft_button_style())
        
        self.excel_path_label = QLabel("请选择包含数据的Excel文件")
        self.excel_path_label.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 8px 12px;
                background-color: #f3f2f1;
                border: 1px solid #edebe9;
                border-radius: 2px;
            }
        """)
        
        file_grid.addWidget(excel_label, 1, 0)
        file_grid.addWidget(excel_btn, 1, 1)
        file_grid.addWidget(self.excel_path_label, 1, 2)
        
        # 输出文件夹行
        output_label = QLabel("输出位置:")
        output_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-width: 80px;
            }
        """)
        
        output_btn = QPushButton("选择输出文件夹")
        output_btn.clicked.connect(self.select_output_directory)
        output_btn.setStyleSheet(self.get_microsoft_button_style())
        
        self.output_path_label = QLabel("请选择文档保存位置")
        self.output_path_label.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 8px 12px;
                background-color: #f3f2f1;
                border: 1px solid #edebe9;
                border-radius: 2px;
            }
        """)
        
        file_grid.addWidget(output_label, 2, 0)
        file_grid.addWidget(output_btn, 2, 1)
        file_grid.addWidget(self.output_path_label, 2, 2)
        
        # 设置列宽比例
        file_grid.setColumnStretch(0, 0)  # 标签列固定宽度
        file_grid.setColumnStretch(1, 0)  # 按钮列固定宽度
        file_grid.setColumnStretch(2, 1)  # 路径显示列自适应
        
        content_layout.addLayout(file_grid)
        
        # 配置按钮
        config_layout = QHBoxLayout()
        config_layout.setContentsMargins(0, 16, 0, 0)
        
        self.config_btn = QPushButton("配置字段映射")
        self.config_btn.clicked.connect(self.open_config_dialog)
        self.config_btn.setStyleSheet("""
            QPushButton {
                background-color: #f3f2f1;
                color: #323130;
                border: 1px solid #edebe9;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                border-radius: 2px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #edebe9;
                border-color: #d2d0ce;
            }
            QPushButton:disabled {
                background-color: #f3f2f1;
                color: #a19f9d;
                border-color: #edebe9;
            }
        """)
        self.config_btn.setEnabled(False)
        
        config_layout.addWidget(self.config_btn)
        config_layout.addStretch()
        content_layout.addLayout(config_layout)
        
        main_layout.addWidget(content_widget)
        
        # 生成按钮
        generate_layout = QHBoxLayout()
        generate_layout.setContentsMargins(0, 0, 0, 0)
        
        self.generate_btn = QPushButton("开始生成文档")
        self.generate_btn.clicked.connect(self.generate_documents)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: 1px solid #0078d4;
                padding: 16px 32px;
                font-size: 16px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                border-radius: 2px;
                min-height: 48px;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #106ebe;
                border-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
                border-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #f3f2f1;
                color: #a19f9d;
                border-color: #edebe9;
            }
        """)
        
        generate_layout.addStretch()
        generate_layout.addWidget(self.generate_btn)
        generate_layout.addStretch()
        main_layout.addLayout(generate_layout)
        
        # 进度区域
        progress_widget, progress_layout = self.create_microsoft_section("生成进度")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet(self.get_microsoft_progress_style())
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_detail_label = QLabel("等待开始...")
        self.progress_detail_label.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-top: 8px;
            }
        """)
        progress_layout.addWidget(self.progress_detail_label)
        
        main_layout.addWidget(progress_widget)
        
        # 日志区域
        log_widget, log_layout = self.create_microsoft_section("操作日志")
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(self.get_microsoft_log_style())
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        
        main_layout.addWidget(log_widget)
        
        # 添加弹性空间
        main_layout.addStretch()
    
    def select_directory(self, *args):
        """选择要处理的文件夹以批量提取信息"""
        dir_path = QFileDialog.getExistingDirectory(self, '选择包含Word文件的文件夹')

        if dir_path:
            self.log_text.append(f'开始提取信息: {dir_path}')

            # 重置进度条
            self.progress_bar.setValue(0)

            try:
                self.extract_worker = ExtractWorker(self.processor, dir_path)
                self.extract_worker.log.connect(self.log_text.append)
                self.extract_worker.progress.connect(self.update_progress)
                self.extract_worker.finished.connect(self.processing_finished)
                self.extract_worker.start()
            except Exception as e:
                self.logger.error(f"提取失败: {str(e)}")
                QMessageBox.warning(self, "错误", f"提取失败: {str(e)}")
    
    def processing_finished(self):
        """信息提取完成回调"""
        QMessageBox.information(self, "完成", "信息提取完成！")
        self.progress_bar.setValue(100)
    
    @ErrorHandler.handle_gui
    def select_template(self, *args):
        """选择Word模板"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择Word模板', '', 'Word Files (*.docx *.doc)')
        
        if file_path:
            self.template_path = file_path
            if self.generator.load_template(file_path):
                self.template_path_label.setText(os.path.basename(file_path))
                self.template_path_label.setStyleSheet("""
                    QLabel {
                        color: #107c10;
                        font-size: 14px;
                        font-family: 'Segoe UI', Arial, sans-serif;
                        font-weight: 600;
                        padding: 8px 12px;
                        background-color: #ddf4dd;
                        border: 1px solid #c3e6c3;
                        border-radius: 2px;
                    }
                """)
                self.log_text.append(f'✅ 已加载模板: {os.path.basename(file_path)}')
                self.check_config_ready()
            else:
                QMessageBox.warning(self, "错误", "加载模板失败")
                self.template_path = None
                self.template_path_label.setText('模板加载失败')
                self.template_path_label.setStyleSheet("""
                    QLabel {
                        color: #d13438;
                        font-size: 14px;
                        font-family: 'Segoe UI', Arial, sans-serif;
                        padding: 8px 12px;
                        background-color: #fde7e9;
                        border: 1px solid #f1aeb5;
                        border-radius: 2px;
                    }
                """)
    
    @ErrorHandler.handle_gui
    def select_excel_file(self, *args):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择Excel文件', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            try:
                self.excel_path = file_path
                self.excel_df = pd.read_excel(file_path)
                self.excel_path_label.setText(os.path.basename(file_path))
                self.excel_path_label.setStyleSheet("""
                    QLabel {
                        color: #107c10;
                        font-size: 14px;
                        font-family: 'Segoe UI', Arial, sans-serif;
                        font-weight: 600;
                        padding: 8px 12px;
                        background-color: #ddf4dd;
                        border: 1px solid #c3e6c3;
                        border-radius: 2px;
                    }
                """)
                self.log_text.append(f'✅ 已加载Excel文件: {os.path.basename(file_path)}')
                self.check_config_ready()
            except Exception as e:
                QMessageBox.warning(self, "错误", f"加载Excel文件失败: {str(e)}")
                self.excel_path = None
                self.excel_df = None
                self.excel_path_label.setText('Excel文件加载失败')
                self.excel_path_label.setStyleSheet("""
                    QLabel {
                        color: #d13438;
                        font-size: 14px;
                        font-family: 'Segoe UI', Arial, sans-serif;
                        padding: 8px 12px;
                        background-color: #fde7e9;
                        border: 1px solid #f1aeb5;
                        border-radius: 2px;
                    }
                """)
    
    def check_config_ready(self):
        """检查模板和Excel是否都已加载"""
        if self.template_path and self.excel_path:
            self.config_btn.setEnabled(True)
            self.log_text.append("模板和Excel文件均已就绪，可以进行配置。")

    @ErrorHandler.handle_gui
    def open_config_dialog(self, *args):
        """打开配置对话框"""
        if not self.template_path or self.excel_df is None:
            QMessageBox.warning(self, "提示", "请先选择模板和Excel文件")
            return

        # 准备模板信息
        template_info = {
            'placeholders': list(self.generator.placeholders.keys()),
            # 目前不支持多表格预览，简化处理
            'tables': [] 
        }

        dialog = FieldConfigDialog(template_info, self.excel_df, self)
        if dialog.exec_() == QDialog.Accepted:
            self.field_mapping = dialog.get_field_mapping()
            self.naming_rule = dialog.get_naming_rule()
            self.log_text.append("配置已更新并保存。")

    @ErrorHandler.handle_gui
    def select_output_directory(self, *args):
        """选择输出文件夹"""
        dir_path = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        if dir_path:
            self.output_dir = dir_path
            self.output_path_label.setText(dir_path)
            self.output_path_label.setStyleSheet("""
                QLabel {
                    color: #107c10;
                    font-size: 14px;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    font-weight: 600;
                    padding: 8px 12px;
                    background-color: #ddf4dd;
                    border: 1px solid #c3e6c3;
                    border-radius: 2px;
                }
            """)
            self.log_text.append(f'✅ 已选择输出文件夹: {dir_path}')
    
    @ErrorHandler.handle_gui
    def generate_documents(self, *args):
        """生成文档"""
        if not self.generator or self.excel_df is None or not self.output_dir:
            QMessageBox.warning(self, "错误", "请先完成所有选择和配置")
            return

        # 如果用户没有打开过配置界面，就使用默认/已保存的配置
        if self.field_mapping is None or self.naming_rule is None:
            self.log_text.append("未进行手动配置，尝试从 'field_config.json' 加载配置...")
            self.load_config_from_file()

        # 禁用按钮，准备开始
        self.generate_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_detail_label.setText("正在初始化...")
        
        total_docs = len(self.excel_df)
        self.log_text.append(f"开始生成 {total_docs} 个文档...")

        # 根据文档数量选择处理模式
        use_parallel = total_docs > 5  # 文档数量大于5时使用并行处理
        max_workers = min(4, max(2, total_docs // 10))  # 动态计算线程数
        
        if use_parallel:
            self.log_text.append(f"使用并行模式，启动 {max_workers} 个工作线程")
            # 创建并启动并行工作线程
            self.gen_worker = ParallelGenerateWorker(
                self.generator, self.excel_df, self.output_dir,
                self.field_mapping, self.naming_rule, max_workers
            )
        else:
            self.log_text.append("使用串行模式处理")
            # 创建串行工作线程（原有的简单模式）
            self.gen_worker = GenerateDocsWorker(
                self.generator, self.excel_df, self.output_dir,
                self.field_mapping, self.naming_rule
            )
        
        # 连接信号
        self.gen_worker.progress.connect(self.update_progress)
        self.gen_worker.log.connect(self.log_text.append)
        self.gen_worker.finished.connect(self.generation_finished)
        self.gen_worker.start()

    def generation_finished(self):
        """生成完成后的收尾工作"""
        self.generate_btn.setEnabled(True)
        self.progress_bar.setValue(100)
        self.progress_detail_label.setText("生成完成")
        
        total = len(self.excel_df) if self.excel_df is not None else 0
        
        self.log_text.append("=" * 40)
        self.log_text.append("文档生成任务完成！")
        self.log_text.append(f"总计: {total} 个文档")
        self.log_text.append("=" * 40)
        
        # 弹出完成提示
        QMessageBox.information(self, "生成完成", f"文档生成完成！\n总计: {total} 个文档")

    def update_progress(self, value: int):
        """更新进度条"""
        self.progress_bar.setValue(value)
        self.progress_detail_label.setText(f"进度: {value}%")



    def load_config_from_file(self):
        """从文件加载映射和命名规则"""
        config_path = "field_config.json"
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                    self.field_mapping = cfg.get('mapping', {})
                    self.naming_rule = cfg.get('naming_rule', {})
                    self.log_text.append("成功加载已保存的配置。")
            except Exception as e:
                self.log_text.append(f"加载配置文件失败: {str(e)}")
                self.field_mapping = {}
                self.naming_rule = {}
        else:
            self.log_text.append("未找到配置文件，将使用默认映射。")
            self.field_mapping = {}
            self.naming_rule = {}

    @ErrorHandler.handle_gui
    def batch_convert_docs(self, *args):
        """批量转换DOC文件到DOCX"""
        dir_path = QFileDialog.getExistingDirectory(self, '选择包含DOC文件的文件夹')
        
        if dir_path:
            self.log_text.append(f'开始转换文件夹: {dir_path}')
            # 重置进度条
            self.progress_bar.setValue(0)
            
            # 创建转换线程
            self.convert_worker = ConvertWorkerThread(dir_path)
            self.convert_worker.log.connect(self.log_text.append)
            self.convert_worker.progress.connect(self.update_progress)
            self.convert_worker.finished.connect(self.conversion_finished)
            self.convert_worker.start()
            
    def conversion_finished(self):
        """转换完成的回调"""
        QMessageBox.information(self, "完成", "文档转换完成")
        self.progress_bar.setValue(100)

# 添加新的转换工作线程类
class ConvertWorkerThread(QThread):
    """文档转换工作线程"""
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    
    def __init__(self, directory):
        super().__init__()
        self.directory = directory
        self.word_app = None
    
    def clean_path(self, path: str) -> str:
        """清理文件路径"""
        # 替换路径中的正斜杠为反斜杠
        path = path.replace('/', '\\')
        # 确保路径是绝对路径
        path = os.path.abspath(path)
        # 处理特殊字符
        path = path.replace('（', '(').replace('）', ')')
        path = path.replace('，', ',').replace('：', ':')
        # 如果路径包含空格或特殊字符，添加引号
        if any(c in path for c in ' -()[]（）'):
            path = f'"{path}"'
        return path
    
    def run(self):
        try:
            # 初始化 Word 应用
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = False  # 禁用警告弹窗
            
            # 获取所有 .doc 文件
            doc_files = []
            for root, dirs, files in os.walk(self.directory):
                for file in files:
                    if file.endswith('.doc') and not file.startswith('~$'):
                        doc_files.append(os.path.join(root, file))
            
            total_files = len(doc_files)
            converted_count = 0
            failed_count = 0
            
            for i, doc_path in enumerate(doc_files, 1):
                try:
                    # 清理并构造路径
                    clean_doc_path = self.clean_path(doc_path)
                    docx_path = clean_doc_path.rstrip('"') + 'x'  # 添加 x 后缀
                    if clean_doc_path.startswith('"'):
                        docx_path = f'"{docx_path}"'
                    
                    # 创建临时目录
                    temp_dir = os.path.join(os.path.dirname(doc_path), 'temp')
                    os.makedirs(temp_dir, exist_ok=True)
                    
                    # 在临时目录中创建文件副本，使用简单的文件名
                    temp_doc = os.path.join(temp_dir, f'temp_{i}.doc')
                    temp_docx = os.path.join(temp_dir, f'temp_{i}.docx')
                    
                    # 复制原文件到临时目录
                    import shutil
                    shutil.copy2(doc_path, temp_doc)
                    
                    try:
                        # 转换文件
                        doc = self.word_app.Documents.Open(temp_doc)
                        doc.SaveAs2(temp_docx, FileFormat=16)  # 16 表示 docx 格式
                        doc.Close()
                        
                        # 将转换后的文件复制到目标位置
                        target_path = docx_path.strip('"')
                        shutil.copy2(temp_docx, target_path)
                        
                        converted_count += 1
                        self.log.emit(f'已转换 ({converted_count}/{total_files}): {os.path.basename(doc_path)}')
                        self.progress.emit(int(i * 100 / total_files))
                        
                    finally:
                        # 清理临时文件
                        try:
                            shutil.rmtree(temp_dir)
                        except:
                            pass
                    
                except Exception as e:
                    failed_count += 1
                    self.log.emit(f'转换失败 ({failed_count}): {os.path.basename(doc_path)}\n错误: {str(e)}')
                    
            # 显示最终统计
            self.log.emit(f'\n转换完成:\n成功: {converted_count}\n失败: {failed_count}\n总计: {total_files}')
            
        except Exception as e:
            self.log.emit(f'转换过程出错: {str(e)}')
            
        finally:
            if self.word_app:
                try:
                    self.word_app.DisplayAlerts = True  # 恢复警告设置
                    self.word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize() 