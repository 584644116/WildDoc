import sys
import os
import time
import pandas as pd
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QProgressBar, QTextEdit, 
                            QFileDialog, QMessageBox, QFrame, QGridLayout,
                            QScrollArea, QSizePolicy, QSpacerItem, QDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve, QRect
from PyQt5.QtGui import QFont, QPalette, QColor, QPainter, QPainterPath
from PyQt5.QtCore import QTimer

# 导入现有的业务逻辑
from core.document_generator import DocumentGenerator
from core.batch_processor import BatchProcessor
from gui.modern_field_config_dialog import ModernFieldConfigDialog
from utils.logger import Logger
from utils.error_handler import ErrorHandler


class ModernCard(QFrame):
    """现代化卡片组件"""
    
    def __init__(self, title="", subtitle="", parent=None):
        super().__init__(parent)
        self.title = title
        self.subtitle = subtitle
        self.setup_ui()
        self.setup_style()
    
    def setup_ui(self):
        """设置UI"""
        self.setFixedHeight(100)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(6)
        
        # 标题
        if self.title:
            title_label = QLabel(self.title)
            title_label.setObjectName("cardTitle")
            layout.addWidget(title_label)
        
        # 副标题
        if self.subtitle:
            subtitle_label = QLabel(self.subtitle)
            subtitle_label.setObjectName("cardSubtitle")
            layout.addWidget(subtitle_label)
        
        # 内容区域
        self.content_layout = QVBoxLayout()
        self.content_layout.setAlignment(Qt.AlignCenter)  # 内容居中
        layout.addLayout(self.content_layout)

        # 弹性空间
        layout.addStretch()
    
    def setup_style(self):
        """设置样式"""
        self.setObjectName("modernCard")
        self.setStyleSheet("""
            QFrame#modernCard {
                background-color: #ffffff;
                border: 1px solid #e1dfdd;
                border-radius: 8px;
                margin: 4px;
            }
            QFrame#modernCard:hover {
                border-color: #c7c5c4;
            }
            QLabel#cardTitle {
                color: #201f1e;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel#cardSubtitle {
                color: #605e5c;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)
    
    def add_content(self, widget):
        """添加内容到卡片"""
        self.content_layout.addWidget(widget)


class StepIndicator(QWidget):
    """步骤指示器"""
    
    def __init__(self, steps, parent=None):
        super().__init__(parent)
        self.steps = steps
        self.current_step = 0
        self.setup_ui()
    
    def setup_ui(self):
        """设置UI"""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(20)
        
        self.step_widgets = []
        
        for i, step in enumerate(self.steps):
            # 步骤容器
            step_container = QWidget()
            step_layout = QVBoxLayout(step_container)
            step_layout.setContentsMargins(0, 0, 0, 0)
            step_layout.setSpacing(8)
            
            # 步骤圆圈
            step_circle = QLabel(str(i + 1))
            step_circle.setObjectName("stepCircle")
            step_circle.setAlignment(Qt.AlignCenter)
            step_circle.setFixedSize(24, 24)
            
            # 步骤标题
            step_title = QLabel(step)
            step_title.setObjectName("stepTitle")
            step_title.setAlignment(Qt.AlignCenter)
            
            step_layout.addWidget(step_circle)
            step_layout.addWidget(step_title)
            
            self.step_widgets.append((step_circle, step_title))
            layout.addWidget(step_container)
            
            # 添加连接线（除了最后一个步骤）
            if i < len(self.steps) - 1:
                line = QFrame()
                line.setFrameShape(QFrame.HLine)
                line.setObjectName("stepLine")
                line.setFixedHeight(2)
                layout.addWidget(line)
        
        self.update_style()
    
    def set_current_step(self, step):
        """设置当前步骤"""
        self.current_step = step
        self.update_style()
    
    def update_style(self):
        """更新样式"""
        style = """
            QLabel#stepCircle {
                border-radius: 12px;
                font-size: 11px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel#stepTitle {
                color: #605e5c;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QFrame#stepLine {
                background-color: #edebe9;
                border: none;
            }
        """
        
        # 为每个步骤添加特定样式
        for i, (circle, title) in enumerate(self.step_widgets):
            if i < self.current_step:
                # 已完成步骤
                circle.setStyleSheet(style + """
                    QLabel#stepCircle {
                        background-color: #107c10;
                        color: white;
                    }
                """)
                title.setStyleSheet(style + """
                    QLabel#stepTitle {
                        color: #107c10;
                        font-weight: 600;
                    }
                """)
            elif i == self.current_step:
                # 当前步骤
                circle.setStyleSheet(style + """
                    QLabel#stepCircle {
                        background-color: #0078d4;
                        color: white;
                    }
                """)
                title.setStyleSheet(style + """
                    QLabel#stepTitle {
                        color: #0078d4;
                        font-weight: 600;
                    }
                """)
            else:
                # 未完成步骤
                circle.setStyleSheet(style + """
                    QLabel#stepCircle {
                        background-color: #f3f2f1;
                        color: #605e5c;
                        border: 2px solid #edebe9;
                    }
                """)
        
        self.setStyleSheet(style)


class ModernButton(QPushButton):
    """现代化按钮"""
    
    def __init__(self, text="", button_type="primary", parent=None):
        super().__init__(text, parent)
        self.button_type = button_type
        self.setup_style()
    
    def setup_style(self):
        """设置样式"""
        base_style = """
            QPushButton {
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 32px;
                text-align: center;
            }
        """
        
        if self.button_type == "primary":
            style = base_style + """
                QPushButton {
                    background-color: #0078d4;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #106ebe;
                }
                QPushButton:pressed {
                    background-color: #005a9e;
                }
                QPushButton:disabled {
                    background-color: #f3f2f1;
                    color: #a19f9d;
                }
            """
        elif self.button_type == "secondary":
            style = base_style + """
                QPushButton {
                    background-color: #ffffff;
                    color: #323130;
                    border: 1px solid #8a8886;
                }
                QPushButton:hover {
                    background-color: #f3f2f1;
                    border-color: #605e5c;
                }
                QPushButton:pressed {
                    background-color: #edebe9;
                }
                QPushButton:disabled {
                    background-color: #f3f2f1;
                    color: #a19f9d;
                    border-color: #edebe9;
                }
            """
        else:  # outline
            style = base_style + """
                QPushButton {
                    background-color: transparent;
                    color: #0078d4;
                    border: 1px solid #0078d4;
                }
                QPushButton:hover {
                    background-color: #f3f2f1;
                }
                QPushButton:pressed {
                    background-color: #edebe9;
                }
                QPushButton:disabled {
                    color: #a19f9d;
                    border-color: #edebe9;
                }
            """
        
        self.setStyleSheet(style)


class ModernProgressBar(QProgressBar):
    """现代化进度条"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_style()
    
    def setup_style(self):
        """设置样式"""
        self.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 4px;
                text-align: center;
                font-size: 13px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f3f2f1;
                height: 8px;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 4px;
            }
        """)


class ModernMainWindow(QMainWindow):
    """现代化主窗口"""
    
    def __init__(self):
        super().__init__()
        self.processor = BatchProcessor()
        self.logger = Logger('modern_main_window')
        
        # 初始化状态变量
        self.template_path = None
        self.excel_path = None
        self.output_dir = None
        self.excel_df = None
        self.field_mapping = None
        self.naming_rule = None
        self.generator = DocumentGenerator()
        
        # 步骤状态
        self.current_step = 0
        self.steps = ["选择模板", "导入数据", "配置字段", "设置输出", "生成文档"]
        
        self.init_ui()
        self.setup_connections()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle('Word文档生成工具 - 现代版')
        self.setGeometry(100, 100, 1000, 650)
        self.setMinimumSize(900, 600)
        
        # 设置主窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #faf9f8;
            }
        """)
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 16, 20, 16)
        main_layout.setSpacing(12)
        
        # 标题区域
        self.create_header(main_layout)
        
        # 步骤指示器
        self.step_indicator = StepIndicator(self.steps)
        main_layout.addWidget(self.step_indicator)
        
        # 功能卡片区域
        self.create_cards_area(main_layout)
        
        # 进度和日志区域
        self.create_progress_area(main_layout)
    
    def create_header(self, layout):
        """创建标题区域"""
        header_layout = QVBoxLayout()
        
        # 主标题
        title = QLabel("Word文档生成工具")
        title.setStyleSheet("""
            QLabel {
                color: #201f1e;
                font-size: 20px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-bottom: 2px;
            }
        """)
        header_layout.addWidget(title)

        # 副标题
        subtitle = QLabel("快速生成个性化Word文档")
        subtitle.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 13px;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)
        header_layout.addWidget(subtitle)
        
        layout.addLayout(header_layout)

    def create_cards_area(self, layout):
        """创建功能卡片区域（可滚动，避免小屏/高DPI遮挡）"""
        from PyQt5.QtWidgets import QScrollArea

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        container = QWidget()
        grid = QGridLayout(container)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(12)

        # 创建功能卡片 - 2x3布局
        self.create_template_card(grid, 0, 0)
        self.create_excel_card(grid, 0, 1)
        self.create_config_card(grid, 1, 0)
        self.create_output_card(grid, 1, 1)
        self.create_generate_card(grid, 2, 0, 2)  # 跨两列

        scroll.setWidget(container)
        layout.addWidget(scroll)

    def create_template_card(self, layout, row, col):
        """创建模板选择卡片"""
        card = ModernCard("1. 选择模板", "")

        # 按钮和状态的水平布局
        content_layout = QHBoxLayout()
        content_layout.setSpacing(8)

        # 选择按钮 - 紧凑
        self.template_btn = ModernButton("选择模板", "primary")
        self.template_btn.clicked.connect(self.select_template)
        self.template_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 6px 12px;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 24px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)
        content_layout.addWidget(self.template_btn)

        # 状态标签 - 紧凑
        self.template_status = QLabel("请选择Word模板文件")
        self.template_status.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 11px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 4px 8px;
                background-color: #f3f2f1;
                border-radius: 3px;
            }
        """)
        self.template_status.setWordWrap(True)
        content_layout.addWidget(self.template_status, 1)

        card.content_layout.addLayout(content_layout)
        layout.addWidget(card, row, col)

    def create_excel_card(self, layout, row, col):
        """创建Excel数据卡片"""
        card = ModernCard("2. 导入数据", "")

        # 按钮和状态的水平布局
        content_layout = QHBoxLayout()
        content_layout.setSpacing(8)

        # 选择按钮
        self.excel_btn = ModernButton("选择Excel", "primary")
        self.excel_btn.clicked.connect(self.select_excel_file)
        self.excel_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 6px 12px;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 24px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)
        content_layout.addWidget(self.excel_btn)

        # 状态标签
        self.excel_status = QLabel("请选择Excel文件")
        self.excel_status.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 11px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 4px 8px;
                background-color: #f3f2f1;
                border-radius: 3px;
            }
        """)
        self.excel_status.setWordWrap(True)
        content_layout.addWidget(self.excel_status, 1)

        card.content_layout.addLayout(content_layout)
        layout.addWidget(card, row, col)

    def create_config_card(self, layout, row, col):
        """创建字段配置卡片"""
        card = ModernCard("3. 配置字段", "")

        # 按钮和状态的水平布局
        content_layout = QHBoxLayout()
        content_layout.setSpacing(8)

        # 配置按钮
        self.config_btn = ModernButton("字段配置", "secondary")
        self.config_btn.clicked.connect(self.open_config_dialog)
        self.config_btn.setEnabled(False)
        self.config_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #323130;
                border: 1px solid #8a8886;
                border-radius: 3px;
                padding: 6px 12px;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 24px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #f3f2f1;
                border-color: #605e5c;
            }
            QPushButton:pressed {
                background-color: #edebe9;
            }
            QPushButton:disabled {
                background-color: #f3f2f1;
                color: #a19f9d;
                border-color: #edebe9;
            }
        """)
        content_layout.addWidget(self.config_btn)

        # 状态标签
        self.config_status = QLabel("需要模板和Excel")
        self.config_status.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 11px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 4px 8px;
                background-color: #f3f2f1;
                border-radius: 3px;
            }
        """)
        self.config_status.setWordWrap(True)
        content_layout.addWidget(self.config_status, 1)

        card.content_layout.addLayout(content_layout)
        layout.addWidget(card, row, col)

    def create_output_card(self, layout, row, col):
        """创建输出设置卡片"""
        card = ModernCard("4. 设置输出", "")

        # 按钮和状态的水平布局
        content_layout = QHBoxLayout()
        content_layout.setSpacing(8)

        # 选择按钮
        self.output_btn = ModernButton("选择目录", "primary")
        self.output_btn.clicked.connect(self.select_output_directory)
        self.output_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 6px 12px;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 24px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)
        content_layout.addWidget(self.output_btn)

        # 状态标签
        self.output_status = QLabel("请选择保存位置")
        self.output_status.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 11px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 4px 8px;
                background-color: #f3f2f1;
                border-radius: 3px;
            }
        """)
        self.output_status.setWordWrap(True)
        content_layout.addWidget(self.output_status, 1)

        card.content_layout.addLayout(content_layout)
        layout.addWidget(card, row, col)

    def create_generate_card(self, layout, row, col, colspan=1):
        """创建生成控制卡片"""
        card = ModernCard("5. 生成文档", "")
        card.setFixedHeight(100)  # 增加高度以容纳标题和按钮

        # 使用垂直布局，确保按钮居中
        content_layout = QVBoxLayout()
        content_layout.setAlignment(Qt.AlignCenter)

        # 生成按钮
        self.generate_btn = ModernButton("开始生成", "primary")
        self.generate_btn.clicked.connect(self.generate_documents)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 28px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #f3f2f1;
                color: #a19f9d;
            }
        """)

        # 水平居中布局
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(self.generate_btn)
        btn_layout.addStretch()

        content_layout.addLayout(btn_layout)
        card.content_layout.addLayout(content_layout)

        layout.addWidget(card, row, col, 1, colspan)

    def create_progress_area(self, layout):
        """创建进度和日志区域"""
        # 创建底部区域容器
        bottom_container = QWidget()
        bottom_layout = QVBoxLayout(bottom_container)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(12)

        # 进度区域 - 更紧凑
        progress_widget = QWidget()
        progress_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 1px solid #e1dfdd;
                border-radius: 8px;
            }
        """)
        progress_layout = QHBoxLayout(progress_widget)
        progress_layout.setContentsMargins(16, 12, 16, 12)
        progress_layout.setSpacing(16)

        # 进度标题
        progress_title = QLabel("进度:")
        progress_title.setStyleSheet("""
            QLabel {
                color: #201f1e;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                min-width: 50px;
            }
        """)
        progress_layout.addWidget(progress_title)

        # 进度条
        self.progress_bar = ModernProgressBar()
        self.progress_bar.setMaximumHeight(8)
        progress_layout.addWidget(self.progress_bar, 1)

        # 进度详情
        self.progress_detail = QLabel("等待开始...")
        self.progress_detail.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 13px;
                font-family: 'Segoe UI', Arial, sans-serif;
                min-width: 120px;
            }
        """)
        progress_layout.addWidget(self.progress_detail)

        bottom_layout.addWidget(progress_widget)

        # 日志区域
        log_widget = QWidget()
        log_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 1px solid #e1dfdd;
                border-radius: 8px;
            }
        """)
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(16, 12, 16, 12)
        log_layout.setSpacing(8)

        # 日志标题
        log_title = QLabel("操作日志")
        log_title.setStyleSheet("""
            QLabel {
                color: #201f1e;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)
        log_layout.addWidget(log_title)

        # 日志文本
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #faf9f8;
                border: 1px solid #edebe9;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                line-height: 1.3;
                color: #323130;
            }
        """)
        log_layout.addWidget(self.log_text)

        bottom_layout.addWidget(log_widget)
        layout.addWidget(bottom_container)

    def setup_connections(self):
        """设置信号连接"""
        pass  # 连接已在创建按钮时设置

    def update_step(self, step):
        """更新当前步骤"""
        self.current_step = step
        self.step_indicator.set_current_step(step)
        self.check_ready_state()

    def check_ready_state(self):
        """检查是否可以进行下一步"""
        # 检查配置按钮状态
        if self.template_path and self.excel_df is not None:
            self.config_btn.setEnabled(True)
            self.config_status.setText("可以配置字段映射")
            self.config_status.setStyleSheet("""
                QLabel {
                    color: #107c10;
                    font-size: 13px;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    padding: 8px;
                    background-color: #ddf4dd;
                    border-radius: 4px;
                    margin-top: 8px;
                }
            """)

        # 检查生成按钮状态
        if (self.template_path and self.excel_df is not None and
            self.output_dir and self.field_mapping is not None):
            self.generate_btn.setEnabled(True)

    @ErrorHandler.handle_gui
    def select_template(self, *args):
        """选择模板文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择Word模板', '', 'Word Files (*.docx *.doc)')

        if file_path:
            try:
                self.template_path = file_path
                if self.generator.load_template(file_path):
                    filename = os.path.basename(file_path)
                    self.template_status.setText(f"已选择: {filename}")
                    self.template_status.setStyleSheet("""
                        QLabel {
                            color: #107c10;
                            font-size: 13px;
                            font-family: 'Segoe UI', Arial, sans-serif;
                            padding: 8px;
                            background-color: #ddf4dd;
                            border-radius: 4px;
                            margin-top: 8px;
                        }
                    """)
                    self.log_text.append(f"✓ 已加载模板: {filename}")
                    self.log_text.append(f"  发现 {len(self.generator.placeholders)} 个占位符")
                    self.update_step(1)
                else:
                    QMessageBox.warning(self, "错误", "模板加载失败")
            except Exception as e:
                self.logger.error(f"选择模板失败: {str(e)}")
                QMessageBox.warning(self, "错误", f"选择模板失败: {str(e)}")

    @ErrorHandler.handle_gui
    def select_excel_file(self, *args):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择Excel文件', '', 'Excel Files (*.xlsx *.xls)')

        if file_path:
            try:
                self.excel_path = file_path
                self.excel_df = pd.read_excel(file_path)
                filename = os.path.basename(file_path)
                self.excel_status.setText(f"已选择: {filename} ({len(self.excel_df)} 行数据)")
                self.excel_status.setStyleSheet("""
                    QLabel {
                        color: #107c10;
                        font-size: 13px;
                        font-family: 'Segoe UI', Arial, sans-serif;
                        padding: 8px;
                        background-color: #ddf4dd;
                        border-radius: 4px;
                        margin-top: 8px;
                    }
                """)
                self.log_text.append(f"✓ 已加载Excel: {filename}")
                self.log_text.append(f"  包含 {len(self.excel_df)} 行数据，{len(self.excel_df.columns)} 列")
                self.update_step(2)
            except Exception as e:
                self.logger.error(f"读取Excel失败: {str(e)}")
                QMessageBox.warning(self, "错误", f"读取Excel失败: {str(e)}")

    @ErrorHandler.handle_gui
    def open_config_dialog(self, *args):
        """打开配置对话框"""
        if not self.template_path or self.excel_df is None:
            QMessageBox.warning(self, "提示", "请先选择模板和Excel文件")
            return

        # 准备模板信息
        template_info = {
            'placeholders': list(self.generator.placeholders.keys()),
            'tables': []
        }

        dialog = ModernFieldConfigDialog(template_info, self.excel_df, self)
        if dialog.exec_() == QDialog.Accepted:
            self.field_mapping = dialog.get_field_mapping()
            self.naming_rule = dialog.get_naming_rule()
            self.config_status.setText("字段映射配置完成")
            self.config_status.setStyleSheet("""
                QLabel {
                    color: #107c10;
                    font-size: 13px;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    padding: 8px;
                    background-color: #ddf4dd;
                    border-radius: 4px;
                    margin-top: 8px;
                }
            """)
            self.log_text.append("✓ 字段映射配置已完成")
            self.update_step(3)

    @ErrorHandler.handle_gui
    def select_output_directory(self, *args):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, '选择输出文件夹')

        if dir_path:
            self.output_dir = dir_path
            self.output_status.setText(f"输出到: {os.path.basename(dir_path)}")
            self.output_status.setStyleSheet("""
                QLabel {
                    color: #107c10;
                    font-size: 13px;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    padding: 8px;
                    background-color: #ddf4dd;
                    border-radius: 4px;
                    margin-top: 8px;
                }
            """)
            self.log_text.append(f"✓ 已设置输出目录: {dir_path}")
            self.update_step(4)

    def load_config_from_file(self):
        """从文件加载配置"""
        config_file = "field_config.json"
        if os.path.exists(config_file):
            try:
                import json
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.field_mapping = config.get('mapping', {})
                    self.naming_rule = config.get('naming_rule', {'pattern': '[姓名]_文档'})
                    self.log_text.append("✓ 已从配置文件加载字段映射")
                    return True
            except Exception as e:
                self.logger.error(f"加载配置文件失败: {str(e)}")
        return False

    @ErrorHandler.handle_gui
    def generate_documents(self, *args):
        """生成文档"""
        if not self.generator or self.excel_df is None or not self.output_dir:
            QMessageBox.warning(self, "错误", "请先完成所有选择和配置")
            return

        # 如果用户没有配置，尝试加载默认配置
        if self.field_mapping is None or self.naming_rule is None:
            self.log_text.append("未进行手动配置，尝试加载默认配置...")
            if not self.load_config_from_file():
                # 使用默认映射
                self.field_mapping = {p: p for p in self.generator.placeholders.keys()}
                self.naming_rule = {'pattern': '[姓名]_文档'}
                self.log_text.append("使用默认字段映射配置")

        # 设置生成器配置
        self.generator.set_mapping(self.field_mapping)
        self.generator.set_naming_rule(self.naming_rule)

        # 禁用生成按钮
        self.generate_btn.setEnabled(False)
        self.generate_btn.setText("生成中...")

        # 重置进度
        self.progress_bar.setValue(0)
        self.progress_detail.setText("准备生成文档...")

        # 启动生成线程
        self.generate_worker = GenerateWorker(
            self.generator, self.excel_df, self.output_dir)
        self.generate_worker.progress.connect(self.update_progress)
        self.generate_worker.log.connect(self.log_text.append)
        self.generate_worker.finished.connect(self.generation_finished)
        self.generate_worker.start()

        self.log_text.append(f"开始生成 {len(self.excel_df)} 个文档...")
        self.update_step(5)

    def update_progress(self, value):
        """更新进度"""
        self.progress_bar.setValue(value)
        if value < 100:
            self.progress_detail.setText(f"正在生成文档... {value}%")
        else:
            self.progress_detail.setText("生成完成！")

    def generation_finished(self):
        """生成完成"""
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("开始生成文档")
        self.progress_bar.setValue(100)
        self.progress_detail.setText("所有文档生成完成！")
        self.log_text.append("✓ 所有文档生成完成！")
        QMessageBox.information(self, "完成", "所有文档生成完成！")


class GenerateWorker(QThread):
    """文档生成工作线程"""

    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, generator, df, output_dir):
        super().__init__()
        self.generator = generator
        self.df = df
        self.output_dir = output_dir

    def run(self):
        """执行生成"""
        try:
            total_rows = len(self.df)

            for index, (row_idx, row) in enumerate(self.df.iterrows()):
                success, result = self.generator.generate_single_document(row, self.output_dir)

                if success:
                    self.log.emit(f"✓ 已生成: {os.path.basename(result)}")
                else:
                    self.log.emit(f"✗ 生成失败: {result}")

                # 更新进度
                progress = int((index + 1) * 100 / total_rows)
                self.progress.emit(progress)

            self.finished.emit()

        except Exception as e:
            self.log.emit(f"✗ 生成过程出错: {str(e)}")
            self.finished.emit()
