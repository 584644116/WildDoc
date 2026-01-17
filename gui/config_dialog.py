from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
                            QTableWidgetItem, QPushButton, QLabel, QHeaderView,
                            QComboBox, QCheckBox, QMessageBox, QTabWidget,
                            QWidget, QGroupBox, QMenu, QApplication)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDrag, QDragEnterEvent, QDropEvent, QColor
import pandas as pd
import os
import sys

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)

# 使用绝对导入
from utils.error_handler import ErrorHandler
from utils.logger import Logger
from utils.validator import ConfigValidator

class DraggableTableWidget(QTableWidget):
    """支持拖拽的表格控件"""
    
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setDragDropMode(QTableWidget.InternalMove)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasText():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasText():
            text = event.mimeData().text()
            row = self.rowAt(event.pos().y())
            if row >= 0:
                self.item(row, 1).setText(text)
                event.acceptProposedAction()
        else:
            super().dropEvent(event)

class ConfigDialog(QDialog):
    """配置对话框"""
    
    def __init__(self, template_info, config_manager, parent=None):
        super().__init__(parent)
        self.logger = Logger('config_dialog')
        self.template_info = template_info
        self.initUI()
    
    def initUI(self):
        """初始化UI"""
        self.setWindowTitle('提取字段预览')
        self.setGeometry(100, 100, 800, 600)
        
        layout = QVBoxLayout(self)
        
        # 创建预览表格
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(2)
        self.preview_table.setHorizontalHeaderLabels(['字段名称', '提取内容'])
        
        # 设置列宽
        header = self.preview_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        
        # 添加提取的字段内容
        fields = self.template_info['fields']
        self.preview_table.setRowCount(len(fields))
        
        for i, (field_name, value) in enumerate(fields.items()):
            # 字段名称
            name_item = QTableWidgetItem(field_name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)  # 设置为只读
            self.preview_table.setItem(i, 0, name_item)
            
            # 字段内容
            value_item = QTableWidgetItem(str(value))
            value_item.setFlags(value_item.flags() & ~Qt.ItemIsEditable)  # 设置为只读
            self.preview_table.setItem(i, 1, value_item)
        
        layout.addWidget(self.preview_table)
        
        # 添加确认按钮
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton('确认')
        ok_btn.clicked.connect(self.accept)
        btn_layout.addStretch()
        btn_layout.addWidget(ok_btn)
        
        layout.addLayout(btn_layout) 