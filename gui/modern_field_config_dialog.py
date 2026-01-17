from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
                            QTableWidgetItem, QPushButton, QHeaderView, QMessageBox,
                            QComboBox, QLabel, QLineEdit, QGroupBox, QFrame,
                            QScrollArea, QWidget, QSizePolicy, QGridLayout)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
import json
import os
import pandas as pd


class FieldMappingCard(QFrame):
    """字段映射卡片"""

    def __init__(self, placeholder, excel_columns, parent=None):
        super().__init__(parent)
        self.placeholder = placeholder
        self.excel_columns = excel_columns
        self.setup_ui()

    def setup_ui(self):
        """设置UI"""
        # 固定卡片宽度，便于动态排列
        self.setFixedSize(200, 100)
        self.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e1dfdd;
                border-radius: 6px;
                margin: 2px;
            }
            QFrame:hover {
                border-color: #0078d4;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(4)

        # 上半部分：模板占位符
        placeholder_label = QLabel(self.placeholder)
        placeholder_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 13px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f3f2f1;
                padding: 6px 8px;
                border-radius: 4px;
            }
        """)
        placeholder_label.setAlignment(Qt.AlignCenter)
        placeholder_label.setWordWrap(True)
        layout.addWidget(placeholder_label)

        # 下半部分：Excel列选择
        self.column_combo = QComboBox()
        self.column_combo.addItems([""] + self.excel_columns)
        self.column_combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #8a8886;
                border-radius: 4px;
                padding: 4px 8px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
                min-height: 20px;
            }
            QComboBox:hover {
                border-color: #605e5c;
            }
            QComboBox:focus {
                border-color: #0078d4;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 4px solid #605e5c;
                margin-right: 6px;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                border: 1px solid #8a8886;
                border-radius: 4px;
                selection-background-color: #deecf9;
                selection-color: #323130;
                color: #323130;
                outline: none;
            }
            QComboBox QAbstractItemView::item {
                padding: 4px 8px;
                color: #323130;
                background-color: #ffffff;
            }
            QComboBox QAbstractItemView::item:selected {
                background-color: #deecf9;
                color: #323130;
            }
        """)

        # 智能匹配
        if self.placeholder in self.excel_columns:
            self.column_combo.setCurrentText(self.placeholder)

        layout.addWidget(self.column_combo)

    def get_mapping(self):
        """获取映射关系"""
        return self.placeholder, self.column_combo.currentText()

    def set_mapping(self, column_name):
        """设置映射关系"""
        self.column_combo.setCurrentText(column_name)


class ModernFieldConfigDialog(QDialog):
    """现代化字段配置对话框"""
    
    def __init__(self, template_info, excel_df, parent=None):
        super().__init__(parent)
        self.template_info = template_info
        self.excel_df = excel_df
        self.config_file = "field_config.json"
        
        # 命名规则默认值
        self.naming_rule = {
            'columns': [],
            'separator': '、',
            'fixed_text': ''
        }
        
        self.init_ui()
        self.populate_data()
        self.load_config()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("字段配置")
        self.setGeometry(150, 150, 900, 650)
        self.setMinimumSize(800, 600)
        
        # 设置对话框样式
        self.setStyleSheet("""
            QDialog {
                background-color: #faf9f8;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 16, 20, 16)
        main_layout.setSpacing(8)

        # 简化的标题
        title = QLabel("字段配置")
        title.setStyleSheet("""
            QLabel {
                color: #201f1e;
                font-size: 18px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-bottom: 8px;
            }
        """)
        main_layout.addWidget(title)

        # 字段映射区域
        self.create_mapping_section(main_layout)

        # 命名规则区域
        self.create_naming_section(main_layout)

        # 按钮区域
        self.create_buttons(main_layout)
    

    
    def create_mapping_section(self, layout):
        """创建字段映射区域"""
        # 区域标题
        section_title = QLabel("字段映射配置")
        section_title.setStyleSheet("""
            QLabel {
                color: #201f1e;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-bottom: 4px;
            }
        """)
        layout.addWidget(section_title)

        # 说明文字
        description = QLabel("将模板占位符映射到Excel列")
        description.setStyleSheet("""
            QLabel {
                color: #605e5c;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-bottom: 8px;
                line-height: 1.0;
            }
        """)
        layout.addWidget(description)

        # 滚动区域包含卡片
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: 1px solid #e1dfdd;
                border-radius: 6px;
                background-color: #faf9f8;
            }
        """)

        # 卡片容器
        cards_widget = QWidget()
        cards_widget.setStyleSheet("""
            QWidget {
                background-color: #faf9f8;
            }
        """)

        # 使用网格布局排列卡片
        self.cards_layout = QGridLayout(cards_widget)
        self.cards_layout.setContentsMargins(12, 12, 12, 12)
        self.cards_layout.setSpacing(10)

        # 创建字段映射卡片
        self.mapping_cards = []
        self.setup_mapping_cards()

        scroll_area.setWidget(cards_widget)
        layout.addWidget(scroll_area, 1)  # 占用主要空间

    def setup_mapping_cards(self):
        """设置字段映射卡片"""
        placeholders = self.template_info.get('placeholders', [])
        excel_columns = list(self.excel_df.columns) if self.excel_df is not None else []

        # 创建所有卡片
        for placeholder in placeholders:
            card = FieldMappingCard(placeholder, excel_columns)
            self.mapping_cards.append(card)

        # 初始排列
        self.rearrange_cards()

    def rearrange_cards(self):
        """重新排列卡片"""
        if not self.mapping_cards:
            return

        # 清除现有布局
        for i in reversed(range(self.cards_layout.count())):
            item = self.cards_layout.itemAt(i)
            if item:
                widget = item.widget()
                if widget:
                    self.cards_layout.removeWidget(widget)

        # 计算每行卡片数量
        card_width = 210  # 卡片宽度 + 间距
        available_width = self.width() - 80  # 减去边距和滚动条
        cols_per_row = max(3, available_width // card_width)

        # 重新添加卡片
        for i, card in enumerate(self.mapping_cards):
            row = i // cols_per_row
            col = i % cols_per_row
            self.cards_layout.addWidget(card, row, col)

        # 添加弹性空间
        if len(self.mapping_cards) % cols_per_row != 0:
            last_row = len(self.mapping_cards) // cols_per_row
            for col in range(len(self.mapping_cards) % cols_per_row, cols_per_row):
                spacer = QWidget()
                spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
                self.cards_layout.addWidget(spacer, last_row, col)

    def resizeEvent(self, event):
        """窗口大小变化事件"""
        super().resizeEvent(event)
        if hasattr(self, 'mapping_cards') and self.mapping_cards:
            # 延迟重新排列，避免频繁调用
            QTimer.singleShot(100, self.rearrange_cards)
    
    def create_naming_section(self, layout):
        """创建命名规则区域"""
        # 区域标题
        section_title = QLabel("文件命名规则")
        section_title.setStyleSheet("""
            QLabel {
                color: #201f1e;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                margin-top: 8px;
                margin-bottom: 4px;
            }
        """)
        layout.addWidget(section_title)

        # 命名模式输入 - 水平布局
        pattern_layout = QHBoxLayout()
        pattern_layout.setSpacing(8)

        pattern_label = QLabel("命名模式:")
        pattern_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 12px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                min-width: 70px;
            }
        """)
        pattern_layout.addWidget(pattern_label)

        self.naming_pattern_edit = QLineEdit("[姓名]_岗位能力确认表")
        self.naming_pattern_edit.setStyleSheet("""
            QLineEdit {
                background-color: #ffffff;
                border: 1px solid #8a8886;
                border-radius: 3px;
                padding: 4px 8px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
                min-height: 16px;
            }
            QLineEdit:hover {
                border-color: #605e5c;
            }
            QLineEdit:focus {
                border-color: #0078d4;
            }
        """)
        pattern_layout.addWidget(self.naming_pattern_edit, 1)

        layout.addLayout(pattern_layout)

        # 命名示例
        example_layout = QHBoxLayout()
        example_layout.setSpacing(8)

        example_label = QLabel("示例:")
        example_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 12px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                min-width: 70px;
            }
        """)
        example_layout.addWidget(example_label)

        # 动态生成示例
        self.example_text = QLabel()
        self.example_text.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 4px 8px;
                background-color: #deecf9;
                border-radius: 3px;
            }
        """)
        example_layout.addWidget(self.example_text, 1)

        layout.addLayout(example_layout)

        # 连接输入框变化事件
        self.naming_pattern_edit.textChanged.connect(self.update_naming_example)

        # 可用列提示 - 增大文字
        if self.excel_df is not None:
            available_columns = ", ".join(self.excel_df.columns)
            columns_label = QLabel(f"可用列: {available_columns}")
            columns_label.setWordWrap(True)
            columns_label.setStyleSheet("""
                QLabel {
                    color: #605e5c;
                    font-size: 12px;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    background-color: #f3f2f1;
                    padding: 8px 12px;
                    border-radius: 4px;
                    margin-top: 6px;
                    line-height: 1.2;
                }
            """)
            layout.addWidget(columns_label)

        # 使用说明
        help_text = QLabel("使用说明: 在命名模式中使用 [列名] 作为占位符，如 [姓名]_报告.docx")
        help_text.setWordWrap(True)
        help_text.setStyleSheet("""
            QLabel {
                color: #797775;
                font-size: 11px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding: 6px 8px;
                background-color: #f8f7f6;
                border-radius: 3px;
                margin-top: 4px;
                line-height: 1.3;
            }
        """)
        layout.addWidget(help_text)

        # 初始化示例
        self.update_naming_example()

    def update_naming_example(self):
        """更新命名示例"""
        pattern = self.naming_pattern_edit.text()
        if not pattern:
            self.example_text.setText("请输入命名模式")
            return

        # 生成示例文件名
        example_name = pattern

        # 如果有Excel数据，使用第一行数据作为示例
        if self.excel_df is not None and len(self.excel_df) > 0:
            first_row = self.excel_df.iloc[0]

            # 替换所有可能的占位符
            for column in self.excel_df.columns:
                placeholder = f"[{column}]"
                if placeholder in example_name:
                    value = str(first_row[column])
                    # 显示完整值，不省略
                    example_name = example_name.replace(placeholder, value)
        else:
            # 没有数据时使用通用示例
            import re
            placeholders = re.findall(r'\[([^\]]+)\]', pattern)
            for placeholder in placeholders:
                example_name = example_name.replace(f"[{placeholder}]", f"示例{placeholder}")

        # 确保有文件扩展名
        if not example_name.endswith('.docx') and not example_name.endswith('.doc'):
            example_name += '.docx'

        self.example_text.setText(f"示例文件名: {example_name}")
    
    def populate_data(self):
        """填充数据"""
        pass  # 数据填充已在setup_mapping_cards中完成

    def create_buttons(self, layout):
        """创建按钮区域"""
        button_frame = QFrame()
        button_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e1dfdd;
                border-radius: 8px;
                padding: 16px;
            }
        """)

        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(16, 16, 16, 16)
        button_layout.setSpacing(12)

        # 预览按钮
        preview_btn = QPushButton("预览配置")
        preview_btn.clicked.connect(self.preview_config)
        preview_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #0078d4;
                border: 1px solid #0078d4;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 32px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #f3f2f1;
            }
            QPushButton:pressed {
                background-color: #edebe9;
            }
        """)

        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #323130;
                border: 1px solid #8a8886;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 32px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #f3f2f1;
                border-color: #605e5c;
            }
            QPushButton:pressed {
                background-color: #edebe9;
            }
        """)

        # 确认按钮
        ok_btn = QPushButton("确认并保存")
        ok_btn.clicked.connect(self.save_and_accept)
        ok_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-weight: 600;
                min-height: 32px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)

        button_layout.addWidget(preview_btn)
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(ok_btn)

        layout.addWidget(button_frame)

    def load_config(self):
        """加载已保存的配置"""
        if not os.path.exists(self.config_file):
            return

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)

                # 加载字段映射
                mapping_cfg = config.get('mapping', {})
                for card in self.mapping_cards:
                    placeholder, _ = card.get_mapping()
                    if placeholder in mapping_cfg:
                        col_name = mapping_cfg[placeholder]
                        card.set_mapping(col_name)

                # 加载命名规则
                if 'naming_rule' in config and 'pattern' in config['naming_rule']:
                    self.naming_pattern_edit.setText(config['naming_rule']['pattern'])
                    self.update_naming_example()  # 更新示例

        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载配置失败: {str(e)}")

    def save_config(self):
        """保存配置"""
        # 获取字段映射
        mapping_cfg = self.get_field_mapping()

        # 获取命名规则
        self.naming_rule = {
            'pattern': self.naming_pattern_edit.text()
        }

        config = {
            'mapping': mapping_cfg,
            'naming_rule': self.naming_rule
        }

        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存配置失败: {str(e)}")

    def save_and_accept(self):
        """保存并接受"""
        self.save_config()
        self.accept()

    def preview_config(self):
        """预览配置"""
        mapping = self.get_field_mapping()
        preview = "字段映射预览：\n\n"

        for placeholder, col in mapping.items():
            if col:
                preview += f"[{placeholder}] → {col}\n"
            else:
                preview += f"[{placeholder}] → 未选择\n"

        pattern = self.naming_pattern_edit.text()
        preview += f"\n文件命名规则：\n{pattern}\n"

        # 创建预览对话框
        preview_dialog = QMessageBox(self)
        preview_dialog.setWindowTitle("配置预览")
        preview_dialog.setText(preview)
        preview_dialog.setIcon(QMessageBox.Information)
        preview_dialog.setStyleSheet("""
            QMessageBox {
                background-color: #ffffff;
                color: #323130;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QMessageBox QLabel {
                font-size: 14px;
                color: #323130;
            }
        """)
        preview_dialog.exec_()

    def get_field_mapping(self):
        """获取字段映射"""
        mapping = {}
        for card in self.mapping_cards:
            placeholder, column = card.get_mapping()
            mapping[placeholder] = column
        return mapping

    def get_naming_rule(self):
        """获取命名规则"""
        return {
            'pattern': self.naming_pattern_edit.text()
        }
