# -*- coding: utf-8 -*-
"""
UI模块 - 论文格式助手主界面
"""

import os
import subprocess
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QHeaderView, QPushButton,
    QLabel, QLineEdit, QProgressBar, QFileDialog, QMessageBox,
    QFrame, QCheckBox, QComboBox, QStyleFactory
)
from PySide6.QtCore import Qt, QThread, Signal

from utils.config import (
    CHINESE_FONTS, ENGLISH_FONTS, FONT_SIZE_NAMES,
    ALIGNMENT_NAMES, LINE_SPACING_NAMES, SPACINGS,
    FIRST_LINE_INDENTS, OVERALL_INDENTS,
    PARTS, TABLE_COLUMNS, DEFAULT_FORMATS,
    DEFAULT_OUTPUT_DIR_HINT, DEFAULT_FILENAME_HINT
)
from services.doc_service import DocumentProcessingService


# ==================== 自定义组合框和复选框 widget ====================

class TableComboBox(QWidget):
    """表格中使用的组合框"""
    def __init__(self, items, default_item="", parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)

        self.combo = QComboBox()
        self.combo.addItems(items)
        if default_item in items:
            self.combo.setCurrentText(default_item)

        # 匹配图中的精致下拉框：纯白背景，浅灰边框，圆角，字号适中
        self.combo.setStyleSheet("""
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px 4px;
                font-family: "Microsoft YaHei UI";
                font-size: 13px;
                color: #334155;
            }
            QComboBox:hover {
                border: 1px solid #93C5FD;
                background-color: #F8FAFC;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 0px;
                border-left: none;
            }
            QComboBox::down-arrow {
                width: 0px;
                height: 0px;
            }
        """)

        self.combo.installEventFilter(self)
        layout.addWidget(self.combo)

    def text(self): return self.combo.currentText()
    def setText(self, text):
        index = self.combo.findText(text)
        if index >= 0: self.combo.setCurrentIndex(index)
    def eventFilter(self, obj, event):
        from PySide6.QtCore import QEvent
        if event.type() == QEvent.Wheel: return True
        return super().eventFilter(obj, event)


class TableCheckBox(QWidget):
    """表格中使用的复选框"""
    def __init__(self, checked=False, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(checked)

        self.checkbox.setStyleSheet("""
            QCheckBox { spacing: 0px; }
            QCheckBox::indicator {
                width: 16px; height: 16px;
                border: 1px solid #CBD5E1; border-radius: 4px; background-color: #FFFFFF;
            }
            QCheckBox::indicator:checked { background-color: #3B82F6; border: 1px solid #3B82F6; }
        """)
        layout.addWidget(self.checkbox)

    def isChecked(self): return self.checkbox.isChecked()
    def setChecked(self, checked): self.checkbox.setChecked(checked)


# ==================== 格式化工作线程 ====================

class FormatThread(QThread):
    """格式化工作线程"""
    progress_update = Signal(int, str)
    finished = Signal(bool, str, str)
    error = Signal(str)

    def __init__(self, task_snapshot):
        super().__init__()
        self.task_snapshot = task_snapshot

    def run(self):
        service = DocumentProcessingService(progress_callback=self.progress_update.emit)
        result = service.process(self.task_snapshot)

        output_path = result.get("output_path", "")
        report_path = result.get("report_path", "")
        error_msg = result.get("error", "")

        if output_path:
            self.finished.emit(result.get("success", False), output_path, report_path)
            return

        if report_path:
            error_msg = f"{error_msg}\n错误报告：{report_path}" if error_msg else f"错误报告：{report_path}"
        self.error.emit(error_msg or "处理失败")


class ThesisFormatterApp(QMainWindow):
    """论文格式助手主程序"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("论文格式助手 v1.0")
        self.resize(1360, 860)
        self.setMinimumSize(1220, 780)

        # 设置应用样式
        self.setup_styles()

        # 存储格式规则
        self.format_vars = {}
        self.format_widgets = {}

        # 存储文件路径
        self.last_output_path = None
        self.last_report_path = None

        # 创建UI
        self.create_ui()

    def setup_styles(self):
        """设置应用样式 - 像素级复刻 Win11 现代卡片风"""
        QApplication.setStyle(QStyleFactory.create("Fusion"))

        stylesheet = """
            /* 1. 整体窗口背景：浅蓝灰色 */
            QMainWindow { background-color: #F0F4F8; }

            /* 2. 全局文字 */
            QWidget {
                color: #334155;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
                font-size: 13px;
            }

            /* 3. 两大块卡片背景（纯白+圆角+极淡的边框模拟阴影） */
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
            }

            /* 4. 标题文字大小 */
            QLabel { border: none; background: transparent; }

            /* 5. 输入框样式 */
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 8px 12px;
                color: #334155;
            }
            QLineEdit:focus { border: 2px solid #60A5FA; }

            /* 6. 核心按钮："开始排版"（亮蓝色渐变） */
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3B82F6, stop:1 #2563EB);
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #60A5FA, stop:1 #3B82F6); }
            QPushButton:disabled { background: #CBD5E1; color: #94A3B8; }

            /* 7. 浏览按钮（浅蓝色） */
            QPushButton[objectName="browse"] {
                background-color: #93C5FD;
                color: white;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton[objectName="browse"]:hover { background-color: #60A5FA; }

            /* 8. 底部小按钮（浅灰色） */
            QPushButton[objectName="action"] {
                background-color: #E2E8F0;
                color: #475569;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton[objectName="action"]:hover { background-color: #CBD5E1; }
            QPushButton[objectName="action"]:disabled { background-color: #F1F5F9; color: #94A3B8; }

            /* 9. 表格整体样式（无网格线，斑马纹，直角） */
            QTableWidget {
                background-color: #FFFFFF;
                border: none;
                border-radius: 0px;
                gridline-color: transparent;
                alternate-background-color: #F8FAFC;
                selection-background-color: #EFF6FF;
                selection-color: #1E3A8A;
            }
            QTableWidget::item { border-bottom: 1px solid #F1F5F9; }

            /* 10. 表头样式（浅蓝灰底色） */
            QHeaderView::section {
                background-color: #E2EBF5;
                color: #475569;
                font-weight: bold;
                border: none;
                border-radius: 0px;
                padding: 8px;
            }

            /* 左上角交界处样式 */
            QTableWidget QTableCornerButton::section {
                background-color: #E2EBF5;
                border: none;
                border-radius: 0px;
            }

            /* 11. 进度条（胶囊形状） */
            QProgressBar {
                border: none;
                background-color: #E2E8F0;
                border border-radius: 4px;
                height: 8px;
                text-align: center;
                color: transparent;
            }
            QProgressBar::chunk { background-color: #3B82F6; border-radius: 4px; }

            /* 12. 滚动条美化 */
            QScrollBar:vertical {
                border: none;
                background: #F8FAFC;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #CBD5E1;
                min-height: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94A3B8;
            }

            QScrollBar:horizontal {
                border: none;
                background: #F8FAFC;
                height: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal {
                background: #CBD5E1;
                min-width: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #94A3B8;
            }
        """
        self.setStyleSheet(stylesheet)

    def create_ui(self):
        """创建主界面 - Notion风格极简设计"""
        # 主窗口
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(28, 28, 28, 28)

        # 标题区域
        self.create_header(main_layout)

        # 内容区域（左右布局）
        content_layout = QHBoxLayout()
        content_layout.setSpacing(20)

        # 左侧：格式规则表格
        format_card = self.create_card("格式规则设置")
        self.create_format_panel(format_card.layout())
        content_layout.addWidget(format_card, 3)

        # 右侧：文件操作面板
        file_card = self.create_card("文件操作")
        self.create_file_panel(file_card.layout())
        content_layout.addWidget(file_card, 1)

        main_layout.addLayout(content_layout)

    def create_header(self, parent_layout):
        """创建标题区域 - 清新学术风格"""
        header_widget = QWidget()
        header_layout = QVBoxLayout(header_widget)
        header_layout.setSpacing(8)
        header_layout.setContentsMargins(0, 0, 0, 0)

        # 主标题 - 使用明亮的蓝色和清晰的现代字体
        title_label = QLabel("论文格式助手")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 26px;
                font-weight: 700;
                color: #1e293b;
                letter-spacing: 2px;
            }
        """)

        header_layout.addWidget(title_label)
        parent_layout.addWidget(header_widget)
        parent_layout.addSpacing(16)

    def create_card(self, title):
        """创建卡片容器 - 清新学术风格"""
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: #f9fafb;
                border: 1px solid #e5e7eb;
                border-radius: 12px;
            }
        """)

        layout = QVBoxLayout(card)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # 卡片标题
        if title:
            title_label = QLabel(title)
            title_label.setStyleSheet("""
                QLabel {
                    font-size: 16px;
                    font-weight: 600;
                    color: #3b82f6;
                    padding-bottom: 4px;
                    background-color: transparent;
                    border: none;
                    letter-spacing: 0.5px;
                }
            """)
            layout.addWidget(title_label)

        return card

    def create_format_panel(self, parent_layout):
        """创建格式设置面板"""
        # 创建表格
        self.table = QTableWidget()
        self.table.setFocusPolicy(Qt.NoFocus)

        # 设置表格列数和行数（减1因为第一列移到垂直表头）
        self.table.setColumnCount(len(TABLE_COLUMNS) - 1)
        self.table.setRowCount(len(PARTS))

        # 设置水平表头（跳过第一列）
        headers = [col[1] for col in TABLE_COLUMNS[1:]]
        self.table.setHorizontalHeaderLabels(headers)

        # 固定表头
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # 表头稍微调高一点
        self.table.horizontalHeader().setFixedHeight(45)

        # 设置垂直表头（显示论文部分名称）
        self.table.verticalHeader().setVisible(True)
        vertical_headers = [part_name for part_name, _ in PARTS]
        self.table.setVerticalHeaderLabels(vertical_headers)
        self.table.verticalHeader().setFixedWidth(150)  # 设置为原第一列的宽度
        self.table.verticalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        # 设置列宽（跳过第一列）
        for col_idx, (_, _, width) in enumerate(TABLE_COLUMNS[1:]):
            self.table.setColumnWidth(col_idx, width)

        # 设置行高（从原来的40改成了52，避免下拉框太挤）
        self.table.verticalHeader().setDefaultSectionSize(52)

        # 填充表格内容
        self._fill_format_table()

        parent_layout.addWidget(self.table)

    def _fill_format_table(self):
        """填充格式表格"""
        default_formats = DEFAULT_FORMATS

        for row_idx, (part_name, part_key) in enumerate(PARTS):
            self.format_vars[part_key] = {}
            self.format_widgets[part_key] = {}

            defaults = default_formats.get(part_key, {})

            # 从第二列开始填充（第一列已经在垂直表头中）
            for col_idx, (field_name, _, _) in enumerate(TABLE_COLUMNS[1:]):
                widget = self._create_format_widget(part_key, field_name, defaults)
                if widget:
                    self.table.setCellWidget(row_idx, col_idx, widget)
                    self.format_widgets[part_key][field_name] = widget

    def _create_format_widget(self, part_key, field_name, defaults):
        """创建格式设置组件"""
        widget = None
        default_value = defaults.get(field_name, "")

        if field_name == "chinese_font":
            widget = TableComboBox(CHINESE_FONTS, default_value)
            self.format_vars[part_key]["chinese_font"] = widget.combo
        elif field_name == "english_font":
            widget = TableComboBox(ENGLISH_FONTS, default_value)
            self.format_vars[part_key]["english_font"] = widget.combo
        elif field_name == "font_size":
            widget = TableComboBox(FONT_SIZE_NAMES, default_value)
            self.format_vars[part_key]["font_size"] = widget.combo
        elif field_name == "alignment":
            widget = TableComboBox(ALIGNMENT_NAMES, default_value)
            self.format_vars[part_key]["alignment"] = widget.combo
        elif field_name == "line_spacing":
            widget = TableComboBox(LINE_SPACING_NAMES, default_value)
            self.format_vars[part_key]["line_spacing"] = widget.combo
        elif field_name == "space_before":
            widget = TableComboBox(SPACINGS, default_value)
            self.format_vars[part_key]["space_before"] = widget.combo
        elif field_name == "space_after":
            widget = TableComboBox(SPACINGS, default_value)
            self.format_vars[part_key]["space_after"] = widget.combo
        elif field_name == "first_line_indent":
            widget = TableComboBox(FIRST_LINE_INDENTS, default_value)
            self.format_vars[part_key]["first_line_indent"] = widget.combo
        elif field_name == "overall_indent":
            widget = TableComboBox(OVERALL_INDENTS, default_value)
            self.format_vars[part_key]["overall_indent"] = widget.combo
        elif field_name == "bold":
            widget = TableCheckBox(default_value)
            self.format_vars[part_key]["bold"] = widget.checkbox
        elif field_name == "italic":
            widget = TableCheckBox(default_value)
            self.format_vars[part_key]["italic"] = widget.checkbox

        return widget

    def create_file_panel(self, parent_layout):
        """创建文件操作面板"""
        parent_layout.setSpacing(12)

        # 文件选择
        self._create_file_input(parent_layout, "选择论文文件", "input_file_entry", "input_file_btn")

        # 输出目录
        self._create_file_input(parent_layout, "输出文件位置", "output_dir_entry", "output_dir_btn", DEFAULT_OUTPUT_DIR_HINT)

        # 自定义文件名
        self._create_file_input(parent_layout, "自定义文件名", "filename_entry", None, DEFAULT_FILENAME_HINT)

        # AI API Key
        self._create_file_input(
            parent_layout,
            "AI API Key（可选）",
            "ai_api_key_entry",
            None,
            "留空则仅规则识别"
        )
        self.ai_api_key_entry.setEchoMode(QLineEdit.Password)
        self.ai_api_key_entry.setClearButtonEnabled(True)

        parent_layout.addSpacing(16)

        # 开始排版按钮
        self.start_btn = QPushButton("开始排版")
        self.start_btn.setMinimumHeight(42)
        self.start_btn.clicked.connect(self.start_formatting)
        parent_layout.addWidget(self.start_btn)

        parent_layout.addSpacing(20)

        # 输出操作按钮
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(10)

        self.open_doc_btn = QPushButton("打开文档")
        self.open_doc_btn.setEnabled(False)
        self.open_doc_btn.setObjectName("action")
        self.open_doc_btn.clicked.connect(self.open_output_document)
        actions_layout.addWidget(self.open_doc_btn)

        self.open_folder_btn = QPushButton("打开文件夹")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.setObjectName("action")
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        actions_layout.addWidget(self.open_folder_btn)

        parent_layout.addLayout(actions_layout)

        parent_layout.addSpacing(20)

        # 进度区域
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setMinimumHeight(4)
        parent_layout.addWidget(self.progress)

        self.progress_text = QLabel("准备就绪，请选择文件开始处理")
        self.progress_text.setStyleSheet("""
            QLabel {
                font-size: 10px;
                color: #9ca3af;
                background-color: transparent;
                border: none;
            }
        """)
        parent_layout.addWidget(self.progress_text)

    def _create_file_input(self, parent_layout, label_text, entry_name, btn_name, placeholder=""):
        """创建文件输入控件 - 清新学术风格"""
        field_layout = QVBoxLayout()
        field_layout.setSpacing(6)
        field_layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel(label_text)
        label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #475569;
                font-weight: 700;
                background-color: transparent;
                border: none;
                letter-spacing: 0.5px;
            }
        """)
        field_layout.addWidget(label)

        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)
        input_layout.setContentsMargins(0, 0, 0, 0)

        entry = QLineEdit()
        entry.setPlaceholderText(placeholder)
        setattr(self, entry_name, entry)
        input_layout.addWidget(entry)

        if btn_name:
            btn = QPushButton("浏览")
            btn.setMinimumWidth(70)
            btn.setMinimumHeight(36)
            btn.setObjectName("browse")
            if btn_name == "input_file_btn":
                btn.clicked.connect(self.select_file)
            elif btn_name == "output_dir_btn":
                btn.clicked.connect(self.select_output)
            setattr(self, btn_name, btn)
            input_layout.addWidget(btn)

        field_layout.addLayout(input_layout)
        parent_layout.addLayout(field_layout)

    def select_file(self):
        """选择输入文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择论文文件",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        if file_path:
            self.input_file_entry.setText(file_path)

    def select_output(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if dir_path:
            self.output_dir_entry.setText(dir_path)

    def get_format_rules(self):
        """获取格式规则"""
        rules = {}
        for part_key, vars_dict in self.format_vars.items():
            rules[part_key] = {
                "chinese_font": vars_dict["chinese_font"].currentText(),
                "english_font": vars_dict["english_font"].currentText(),
                "font_size": vars_dict["font_size"].currentText(),
                "alignment": vars_dict["alignment"].currentText(),
                "line_spacing": vars_dict["line_spacing"].currentText(),
                "space_before": vars_dict["space_before"].currentText(),
                "space_after": vars_dict["space_after"].currentText(),
                "first_line_indent": vars_dict["first_line_indent"].currentText(),
                "overall_indent": vars_dict["overall_indent"].currentText(),
                "bold": vars_dict["bold"].isChecked(),
                "italic": vars_dict["italic"].isChecked(),
            }
        return rules

    def _normalize_output_dir(self, output_dir):
        if not output_dir or output_dir == DEFAULT_OUTPUT_DIR_HINT:
            return ""
        return output_dir

    def _normalize_custom_filename(self, custom_filename):
        if not custom_filename or custom_filename == DEFAULT_FILENAME_HINT:
            return ""
        return custom_filename

    def _build_output_path(self, input_path, output_dir, custom_filename):
        normalized_output_dir = self._normalize_output_dir(output_dir)
        normalized_filename = self._normalize_custom_filename(custom_filename)

        if normalized_output_dir:
            if normalized_filename:
                if not normalized_filename.endswith(".docx"):
                    normalized_filename += ".docx"
                return os.path.join(normalized_output_dir, normalized_filename)
            return os.path.join(normalized_output_dir, os.path.basename(input_path))

        name, ext = os.path.splitext(input_path)
        if normalized_filename:
            if not normalized_filename.endswith(".docx"):
                normalized_filename += ".docx"
            return os.path.join(os.path.dirname(input_path), normalized_filename)
        return f"{name}_formatted{ext}"

    def start_formatting(self):
        """开始排版"""
        input_path = self.input_file_entry.text().strip()
        if not input_path:
            QMessageBox.warning(self, "提示", "请先选择论文文件")
            return

        if not os.path.exists(input_path):
            QMessageBox.critical(self, "错误", "文件不存在")
            return

        output_dir = self.output_dir_entry.text().strip()
        normalized_output_dir = self._normalize_output_dir(output_dir)
        if normalized_output_dir and not os.path.isdir(normalized_output_dir):
            QMessageBox.critical(self, "错误", "输出文件夹不存在，请重新选择。")
            return

        custom_filename = self.filename_entry.text().strip()
        ai_api_key = self.ai_api_key_entry.text().strip()

        task_snapshot = {
            "input_path": input_path,
            "output_dir": output_dir,
            "custom_filename": custom_filename,
            "ai_api_key": ai_api_key,
            "rules": self.get_format_rules(),
            "output_path": self._build_output_path(input_path, output_dir, custom_filename),
        }

        self.last_output_path = None
        self.last_report_path = None
        self.start_btn.setEnabled(False)
        self.open_doc_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)
        self.progress.setValue(0)
        self.progress_text.setText("准备开始处理...")

        # 创建并启动工作线程
        self.format_thread = FormatThread(task_snapshot)
        self.format_thread.progress_update.connect(self._on_progress_update)
        self.format_thread.finished.connect(self._on_format_finished)
        self.format_thread.error.connect(self._on_format_error)
        self.format_thread.start()

    def _on_progress_update(self, value, text):
        """进度更新回调"""
        self.progress.setValue(value)
        self.progress_text.setText(text)

    def _on_format_finished(self, success, output_path, report_path):
        """格式完成回调"""
        self.last_output_path = output_path
        self.last_report_path = report_path

        if success:
            self.progress_text.setText("排版完成！")
        else:
            QMessageBox.warning(
                self,
                "处理完成",
                f"文档已生成，但处理过程中发现问题。\n错误报告已保存到：\n{report_path}",
            )
            self.progress_text.setText(f"处理完成，发现问题，已生成错误报告。")

        self.start_btn.setEnabled(True)
        self.open_doc_btn.setEnabled(True)
        self.open_folder_btn.setEnabled(True)

    def _on_format_error(self, error_msg):
        """格式错误回调"""
        QMessageBox.critical(
            self,
            "错误",
            f"处理失败: {error_msg}",
        )
        self.progress_text.setText("处理失败")
        self.start_btn.setEnabled(True)
        self.open_doc_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)

    def open_output_document(self):
        """打开最后一次输出的文档"""
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            QMessageBox.warning(self, "提示", "暂时没有可打开的输出文档")
            return

        try:
            os.startfile(self.last_output_path)
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"打开文档失败: {exc}")

    def open_output_folder(self):
        """打开最后一次输出文件所在的文件夹"""
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            QMessageBox.warning(self, "提示", "暂时没有可打开的输出文件夹")
            return

        try:
            subprocess.Popen(f'explorer /select,"{os.path.normpath(self.last_output_path)}"')
        except Exception:
            try:
                os.startfile(os.path.dirname(self.last_output_path))
            except Exception as exc:
                QMessageBox.critical(self, "错误", f"打开文件夹失败: {exc}")


def main():
    app = QApplication([])

    # 强制设置全局抗锯齿字体，让文字像图片里一样清晰
    font = app.font()
    font.setFamily("Microsoft YaHei UI")
    font.setPointSize(10)
    app.setFont(font)

    window = ThesisFormatterApp()
    window.show()
    app.exec()
