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
    DEFAULT_OUTPUT_DIR_HINT, DEFAULT_FILENAME_HINT,
    AI_MODEL_OPTIONS, DEFAULT_AI_MODEL
)
from services.doc_service import DocumentProcessingService


def _asset_path(*parts: str) -> str:
    return os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        *parts,
    ).replace("\\", "/")


def build_combobox_stylesheet(font_size: int = 13, compact: bool = False) -> str:
    arrow_path = _asset_path("assets", "icons", "chevron_down_12.svg")
    control_padding = "3px 18px 3px 6px" if compact else "6px 28px 6px 10px"
    drop_down_width = "18px" if compact else "24px"
    arrow_right_margin = "4px" if compact else "8px"
    item_height = "28px" if compact else "30px"
    item_padding = "3px 10px" if compact else "4px 10px"

    qss = """
        QComboBox {
            background-color: #FFFFFF;
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            padding: __CONTROL_PADDING__;
            color: #334155;
            font-family: "Microsoft YaHei UI";
            font-size: __FONT_SIZE__px;
        }
        QComboBox:hover {
            border: 1px solid #93C5FD;
            background-color: #F8FAFC;
        }
        QComboBox:focus {
            border: 2px solid #60A5FA;
            background-color: #FFFFFF;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: __DROP_DOWN_WIDTH__;
            border: none;
            background: transparent;
            margin: 0px;
            padding: 0px;
        }
        QComboBox::down-arrow {
            image: url("__ARROW_PATH__");
            width: 12px;
            height: 12px;
            margin-right: __ARROW_RIGHT_MARGIN__;
        }
        QComboBox QAbstractItemView {
            border: none;
            background-color: #FFFFFF;
            color: #334155;
            outline: 0;
            selection-background-color: #7A7A7A;
            selection-color: #FFFFFF;
            padding: 4px 0px;
        }
        QComboBox QAbstractItemView:focus {
            outline: none;
        }
        QComboBox QAbstractItemView::item {
            min-height: __ITEM_HEIGHT__;
            padding: __ITEM_PADDING__;
            border: none;
            color: #334155;
            background-color: transparent;
        }
        QComboBox QAbstractItemView::item:selected {
            background-color: #7A7A7A;
            color: #FFFFFF;
        }
        QComboBox QAbstractItemView::item:hover:!selected {
            background-color: #F1F5F9;
        }
    """
    return (
        qss.replace("__ARROW_PATH__", arrow_path)
        .replace("__CONTROL_PADDING__", control_padding)
        .replace("__DROP_DOWN_WIDTH__", drop_down_width)
        .replace("__ARROW_RIGHT_MARGIN__", arrow_right_margin)
        .replace("__FONT_SIZE__", str(font_size))
        .replace("__ITEM_HEIGHT__", item_height)
        .replace("__ITEM_PADDING__", item_padding)
    )


# ==================== 自定义组合框和复选框 widget ====================

class TableComboBox(QWidget):
    """表格中使用的组合框"""
    def __init__(self, items, default_item="", parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)

        self.combo = QComboBox()
        self.combo.addItems(items)
        if default_item in items:
            self.combo.setCurrentText(default_item)

        self.combo.setStyleSheet(build_combobox_stylesheet(font_size=13, compact=True))

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
        check_icon_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "assets",
            "icons",
            "check_white_16.svg",
        ).replace("\\", "/")

        checkbox_qss = """
            QCheckBox { spacing: 0px; }
            QCheckBox::indicator {
                width: 16px; height: 16px;
                border: 1px solid #CBD5E1; border-radius: 4px; background-color: #FFFFFF;
            }
            QCheckBox::indicator:checked {
                background-color: #3B82F6;
                border: 1px solid #3B82F6;
                image: url("__CHECK_ICON_PATH__");
            }
        """.replace("__CHECK_ICON_PATH__", check_icon_path)
        self.checkbox.setStyleSheet(checkbox_qss)
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


class BaseFormatterApp(QMainWindow):
    """论文格式助手主程序"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("论文格式助手 v2.0")
        self.resize(1360, 860)
        self.setMinimumSize(1220, 780)

        # 设置应用样式
        self.setup_styles()

        # 存储格式规则
        self.format_vars = {}
        self.format_widgets = {}
        self.ai_model_combo = None

        # 存储文件路径
        self.last_output_path = None
        self.last_report_path = None

        # 创建UI
        self.create_ui()

    def setup_styles(self):
        raise NotImplementedError("UI style should be implemented by subclass.")

    def create_ui(self):
        raise NotImplementedError("UI layout should be implemented by subclass.")

    def create_header(self, parent_layout):
        raise NotImplementedError("Header UI should be implemented by subclass.")

    def create_card(self, title, header_object_name):
        raise NotImplementedError("Card UI should be implemented by subclass.")

    def create_format_panel(self, parent_layout):
        raise NotImplementedError("Format panel UI should be implemented by subclass.")

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
        raise NotImplementedError("File panel UI should be implemented by subclass.")

    def _create_file_input(self, parent_layout, label_text, entry_name, btn_name, placeholder=""):
        raise NotImplementedError("File input UI should be implemented by subclass.")

    def _create_ai_model_selector(self, parent_layout):
        """创建 AI 模型选择下拉框（仅可选，不可手输）"""
        field_layout = QVBoxLayout()
        field_layout.setSpacing(6)
        field_layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel("AI 模型")
        label.setObjectName("FieldLabel")
        field_layout.addWidget(label)

        self.ai_model_combo = QComboBox()
        self.ai_model_combo.addItems(AI_MODEL_OPTIONS)
        self.ai_model_combo.setEditable(False)
        self.ai_model_combo.setCurrentText(DEFAULT_AI_MODEL)
        self.ai_model_combo.setMinimumHeight(36)
        self.ai_model_combo.setStyleSheet(build_combobox_stylesheet(font_size=13, compact=False))
        field_layout.addWidget(self.ai_model_combo)

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
        input_dir = os.path.dirname(input_path)
        input_basename = os.path.basename(input_path)
        input_stem, input_ext = os.path.splitext(input_basename)

        if normalized_filename:
            if not normalized_filename.endswith(".docx"):
                normalized_filename += ".docx"
            target_dir = normalized_output_dir or input_dir
            return os.path.join(target_dir, normalized_filename)

        default_filename = f"{input_stem}_formatted{input_ext}"
        target_dir = normalized_output_dir or input_dir
        return os.path.join(target_dir, default_filename)

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
        ai_model = DEFAULT_AI_MODEL
        if self.ai_model_combo is not None:
            ai_model = (self.ai_model_combo.currentText() or DEFAULT_AI_MODEL).strip() or DEFAULT_AI_MODEL

        task_snapshot = {
            "input_path": input_path,
            "output_dir": output_dir,
            "custom_filename": custom_filename,
            "ai_api_key": ai_api_key,
            "ai_model": ai_model,
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


class ThesisFormatterApp(BaseFormatterApp):
    """新版 UI：保持业务逻辑不变，仅重绘界面层。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("论文格式助手 v2.0")

    def setup_styles(self):
        QApplication.setStyle(QStyleFactory.create("Fusion"))

        stylesheet = """
            QMainWindow {
                background-color: #F3F5F9;
            }

            QWidget {
                color: #344256;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
                font-size: 13px;
            }

            QFrame#TopBar {
                background-color: #FFFFFF;
                border: none;
                border-bottom: 1px solid #E6EBF2;
            }

            QLabel#TopBarIcon {
                background-color: #2F6CF5;
                color: #FFFFFF;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 700;
                padding: 6px;
            }

            QLabel#TopBarTitle {
                color: #0F1E3A;
                font-size: 20px;
                font-weight: 700;
            }

            QLabel#TopBarVersion {
                color: #6F7D94;
                font-size: 12px;
            }

            QFrame#Card {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F1;
                border-radius: 14px;
            }

            QFrame#FormatCardHeader {
                background-color: #EDF2FF;
                border: none;
                border-top-left-radius: 14px;
                border-top-right-radius: 14px;
                border-bottom: 1px solid #E2E8F1;
            }

            QFrame#FileCardHeader {
                background-color: #F2EFFC;
                border: none;
                border-top-left-radius: 14px;
                border-top-right-radius: 14px;
                border-bottom: 1px solid #E2E8F1;
            }

            QLabel#CardTitle {
                color: #4050BC;
                font-size: 16px;
                font-weight: 700;
            }

            QLabel#FieldLabel {
                color: #3C4A63;
                font-size: 14px;
                font-weight: 700;
            }

            QLineEdit#FieldInput {
                background-color: #F7F9FC;
                border: 1px solid #DDE3EF;
                border-radius: 6px;
                padding: 8px 12px;
                color: #334155;
            }
            QLineEdit#FieldInput:focus {
                border: 2px solid #4D7CF8;
                background-color: #FFFFFF;
            }

            QPushButton#PrimaryButton {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3563EA, stop:1 #4A3DE4
                );
                color: #FFFFFF;
                border: none;
                border-radius: 9px;
                padding: 10px;
                font-weight: 700;
                font-size: 16px;
            }
            QPushButton#PrimaryButton:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3E71F3, stop:1 #5A4EF0
                );
            }
            QPushButton#PrimaryButton:disabled {
                background: #C5D0EA;
                color: #F4F7FF;
            }

            QPushButton#BrowseButton {
                background-color: #4D7CF8;
                color: #FFFFFF;
                border-radius: 6px;
                border: none;
                font-weight: 700;
                font-size: 14px;
                padding: 6px 16px;
            }
            QPushButton#BrowseButton:hover {
                background-color: #6690FB;
            }

            QPushButton#ActionButton {
                background-color: #FFFFFF;
                color: #24344E;
                border: 1px solid #DCE3EF;
                border-radius: 6px;
                font-weight: 700;
                font-size: 14px;
                padding: 8px 12px;
            }
            QPushButton#ActionButton:hover {
                background-color: #F7F9FD;
            }
            QPushButton#ActionButton:disabled {
                color: #9AA7BC;
                border: 1px solid #E6EBF2;
                background-color: #FBFCFF;
            }

            QTableWidget#RulesTable {
                background-color: #FFFFFF;
                border: 1px solid #E8EDF4;
                border-radius: 8px;
                gridline-color: #EEF2F8;
                alternate-background-color: #FBFCFF;
                selection-background-color: #EFF6FF;
                selection-color: #1F3152;
            }
            QTableWidget#RulesTable::item {
                border-bottom: 1px solid #EEF2F8;
                padding-left: 6px;
                padding-right: 6px;
            }

            QHeaderView::section {
                background-color: #F2F4F8;
                color: #5D6880;
                font-weight: 600;
                border: none;
                border-right: 1px solid #E8EDF4;
                border-bottom: 1px solid #E8EDF4;
                padding: 9px 8px;
            }

            QTableWidget QTableCornerButton::section {
                background-color: #F2F4F8;
                border: none;
                border-right: 1px solid #E8EDF4;
                border-bottom: 1px solid #E8EDF4;
            }

            QProgressBar#ProgressBar {
                border: none;
                background-color: #E5EBF3;
                border-radius: 4px;
                height: 8px;
            }
            QProgressBar#ProgressBar::chunk {
                background-color: #3C6FF5;
                border-radius: 4px;
            }

            QLabel#ProgressHint {
                color: #8A96AA;
                background-color: transparent;
                border: none;
                font-size: 12px;
                padding: 0px;
            }

            QScrollBar:vertical {
                border: none;
                background: #F4F7FC;
                width: 10px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #C4CDDC;
                min-height: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #A3AEC2;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }

            QScrollBar:horizontal {
                border: none;
                background: #F4F7FC;
                height: 10px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal {
                background: #C4CDDC;
                min-width: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #A3AEC2;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
        """
        self.setStyleSheet(stylesheet)

    def create_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setSpacing(22)
        content_layout.setContentsMargins(24, 20, 24, 20)

        format_card = self.create_card("格式规则设置", "FormatCardHeader")
        self.create_format_panel(format_card.body_layout)
        content_layout.addWidget(format_card, 66)

        file_card = self.create_card("文件操作", "FileCardHeader")
        file_card.setMinimumWidth(440)
        self.create_file_panel(file_card.body_layout)
        content_layout.addWidget(file_card, 34)

        main_layout.addWidget(content_widget, 1)

    def create_header(self, parent_layout):
        header_frame = QFrame()
        header_frame.setObjectName("TopBar")
        header_frame.setFixedHeight(82)

        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(24, 12, 24, 12)
        header_layout.setSpacing(10)

        icon_label = QLabel("文")
        icon_label.setObjectName("TopBarIcon")
        icon_label.setFixedSize(32, 32)
        icon_label.setAlignment(Qt.AlignCenter)

        title_layout = QVBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(2)

        title_label = QLabel("论文格式助手")
        title_label.setObjectName("TopBarTitle")
        version_label = QLabel("v1.0")
        version_label.setObjectName("TopBarVersion")

        title_layout.addWidget(title_label)
        title_layout.addWidget(version_label)

        header_layout.addWidget(icon_label)
        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        parent_layout.addWidget(header_frame)

    def create_card(self, title, header_object_name):
        card = QFrame()
        card.setObjectName("Card")

        outer_layout = QVBoxLayout(card)
        outer_layout.setSpacing(0)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        header = QFrame()
        header.setObjectName(header_object_name)
        header.setMinimumHeight(74)

        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(22, 0, 22, 0)
        header_layout.setSpacing(0)

        title_label = QLabel(title)
        title_label.setObjectName("CardTitle")
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        body_widget = QWidget()
        body_layout = QVBoxLayout(body_widget)
        body_layout.setSpacing(12)
        body_layout.setContentsMargins(22, 18, 22, 18)

        outer_layout.addWidget(header)
        outer_layout.addWidget(body_widget, 1)
        card.body_layout = body_layout
        return card

    def create_format_panel(self, parent_layout):
        self.table = QTableWidget()
        self.table.setObjectName("RulesTable")
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.setAlternatingRowColors(True)

        self.table.setColumnCount(len(TABLE_COLUMNS) - 1)
        self.table.setRowCount(len(PARTS))

        headers = [col[1] for col in TABLE_COLUMNS[1:]]
        self.table.setHorizontalHeaderLabels(headers)

        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setFixedHeight(45)

        self.table.verticalHeader().setVisible(True)
        vertical_headers = [part_name for part_name, _ in PARTS]
        self.table.setVerticalHeaderLabels(vertical_headers)
        self.table.verticalHeader().setFixedWidth(150)
        self.table.verticalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        for col_idx, (_, _, width) in enumerate(TABLE_COLUMNS[1:]):
            self.table.setColumnWidth(col_idx, width)

        self.table.verticalHeader().setDefaultSectionSize(50)
        self._fill_format_table()
        parent_layout.addWidget(self.table)

    def create_file_panel(self, parent_layout):
        parent_layout.setSpacing(12)

        self._create_file_input(parent_layout, "选择论文文件", "input_file_entry", "input_file_btn")
        self._create_file_input(parent_layout, "输出文件位置", "output_dir_entry", "output_dir_btn", DEFAULT_OUTPUT_DIR_HINT)
        self._create_file_input(parent_layout, "自定义文件名", "filename_entry", None, DEFAULT_FILENAME_HINT)
        self._create_file_input(parent_layout, "AI API Key (可选)", "ai_api_key_entry", None, "留空则仅规则识别")

        self.ai_api_key_entry.setEchoMode(QLineEdit.Password)
        self.ai_api_key_entry.setClearButtonEnabled(True)
        self._create_ai_model_selector(parent_layout)

        parent_layout.addSpacing(16)

        self.start_btn = QPushButton("开始排版")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.setMinimumHeight(48)
        self.start_btn.clicked.connect(self.start_formatting)
        parent_layout.addWidget(self.start_btn)

        parent_layout.addSpacing(16)

        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(10)

        self.open_doc_btn = QPushButton("打开文档")
        self.open_doc_btn.setEnabled(False)
        self.open_doc_btn.setObjectName("ActionButton")
        self.open_doc_btn.setMinimumHeight(40)
        self.open_doc_btn.clicked.connect(self.open_output_document)
        actions_layout.addWidget(self.open_doc_btn)

        self.open_folder_btn = QPushButton("打开文件夹")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.setObjectName("ActionButton")
        self.open_folder_btn.setMinimumHeight(40)
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        actions_layout.addWidget(self.open_folder_btn)

        parent_layout.addLayout(actions_layout)
        parent_layout.addSpacing(16)

        self.progress = QProgressBar()
        self.progress.setObjectName("ProgressBar")
        self.progress.setTextVisible(False)
        self.progress.setMinimumHeight(8)
        parent_layout.addWidget(self.progress)
        parent_layout.addSpacing(4)

        self.progress_text = QLabel("请选择文档，调整文件格式或排版")
        self.progress_text.setObjectName("ProgressHint")
        self.progress_text.setText("请选择文档")
        self.progress_text.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.progress_text.setMinimumHeight(18)
        parent_layout.addWidget(self.progress_text)

    def _create_file_input(self, parent_layout, label_text, entry_name, btn_name, placeholder=""):
        field_layout = QVBoxLayout()
        field_layout.setSpacing(6)
        field_layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel(label_text)
        label.setObjectName("FieldLabel")
        field_layout.addWidget(label)

        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)
        input_layout.setContentsMargins(0, 0, 0, 0)

        entry = QLineEdit()
        entry.setObjectName("FieldInput")
        entry.setMinimumHeight(36)
        entry.setPlaceholderText(placeholder)
        setattr(self, entry_name, entry)
        input_layout.addWidget(entry)

        if btn_name:
            btn = QPushButton("浏览")
            btn.setMinimumWidth(74)
            btn.setMinimumHeight(36)
            btn.setObjectName("BrowseButton")
            if btn_name == "input_file_btn":
                btn.clicked.connect(self.select_file)
            elif btn_name == "output_dir_btn":
                btn.clicked.connect(self.select_output)
            setattr(self, btn_name, btn)
            input_layout.addWidget(btn)

        field_layout.addLayout(input_layout)
        parent_layout.addLayout(field_layout)


def main():
    app = QApplication([])

    font = app.font()
    font.setFamily("Microsoft YaHei UI")
    font.setPointSize(10)
    app.setFont(font)

    window = ThesisFormatterApp()
    window.show()
    app.exec()
