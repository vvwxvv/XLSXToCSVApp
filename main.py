import sys
import os
import glob
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QCheckBox,
    QMessageBox,
    QFileDialog,
)
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QPixmap

from src.assets.xlsx_to_csv import xlsx_to_csv, convert_folder


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ── Style constants ────────────────────────────────────────────────────────────

BTN_ACTIVE = """
    QPushButton {
        background-color: #4A90E2;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px;
        margin: 5px 10px;
    }
    QPushButton:hover { background-color: #357ABD; }
"""

BTN_NORMAL = """
    QPushButton {
        background-color: #CDEBF0;
        color: black;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px;
        margin: 5px 10px;
    }
    QPushButton:hover { background-color: #BEE0E8; }
"""

CHECKBOX_STYLE = """
    QCheckBox {
        background-color: #CDEBF0;
        color: black;
        font-weight: bold;
        padding: 10px;
        margin: 5px;
        border-radius: 8px;
        border: 2px solid #BEE0E8;
    }
    QCheckBox:hover { background-color: #BEE0E8; }
    QCheckBox::indicator {
        background-color: white;
        border: 2px solid #BEE0E8;
        width: 16px;
        height: 16px;
        border-radius: 3px;
    }
    QCheckBox::indicator:checked {
        background-color: #4A90E2;
        border: 2px solid #4A90E2;
    }
"""


class XLSXToCSVApp(QWidget):
    def __init__(self):
        super().__init__()
        # ── state ──────────────────────────────────────────────────────────────
        self.input_xlsx_file_path = None
        self.input_xlsx_folder_path = None
        self.output_directory = None
        self.use_utf8_bom = True
        self.mode = "file"          # "file" | "folder"
        # ── window ────────────────────────────────────────────────────────────
        self.setMouseTracking(True)
        self.oldPos = self.pos()
        self.init_ui()

    # ── UI setup ───────────────────────────────────────────────────────────────

    def init_ui(self):
        """初始化用户界面。"""
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setObjectName("App")

        self.setStyleSheet(
            """
            QWidget#App {
                font-family: 'Arial';
                background-color: transparent;
                border: 2px solid #CDEBF0;
                border-radius: 20px;
            }
            QWidget#Section {
                background-color: transparent;
                border: none;
            }
            QPushButton {
                background-color: #CDEBF0;
                color: black;
                font-weight: bold;
                border-radius: 8px;
                padding: 10px;
                margin: 10px;
            }
            QPushButton:hover { background-color: #BEE0E8; }
            QLabel#Logo { background-color: transparent; border: none; }
            QMessageBox {
                background-color: #CDEBF0;
                color: black;
                font-size: 16px;
                border: 2px solid #BEE0E8;
                border-radius: 12px;
            }
            QMessageBox QPushButton {
                background-color: #CDEBF0;
                color: black;
                font-weight: bold;
                border: 2px solid #BEE0E8;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
                min-width: 80px;
                min-height: 35px;
            }
            QMessageBox QPushButton:hover { background-color: #BEE0E8; }
            """
        )

        self.mainLayout = QVBoxLayout(self)
        self.mainLayout.setContentsMargins(5, 5, 5, 5)
        self.mainLayout.setSpacing(8)

        self.mainLayout.addLayout(self._create_title_bar())
        self.mainLayout.addWidget(self._create_logo_label())
        self.mainLayout.addLayout(self._create_mode_toggle())

        # ── File section (shown in "file" mode) ───────────────────────────────
        self.file_section = QWidget(self)
        self.file_section.setObjectName("Section")
        fs = QVBoxLayout(self.file_section)
        fs.setContentsMargins(0, 0, 0, 0)
        fs.setSpacing(2)

        self.xlsx_file_button = self._make_button(
            "Select XLSX / XLS File", self.select_xlsx_file
        )
        fs.addWidget(self.xlsx_file_button)

        self.input_file_label = self._make_path_label("No file selected")
        fs.addWidget(self.input_file_label)

        self.mainLayout.addWidget(self.file_section)

        # ── Folder section (shown in "folder" mode) ───────────────────────────
        self.folder_section = QWidget(self)
        self.folder_section.setObjectName("Section")
        fos = QVBoxLayout(self.folder_section)
        fos.setContentsMargins(0, 0, 0, 0)
        fos.setSpacing(2)

        self.xlsx_folder_button = self._make_button(
            "Select XLSX Folder", self.select_xlsx_folder
        )
        fos.addWidget(self.xlsx_folder_button)

        self.input_folder_label = self._make_path_label("No folder selected")
        fos.addWidget(self.input_folder_label)

        self.mainLayout.addWidget(self.folder_section)
        self.folder_section.hide()      # hidden until folder mode activated

        # ── Output directory (shared) ──────────────────────────────────────────
        self.output_dir_button = self._make_button(
            "Select Output Directory", self.select_output_directory
        )
        self.mainLayout.addWidget(self.output_dir_button)

        self.output_dir_label = self._make_path_label("No directory selected")
        self.mainLayout.addWidget(self.output_dir_label)

        # ── Output prefix (shared) ─────────────────────────────────────────────
        self.prefix_input = self._make_line_edit(
            "Output filename prefix  (optional)"
        )
        self.mainLayout.addWidget(self.prefix_input)

        # ── Checkboxes ─────────────────────────────────────────────────────────
        cb_row = QHBoxLayout()

        self.utf8_bom_checkbox = QCheckBox("UTF-8 BOM Encoding", self)
        self.utf8_bom_checkbox.setStyleSheet(CHECKBOX_STYLE)
        self.utf8_bom_checkbox.setChecked(True)
        self.utf8_bom_checkbox.stateChanged.connect(self._toggle_utf8_bom)
        cb_row.addWidget(self.utf8_bom_checkbox)

        self.mainLayout.addLayout(cb_row)

        # ── Start button ───────────────────────────────────────────────────────
        self.start_button = self._make_button(
            "▶   Start Converting", self.run_workflow
        )
        self.mainLayout.addWidget(self.start_button)

        self.setLayout(self.mainLayout)
        self.resize(540, 820)
        self._refresh_mode_buttons()

    # ── Factory helpers ────────────────────────────────────────────────────────

    def _make_button(self, text, slot, style=None):
        """创建按钮并设置点击事件。"""
        btn = QPushButton(text, self)
        btn.clicked.connect(slot)
        btn.setStyleSheet(style or BTN_NORMAL)
        return btn

    def _make_line_edit(self, placeholder, style=None):
        """创建输入框。"""
        le = QLineEdit(self)
        le.setPlaceholderText(placeholder)
        le.setStyleSheet(style or """
            QLineEdit {
                border: 2px solid #ccc;
                border-radius: 8px;
                padding: 8px;
                margin: 5px 10px;
                color: black;
                background-color: white;
            }
        """)
        return le

    def _make_path_label(self, text):
        """创建路径显示标签。"""
        lbl = QLabel(text, self)
        lbl.setWordWrap(True)
        lbl.setObjectName("Section")
        lbl.setStyleSheet("""
            QLabel {
                color: #555;
                font-size: 11px;
                padding: 2px 18px 6px 18px;
                border: none;
                background-color: transparent;
            }
        """)
        return lbl

    def _create_title_bar(self):
        """创建标题栏。"""
        row = QHBoxLayout()

        title = QLabel("XLSX → CSV Converter", self)
        title.setObjectName("Section")
        title.setStyleSheet("""
            QLabel {
                color: #333;
                font-size: 14px;
                font-weight: bold;
                border: none;
                background-color: transparent;
                padding-left: 10px;
            }
        """)
        row.addWidget(title, alignment=Qt.AlignLeft)

        close_btn = QPushButton("✕", self)
        close_btn.setFixedSize(30, 30)
        close_btn.clicked.connect(self.close)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF6B6B;
                color: white;
                font-weight: bold;
                border-radius: 15px;
                padding: 0px;
                margin: 5px;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #FF4444; }
        """)
        row.addWidget(close_btn, alignment=Qt.AlignRight)
        return row

    def _create_logo_label(self):
        """创建 Logo 标签。"""
        logo = QLabel(self)
        cover_path = get_resource_path(os.path.join("static", "cover.png"))
        pixmap = QPixmap(cover_path).scaled(500, 800)
        logo.setPixmap(pixmap)
        logo.setAlignment(Qt.AlignCenter)
        logo.setObjectName("Logo")
        return logo

    def _create_mode_toggle(self):
        """创建模式切换行 (单文件 / 文件夹)。"""
        row = QHBoxLayout()
        row.setContentsMargins(10, 4, 10, 4)

        self.file_mode_btn = QPushButton("■ - Single File", self)
        self.file_mode_btn.clicked.connect(lambda: self._set_mode("file"))

        self.folder_mode_btn = QPushButton("■ - Folder Mode", self)
        self.folder_mode_btn.clicked.connect(lambda: self._set_mode("folder"))

        row.addWidget(self.file_mode_btn)
        row.addWidget(self.folder_mode_btn)
        return row

    # ── Mode management ────────────────────────────────────────────────────────

    def _set_mode(self, mode):
        """切换模式并刷新界面。"""
        self.mode = mode
        if mode == "file":
            self.file_section.show()
            self.folder_section.hide()
        else:
            self.file_section.hide()
            self.folder_section.show()
        self._refresh_mode_buttons()

    def _refresh_mode_buttons(self):
        """根据当前模式高亮对应按钮。"""
        if self.mode == "file":
            self.file_mode_btn.setStyleSheet(BTN_ACTIVE)
            self.folder_mode_btn.setStyleSheet(BTN_NORMAL)
        else:
            self.file_mode_btn.setStyleSheet(BTN_NORMAL)
            self.folder_mode_btn.setStyleSheet(BTN_ACTIVE)

    # ── Slot handlers ──────────────────────────────────────────────────────────

    def select_xlsx_file(self):
        """选择单个 Excel 文件。"""
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.input_xlsx_file_path = path
            self.input_file_label.setText(
                f"📄  {os.path.basename(path)}\n    {path}"
            )

    def select_xlsx_folder(self):
        """选择包含 Excel 文件的文件夹。"""
        path = QFileDialog.getExistingDirectory(
            self, "Select Folder Containing XLSX Files"
        )
        if path:
            self.input_xlsx_folder_path = path
            # Count how many Excel files are inside for immediate feedback
            found = (
                glob.glob(os.path.join(path, "*.xlsx"))
                + glob.glob(os.path.join(path, "*.xls"))
            )
            count = len(found)
            color = "#2E7D32" if count > 0 else "#C62828"
            self.input_folder_label.setText(
                f"📁  {path}"
            )
            # Inline count badge
            badge_text = (
                f"   ✅  {count} Excel file(s) ready to convert"
                if count > 0
                else "   ⚠️  No Excel files found in this folder"
            )
            self.input_folder_label.setText(
                f"📁  {path}\n{badge_text}"
            )

    def select_output_directory(self):
        """选择输出目录。"""
        path = QFileDialog.getExistingDirectory(
            self, "Select Output Directory"
        )
        if path:
            self.output_directory = path
            self.output_dir_label.setText(f"📁  {path}")

    def _toggle_utf8_bom(self, state):
        """切换 UTF-8 BOM 编码。"""
        self.use_utf8_bom = state == Qt.Checked

    # ── Workflow ───────────────────────────────────────────────────────────────

    def run_workflow(self):
        """根据当前模式运行转换工作流程。"""
        if not self.output_directory:
            QMessageBox.warning(self, "Error", "Please select an output directory.")
            return

        encoding = "utf-8-sig" if self.use_utf8_bom else "utf-8"
        prefix = self.prefix_input.text().strip()

        if self.mode == "file":
            self._run_single_file(encoding, prefix)
        else:
            self._run_folder(encoding, prefix)

    def _run_single_file(self, encoding, prefix):
        """转换单个 Excel 文件。"""
        if not self.input_xlsx_file_path:
            QMessageBox.warning(self, "Error", "Please select an Excel file.")
            return

        base = os.path.splitext(os.path.basename(self.input_xlsx_file_path))[0]
        out_name = f"{prefix}{base}" if prefix else base
        out_path = os.path.join(self.output_directory, f"{out_name}.csv")

        try:
            results = xlsx_to_csv(
                self.input_xlsx_file_path, out_path, encoding=encoding
            )
            if not results:
                QMessageBox.warning(
                    self, "Error",
                    "Conversion failed. Check the console for details."
                )
                return

            lines = "\n".join(
                f"  • {os.path.basename(p)}  ({r} rows × {c} cols)"
                for p, r, c in results
            )
            QMessageBox.information(
                self, "Success",
                f"Conversion complete!\n\nExported {len(results)} file(s):\n{lines}",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred:\n{str(e)}")

    def _run_folder(self, encoding, prefix):
        """转换文件夹中的全部 Excel 文件。"""
        if not self.input_xlsx_folder_path:
            QMessageBox.warning(self, "Error", "Please select an XLSX folder.")
            return

        try:
            results, success, total = convert_folder(
                self.input_xlsx_folder_path,
                output_dir=self.output_directory,
                prefix=prefix,
                encoding=encoding,
            )

            if total == 0:
                QMessageBox.warning(
                    self, "No Files Found",
                    f"No Excel files (.xlsx / .xls) were found in:\n"
                    f"{self.input_xlsx_folder_path}",
                )
                return

            lines = "\n".join(
                f"  • {os.path.basename(p)}  ({r} rows × {c} cols)"
                for p, r, c in results
            )
            failed = total - success
            summary = (
                f"Converted  {success} / {total}  Excel file(s)\n"
                f"Exported   {len(results)}  CSV file(s)"
            )
            if failed:
                summary += f"\n⚠️  {failed} file(s) failed — check console for details."

            QMessageBox.information(
                self, "Folder Conversion Complete",
                f"{summary}\n\n{lines}",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred:\n{str(e)}")

    # ── Drag to move ───────────────────────────────────────────────────────────

    def mousePressEvent(self, event):
        """鼠标按下事件。"""
        if event.button() == Qt.LeftButton:
            self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        """鼠标移动事件。"""
        if event.buttons() == Qt.LeftButton:
            delta = QPoint(event.globalPos() - self.oldPos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.oldPos = event.globalPos()

    def closeEvent(self, event):
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = XLSXToCSVApp()
    window.show()
    sys.exit(app.exec_())