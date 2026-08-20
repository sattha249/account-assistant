import sys
import os
import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QLineEdit, QTextEdit, QFileDialog,
    QMessageBox, QFrame, QGroupBox
)
from PyQt6.QtCore import Qt, QRectF
from PyQt6.QtGui import QFont, QIcon, QPixmap, QPainter, QPainterPath, QColor
import excel_processor


def get_rounded_pixmap(image_path: str, size: int = 76, corner_radius: int = 14) -> QPixmap:
    """
    Crops image to square center and paints rounded rectangle with anti-aliasing.
    """
    if not os.path.exists(image_path):
        return QPixmap()
    pixmap = QPixmap(image_path)
    if pixmap.isNull():
        return QPixmap()

    w, h = pixmap.width(), pixmap.height()
    min_dim = min(w, h)
    crop_rect = QRectF((w - min_dim) / 2, (h - min_dim) / 2, min_dim, min_dim)
    cropped = pixmap.copy(crop_rect.toRect())
    scaled = cropped.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

    out_pixmap = QPixmap(size, size)
    out_pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(out_pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(0, 0, size, size, corner_radius, corner_radius)
    painter.setClipPath(path)
    painter.drawPixmap(0, 0, scaled)
    painter.end()

    return out_pixmap


class AccountAssistantGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ระบบผู้ช่วยบัญชีและฟาร์มสุกร (Account Assistant - Signature Edition)")
        self.resize(780, 600)
        self.setMinimumSize(720, 540)

        # Set Window Icon to Desktop Wallpaper Signature Logo
        logo_path = "signature_logo.jpg"
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))

        # Apply Modern Dark Theme Stylesheet (QSS)
        self.apply_dark_theme()

        # Main Widget and Layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # 1. Prominent Signature Header Banner
        header_frame = QFrame()
        header_frame.setObjectName("HeaderFrame")
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(18, 14, 18, 14)
        header_layout.setSpacing(16)

        # Signature Logo Image Label (Enlarged)
        logo_size = 115
        self.logo_label = QLabel()
        self.logo_label.setObjectName("LogoLabel")
        rounded_logo = get_rounded_pixmap(logo_path, size=logo_size, corner_radius=20)
        if not rounded_logo.isNull():
            self.logo_label.setPixmap(rounded_logo)
            self.logo_label.setFixedSize(logo_size, logo_size)
        else:
            self.logo_label.setText("🐷")
            self.logo_label.setFont(QFont("Calibri", 42))
            self.logo_label.setFixedSize(logo_size, logo_size)

        # Header Text Box
        header_text_box = QVBoxLayout()
        header_text_box.setSpacing(6)
        header_text_box.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        title_label = QLabel("🐷 Pig Farm Accounting System")
        title_label.setObjectName("HeaderTitle")
        title_label.setFont(QFont("Calibri", 20, QFont.Weight.Bold))

        subtitle_label = QLabel("ระบบสรุปยอดขายสุกร และสรุปยอดแยกเล้าอัตโนมัติ (Desktop GUI)")
        subtitle_label.setObjectName("HeaderSubtitle")
        subtitle_label.setFont(QFont("Calibri", 12))

        header_text_box.addWidget(title_label)
        header_text_box.addWidget(subtitle_label)

        header_layout.addWidget(self.logo_label)
        header_layout.addLayout(header_text_box)
        header_layout.addStretch()

        layout.addWidget(header_frame)

        # 2. File Selection Card
        file_group = QGroupBox("📂 เลือกไฟล์ Excel (.xlsx)")
        file_layout = QVBoxLayout(file_group)
        file_layout.setSpacing(10)

        # Dropdown row
        combo_layout = QHBoxLayout()
        lbl_combo = QLabel("ไฟล์ในโฟลเดอร์ปัจจุบัน:")
        lbl_combo.setFixedWidth(150)
        self.file_combo = QComboBox()
        self.file_combo.currentIndexChanged.connect(self.on_combo_file_selected)

        btn_refresh = QPushButton("🔄 รีเฟรช")
        btn_refresh.setFixedWidth(90)
        btn_refresh.clicked.connect(self.refresh_file_list)

        combo_layout.addWidget(lbl_combo)
        combo_layout.addWidget(self.file_combo)
        combo_layout.addWidget(btn_refresh)
        file_layout.addLayout(combo_layout)

        # File Path Input & Browse Row
        path_layout = QHBoxLayout()
        lbl_path = QLabel("Path ไฟล์ที่เลือก:")
        lbl_path.setFixedWidth(150)

        self.txt_file_path = QLineEdit()
        self.txt_file_path.setPlaceholderText("กรุณาเลือกไฟล์ Excel (.xlsx) ที่ต้องการประมวลผล...")

        btn_browse = QPushButton("📁 Browse...")
        btn_browse.setFixedWidth(100)
        btn_browse.clicked.connect(self.browse_file)

        path_layout.addWidget(lbl_path)
        path_layout.addWidget(self.txt_file_path)
        path_layout.addWidget(btn_browse)
        file_layout.addLayout(path_layout)

        layout.addWidget(file_group)

        # 3. Action Buttons Card
        action_group = QGroupBox("⚙️ เลือกรายการประมวลผล")
        action_layout = QHBoxLayout(action_group)
        action_layout.setSpacing(15)

        self.btn_menu1 = QPushButton("📊  สร้างชีทสรุปยอดสุกร (สรุปยอด)")
        self.btn_menu1.setObjectName("BtnMenu1")
        self.btn_menu1.setFixedHeight(46)
        self.btn_menu1.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        self.btn_menu1.clicked.connect(self.process_menu_1)

        self.btn_menu2 = QPushButton("🏠  สร้างชีทสรุปยอดสุกรแยกเล้า (สรุปยอดแยกเล้า)")
        self.btn_menu2.setObjectName("BtnMenu2")
        self.btn_menu2.setFixedHeight(46)
        self.btn_menu2.setFont(QFont("Calibri", 11, QFont.Weight.Bold))
        self.btn_menu2.clicked.connect(self.process_menu_2)

        action_layout.addWidget(self.btn_menu1)
        action_layout.addWidget(self.btn_menu2)
        layout.addWidget(action_group)

        # 4. Status & Console Output Card
        output_group = QGroupBox("📋 รายงานสถานะการทำงาน (Log Output)")
        output_layout = QVBoxLayout(output_group)

        self.txt_output = QTextEdit()
        self.txt_output.setReadOnly(True)
        self.txt_output.setObjectName("ConsoleOutput")
        output_layout.addWidget(self.txt_output)

        layout.addWidget(output_group)

        # Footer Label
        footer_lbl = QLabel("Account Assistant GUI v2.0 • Signature Desktop Edition")
        footer_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_lbl.setStyleSheet("color: #6C7086; font-size: 11px;")
        layout.addWidget(footer_lbl)

        # Load initial file list
        self.refresh_file_list()
        self.log("ยินดีต้อนรับสู่ระบบผู้ช่วยบัญชีและฟาร์มสุกร (Signature Edition)")
        self.log("กรุณาเลือกไฟล์ Excel และกดปุ่มประมวลผลรายการที่ต้องการ")

    def apply_dark_theme(self):
        """Applies a clean, modern dark CSS palette to the application."""
        qss = """
        QMainWindow {
            background-color: #1E1E2E;
        }
        QWidget {
            color: #CDD6F4;
            font-family: "Calibri", "Segoe UI", sans-serif;
        }
        QFrame#HeaderFrame {
            background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #18365B, stop:1 #2F5597);
            border-radius: 12px;
            border: 1px solid #45475A;
        }
        QLabel#LogoLabel {
            border: 2px solid #89B4FA;
            border-radius: 20px;
            background-color: #11111B;
        }
        QLabel#HeaderTitle {
            color: #FFFFFF;
        }
        QLabel#HeaderSubtitle {
            color: #D9E1F2;
        }
        QGroupBox {
            font-weight: bold;
            font-size: 13px;
            color: #89B4FA;
            border: 1px solid #45475A;
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 14px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 8px;
            background-color: #1E1E2E;
        }
        QLineEdit, QComboBox, QTextEdit {
            background-color: #11111B;
            border: 1px solid #45475A;
            border-radius: 6px;
            padding: 6px 10px;
            color: #F5E0DC;
            font-size: 12px;
        }
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus {
            border: 1px solid #89B4FA;
        }
        QComboBox::drop-down {
            border: none;
            width: 24px;
        }
        QPushButton {
            background-color: #313244;
            color: #CDD6F4;
            border: 1px solid #45475A;
            border-radius: 6px;
            padding: 6px 14px;
            font-size: 12px;
        }
        QPushButton:hover {
            background-color: #45475A;
            color: #FFFFFF;
        }
        QPushButton#BtnMenu1 {
            background-color: #1E66F5;
            color: #FFFFFF;
            border: none;
            border-radius: 6px;
        }
        QPushButton#BtnMenu1:hover {
            background-color: #2A72FF;
        }
        QPushButton#BtnMenu2 {
            background-color: #179299;
            color: #FFFFFF;
            border: none;
            border-radius: 6px;
        }
        QPushButton#BtnMenu2:hover {
            background-color: #20A6AE;
        }
        QTextEdit#ConsoleOutput {
            background-color: #11111B;
            color: #A6E3A1;
            font-family: "Courier New", monospace;
            font-size: 12px;
        }
        """
        self.setStyleSheet(qss)

    def log(self, message: str):
        """Appends a timestamped log message to the text output area."""
        now = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_output.append(f"[{now}] {message}")

    def refresh_file_list(self):
        """Finds all .xlsx files in CWD and populates the dropdown."""
        self.file_combo.clear()
        files = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~") and not f.startswith(".~")]
        files.sort()

        if files:
            self.file_combo.addItems(files)
            self.txt_file_path.setText(files[0])
            self.log(f"พบไฟล์ Excel ในโฟลเดอร์ปัจจุบัน {len(files)} ไฟล์")
        else:
            self.file_combo.addItem("ไม่พบไฟล์ .xlsx ในโฟลเดอร์นี้")
            self.txt_file_path.clear()
            self.log("ไม่พบไฟล์ .xlsx ในโฟลเดอร์ปัจจุบัน")

    def on_combo_file_selected(self, index: int):
        """Updates file path input when dropdown selection changes."""
        text = self.file_combo.currentText()
        if text and text != "ไม่พบไฟล์ .xlsx ในโฟลเดอร์นี้":
            self.txt_file_path.setText(text)

    def browse_file(self):
        """Opens OS File Dialog to pick an Excel file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "เลือกไฟล์ Excel",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.txt_file_path.setText(file_path)
            self.log(f"เลือกไฟล์: {file_path}")

    def get_selected_file(self) -> str:
        """Returns and validates the selected file path."""
        path = self.txt_file_path.text().strip()
        if not path:
            QMessageBox.warning(self, "คำเตือน", "กรุณาเลือกไฟล์ Excel (.xlsx) ก่อนทำการประมวลผล")
            return None
        if not os.path.exists(path):
            QMessageBox.critical(self, "เกิดข้อผิดพลาด", f"ไม่พบไฟล์: {path}")
            return None
        return path

    def process_menu_1(self):
        """Executes Menu 1: Generate Summary Sheet."""
        file_path = self.get_selected_file()
        if not file_path:
            return

        self.log(f"กำลังประมวลผลเมนูที่ 1 (สรุปยอด) สำหรับไฟล์: {file_path}...")
        res = excel_processor.add_summary_sheet(file_path)

        if res["success"]:
            msg = f"สร้างชีทสรุปยอดสำเร็จ!\n\nชื่อชีท: {res['sheet_name']}\nไฟล์: {res['file_path']}"
            self.log(f"✅ สำเร็จ! ชีทที่ถูกสร้าง: {res['sheet_name']}")
            QMessageBox.information(self, "สำเร็จ", msg)
        else:
            self.log(f"❌ เกิดข้อผิดพลาด: {res['error']}")
            QMessageBox.critical(self, "เกิดข้อผิดพลาด", res["error"])

    def process_menu_2(self):
        """Executes Menu 2: Generate Barn Summary Sheet."""
        file_path = self.get_selected_file()
        if not file_path:
            return

        self.log(f"กำลังประมวลผลเมนูที่ 2 (สรุปยอดแยกเล้า) สำหรับไฟล์: {file_path}...")
        res = excel_processor.add_barn_summary_sheet(file_path)

        if res["success"]:
            msg = f"สร้างชีทสรุปยอดแยกเล้าสำเร็จ!\n\nชื่อชีท: {res['sheet_name']}\nไฟล์: {res['file_path']}"
            self.log(f"✅ สำเร็จ! ชีทที่ถูกสร้าง: {res['sheet_name']}")
            QMessageBox.information(self, "สำเร็จ", msg)
        else:
            self.log(f"❌ เกิดข้อผิดพลาด: {res['error']}")
            QMessageBox.critical(self, "เกิดข้อผิดพลาด", res["error"])


def launch_gui():
    """Launches the PyQt6 Desktop GUI Application."""
    app = QApplication(sys.argv)
    window = AccountAssistantGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    launch_gui()
