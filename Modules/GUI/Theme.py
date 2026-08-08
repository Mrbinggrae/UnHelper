from __future__ import annotations

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication


APP_STYLESHEET = """
QWidget {
    color: #F4F4F5;
    background: #050505;
    font-family: "Malgun Gothic", "Segoe UI";
    font-size: 11pt;
}
QMainWindow, QWidget#Root {
    background: #050505;
}
QFrame#Header {
    background: #090909;
    border-bottom: 1px solid #242424;
}
QLabel#Title {
    color: #FFFFFF;
    font-size: 24pt;
    font-weight: 800;
}
QLabel#Version {
    color: #9CA3AF;
    font-size: 9.5pt;
}
QLabel#Status {
    color: #7DD3FC;
    font-weight: 700;
}
QTabWidget::pane {
    border: 0;
    background: #050505;
}
QTabBar::tab {
    background: transparent;
    color: #D4D4D8;
    padding: 12px 20px 9px 20px;
    margin-right: 6px;
    border-bottom: 3px solid transparent;
    font-size: 14pt;
    font-weight: 800;
}
QTabBar::tab:selected {
    color: #FFFFFF;
    border-bottom: 3px solid #45C42E;
}
QTabBar::tab:hover {
    color: #FFFFFF;
}
QTableWidget {
    background: #DDE1E5;
    alternate-background-color: #C9CFD5;
    color: #111827;
    border: 1px solid #E5E7EB;
    gridline-color: #FFFFFF;
    selection-background-color: #60A5FA;
    selection-color: #111827;
}
QHeaderView::section {
    background: #166987;
    color: #FFFFFF;
    border: 1px solid #FFFFFF;
    padding: 6px;
    font-size: 13pt;
    font-weight: 800;
}
QPushButton {
    background: #171717;
    border: 1px solid #3F3F46;
    border-radius: 7px;
    color: #F4F4F5;
    padding: 8px 14px;
    font-weight: 700;
}
QPushButton:hover {
    background: #262626;
    border-color: #60A5FA;
}
QPushButton:disabled {
    background: #111111;
    border-color: #262626;
    color: #52525B;
}
QPushButton#PrimaryButton {
    background: #166987;
    border: 2px solid #0D4359;
    border-radius: 10px;
    color: #FFFFFF;
    font-size: 15pt;
    padding: 12px 34px;
}
QPushButton#PrimaryButton:hover {
    background: #1E7E9F;
}
QPushButton#StopButton {
    background: #7F1D1D;
    border-color: #991B1B;
}
QPushButton#CategoryButton {
    background: #39A91E;
    border-color: #45C42E;
    font-size: 11pt;
    padding: 5px 12px;
}
QPushButton#CategoryButton:hover {
    background: #45C42E;
}
QPushButton#CategoryButton[classification="중량"] {
    background: #B45309;
    border-color: #F59E0B;
}
QPushButton#CategoryButton[classification="중량"]:hover {
    background: #D97706;
}
QPushButton#CategoryButton[classification="고단"] {
    background: #6D28D9;
    border-color: #8B5CF6;
}
QPushButton#CategoryButton[classification="고단"]:hover {
    background: #7C3AED;
}
QPushButton#CategoryButton[classification="?"] {
    background: #3F3F46;
    border-color: #71717A;
}
QPushButton#SettingsButton {
    background: #F5F5F4;
    color: #27272A;
    border: 1px solid #A1A1AA;
    border-radius: 5px;
    font-size: 22pt;
    padding: 0;
}
QPlainTextEdit {
    background: #09090B;
    border: 1px solid #27272A;
    border-radius: 6px;
    color: #D4D4D8;
    font-family: "Consolas", "Malgun Gothic";
    font-size: 9pt;
    padding: 7px;
}
QDialog, QMessageBox {
    background: #111111;
}
QLineEdit {
    background: #18181B;
    border: 1px solid #3F3F46;
    border-radius: 5px;
    color: #F4F4F5;
    padding: 7px;
}
QCheckBox {
    spacing: 8px;
}
QProgressBar {
    background: #18181B;
    border: 1px solid #3F3F46;
    border-radius: 5px;
    text-align: center;
}
QProgressBar::chunk {
    background: #39A91E;
    border-radius: 4px;
}
"""


def apply_dark_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#050505"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#F4F4F5"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#111111"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#18181B"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#F4F4F5"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#171717"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#F4F4F5"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#39A91E"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor("#71717A"))
    app.setPalette(palette)
    app.setStyleSheet(APP_STYLESHEET)
