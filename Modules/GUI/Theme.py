from __future__ import annotations

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication


COLORS = {
    "canvas": "#0B0F14",
    "surface": "#111821",
    "raised": "#16212C",
    "hover": "#1B2836",
    "border": "#253244",
    "border_strong": "#34475A",
    "text": "#F3F7FA",
    "secondary": "#A8B4C2",
    "muted": "#738294",
    "accent": "#38BDF8",
    "accent_hover": "#60CAFA",
    "accent_pressed": "#0EA5E9",
    "primary": "#0369A1",
    "primary_hover": "#0E7490",
    "selection": "#174B63",
    "success": "#22C55E",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "violet": "#A78BFA",
}


APP_STYLESHEET = r"""
QWidget {
    color: #F3F7FA;
    font-family: "Segoe UI", "Malgun Gothic";
    font-size: 10pt;
}
QWidget:disabled {
    color: #58677A;
}
QMainWindow, QWidget#Root, QWidget#Page, QDialog, QMessageBox {
    background: #0B0F14;
}
QLabel {
    background: transparent;
}
QFrame#Header {
    background: #0E151E;
    border: 0;
    border-bottom: 1px solid #253244;
}
QFrame#Footer {
    background: #0E151E;
    border: 0;
    border-top: 1px solid #1C2836;
}
QFrame#DataCard, QFrame[card="true"] {
    background: #111821;
    border: 1px solid #253244;
    border-radius: 12px;
}
QScrollArea#SettingsScroll, QWidget#SettingsViewport, QWidget#SettingsContent {
    background: transparent;
    border: 0;
}
QLabel#BrandMark {
    background: #0C2B3A;
    border: 1px solid #1D5269;
    border-radius: 11px;
    color: #7DD3FC;
    font-size: 14pt;
    font-weight: 800;
}
QLabel#Title {
    color: #FFFFFF;
    font-size: 21pt;
    font-weight: 800;
}
QLabel#Version {
    color: #8FA0B3;
    font-size: 9.5pt;
}
QLabel#DialogTitle {
    color: #F8FAFC;
    font-size: 18pt;
    font-weight: 800;
}
QLabel#DialogHeading, QLabel#SectionTitle {
    color: #F8FAFC;
    font-size: 14pt;
    font-weight: 700;
}
QLabel#FieldLabel, QLabel#DetailLabel {
    color: #DCE5EE;
    font-weight: 700;
}
QLabel#SectionDescription, QLabel#HelpText, QLabel#MutedText {
    color: #8291A3;
    font-size: 9.25pt;
}
QLabel#PlaceholderText {
    color: #64748B;
    font-size: 12pt;
}
QLabel#Status {
    color: #BAE6FD;
    font-weight: 700;
}
QLabel#StatusDot {
    color: #38BDF8;
    font-size: 9pt;
}
QLabel#ErrorSummary {
    background: #241419;
    border: 1px solid #5F2631;
    border-radius: 8px;
    color: #FECACA;
    padding: 10px 12px;
}

QTabWidget::pane {
    background: transparent;
    border: 0;
}
QTabBar#MainTabBar {
    background: #0D141D;
    border-bottom: 1px solid #253244;
}
QTabBar#MainTabBar::tab {
    background: transparent;
    border: 0;
    border-bottom: 2px solid transparent;
    color: #8FA0B3;
    min-width: 86px;
    padding: 12px 20px 10px 20px;
    margin: 0 3px;
    font-size: 11.5pt;
    font-weight: 700;
}
QTabBar#MainTabBar::tab:selected {
    background: #131E2A;
    border-bottom-color: #38BDF8;
    color: #F8FAFC;
}
QTabBar#MainTabBar::tab:hover:!selected {
    background: #111B26;
    color: #DCE5EE;
}
QTabBar#SubTabBar {
    background: transparent;
}
QTabBar#SubTabBar::tab {
    background: #111821;
    border: 1px solid #253244;
    border-radius: 8px;
    color: #8FA0B3;
    min-width: 72px;
    padding: 7px 15px;
    margin: 8px 4px 6px 4px;
    font-size: 10pt;
    font-weight: 650;
}
QTabBar#SubTabBar::tab:selected {
    background: #0C2B3A;
    border-color: #286079;
    color: #BAE6FD;
}
QTabBar#SubTabBar::tab:hover:!selected {
    background: #172330;
    border-color: #34475A;
    color: #DCE5EE;
}

QTableWidget, QTableView {
    background: #0F1620;
    alternate-background-color: #131C26;
    color: #E6EDF5;
    border: 1px solid #253244;
    border-radius: 10px;
    gridline-color: #202C3A;
    outline: 0;
    selection-background-color: #174B63;
    selection-color: #F8FAFC;
}
QFrame#DataCard QTableWidget#RawTable {
    border: 0;
    border-radius: 11px;
}
QTableWidget::item, QTableView::item {
    border: 0;
    padding: 7px 9px;
}
QTableWidget::item:hover, QTableView::item:hover {
    background: #172431;
}
QTableWidget::item:selected, QTableView::item:selected {
    background: #174B63;
    color: #F8FAFC;
}
QHeaderView {
    background: transparent;
}
QHeaderView::section {
    background: #182432;
    color: #DCE5EE;
    border: 0;
    border-right: 1px solid #253244;
    border-bottom: 1px solid #34475A;
    padding: 8px 9px;
    font-size: 9.75pt;
    font-weight: 700;
}
QHeaderView::section:hover {
    background: #1D2C3C;
}
QTableCornerButton::section {
    background: #182432;
    border: 0;
    border-right: 1px solid #253244;
    border-bottom: 1px solid #34475A;
}

QPushButton {
    background: #16212C;
    border: 1px solid #34475A;
    border-radius: 8px;
    color: #E7EEF6;
    min-height: 22px;
    padding: 7px 14px;
    font-weight: 650;
}
QPushButton:hover {
    background: #1B2836;
    border-color: #4A6178;
    color: #FFFFFF;
}
QPushButton:pressed {
    background: #101923;
    border-color: #38BDF8;
}
QPushButton:focus {
    border: 2px solid #38BDF8;
}
QPushButton:disabled {
    background: #10171F;
    border-color: #1E2A38;
    color: #58677A;
}
QPushButton#PrimaryButton {
    background: #0369A1;
    border: 1px solid #38BDF8;
    border-radius: 9px;
    color: #FFFFFF;
    min-height: 28px;
    padding: 9px 24px;
    font-size: 11pt;
    font-weight: 750;
}
QPushButton#PrimaryButton:hover {
    background: #0E7490;
    border-color: #7DD3FC;
}
QPushButton#PrimaryButton:pressed {
    background: #075985;
}
QPushButton#StopButton, QPushButton#DangerButton {
    background: #3B171C;
    border-color: #7F2833;
    color: #FECACA;
}
QPushButton#StopButton:hover, QPushButton#DangerButton:hover {
    background: #551D25;
    border-color: #EF4444;
    color: #FFFFFF;
}
QPushButton#SettingsButton {
    background: #121C27;
    border: 1px solid #34475A;
    border-radius: 9px;
    color: #DCE5EE;
    padding: 6px 14px;
    font-size: 10pt;
    font-weight: 700;
}
QPushButton#SettingsButton:hover {
    background: #1B2836;
    border-color: #38BDF8;
    color: #FFFFFF;
}
QPushButton#CategoryButton {
    background: #273444;
    border: 1px solid #506176;
    border-radius: 7px;
    color: #E2E8F0;
    min-height: 18px;
    padding: 4px 11px;
    font-size: 9.5pt;
}
QPushButton#CategoryButton[classification="경량"] {
    background: #123523;
    border-color: #247848;
    color: #BBF7D0;
}
QPushButton#CategoryButton[classification="경량"]:hover {
    background: #18502F;
    border-color: #22C55E;
}
QPushButton#CategoryButton[classification="중량"] {
    background: #3B2A0E;
    border-color: #8A5A12;
    color: #FDE68A;
}
QPushButton#CategoryButton[classification="중량"]:hover {
    background: #563B10;
    border-color: #F59E0B;
}
QPushButton#CategoryButton[classification="고단"] {
    background: #2B1D49;
    border-color: #6445A5;
    color: #DDD6FE;
}
QPushButton#CategoryButton[classification="고단"]:hover {
    background: #3B2765;
    border-color: #A78BFA;
}
QPushButton#CategoryButton[classification="?"] {
    background: #263241;
    border-color: #53657A;
    color: #CBD5E1;
}

QLineEdit, QComboBox {
    background: #16212C;
    border: 1px solid #34475A;
    border-radius: 8px;
    color: #F3F7FA;
    min-height: 22px;
    padding: 7px 10px;
    selection-background-color: #174B63;
    selection-color: #FFFFFF;
}
QLineEdit:hover, QComboBox:hover {
    border-color: #4A6178;
}
QLineEdit:focus, QComboBox:focus {
    border: 2px solid #38BDF8;
    padding: 6px 9px;
}
QLineEdit:read-only {
    background: #0F1620;
    color: #A8B4C2;
}
QLineEdit::placeholder {
    color: #64748B;
}
QComboBox::drop-down {
    border: 0;
    width: 26px;
}
QComboBox QAbstractItemView {
    background: #16212C;
    border: 1px solid #34475A;
    color: #F3F7FA;
    outline: 0;
    padding: 4px;
    selection-background-color: #174B63;
    selection-color: #FFFFFF;
}
QCheckBox {
    color: #D5DEE8;
    spacing: 8px;
}

QPlainTextEdit, QTextEdit {
    background: #0F1620;
    border: 1px solid #253244;
    border-radius: 8px;
    color: #C9D5E2;
    padding: 8px;
    selection-background-color: #174B63;
}
QPlainTextEdit#LogView {
    background: #080C11;
    color: #91A4B7;
    font-family: "Cascadia Mono", "Consolas", "Malgun Gothic";
    font-size: 8.75pt;
}

QProgressBar {
    background: #101821;
    border: 1px solid #253244;
    border-radius: 7px;
    color: #DCE5EE;
    min-height: 16px;
    text-align: center;
}
QProgressBar::chunk {
    background: #0EA5E9;
    border-radius: 6px;
}

QScrollBar:vertical {
    background: transparent;
    border: 0;
    margin: 2px;
    width: 11px;
}
QScrollBar:horizontal {
    background: transparent;
    border: 0;
    margin: 2px;
    height: 11px;
}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
    background: #34475A;
    border-radius: 5px;
    min-height: 32px;
    min-width: 32px;
}
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {
    background: #486079;
}
QScrollBar::add-line, QScrollBar::sub-line {
    width: 0;
    height: 0;
}
QScrollBar::add-page, QScrollBar::sub-page {
    background: transparent;
}

QToolTip {
    background: #1B2836;
    border: 1px solid #486079;
    border-radius: 5px;
    color: #F3F7FA;
    padding: 6px 8px;
}
QFrame[frameShape="4"], QFrame[frameShape="5"] {
    color: #253244;
}
"""


def apply_dark_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(COLORS["canvas"]))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(COLORS["text"]))
    palette.setColor(QPalette.ColorRole.Base, QColor("#0F1620"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#131C26"))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(COLORS["hover"]))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(COLORS["text"]))
    palette.setColor(QPalette.ColorRole.Text, QColor(COLORS["text"]))
    palette.setColor(QPalette.ColorRole.Button, QColor(COLORS["raised"]))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(COLORS["text"]))
    palette.setColor(QPalette.ColorRole.BrightText, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.Link, QColor(COLORS["accent"]))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(COLORS["selection"]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(COLORS["muted"]))
    palette.setColor(
        QPalette.ColorGroup.Disabled,
        QPalette.ColorRole.Text,
        QColor("#58677A"),
    )
    palette.setColor(
        QPalette.ColorGroup.Disabled,
        QPalette.ColorRole.ButtonText,
        QColor("#58677A"),
    )
    app.setPalette(palette)
    app.setStyleSheet(APP_STYLESHEET)
