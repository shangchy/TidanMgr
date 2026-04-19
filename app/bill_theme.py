"""Qt 样式表：与「提单管理原型」一致的深色 / 浅色高端界面。"""

STYLESHEET_DARK = """
QMainWindow#BillAppMain { background-color: #0a0c10; }
QWidget { color: #e9ecf4; font-size: 13px; }
QDialog, QMessageBox { background-color: #1a2030; color: #e9ecf4; }

#sidebar {
  background-color: #12151e;
  border-right: 1px solid rgba(255, 255, 255, 0.08);
}
#sidebarBrand {
  padding: 4px 4px 8px 4px;
}
QLabel#sidebarLogoTitle {
  font-size: 16px;
  font-weight: 700;
  letter-spacing: 0.03em;
  color: #eef1f7;
}
QLabel#sidebarLogoSub {
  font-size: 11px;
  font-weight: 600;
  color: rgba(178, 186, 206, 0.92);
}
QPushButton#navActive {
  text-align: left;
  padding: 13px 16px;
  margin: 2px 0px;
  font-size: 16px;
  font-weight: 700;
  background-color: rgba(124, 158, 255, 0.14);
  color: #ffffff;
  border: none;
  border-left: 3px solid #a8bfff;
  border-radius: 6px;
}
QPushButton#navActive:hover { background-color: rgba(124, 158, 255, 0.2); }
QPushButton#navNormal {
  text-align: left;
  padding: 13px 16px;
  margin: 2px 0px;
  font-size: 16px;
  font-weight: 700;
  background-color: transparent;
  color: #f1f4fa;
  border: none;
  border-left: 3px solid transparent;
  border-radius: 6px;
}
QPushButton#navNormal:hover {
  background-color: rgba(255, 255, 255, 0.06);
  color: #ffffff;
}
#navDisabled {
  padding: 12px 16px;
  margin: 2px 0px;
  color: #e2e6ef;
  font-size: 16px;
  font-weight: 700;
}
QPushButton#navSub {
  text-align: left;
  padding: 9px 14px 9px 26px;
  margin: 0px 0px 0px 6px;
  font-size: 14px;
  font-weight: 600;
  background-color: transparent;
  color: rgba(210, 216, 232, 0.92);
  border: none;
  border-left: 2px solid transparent;
  border-radius: 6px;
}
QPushButton#navSub:hover {
  background-color: rgba(255, 255, 255, 0.05);
  color: #ffffff;
}
QPushButton#navSubActive {
  text-align: left;
  padding: 9px 14px 9px 26px;
  margin: 0px 0px 0px 6px;
  font-size: 14px;
  font-weight: 700;
  background-color: rgba(124, 158, 255, 0.12);
  color: #ffffff;
  border: none;
  border-left: 2px solid #a8bfff;
  border-radius: 6px;
}
QPushButton#navSubActive:hover {
  background-color: rgba(124, 158, 255, 0.18);
}
#sideDivider {
  background-color: rgba(255, 255, 255, 0.1);
  max-height: 1px;
  min-height: 1px;
  margin-top: 12px;
  margin-bottom: 12px;
  margin-left: 4px;
  margin-right: 4px;
}
#toolbarSep { background-color: rgba(255, 255, 255, 0.1); max-width: 1px; min-width: 1px; }

QTableWidget#tableFrozenCol {
  border: none;
  border-right: 1px solid rgba(255, 255, 255, 0.1);
  gridline-color: rgba(255, 255, 255, 0.06);
}
QTableWidget#tableScrollPart, QTableWidget#historyScrollPart {
  border: none;
}

#contentStack {
  background-color: #141820;
  border-left: 1px solid rgba(255, 255, 255, 0.05);
}
#stackBillPage, #stackHistoryPage, #stackPrintRecordsPage, #stackSettingsPage {
  background-color: #141820;
}

QLabel#sectionTitle {
  font-size: 14px;
  font-weight: 700;
  letter-spacing: 0.5px;
  color: #f1f4fa;
}
QLabel#settingsTitle {
  font-size: 17px;
  font-weight: 700;
  color: #b4c6ff;
  letter-spacing: 0.5px;
}
QLabel#hintLabel {
  color: #e8b86a;
  font-weight: 700;
  font-size: 11px;
}

QLineEdit, QTextEdit {
  background-color: #141822;
  color: #e9ecf4;
  border: 1px solid rgba(255, 255, 255, 0.12);
  border-radius: 6px;
  padding: 6px 10px;
  selection-background-color: #3d5a99;
  selection-color: #ffffff;
}
QLineEdit:focus, QTextEdit:focus {
  border: 1px solid #6b8cff;
}

QComboBox {
  background-color: #141822;
  color: #e9ecf4;
  border: 1px solid rgba(255, 255, 255, 0.12);
  border-radius: 6px;
  padding: 4px 10px;
  min-height: 1.2em;
}
QComboBox:focus { border: 1px solid #6b8cff; }
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView {
  background-color: #1e2433;
  color: #e9ecf4;
  selection-background-color: rgba(124, 158, 255, 0.35);
  border: 1px solid rgba(255, 255, 255, 0.12);
}

QTableWidget {
  background-color: #141822;
  color: #eef1f6;
  alternate-background-color: #11151c;
  gridline-color: rgba(255, 255, 255, 0.06);
  border: 1px solid rgba(255, 255, 255, 0.08);
  border-radius: 8px;
}
QHeaderView::section {
  background-color: #1e2433;
  color: #f4f6fb;
  padding: 14px 15px;
  border: none;
  border-right: 1px solid rgba(255, 255, 255, 0.06);
  border-bottom: 1px solid rgba(255, 255, 255, 0.1);
  font-weight: 700;
  font-size: 16px;
}
QTableWidget::item {
  padding: 6px 8px;
  color: #eef1f6;
}
QTableWidget::item:selected {
  background-color: rgba(124, 158, 255, 0.24);
  color: #ffffff;
}
QTableWidget::item:alternate:selected {
  background-color: rgba(124, 158, 255, 0.24);
  color: #ffffff;
}

QListWidget {
  background-color: #141822;
  color: #e9ecf4;
  border: 1px solid rgba(255, 255, 255, 0.12);
  border-radius: 6px;
  outline: none;
}
QListWidget::item { padding: 6px 8px; }
QListWidget::item:selected { background-color: rgba(124, 158, 255, 0.22); color: #ffffff; }
QListWidget::item:hover { background-color: rgba(255, 255, 255, 0.05); }

QPushButton {
  background-color: #252b3a;
  color: #f1f4fa;
  border: 1px solid rgba(255, 255, 255, 0.12);
  border-radius: 6px;
  padding: 6px 14px;
  font-weight: 600;
}
QPushButton:hover { background-color: #2e3548; color: #ffffff; }
QPushButton:pressed { background-color: #1e2433; }

QPushButton#btnAccent {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6b8cff, stop:1 #516fd9);
  color: #ffffff;
  border: 1px solid rgba(107, 140, 255, 0.55);
}
QPushButton#btnAccent:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7b9aff, stop:1 #5c7de8);
}
QPushButton#btnSuccess {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2fa882, stop:1 #25856a);
  color: #ffffff;
  border: 1px solid rgba(45, 157, 120, 0.5);
}
QPushButton#btnSuccess:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #34b88a, stop:1 #2a9678);
}
QPushButton#btnDanger {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f06676, stop:1 #c93042);
  color: #ffffff;
  border: 1px solid rgba(224, 82, 99, 0.55);
}
QPushButton#btnDanger:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f57a88, stop:1 #d94455);
}
QPushButton#btnGhost {
  background-color: transparent;
  color: #f1f4fa;
  border: 1px solid rgba(255, 255, 255, 0.12);
}
QPushButton#btnGhost:hover {
  background-color: rgba(255, 255, 255, 0.08);
  color: #ffffff;
}

QCheckBox { spacing: 8px; color: #f1f4fa; font-weight: 700; }
QCheckBox::indicator { width: 16px; height: 16px; }

QStatusBar {
  background-color: #e4e7ee;
  color: #0a0a0a;
  border-top: 1px solid rgba(0, 0, 0, 0.1);
  font-weight: 700;
  font-size: 11px;
}
QStatusBar QLabel { color: #0a0a0a; font-weight: 700; }

QScrollBar:vertical { width: 9px; background: transparent; margin: 2px; }
QScrollBar::handle:vertical {
  background: rgba(255, 255, 255, 0.14);
  border-radius: 4px;
  min-height: 28px;
}
QScrollBar::handle:vertical:hover { background: rgba(255, 255, 255, 0.22); }
QScrollBar:horizontal { height: 9px; background: transparent; margin: 2px; }
QScrollBar::handle:horizontal {
  background: rgba(255, 255, 255, 0.14);
  border-radius: 4px;
  min-width: 28px;
}
QScrollBar::handle:horizontal:hover { background: rgba(255, 255, 255, 0.22); }
"""

STYLESHEET_LIGHT = """
QMainWindow#BillAppMain { background-color: #e8ebf2; }
QWidget { color: #000000; font-size: 13px; }
QDialog, QMessageBox { background-color: #ffffff; color: #000000; }

#sidebar {
  background-color: #e8ecf4;
  border-right: 1px solid rgba(15, 23, 42, 0.1);
}
#sidebarBrand {
  padding: 4px 4px 8px 4px;
}
QLabel#sidebarLogoTitle {
  font-size: 16px;
  font-weight: 700;
  letter-spacing: 0.03em;
  color: #0f172a;
}
QLabel#sidebarLogoSub {
  font-size: 11px;
  font-weight: 600;
  color: rgba(71, 85, 105, 0.95);
}
QPushButton#navActive {
  text-align: left;
  padding: 13px 16px;
  margin: 2px 0px;
  font-size: 16px;
  font-weight: 700;
  background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 rgba(37, 99, 235, 0.12), stop:1 rgba(37, 99, 235, 0.02));
  color: #000000;
  border: none;
  border-left: 3px solid #2563eb;
  border-radius: 6px;
}
QPushButton#navActive:hover { background-color: rgba(37, 99, 235, 0.1); }
QPushButton#navNormal {
  text-align: left;
  padding: 13px 16px;
  margin: 2px 0px;
  font-size: 16px;
  font-weight: 700;
  background-color: transparent;
  color: #000000;
  border: none;
  border-left: 3px solid transparent;
  border-radius: 6px;
}
QPushButton#navNormal:hover {
  background-color: rgba(15, 23, 42, 0.06);
  color: #000000;
}
#navDisabled {
  padding: 12px 16px;
  margin: 2px 0px;
  color: #404040;
  font-size: 16px;
  font-weight: 700;
}
QPushButton#navSub {
  text-align: left;
  padding: 9px 14px 9px 26px;
  margin: 0px 0px 0px 6px;
  font-size: 14px;
  font-weight: 600;
  background-color: transparent;
  color: rgba(51, 65, 85, 0.95);
  border: none;
  border-left: 2px solid transparent;
  border-radius: 6px;
}
QPushButton#navSub:hover {
  background-color: rgba(15, 23, 42, 0.06);
  color: #000000;
}
QPushButton#navSubActive {
  text-align: left;
  padding: 9px 14px 9px 26px;
  margin: 0px 0px 0px 6px;
  font-size: 14px;
  font-weight: 700;
  background-color: rgba(37, 99, 235, 0.1);
  color: #000000;
  border: none;
  border-left: 2px solid #2563eb;
  border-radius: 6px;
}
QPushButton#navSubActive:hover {
  background-color: rgba(37, 99, 235, 0.14);
}
#sideDivider {
  background-color: rgba(15, 23, 42, 0.12);
  max-height: 1px;
  min-height: 1px;
  margin-top: 12px;
  margin-bottom: 12px;
  margin-left: 4px;
  margin-right: 4px;
}
#toolbarSep { background-color: rgba(15, 23, 42, 0.12); max-width: 1px; min-width: 1px; }

QTableWidget#tableFrozenCol {
  border: none;
  border-right: 1px solid rgba(15, 23, 42, 0.12);
  gridline-color: rgba(15, 23, 42, 0.08);
}
QTableWidget#tableScrollPart, QTableWidget#historyScrollPart {
  border: none;
}

#contentStack {
  background-color: #f6f7fb;
  border-left: 1px solid rgba(15, 23, 42, 0.06);
}
#stackBillPage, #stackHistoryPage, #stackPrintRecordsPage, #stackSettingsPage {
  background-color: #f6f7fb;
}

QLabel#sectionTitle {
  font-size: 14px;
  font-weight: 700;
  letter-spacing: 0.5px;
  color: #000000;
}
QLabel#settingsTitle {
  font-size: 17px;
  font-weight: 700;
  color: #1e40af;
  letter-spacing: 0.5px;
}
QLabel#hintLabel {
  color: #b45309;
  font-weight: 700;
  font-size: 11px;
}

QLineEdit, QTextEdit {
  background-color: #ffffff;
  color: #0f172a;
  border: 1px solid rgba(15, 23, 42, 0.14);
  border-radius: 6px;
  padding: 6px 10px;
  selection-background-color: #bfdbfe;
  selection-color: #0f172a;
}
QLineEdit:focus, QTextEdit:focus {
  border: 1px solid #2563eb;
}

QComboBox {
  background-color: #ffffff;
  color: #0f172a;
  border: 1px solid rgba(15, 23, 42, 0.14);
  border-radius: 6px;
  padding: 4px 10px;
  min-height: 1.2em;
}
QComboBox:focus { border: 1px solid #2563eb; }
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView {
  background-color: #ffffff;
  color: #0f172a;
  selection-background-color: #dbeafe;
  border: 1px solid rgba(15, 23, 42, 0.12);
}

QTableWidget {
  background-color: #ffffff;
  color: #0a0a0a;
  alternate-background-color: #f8fafc;
  gridline-color: rgba(15, 23, 42, 0.08);
  border: 1px solid rgba(15, 23, 42, 0.1);
  border-radius: 8px;
}
QHeaderView::section {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f1f5f9, stop:1 #e8edf5);
  color: #000000;
  padding: 14px 15px;
  border: none;
  border-right: 1px solid rgba(15, 23, 42, 0.08);
  border-bottom: 1px solid rgba(15, 23, 42, 0.1);
  font-weight: 700;
  font-size: 16px;
}
QTableWidget::item {
  padding: 6px 8px;
  color: #0a0a0a;
}
QTableWidget::item:selected {
  background-color: rgba(37, 99, 235, 0.14);
  color: #000000;
}
QTableWidget::item:alternate:selected {
  background-color: rgba(37, 99, 235, 0.14);
  color: #000000;
}

QListWidget {
  background-color: #ffffff;
  color: #0f172a;
  border: 1px solid rgba(15, 23, 42, 0.12);
  border-radius: 6px;
  outline: none;
}
QListWidget::item { padding: 6px 8px; }
QListWidget::item:selected { background-color: #dbeafe; color: #0f172a; }
QListWidget::item:hover { background-color: #f1f5f9; }

QPushButton {
  background-color: #ffffff;
  color: #000000;
  border: 1px solid rgba(15, 23, 42, 0.14);
  border-radius: 6px;
  padding: 6px 14px;
  font-weight: 600;
}
QPushButton:hover { background-color: #f1f5f9; color: #000000; }
QPushButton:pressed { background-color: #e2e8f0; color: #0f172a; }

QPushButton#btnAccent {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3b82f6, stop:1 #2563eb);
  color: #ffffff;
  border: 1px solid rgba(37, 99, 235, 0.4);
}
QPushButton#btnAccent:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4f8ff7, stop:1 #2f6fed);
}
QPushButton#btnSuccess {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #14b8a6, stop:1 #0d9488);
  color: #ffffff;
  border: 1px solid rgba(13, 148, 136, 0.45);
}
QPushButton#btnSuccess:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2dd4c1, stop:1 #14b8a6);
}
QPushButton#btnDanger {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ef4444, stop:1 #dc2626);
  color: #ffffff;
  border: 1px solid rgba(220, 38, 38, 0.45);
}
QPushButton#btnDanger:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f87171, stop:1 #e11d48);
}
QPushButton#btnGhost {
  background-color: transparent;
  color: #000000;
  border: 1px solid rgba(15, 23, 42, 0.14);
}
QPushButton#btnGhost:hover {
  background-color: rgba(15, 23, 42, 0.06);
  color: #000000;
}

QCheckBox { spacing: 8px; color: #000000; font-weight: 700; }

QStatusBar {
  background-color: #e8ecf4;
  color: #000000;
  border-top: 1px solid rgba(15, 23, 42, 0.1);
  font-weight: 700;
  font-size: 11px;
}
QStatusBar QLabel { color: #000000; font-weight: 700; }

QScrollBar:vertical { width: 9px; background: transparent; margin: 2px; }
QScrollBar::handle:vertical {
  background: rgba(15, 23, 42, 0.15);
  border-radius: 4px;
  min-height: 28px;
}
QScrollBar::handle:vertical:hover { background: rgba(15, 23, 42, 0.28); }
QScrollBar:horizontal { height: 9px; background: transparent; margin: 2px; }
QScrollBar::handle:horizontal {
  background: rgba(15, 23, 42, 0.15);
  border-radius: 4px;
  min-width: 28px;
}
QScrollBar::handle:horizontal:hover { background: rgba(15, 23, 42, 0.28); }
"""
