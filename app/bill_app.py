import json
import os
import re
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from openpyxl import load_workbook
from PySide6.QtCore import QEvent, QObject, QModelIndex, QPoint, QPointF, QRect, QRectF, QSize, Qt, QTimer, QUrl
from PySide6.QtGui import (
    QBrush,
    QColor,
    QDesktopServices,
    QFont,
    QFontMetrics,
    QFontMetricsF,
    QGuiApplication,
    QIcon,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
    QPalette,
    QPolygon,
    QPolygonF,
)

from bill_constants import (
    ALLOW_PRINT_SCROLL_COL,
    ALLOW_PRINT_URL_DISPLAY_COL,
    CITIES,
    FROZEN_COLUMNS,
    DISPLAY_FIELDS,
    DURATIONS,
    FIELDS,
    HEADERS,
    HISTORY_SCROLL_COLUMNS,
    HISTORY_SCROLL_FIELDS,
    MAIN_SCROLL_COLUMNS,
    OP_MAP,
    PRINT_LOG_COL_COUNT,
    PRINT_LOG_DATA_FIELDS,
    PRINT_LOG_HEADERS,
    PROVINCE_TO_CITIES,
    PROVINCES,
    TYPE_MAP,
    coerce_allow_print_url,
    find_earliest_region_in_left,
    first_url,
    int_to_cn,
    url_lines_for_filter,
    sanitize_filename,
    split_multi,
    cities_under_provinces,
)
from bill_paths import (
    APP_DIR,
    DATA_FILE,
    HISTORY_FILE,
    LICENSE_EXPIRED_MSG,
    PICKER_RECENT_FILE,
    PRINT_RECORDS_FILE,
    TEMPLATE_FILE,
    THEME_FILE,
    is_license_valid,
)
from bill_theme import STYLESHEET_DARK, STYLESHEET_LIGHT
from bill_widgets import (
    AllowPrintUrlCellDelegate,
    ColumnPickFilterPopup,
    HoverFilterHeaderView,
    MultiSelectDialog,
    UrlCellEditor,
    make_sidebar_chevrons_pixmap,
    make_sidebar_logo_pixmap,
    style_combo_centered,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class BillApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setObjectName("BillAppMain")
        self.setWindowTitle("提单管理")
        self.records: list[dict[str, Any]] = []
        self.history_records: list[dict[str, Any]] = []
        self.history_filtered_indices: list[int] = []
        self.filtered_indices: list[int] = []
        self.theme = self.load_theme()
        self.updating_table = False
        self.select_all_state = False
        self.field_errors: dict[tuple[int, str], bool] = {}
        self.field_error_msgs: dict[tuple[int, str], str] = {}
        self.picker_recent: dict[str, list[str]] = {}
        self.header_filters: dict[str, set[str]] = {}
        self.header_sort_field: str | None = None
        self.header_sort_order = Qt.SortOrder.AscendingOrder
        self._table_v_sync = False
        self._table_sel_sync = False
        self.history_header_filters: dict[str, set[str]] = {}
        self.history_sort_field: str | None = None
        self.history_sort_order = Qt.SortOrder.AscendingOrder
        self.print_records: list[dict[str, Any]] = []
        self.print_log_filtered_indices: list[int] = []
        self.print_header_filters: dict[str, set[str]] = {}
        self.print_sort_field: str | None = "printed_at"
        self.print_sort_order = Qt.SortOrder.DescendingOrder
        self._hist_v_sync = False
        self._hist_sel_sync = False
        self._filter_popup: ColumnPickFilterPopup | None = None
        self._license_timer = QTimer(self)
        self._license_timer.timeout.connect(self._check_license_and_exit_if_needed)
        self._license_timer.start(10_000)
        QTimer.singleShot(0, self._check_license_and_exit_if_needed)
        # 须在 _setup_ui 之前：建表时会触发列宽/URL 行高回调，依赖本属性
        _fm0 = QFontMetrics(self.font())
        self._data_row_height = max(44, int((_fm0.height() + max(_fm0.leading(), 0)) * 1.5) + 14)
        self._setup_ui()
        self.load_data()
        self.load_history_data()
        self.load_print_records()
        self.load_picker_recent()
        self.apply_theme()
        fm = QFontMetrics(self.font())
        self._data_row_height = max(44, int((fm.height() + max(fm.leading(), 0)) * 1.5) + 14)
        self.refresh_table()
        # 刷新后子控件尺寸提示已稳定，再套一次初始几何（避免宽表把窗口最小宽度撑满）
        self._apply_initial_window_geometry()
        cw = self.centralWidget()
        if cw is not None:
            cw.setMinimumWidth(0)

    def _check_license_and_exit_if_needed(self):
        if is_license_valid():
            return
        self._license_timer.stop()
        _show_license_expired_dialog(self)
        QApplication.quit()
        sys.exit(0)

    def _setup_ui(self):
        center = QWidget(self)
        center.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        self.setCentralWidget(center)
        root = QHBoxLayout(center)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setAutoFillBackground(True)
        self.sidebar.setFixedWidth(240)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(10, 16, 10, 20)
        sidebar_layout.setSpacing(8)
        self.sidebar_brand = QWidget()
        self.sidebar_brand.setObjectName("sidebarBrand")
        brand_row = QHBoxLayout(self.sidebar_brand)
        brand_row.setContentsMargins(0, 0, 0, 0)
        brand_row.setSpacing(12)
        self._sidebar_logo_icon = QLabel()
        self._sidebar_logo_icon.setFixedSize(44, 44)
        brand_titles = QVBoxLayout()
        brand_titles.setContentsMargins(0, 1, 0, 0)
        brand_titles.setSpacing(5)
        self._sidebar_logo_title = QLabel("提单管理")
        self._sidebar_logo_title.setObjectName("sidebarLogoTitle")
        self._sidebar_logo_sub = QLabel("任务表 · 核对与归档")
        self._sidebar_logo_sub.setObjectName("sidebarLogoSub")
        brand_titles.addWidget(self._sidebar_logo_title)
        brand_titles.addWidget(self._sidebar_logo_sub)
        brand_row.addWidget(self._sidebar_logo_icon, 0, Qt.AlignmentFlag.AlignTop)
        brand_row.addLayout(brand_titles, 1)
        self.nav_bill = QPushButton("📝  提单表")
        self.nav_bill.setObjectName("navActive")
        self.nav_bill.clicked.connect(self.show_bill_page)
        side_item_sub1 = QLabel("📎  配件表（规划中）")
        side_item_sub1.setObjectName("navDisabled")
        self.nav_history = QPushButton("🕘  历史提单表")
        self.nav_history.setObjectName("navNormal")
        self.nav_history.clicked.connect(self.show_history_page)
        self.nav_print_records = QPushButton("📋  打印记录")
        self.nav_print_records.setObjectName("navNormal")
        self.nav_print_records.clicked.connect(self.show_print_records_page)
        self.nav_settings = QPushButton("⚙️  设置")
        self.nav_settings.setObjectName("navNormal")
        self.nav_settings.clicked.connect(self.show_settings_page)
        divider = QFrame()
        divider.setObjectName("sideDivider")
        divider.setFrameShape(QFrame.HLine)
        sidebar_layout.addWidget(self.sidebar_brand)
        sidebar_layout.addWidget(self.nav_bill)
        sidebar_layout.addWidget(self.nav_history)
        sidebar_layout.addWidget(self.nav_print_records)
        sidebar_layout.addWidget(divider)
        sidebar_layout.addWidget(side_item_sub1)
        sidebar_layout.addWidget(self.nav_settings)
        sidebar_layout.addStretch(1)
        root.addWidget(self.sidebar)
        # 窗口 show 之前 sidebar.isVisible() 常为 False，故用逻辑状态驱动侧栏按钮图标/提示
        self._sidebar_visible = True

        main_wrap = QWidget()
        main_wrap.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        main_layout = QVBoxLayout(main_wrap)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        self.content_stack = QStackedWidget()
        self.content_stack.setObjectName("contentStack")
        self.content_stack.setAutoFillBackground(True)
        # 水平 Ignored：避免 QStackedWidget 与各页「最宽子控件」累加成超大窗口最小宽度
        self.content_stack.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        main_layout.addWidget(self.content_stack, 1)

        bill_page = QWidget()
        bill_page.setObjectName("stackBillPage")
        bill_page.setAutoFillBackground(True)
        bill_page.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        bill_page_layout = QVBoxLayout(bill_page)
        bill_page_layout.setContentsMargins(0, 0, 0, 0)
        bill_page_layout.setSpacing(0)
        top = QHBoxLayout()
        top.setContentsMargins(16, 12, 16, 12)
        top.setSpacing(12)
        self.btn_toggle_sidebar = QPushButton()
        self.btn_toggle_sidebar.setObjectName("btnGhost")
        self.btn_toggle_sidebar.setIconSize(QSize(24, 24))
        self.btn_toggle_sidebar.setMinimumSize(36, 30)
        self.btn_toggle_sidebar.clicked.connect(self.toggle_sidebar)
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索任务名...")
        self.search.setFixedWidth(220)
        self.search.textChanged.connect(self.refresh_table)
        btn_search = QPushButton("🔍 搜索")
        btn_search.setObjectName("btnGhost")
        btn_search.clicked.connect(self.on_bill_search_clicked)
        btn_clear = QPushButton("🔄 清空筛选")
        btn_clear.setObjectName("btnGhost")
        btn_clear.clicked.connect(self.clear_bill_search_and_filters)
        self.chk_print_url = QCheckBox("是否打印URL")
        self.lbl_print_url_hint_off = QLabel("未勾选：导出文件中 URL 列为空。")
        self.lbl_print_url_hint_on = QLabel("勾选：仅「允许」为是的行写入 URL。")
        for lb in (self.lbl_print_url_hint_off, self.lbl_print_url_hint_on):
            lb.setObjectName("hintLabel")
            lb.setWordWrap(True)
            lb.setMinimumWidth(0)
            lb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.chk_print_url.toggled.connect(self._update_print_url_hint_lines_style)
        self.btn_print = QPushButton("🖨 打印")
        self.btn_print.setObjectName("btnAccent")
        self.btn_print.clicked.connect(self.export_excel)
        self.btn_add = QPushButton("➕ 新增行")
        self.btn_add.setObjectName("btnSuccess")
        self.btn_add.clicked.connect(self.add_row)
        self.btn_bulk_fill = QPushButton("📋 批量填充")
        self.btn_bulk_fill.setObjectName("btnAccent")
        self.btn_bulk_fill.clicked.connect(self.bulk_fill_selected)
        self.btn_bulk_fill.setParent(bill_page)
        self.btn_bulk_fill.hide()
        self.btn_delete_sel = QPushButton("🗑 删除选中")
        self.btn_delete_sel.setObjectName("btnDanger")
        self.btn_delete_sel.clicked.connect(self.delete_selected)
        print_opt_layout = QVBoxLayout()
        print_opt_layout.setContentsMargins(0, 0, 0, 0)
        print_opt_layout.setSpacing(2)
        print_opt_layout.addWidget(
            self.chk_print_url, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        print_opt_layout.addWidget(self.lbl_print_url_hint_off, 0)
        print_opt_layout.addWidget(self.lbl_print_url_hint_on, 0)
        self._update_print_url_hint_lines_style()
        sep1 = QFrame()
        sep1.setObjectName("toolbarSep")
        sep1.setFrameShape(QFrame.VLine)
        sep2 = QFrame()
        sep2.setObjectName("toolbarSep")
        sep2.setFrameShape(QFrame.VLine)
        for w in [self.btn_toggle_sidebar, self.search, btn_search, btn_clear]:
            top.addWidget(w)
        top.addWidget(sep1)
        top.addLayout(print_opt_layout, 1)
        top.addWidget(sep2)
        for w in [self.btn_print, self.btn_add]:
            top.addWidget(w)
        top.addStretch(1)
        bill_page_layout.addLayout(top)

        table_split = QWidget()
        table_split.setObjectName("tableSplit")
        table_split.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        split_lo = QHBoxLayout(table_split)
        split_lo.setContentsMargins(0, 0, 0, 0)
        split_lo.setSpacing(0)
        self.table_frozen = QTableWidget(0, FROZEN_COLUMNS)
        self.table_frozen.setObjectName("tableFrozenCol")
        # 首列勾选：表头文案由 HoverFilterHeaderView 自绘（图标+全选/取消全选），此处占位为空
        self.table_frozen.setHorizontalHeaderLabels(["", HEADERS["task_name"]])
        self.table_frozen.verticalHeader().setVisible(False)
        self.table_frozen.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_frozen.setAlternatingRowColors(True)
        self.table_frozen.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        # 与右侧主表同时保留横向滚动条占位，避免仅一侧出现横向条时纵向视口高度不一致导致滚动错位
        self.table_frozen.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.table_frozen.horizontalScrollBar().setEnabled(False)
        self.table_frozen.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.table_frozen.setFrameShape(QFrame.NoFrame)
        self.table_frozen.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        self.table_frozen.cellClicked.connect(lambda r, c: self.on_cell_clicked(r, c))
        self.table_frozen.cellDoubleClicked.connect(lambda r, c: self.on_cell_double_click(r, c))
        self.table = QTableWidget(0, MAIN_SCROLL_COLUMNS)
        self.table.setObjectName("tableScrollPart")
        scroll_headers = [HEADERS[x] for x in DISPLAY_FIELDS[FROZEN_COLUMNS:]]
        self.table.setHorizontalHeaderLabels(scroll_headers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.cellDoubleClicked.connect(lambda r, c: self.on_cell_double_click(r, c + FROZEN_COLUMNS))
        self.table.cellClicked.connect(lambda r, c: self.on_cell_clicked(r, c + FROZEN_COLUMNS))
        self.table.currentCellChanged.connect(
            lambda cr, cc, pr, pc: self.on_current_cell_changed(
                cr, cc + FROZEN_COLUMNS if cc >= 0 else -1, pr, pc
            )
        )
        self.table.customContextMenuRequested.connect(self.on_table_context_menu)
        hdr_main_f = HoverFilterHeaderView(self.table_frozen, self, "main_frozen")
        hdr_main_s = HoverFilterHeaderView(self.table, self, "main_scroll")
        self.table_frozen.setHorizontalHeader(hdr_main_f)
        self.table.setHorizontalHeader(hdr_main_s)
        hdr_main_f.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr_main_f.setSectionResizeMode(1, QHeaderView.Interactive)
        hdr_main_s.setSectionResizeMode(QHeaderView.Interactive)
        self.table.setItemDelegateForColumn(ALLOW_PRINT_SCROLL_COL, AllowPrintUrlCellDelegate(self.table))
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        # 列总宽很大时，勿把整表宽度当作窗口最小宽度，否则初始宽度无效且无法横向缩小
        self.table.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        hdr_main_f.sectionClicked.connect(self._on_main_frozen_header_section_clicked)
        hdr_main_s.sectionClicked.connect(self._on_main_scroll_header_clicked)
        self.table.horizontalHeader().sectionResized.connect(self._on_main_scroll_column_resized_for_url)
        self.table.horizontalHeader().sectionResized.connect(self._on_main_scroll_column_resized_sync_check_width)
        self.table_frozen.horizontalHeader().sectionResized.connect(self._on_main_frozen_section_resized)
        self.table.verticalScrollBar().valueChanged.connect(lambda v: self._sync_main_v_scroll(v, "scroll"))
        self.table_frozen.verticalScrollBar().valueChanged.connect(lambda v: self._sync_main_v_scroll(v, "frozen"))
        self.table.itemSelectionChanged.connect(self._sync_selection_main_to_frozen)
        self.table_frozen.itemSelectionChanged.connect(self._sync_selection_frozen_to_main)
        split_lo.addWidget(self.table_frozen, 0)
        split_lo.addWidget(self.table, 1)
        table_wrap = QWidget()
        table_wrap.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        table_layout = QVBoxLayout(table_wrap)
        table_layout.setContentsMargins(16, 10, 16, 16)
        table_layout.setSpacing(8)
        table_layout.addWidget(table_split, 1)
        bottom_actions = QHBoxLayout()
        bottom_actions.addWidget(self.btn_delete_sel)
        bottom_actions.addStretch(1)
        table_layout.addLayout(bottom_actions)
        bill_page_layout.addWidget(table_wrap, 1)
        self.content_stack.addWidget(bill_page)

        history_page = QWidget()
        history_page.setObjectName("stackHistoryPage")
        history_page.setAutoFillBackground(True)
        history_page.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        history_layout = QVBoxLayout(history_page)
        history_layout.setContentsMargins(16, 12, 16, 16)
        history_layout.setSpacing(10)
        history_top = QHBoxLayout()
        self.history_search = QLineEdit()
        self.history_search.setPlaceholderText("搜索任务名...")
        self.history_search.setFixedWidth(220)
        self.history_search.textChanged.connect(self.refresh_history_table)
        btn_history_search = QPushButton("🔍 搜索")
        btn_history_search.setObjectName("btnGhost")
        btn_history_search.clicked.connect(self.on_history_search_clicked)
        btn_history_clear = QPushButton("🔄 清空筛选")
        btn_history_clear.setObjectName("btnGhost")
        btn_history_clear.clicked.connect(self.clear_history_search_and_filters)
        history_top.addWidget(self.history_search)
        history_top.addWidget(btn_history_search)
        history_top.addWidget(btn_history_clear)
        history_top.addStretch(1)
        self.btn_restore_selected = QPushButton("↩ 恢复选中到提单表")
        self.btn_restore_selected.setObjectName("btnAccent")
        self.btn_restore_selected.clicked.connect(self.restore_selected_history)
        history_top.addWidget(self.btn_restore_selected)
        history_layout.addLayout(history_top)
        hist_split = QWidget()
        hist_split.setObjectName("historyTableSplit")
        hist_split.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        hist_lo = QHBoxLayout(hist_split)
        hist_lo.setContentsMargins(0, 0, 0, 0)
        hist_lo.setSpacing(0)
        self.history_table_frozen = QTableWidget(0, FROZEN_COLUMNS)
        self.history_table_frozen.setObjectName("tableFrozenCol")
        self.history_table_frozen.setHorizontalHeaderLabels(["", HEADERS["task_name"]])
        self.history_table_frozen.verticalHeader().setVisible(False)
        self.history_table_frozen.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table_frozen.setAlternatingRowColors(True)
        self.history_table_frozen.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.history_table_frozen.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.history_table_frozen.horizontalScrollBar().setEnabled(False)
        self.history_table_frozen.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.history_table_frozen.setFrameShape(QFrame.NoFrame)
        self.history_table_frozen.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        self.history_table_frozen.cellClicked.connect(lambda r, c: self.on_history_cell_clicked(r, c))
        self.history_table = QTableWidget(0, HISTORY_SCROLL_COLUMNS)
        self.history_table.setObjectName("historyScrollPart")
        history_scroll_headers = [HEADERS[f] for f in HISTORY_SCROLL_FIELDS if f != "deleted_at"] + ["删除时间"]
        self.history_table.setHorizontalHeaderLabels(history_scroll_headers)
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.cellClicked.connect(lambda r, c: self.on_history_cell_clicked(r, c + FROZEN_COLUMNS))
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        hdr_hist_f = HoverFilterHeaderView(self.history_table_frozen, self, "hist_frozen")
        hdr_hist_s = HoverFilterHeaderView(self.history_table, self, "hist_scroll")
        self.history_table_frozen.setHorizontalHeader(hdr_hist_f)
        self.history_table.setHorizontalHeader(hdr_hist_s)
        hdr_hist_f.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr_hist_f.setSectionResizeMode(1, QHeaderView.Interactive)
        hdr_hist_s.setSectionResizeMode(QHeaderView.Interactive)
        self.history_table.setItemDelegateForColumn(ALLOW_PRINT_SCROLL_COL, AllowPrintUrlCellDelegate(self.history_table))
        self.history_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.history_table.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        hdr_hist_f.sectionClicked.connect(self._on_hist_frozen_header_section_clicked)
        hdr_hist_s.sectionClicked.connect(self._on_history_scroll_header_clicked)
        self.history_table.horizontalHeader().sectionResized.connect(self._on_hist_scroll_column_resized_sync_check_width)
        self.history_table_frozen.horizontalHeader().sectionResized.connect(self._on_hist_frozen_section_resized)
        self.history_table.verticalScrollBar().valueChanged.connect(lambda v: self._sync_hist_v_scroll(v, "scroll"))
        self.history_table_frozen.verticalScrollBar().valueChanged.connect(lambda v: self._sync_hist_v_scroll(v, "frozen"))
        self.history_table.itemSelectionChanged.connect(self._sync_selection_hist_to_frozen)
        self.history_table_frozen.itemSelectionChanged.connect(self._sync_selection_hist_frozen_to_scroll)
        hist_lo.addWidget(self.history_table_frozen, 0)
        hist_lo.addWidget(self.history_table, 1)
        history_layout.addWidget(hist_split, 1)
        self.content_stack.addWidget(history_page)

        print_log_page = QWidget()
        print_log_page.setObjectName("stackPrintRecordsPage")
        print_log_page.setAutoFillBackground(True)
        print_log_page.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        pl_layout = QVBoxLayout(print_log_page)
        pl_layout.setContentsMargins(16, 12, 16, 16)
        pl_layout.setSpacing(10)
        pl_top = QHBoxLayout()
        self.print_log_search = QLineEdit()
        self.print_log_search.setPlaceholderText("搜索文件名或路径...")
        self.print_log_search.setFixedWidth(260)
        self.print_log_search.textChanged.connect(self.refresh_print_records_table)
        btn_pl_search = QPushButton("🔍 搜索")
        btn_pl_search.setObjectName("btnGhost")
        btn_pl_search.clicked.connect(self.on_print_log_search_clicked)
        btn_pl_clear = QPushButton("🔄 清空筛选")
        btn_pl_clear.setObjectName("btnGhost")
        btn_pl_clear.clicked.connect(self.clear_print_log_search_and_filters)
        self.btn_print_log_delete = QPushButton("🗑 批量删除选中")
        self.btn_print_log_delete.setObjectName("btnDanger")
        self.btn_print_log_delete.clicked.connect(self.delete_selected_print_records)
        pl_top.addWidget(self.print_log_search)
        pl_top.addWidget(btn_pl_search)
        pl_top.addWidget(btn_pl_clear)
        pl_top.addStretch(1)
        pl_top.addWidget(self.btn_print_log_delete)
        pl_layout.addLayout(pl_top)

        pl_headers = [""] + [PRINT_LOG_HEADERS[k] for k in PRINT_LOG_DATA_FIELDS] + ["操作"]
        self.print_log_table = QTableWidget(0, PRINT_LOG_COL_COUNT)
        self.print_log_table.setObjectName("historyScrollPart")
        self.print_log_table.setHorizontalHeaderLabels(pl_headers)
        self.print_log_table.verticalHeader().setVisible(False)
        self.print_log_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.print_log_table.setAlternatingRowColors(True)
        self.print_log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.print_log_table.cellClicked.connect(self.on_print_log_cell_clicked)
        hdr_print = HoverFilterHeaderView(self.print_log_table, self, "print_rec")
        self.print_log_table.setHorizontalHeader(hdr_print)
        hdr_print.setSectionResizeMode(QHeaderView.Interactive)
        hdr_print.sectionClicked.connect(self._on_print_rec_header_section_clicked)
        self.print_log_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.print_log_table.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
        pl_layout.addWidget(self.print_log_table, 1)
        self.content_stack.addWidget(print_log_page)

        settings_page = QWidget()
        settings_page.setObjectName("stackSettingsPage")
        settings_page.setAutoFillBackground(True)
        settings_layout = QVBoxLayout(settings_page)
        settings_layout.setContentsMargins(24, 24, 24, 24)
        settings_layout.setSpacing(10)
        settings_title = QLabel("设置")
        settings_title.setObjectName("settingsTitle")
        settings_layout.addWidget(settings_title)
        settings_layout.addWidget(QLabel("主题模式"))
        self.btn_theme = QPushButton("🎨 切换主题")
        self.btn_theme.setObjectName("btnGhost")
        self.btn_theme.clicked.connect(self.toggle_theme)
        settings_layout.addWidget(self.btn_theme)
        settings_layout.addStretch(1)
        self.content_stack.addWidget(settings_page)

        status = QStatusBar(self)
        self.setStatusBar(status)
        self.lbl_count = QLabel("提单数: 0")
        self.lbl_time = QLabel()
        self.lbl_save = QLabel("已自动保存")
        status.addPermanentWidget(self.lbl_count)
        status.addPermanentWidget(QLabel("|"))
        status.addPermanentWidget(self.lbl_time)
        status.addPermanentWidget(QLabel("|"))
        status.addPermanentWidget(self.lbl_save)
        timer = QTimer(self)
        timer.timeout.connect(self.update_time)
        timer.start(1000)
        self.update_time()
        root.addWidget(main_wrap, 1)
        self.show_bill_page()

        # 与滚动区列顺序一致：类型…行业、「允许」、URL…；长度须 >= MAIN_SCROLL_COLUMNS+1（含 col_widths[0] 给冻结任务名列）
        # 滚动区列序与 col_widths[1:] 对齐；末四列为 提单时间、打印次数、打印时间、操作（原打印相关宽顺序已随列序调整）
        col_widths = [260, 150, 150, 120, 72, 260, 90, 130, 90, 90, 90, 160, 160, 160, 160, 168, 170, 72, 90, 100]
        self.table_frozen.setColumnWidth(1, col_widths[0])
        for s in range(MAIN_SCROLL_COLUMNS):
            self.table.setColumnWidth(s, col_widths[s + 1])
        self._sync_main_frozen_check_column_width()
        self.history_table_frozen.setColumnWidth(1, col_widths[0])
        hist_widths = col_widths + [160]
        for s in range(HISTORY_SCROLL_COLUMNS):
            self.history_table.setColumnWidth(s, hist_widths[s + 1])
        self._sync_hist_frozen_check_column_width()
        _pw = [300, 168, 88, 100, 88, 380, 168]
        for s, w in enumerate(_pw, start=1):
            self.print_log_table.setColumnWidth(s, w)
        self.table_frozen.itemChanged.connect(self.on_item_changed)
        self._apply_initial_window_geometry()
        self._on_main_frozen_section_resized()
        self._on_hist_frozen_section_resized()

    def _apply_initial_window_geometry(self):
        """初始窗口宽度 1350；高度为可用区估算值再加 50（无屏信息时 910）；窗口可自由拉伸。"""
        w, h = 1350, 910
        screen = QGuiApplication.primaryScreen()
        if screen is not None:
            ag = screen.availableGeometry()
            h = max(520, min(int(ag.height() * 0.88), ag.height() - 8)) + 50
            h = min(h, ag.height() - 8)
            x = ag.left() + max(0, (ag.width() - w) // 2)
            y = ag.top() + max(0, (ag.height() - h) // 2)
            self.setGeometry(x, y, w, h)
        # 统一再设客户区宽高，避免最小宽度约束导致实际宽度大于目标值
        self.resize(w, h)

    def showEvent(self, event):
        super().showEvent(event)
        if getattr(self, "_bill_post_show_geo", False):
            return
        self._bill_post_show_geo = True
        QTimer.singleShot(0, self._apply_initial_window_geometry)

    def _sync_main_frozen_check_column_width(self):
        """冻结首列宽与右侧滚动表中「允许」列一致。"""
        hf = self.table_frozen.horizontalHeader()
        hf.blockSignals(True)
        try:
            w = max(52, self.table.columnWidth(ALLOW_PRINT_SCROLL_COL))
            self.table_frozen.setColumnWidth(0, w)
        finally:
            hf.blockSignals(False)
        self._on_main_frozen_section_resized()
        self._sync_print_log_check_column_width()

    def _sync_print_log_check_column_width(self):
        """打印记录表首列宽与提单页冻结勾选列（「允许」列宽）一致。"""
        w = max(52, self.table.columnWidth(ALLOW_PRINT_SCROLL_COL))
        self.print_log_table.setColumnWidth(0, w)

    def _sync_hist_frozen_check_column_width(self):
        hf = self.history_table_frozen.horizontalHeader()
        hf.blockSignals(True)
        try:
            w = max(52, self.history_table.columnWidth(ALLOW_PRINT_SCROLL_COL))
            self.history_table_frozen.setColumnWidth(0, w)
        finally:
            hf.blockSignals(False)
        self._on_hist_frozen_section_resized()

    def _on_main_scroll_column_resized_sync_check_width(self, logical_index: int, _old: int, _new: int):
        if logical_index == ALLOW_PRINT_SCROLL_COL:
            self._sync_main_frozen_check_column_width()

    def _on_hist_scroll_column_resized_sync_check_width(self, logical_index: int, _old: int, _new: int):
        if logical_index == ALLOW_PRINT_SCROLL_COL:
            self._sync_hist_frozen_check_column_width()

    def _on_main_frozen_section_resized(self, *_args):
        """冻结区仅横向不滚动：首列随「允许」列宽，任务名列可调；总宽随两列之和更新。"""
        hf = self.table_frozen.horizontalHeader()
        hf.blockSignals(True)
        try:
            w_allow = max(52, self.table.columnWidth(ALLOW_PRINT_SCROLL_COL))
            w0 = w_allow
            w1 = max(120, min(self.table_frozen.columnWidth(1), 2000))
            self.table_frozen.setColumnWidth(0, w0)
            self.table_frozen.setColumnWidth(1, w1)
            self.table_frozen.setFixedWidth(w0 + w1)
        finally:
            hf.blockSignals(False)

    def _on_hist_frozen_section_resized(self, *_args):
        hf = self.history_table_frozen.horizontalHeader()
        hf.blockSignals(True)
        try:
            w_allow = max(52, self.history_table.columnWidth(ALLOW_PRINT_SCROLL_COL))
            w0 = w_allow
            w1 = max(120, min(self.history_table_frozen.columnWidth(1), 2000))
            self.history_table_frozen.setColumnWidth(0, w0)
            self.history_table_frozen.setColumnWidth(1, w1)
            self.history_table_frozen.setFixedWidth(w0 + w1)
        finally:
            hf.blockSignals(False)

    def load_theme(self):
        if THEME_FILE.exists():
            try:
                return json.loads(THEME_FILE.read_text(encoding="utf-8")).get("theme", "light")
            except Exception:
                return "light"
        return "light"

    def save_theme(self):
        THEME_FILE.write_text(json.dumps({"theme": self.theme}, ensure_ascii=False, indent=2), encoding="utf-8")

    def apply_theme(self):
        self.setStyleSheet(STYLESHEET_DARK if self.theme == "dark" else STYLESHEET_LIGHT)
        self._refresh_sidebar_logo()
        self._refresh_sidebar_toggle_button()
        self._update_print_url_hint_lines_style()

    def _refresh_sidebar_toggle_button(self):
        """侧栏开：实心 ><（收起）；侧栏关：实心 <>（展开）。"""
        dark = self.theme == "dark"
        visible = bool(getattr(self, "_sidebar_visible", True))
        pm = make_sidebar_chevrons_pixmap(collapse=visible, dark=dark, size=24)
        self.btn_toggle_sidebar.setIcon(QIcon(pm))
        self.btn_toggle_sidebar.setText("")
        if visible:
            self.btn_toggle_sidebar.setToolTip("隐藏侧栏菜单")
        else:
            self.btn_toggle_sidebar.setToolTip("展开侧栏菜单")

    def _refresh_sidebar_logo(self):
        if not hasattr(self, "_sidebar_logo_icon"):
            return
        self._sidebar_logo_icon.setPixmap(make_sidebar_logo_pixmap(dark=self.theme == "dark", size=44))

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self.apply_theme()
        self.save_theme()

    def toggle_sidebar(self):
        self._sidebar_visible = not bool(getattr(self, "_sidebar_visible", True))
        self.sidebar.setVisible(self._sidebar_visible)
        self._refresh_sidebar_toggle_button()

    def _update_print_url_hint_lines_style(self, *_args):
        """未勾选 / 勾选说明各占一行；当前状态对应行加粗。"""
        if not hasattr(self, "lbl_print_url_hint_off"):
            return
        checked = self.chk_print_url.isChecked()
        base = self.font()
        f_on = QFont(base)
        f_on.setBold(True)
        f_off = QFont(base)
        f_off.setBold(False)
        self.lbl_print_url_hint_off.setFont(f_on if not checked else f_off)
        self.lbl_print_url_hint_on.setFont(f_on if checked else f_off)

    def _restyle_sidebar_nav(self, page: str):
        styles = {
            "bill": ("navActive", "navNormal", "navNormal", "navNormal"),
            "history": ("navNormal", "navActive", "navNormal", "navNormal"),
            "print": ("navNormal", "navNormal", "navActive", "navNormal"),
            "settings": ("navNormal", "navNormal", "navNormal", "navActive"),
        }
        names = styles[page]
        for w, oname in zip((self.nav_bill, self.nav_history, self.nav_print_records, self.nav_settings), names):
            w.setObjectName(oname)
            w.style().unpolish(w)
            w.style().polish(w)

    def show_bill_page(self):
        self.content_stack.setCurrentIndex(0)
        self._restyle_sidebar_nav("bill")
        self.apply_theme()

    def show_history_page(self):
        self.content_stack.setCurrentIndex(1)
        self._restyle_sidebar_nav("history")
        self.refresh_history_table()
        self.apply_theme()

    def show_print_records_page(self):
        self.content_stack.setCurrentIndex(2)
        self._restyle_sidebar_nav("print")
        self.refresh_print_records_table()
        self.apply_theme()

    def show_settings_page(self):
        self.content_stack.setCurrentIndex(3)
        self._restyle_sidebar_nav("settings")
        self.apply_theme()

    def update_time(self):
        self.lbl_time.setText(datetime.now().strftime("%Y/%m/%d %H:%M:%S"))

    def default_record(self):
        d = {k: "" for k in FIELDS} | {"checked": False, "created_at": ""}
        d["allow_print_url"] = True
        d["print_count"] = 0
        d["last_printed_at"] = ""
        return d

    def load_data(self):
        if DATA_FILE.exists():
            try:
                self.records = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            except Exception:
                self.records = []
        for rec in self.records:
            rec.setdefault("created_at", "")
            rec["allow_print_url"] = coerce_allow_print_url(rec.get("allow_print_url"))
            rec.setdefault("print_count", 0)
            try:
                rec["print_count"] = int(rec.get("print_count", 0) or 0)
            except (TypeError, ValueError):
                rec["print_count"] = 0
            rec.setdefault("last_printed_at", "")
            rec["last_printed_at"] = str(rec.get("last_printed_at", "") or "")
        if not self.records:
            self.records = [self.default_record()]

    def load_history_data(self):
        if HISTORY_FILE.exists():
            try:
                self.history_records = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            except Exception:
                self.history_records = []
        for rec in self.history_records:
            rec.setdefault("checked", False)
            rec.setdefault("deleted_at", "")
            rec.setdefault("created_at", "")
            rec["allow_print_url"] = coerce_allow_print_url(rec.get("allow_print_url"))
            rec.setdefault("print_count", 0)
            try:
                rec["print_count"] = int(rec.get("print_count", 0) or 0)
            except (TypeError, ValueError):
                rec["print_count"] = 0
            rec.setdefault("last_printed_at", "")
            rec["last_printed_at"] = str(rec.get("last_printed_at", "") or "")

    def load_picker_recent(self):
        if PICKER_RECENT_FILE.exists():
            try:
                data = json.loads(PICKER_RECENT_FILE.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    self.picker_recent = {str(k): list(v) for k, v in data.items()}
            except Exception:
                self.picker_recent = {}

    def save_picker_recent(self):
        PICKER_RECENT_FILE.write_text(json.dumps(self.picker_recent, ensure_ascii=False, indent=2), encoding="utf-8")

    def update_picker_recent(self, field: str, values: list[str]):
        old = self.picker_recent.get(field, [])
        merged = values + [x for x in old if x not in values]
        self.picker_recent[field] = merged[:20]
        self.save_picker_recent()

    def save_data(self):
        stamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        for rec in self.records:
            if self.is_row_saved(rec) and not str(rec.get("created_at", "")).strip():
                rec["created_at"] = stamp
        DATA_FILE.write_text(json.dumps(self.records, ensure_ascii=False, indent=2), encoding="utf-8")
        self.lbl_save.setText("已自动保存")

    def save_history_data(self):
        HISTORY_FILE.write_text(json.dumps(self.history_records, ensure_ascii=False, indent=2), encoding="utf-8")
        self.lbl_save.setText("已自动保存")

    def load_print_records(self):
        if PRINT_RECORDS_FILE.exists():
            try:
                raw = json.loads(PRINT_RECORDS_FILE.read_text(encoding="utf-8"))
                rows = [x for x in raw if isinstance(x, dict)] if isinstance(raw, list) else []
            except Exception:
                rows = []
        else:
            rows = []
        for rec in rows:
            rec.setdefault("id", str(uuid.uuid4()))
            rec.setdefault("path", "")
            p = str(rec.get("path", "") or "").strip()
            if p:
                rec["path"] = p
            fn = str(rec.get("filename", "") or "").strip()
            if not fn and p:
                rec["filename"] = os.path.basename(p)
            elif not fn:
                rec["filename"] = "导出.xlsx"
            try:
                rec["row_count"] = int(rec.get("row_count", 0) or 0)
            except (TypeError, ValueError):
                rec["row_count"] = 0
            if "include_print_url" not in rec:
                rec["include_print_url"] = False
            else:
                rec["include_print_url"] = bool(rec.get("include_print_url"))
        self.print_records = [r for r in rows if str(r.get("path", "") or "").strip()]

    def save_print_records(self):
        clean = [{k: v for k, v in r.items() if k != "checked"} for r in self.print_records]
        PRINT_RECORDS_FILE.write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8")
        self.lbl_save.setText("已自动保存")

    @staticmethod
    def _print_record_file_ok(rec: dict) -> bool:
        p = str(rec.get("path", "") or "").strip()
        if not p:
            return False
        try:
            return Path(p).is_file()
        except OSError:
            return False

    def _print_cell_display_for_filter(self, rec: dict, field: str) -> str:
        if field == "row_count":
            return str(int(rec.get("row_count", 0) or 0))
        if field == "include_print_url":
            return "是" if rec.get("include_print_url") else "否"
        if field == "file_exists":
            return "存在" if self._print_record_file_ok(rec) else "缺失"
        if field == "filename":
            return str(rec.get("filename", "") or "")
        if field == "printed_at":
            return BillApp._created_at_filter_key(rec.get("printed_at"))
        if field == "path":
            return str(rec.get("path", "") or "")
        return ""

    def _print_row_matches_header_filters_except(self, rec: dict, skip_field: str | None) -> bool:
        for field, vals in self.print_header_filters.items():
            if skip_field and field == skip_field:
                continue
            if not vals:
                continue
            if field == "printed_at":
                dk = BillApp._created_at_filter_key(rec.get("printed_at"))
                allowed = {BillApp._created_at_filter_key(x) for x in vals}
                if dk not in allowed:
                    return False
            elif self._print_cell_display_for_filter(rec, field) not in vals:
                return False
        return True

    def _print_row_matches_header_filters(self, rec: dict) -> bool:
        return self._print_row_matches_header_filters_except(rec, None)

    def _unique_display_values_print(self, field: str) -> list[str]:
        query = self.print_log_search.text().strip().lower()
        out: list[str] = []
        seen: set[str] = set()
        for r in self.print_records:
            if query:
                fn = str(r.get("filename", "")).lower()
                pt = str(r.get("path", "")).lower()
                if query not in fn and query not in pt:
                    continue
            if not self._print_row_matches_header_filters_except(r, field):
                continue
            dv = self._print_cell_display_for_filter(r, field)
            if dv not in seen:
                seen.add(dv)
                out.append(dv)
        if field == "printed_at":
            out = BillApp._unique_created_at_filter_options(out)
        else:
            out.sort(key=lambda s: (s == "", s.casefold()))
        return out

    def _sort_key_print_field(self, rec: dict, field: str) -> Any:
        if field == "row_count":
            try:
                return int(rec.get("row_count", 0) or 0)
            except (TypeError, ValueError):
                return -1
        if field == "include_print_url":
            return 1 if rec.get("include_print_url") else 0
        if field == "file_exists":
            return 1 if self._print_record_file_ok(rec) else 0
        if field == "printed_at":
            return str(rec.get("printed_at", "") or "").strip()
        if field == "filename":
            return str(rec.get("filename", "") or "").lower()
        if field == "path":
            return str(rec.get("path", "") or "").lower()
        return ""

    def _toggle_print_sort_for_field(self, field: str):
        if self.print_sort_field == field:
            if self.print_sort_order == Qt.SortOrder.AscendingOrder:
                self.print_sort_order = Qt.SortOrder.DescendingOrder
            else:
                self.print_sort_field = None
        else:
            self.print_sort_field = field
            self.print_sort_order = Qt.SortOrder.AscendingOrder
        self.refresh_print_records_table()

    def _on_print_rec_header_section_clicked(self, section: int):
        if section < 1 or section > len(PRINT_LOG_DATA_FIELDS):
            return
        field = PRINT_LOG_DATA_FIELDS[section - 1]
        self._toggle_print_sort_for_field(field)

    def _update_print_log_header_sort_indicator(self):
        h = self.print_log_table.horizontalHeader()
        h.setSortIndicatorShown(False)
        if not self.print_sort_field or self.print_sort_field not in PRINT_LOG_DATA_FIELDS:
            return
        si = PRINT_LOG_DATA_FIELDS.index(self.print_sort_field) + 1
        h.setSortIndicatorShown(True)
        h.setSortIndicator(si, self.print_sort_order)

    def _update_print_log_header_tooltips(self):
        for s in range(self.print_log_table.columnCount()):
            field = self._field_for_header_section("print_rec", s)
            if not field:
                continue
            title = PRINT_LOG_HEADERS.get(field, field)
            tip = title + "\n单击：排序；悬停右侧 ▼ 筛选"
            vals = self.print_header_filters.get(field)
            if vals:
                tip += f"\n已选 {len(vals)} 项"
            it = self.print_log_table.horizontalHeaderItem(s)
            if it:
                it.setToolTip(tip)

    def update_print_log_header_check(self):
        it = self.print_log_table.horizontalHeaderItem(0)
        if not it:
            return
        info = self._select_all_header_paint_info("print_rec")
        if not self.print_log_filtered_indices:
            it.setToolTip("当前筛选下列表为空")
        elif info["glyph"] == "x":
            it.setToolTip("单击：取消全选（清除当前列表中所有勾选）")
        else:
            it.setToolTip("单击：全选（将当前列表全部行设为勾选；部分勾选时则全选）")
        self.print_log_table.horizontalHeader().viewport().update()

    def _on_print_log_header_col0_clicked(self):
        if not self.print_log_filtered_indices:
            return
        all_checked = all(self.print_records[i].get("checked") for i in self.print_log_filtered_indices)
        for i in self.print_log_filtered_indices:
            self.print_records[i]["checked"] = not all_checked
        self.refresh_print_records_table()

    def on_print_log_cell_clicked(self, row: int, col: int):
        if self.updating_table or col != 0:
            return
        if row < 0 or row >= len(self.print_log_filtered_indices):
            return
        rec_idx = self.print_log_filtered_indices[row]
        cur = bool(self.print_records[rec_idx].get("checked"))
        self.print_records[rec_idx]["checked"] = not cur
        self.refresh_print_records_table()

    def on_print_log_checkbox_changed(self, rec_idx: int, checked: bool):
        if self.updating_table:
            return
        if 0 <= rec_idx < len(self.print_records):
            self.print_records[rec_idx]["checked"] = bool(checked)
            self.refresh_print_records_table()

    def on_print_log_search_clicked(self):
        self.print_header_filters.clear()
        self.refresh_print_records_table()

    def clear_print_log_search_and_filters(self):
        self.print_log_search.setText("")
        self.print_header_filters.clear()
        self.print_sort_field = "printed_at"
        self.print_sort_order = Qt.SortOrder.DescendingOrder
        self.refresh_print_records_table()

    def refresh_print_records_table(self):
        self.updating_table = True
        query = self.print_log_search.text().strip().lower()
        self.print_log_filtered_indices = [
            i
            for i, r in enumerate(self.print_records)
            if (
                not query
                or query in str(r.get("filename", "")).lower()
                or query in str(r.get("path", "")).lower()
            )
            and self._print_row_matches_header_filters(r)
        ]
        if not self.print_sort_field:
            self.print_log_filtered_indices.sort(
                key=lambda i: str(self.print_records[i].get("printed_at", "") or ""), reverse=True
            )
        else:
            f = self.print_sort_field
            rev = self.print_sort_order == Qt.SortOrder.DescendingOrder
            self.print_log_filtered_indices.sort(
                key=lambda i, ff=f: self._sort_key_print_field(self.print_records[i], ff), reverse=rev
            )
        n = len(self.print_log_filtered_indices)
        self.print_log_table.setRowCount(n)
        h = self._data_row_height
        for row, rec_idx in enumerate(self.print_log_filtered_indices):
            rec = self.print_records[rec_idx]
            box_wrap = QWidget()
            box_layout = QHBoxLayout(box_wrap)
            box_layout.setContentsMargins(0, 0, 0, 0)
            box_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox()
            chk.setChecked(bool(rec.get("checked")))
            chk.stateChanged.connect(lambda _s, x=rec_idx, c=chk: self.on_print_log_checkbox_changed(x, c.isChecked()))
            box_layout.addWidget(chk)
            self.print_log_table.setCellWidget(row, 0, box_wrap)
            ph0 = QTableWidgetItem("")
            ph0.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.print_log_table.setItem(row, 0, ph0)

            fn = str(rec.get("filename", "") or "")
            it_fn = QTableWidgetItem(fn)
            it_fn.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            it_fn.setToolTip(fn)
            it_fn.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.print_log_table.setItem(row, 1, it_fn)

            it_time = QTableWidgetItem(str(rec.get("printed_at", "") or ""))
            it_time.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it_time.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.print_log_table.setItem(row, 2, it_time)

            rc = int(rec.get("row_count", 0) or 0)
            it_rc = QTableWidgetItem(str(rc))
            it_rc.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it_rc.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.print_log_table.setItem(row, 3, it_rc)

            iurl = QTableWidgetItem("是" if rec.get("include_print_url") else "否")
            iurl.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            iurl.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.print_log_table.setItem(row, 4, iurl)

            ok = self._print_record_file_ok(rec)
            st = QTableWidgetItem("存在" if ok else "缺失")
            st.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            st.setForeground(QColor("#2e8b57" if ok else "#c93042"))
            st.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.print_log_table.setItem(row, 5, st)

            pth = str(rec.get("path", "") or "")
            disp = pth if len(pth) <= 64 else pth[:30] + "…" + pth[-28:]
            it_p = QTableWidgetItem(disp)
            it_p.setTextAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            it_p.setToolTip(pth)
            it_p.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.print_log_table.setItem(row, 6, it_p)

            act = QWidget()
            al = QHBoxLayout(act)
            al.setContentsMargins(4, 2, 4, 2)
            al.setSpacing(6)
            bp = QPushButton("预览")
            bp.setObjectName("btnGhost")
            bp.clicked.connect(lambda _=False, p=pth: self.preview_print_record_file(p))
            bd = QPushButton("打开")
            bd.setObjectName("btnGhost")
            bd.clicked.connect(lambda _=False, p=pth: self.open_print_record_file(p))
            al.addWidget(bp)
            al.addWidget(bd)
            self.print_log_table.setCellWidget(row, 7, act)

        for r in range(n):
            self.print_log_table.setRowHeight(r, h)
        self._update_print_log_header_sort_indicator()
        self._update_print_log_header_tooltips()
        self.update_print_log_header_check()
        self.updating_table = False

    def preview_print_record_file(self, path: str):
        p = Path(path)
        if not p.is_file():
            QMessageBox.warning(self, "预览", "文件不存在或已被移动，可勾选后批量删除无效记录。")
            return
        if p.suffix.lower() != ".xlsx":
            QMessageBox.information(self, "预览", "暂仅支持预览 .xlsx 表格内容。")
            return
        hint = ""
        try:
            wb = load_workbook(str(p), read_only=True, data_only=True)
            try:
                ws = wb[wb.sheetnames[0]]
                max_rows = min(int(ws.max_row or 1), 200)
                max_cols = min(int(ws.max_column or 1), 64)
                lines: list[str] = []
                for row in ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols, values_only=True):
                    cells: list[str] = []
                    for v in row:
                        if v is None:
                            cells.append("")
                        else:
                            cells.append(str(v).replace("\t", " ").replace("\n", " "))
                    lines.append("\t".join(cells))
                text = "\n".join(lines)
                if (ws.max_row or 0) > 200:
                    hint = f"\n\n…（仅显示前 200 行，共 {ws.max_row} 行）"
            finally:
                wb.close()
        except Exception as e:
            QMessageBox.warning(self, "预览失败", str(e))
            return
        dlg = QDialog(self)
        dlg.setWindowTitle(f"预览 — {p.name}")
        dlg.resize(960, 720)
        dlg.setMinimumSize(640, 400)
        te = QPlainTextEdit()
        te.setReadOnly(True)
        te.setPlainText(text + hint)
        te.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        te.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        te.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        lo = QVBoxLayout(dlg)
        lo.setContentsMargins(8, 8, 8, 8)
        lo.addWidget(te, 1)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        bb.rejected.connect(dlg.reject)
        lo.addWidget(bb)
        dlg.exec()

    def open_print_record_file(self, path: str):
        p = Path(path)
        if not p.is_file():
            QMessageBox.warning(self, "打开", "文件不存在或已被移动。")
            return
        url = QUrl.fromLocalFile(str(p.resolve()))
        if not QDesktopServices.openUrl(url):
            QMessageBox.warning(self, "打开失败", "系统无法用默认应用打开该文件。")

    def delete_selected_print_records(self):
        to_del = [i for i, r in enumerate(self.print_records) if r.get("checked")]
        if not to_del:
            QMessageBox.information(self, "提示", "请先勾选要删除的打印记录")
            return
        if (
            QMessageBox.question(
                self,
                "确认删除",
                f"将删除 {len(to_del)} 条打印记录。仅从列表移除记录，不会删除磁盘上的 Excel 文件。",
            )
            != QMessageBox.Yes
        ):
            return
        for idx in sorted(to_del, reverse=True):
            self.print_records.pop(idx)
        self.save_print_records()
        self.refresh_print_records_table()

    def add_row(self):
        insert_at = self.filtered_indices[0] if self.filtered_indices else 0
        self.records.insert(insert_at, self.default_record())
        self._shift_field_errors_after_insert(insert_at)
        self.save_data()
        self.refresh_table()

    def is_row_saved(self, rec):
        return any(str(rec.get(k, "")).strip() for k in FIELDS if k != "allow_print_url")

    def delete_selected(self):
        to_del = [i for i in range(len(self.records)) if self.records[i].get("checked")]
        if not to_del:
            QMessageBox.information(self, "提示", "请先勾选要删除的数据")
            return
        has_saved = any(self.is_row_saved(self.records[i]) for i in to_del)
        msg = "选中的记录中包含已保存数据，确认删除吗？此操作不可撤销。" if has_saved else "确认删除选中的记录吗？"
        if QMessageBox.question(self, "确认删除", msg) != QMessageBox.Yes:
            return
        for idx in sorted(to_del, reverse=True):
            rec = self.records[idx]
            if self.is_row_saved(rec):
                self.archive_record(rec)
            self.records.pop(idx)
            self._shift_field_errors_after_delete(idx)
        self.save_data()
        self.save_history_data()
        self.refresh_table()

    def delete_row(self, idx):
        if idx < 0 or idx >= len(self.records):
            return
        has_saved = self.is_row_saved(self.records[idx])
        msg = "该行包含已保存数据，确认删除吗？此操作不可撤销。" if has_saved else "确认删除该行吗？"
        if QMessageBox.question(self, "确认删除", msg) != QMessageBox.Yes:
            return
        rec = self.records[idx]
        if self.is_row_saved(rec):
            self.archive_record(rec)
        self.records.pop(idx)
        self._shift_field_errors_after_delete(idx)
        self.save_data()
        self.save_history_data()
        self.refresh_table()

    def archive_record(self, rec):
        hist = dict(rec)
        hist["checked"] = False
        hist["deleted_at"] = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        self.history_records.insert(0, hist)

    def clear_history_search_and_filters(self):
        self.history_search.setText("")
        self.history_header_filters.clear()
        self.history_sort_field = None
        self.refresh_history_table()

    def refresh_history_table(self):
        self.updating_table = True
        query = self.history_search.text().strip().lower()
        self.history_filtered_indices = [
            i
            for i, r in enumerate(self.history_records)
            if (not query or query in str(r.get("task_name", "")).lower()) and self._history_row_matches_header_filters(r)
        ]
        if not self.history_sort_field:
            self.history_filtered_indices.sort(
                key=lambda i: self._created_at_sort_key(self.history_records[i]), reverse=True
            )
        else:
            f = self.history_sort_field
            rev = self.history_sort_order == Qt.SortOrder.DescendingOrder
            self.history_filtered_indices.sort(
                key=lambda i, ff=f: self._sort_key_history_field(self.history_records[i], ff),
                reverse=rev,
            )
        n = len(self.history_filtered_indices)
        self.history_table.setRowCount(n)
        self.history_table_frozen.setRowCount(n)
        for row, rec_idx in enumerate(self.history_filtered_indices):
            rec = self.history_records[rec_idx]
            box_wrap = QWidget()
            box_layout = QHBoxLayout(box_wrap)
            box_layout.setContentsMargins(0, 0, 0, 0)
            box_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox()
            chk.setChecked(bool(rec.get("checked")))
            chk.stateChanged.connect(lambda _=0, x=rec_idx, c=chk: self.on_history_checkbox_changed(x, c.isChecked()))
            box_layout.addWidget(chk)
            self.history_table_frozen.setCellWidget(row, 0, box_wrap)
            ph0 = QTableWidgetItem("")
            ph0.setFlags(Qt.ItemIsEnabled)
            self.history_table_frozen.setItem(row, 0, ph0)
            tn_item = QTableWidgetItem(self.display_val(rec, "task_name"))
            tn_item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            tn_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.history_table_frozen.setItem(row, 1, tn_item)
            for sc, f in enumerate(FIELDS[1:]):
                if f == "allow_print_url":
                    yes = coerce_allow_print_url(rec.get("allow_print_url"))
                    aw_item = QTableWidgetItem("是" if yes else "否")
                    aw_item.setFont(QFont(self.history_table_frozen.font()))
                    aw_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    aw_item.setToolTip("允许打印URL：是" if yes else "允许打印URL：否")
                    aw_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                    self.history_table.setItem(row, sc, aw_item)
                    continue
                item = QTableWidgetItem(self.display_val(rec, f))
                item.setTextAlignment(Qt.AlignCenter)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.history_table.setItem(row, sc, item)
            ca_hist = QTableWidgetItem(str(rec.get("created_at", "")))
            ca_hist.setTextAlignment(Qt.AlignCenter)
            ca_hist.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.history_table.setItem(row, len(FIELDS) - 1, ca_hist)
            pc_hist = QTableWidgetItem(str(int(rec.get("print_count", 0) or 0)))
            pc_hist.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            pc_hist.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.history_table.setItem(row, len(FIELDS), pc_hist)
            pt_hist = QTableWidgetItem(str(rec.get("last_printed_at", "") or ""))
            pt_hist.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            pt_hist.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.history_table.setItem(row, len(FIELDS) + 1, pt_hist)
            deleted_item = QTableWidgetItem(str(rec.get("deleted_at", "")))
            deleted_item.setTextAlignment(Qt.AlignCenter)
            deleted_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.history_table.setItem(row, len(FIELDS) + 2, deleted_item)
        h = self._data_row_height
        for r in range(n):
            self.history_table.setRowHeight(r, h)
            self.history_table_frozen.setRowHeight(r, h)
        self._update_history_header_sort_indicator()
        self._update_history_header_tooltips()
        self.update_history_header_check()
        self.updating_table = False

    def update_history_header_check(self):
        it = self.history_table_frozen.horizontalHeaderItem(0)
        if not it:
            return
        info = self._select_all_header_paint_info("hist_frozen")
        if not self.history_filtered_indices:
            it.setToolTip("当前筛选下列表为空")
        elif info["glyph"] == "x":
            it.setToolTip("单击：取消全选（清除当前列表中所有勾选）")
        else:
            it.setToolTip("单击：全选（将当前列表全部行设为勾选；部分勾选时则全选）")
        self.history_table_frozen.horizontalHeader().viewport().update()

    def on_history_cell_clicked(self, row, col):
        if self.updating_table or col != 0:
            return
        if row < 0 or row >= len(self.history_filtered_indices):
            return
        rec_idx = self.history_filtered_indices[row]
        self.on_history_checkbox_changed(rec_idx, not self.history_records[rec_idx].get("checked", False))

    def on_history_checkbox_changed(self, rec_idx: int, checked: bool):
        if self.updating_table:
            return
        if rec_idx < 0 or rec_idx >= len(self.history_records):
            return
        self.history_records[rec_idx]["checked"] = bool(checked)
        self.save_history_data()
        self.refresh_history_table()

    def restore_selected_history(self):
        restore_idx = [i for i, rec in enumerate(self.history_records) if rec.get("checked")]
        if not restore_idx:
            QMessageBox.information(self, "提示", "请先勾选要恢复的数据")
            return
        if QMessageBox.question(self, "确认恢复", f"确认恢复 {len(restore_idx)} 条历史提单到提单表吗？") != QMessageBox.Yes:
            return
        for idx in sorted(restore_idx):
            src = self.history_records[idx]
            rec = {k: src.get(k, "") for k in FIELDS}
            rec["allow_print_url"] = coerce_allow_print_url(src.get("allow_print_url"))
            try:
                rec["print_count"] = int(src.get("print_count", 0) or 0)
            except (TypeError, ValueError):
                rec["print_count"] = 0
            rec["last_printed_at"] = str(src.get("last_printed_at", "") or "")
            rec["checked"] = False
            rec["created_at"] = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            self.records.append(rec)
        for idx in sorted(restore_idx, reverse=True):
            self.history_records.pop(idx)
        self.save_data()
        self.save_history_data()
        self.refresh_table()
        self.refresh_history_table()
        QMessageBox.information(self, "恢复成功", f"已恢复 {len(restore_idx)} 条数据到提单表")

    def display_val(self, rec, field):
        if field == "type_code":
            return TYPE_MAP.get(rec.get(field, ""), "")
        if field == "operator_code":
            return OP_MAP.get(rec.get(field, ""), "")
        if field == "allow_print_url":
            return "是" if coerce_allow_print_url(rec.get("allow_print_url")) else "否"
        if field == "url":
            return first_url(rec.get(field, ""))
        if field == "print_count":
            return str(int(rec.get("print_count", 0) or 0))
        if field == "last_printed_at":
            return str(rec.get("last_printed_at", "") or "")
        return str(rec.get(field, ""))

    @staticmethod
    def _geo_tokens_invalid(field: str, parts: list[str], rec: dict | None = None) -> list[str]:
        if field in ("province", "exclude_province"):
            allowed = set(PROVINCES)
            return [p for p in parts if p not in allowed]
        allow_cities: set[str] | None = None
        if field == "exclude_city" and rec is not None:
            provs = split_multi(rec.get("province", ""))
            if provs:
                allow_cities = set(cities_under_provinces(provs))
        if allow_cities is not None:
            return [p for p in parts if p not in allow_cities]
        return [p for p in parts if p not in set(CITIES)]

    @staticmethod
    def _normalize_geo_field_value(field: str, raw: str, rec: dict | None = None) -> str | None:
        text = str(raw or "").replace("\n", " ").replace("\r", " ").strip()
        if not text:
            return ""
        parts = split_multi(text.replace("，", "|").replace(",", "|"))
        if BillApp._geo_tokens_invalid(field, parts, rec):
            return None
        return "|".join(parts)

    @staticmethod
    def _city_picker_source(field: str, rec: dict) -> list[str]:
        """地市类弹窗候选项：排除地市在已选省份时仅显示这些省下的地市。"""
        if field == "exclude_city" and split_multi(rec.get("province", "")):
            return cities_under_provinces(split_multi(rec.get("province", "")))
        if field in ("city", "exclude_city"):
            return list(CITIES)
        raise ValueError(field)

    @staticmethod
    def _created_at_sort_key(rec: dict) -> str:
        return str(rec.get("created_at", "") or "").strip()

    @staticmethod
    def _created_at_filter_key(raw: Any) -> str:
        """表头筛选用：只保留 YYYY-MM-DD，同一日期去重；选项中不显示时分秒。
        亦用于「打印时间」last_printed_at / 打印记录 printed_at（支持 YYYY/MM/DD …）。"""
        t = str(raw or "").strip()
        if not t:
            return ""
        # 常见：YYYY-MM-DD 后接空格/T 及时间
        if re.match(r"^\d{4}-\d{2}-\d{2}", t):
            head = t[:10]
            if len(t) == 10 or (len(t) > 10 and t[10] in " Tt"):
                return head
        if "T" in t[:32]:
            left = t.split("T", 1)[0].strip()
            if len(left) >= 10 and re.match(r"^\d{4}-\d{2}-\d{2}$", left[:10]):
                return left[:10]
        m = re.match(r"^(\d{4})[./-](\d{1,2})[./-](\d{1,2})", t)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                return datetime(y, mo, d).strftime("%Y-%m-%d")
            except ValueError:
                pass
        try:
            if "T" in t:
                s = t.replace("Z", "").split("+", 1)[0].split(".", 1)[0][:19]
                return datetime.fromisoformat(s).strftime("%Y-%m-%d")
            if re.search(r"\d{1,2}:\d{2}", t):
                return datetime.strptime(t[:19], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
        except Exception:
            pass
        try:
            return datetime.strptime(t[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            pass
        m2 = re.match(r"^(\d{4}-\d{2}-\d{2})", t)
        if m2:
            return m2.group(1)
        return ""

    def _main_table_sort_key(self, rec_idx: int) -> tuple:
        """未保存（无提单时间）的草稿排在最前；同组内按下标排序（新增行插在首条可见记录之前）。"""
        ca = self._created_at_sort_key(self.records[rec_idx])
        if ca:
            return (0, ca)
        return (1, -rec_idx)

    def on_bill_search_clicked(self):
        """搜索按钮：刷新列表并清空所有表头列筛选。"""
        self.header_filters.clear()
        self.refresh_table()

    def on_history_search_clicked(self):
        """历史页搜索按钮：刷新并清空表头列筛选。"""
        self.history_header_filters.clear()
        self.refresh_history_table()

    def clear_bill_search_and_filters(self):
        self.search.setText("")
        self.header_filters.clear()
        self.header_sort_field = None
        self.refresh_table()

    def _sync_main_v_scroll(self, value: int, source: str):
        if self._table_v_sync:
            return
        self._table_v_sync = True
        try:
            if source == "scroll":
                self.table_frozen.verticalScrollBar().setValue(value)
            else:
                self.table.verticalScrollBar().setValue(value)
        finally:
            self._table_v_sync = False

    def _sync_selection_main_to_frozen(self):
        if self._table_sel_sync or self.table.selectionModel() is None:
            return
        self._table_sel_sync = True
        try:
            self.table_frozen.clearSelection()
            for idx in self.table.selectionModel().selectedRows():
                self.table_frozen.selectRow(idx.row())
        finally:
            self._table_sel_sync = False

    def _sync_selection_frozen_to_main(self):
        if self._table_sel_sync or self.table_frozen.selectionModel() is None:
            return
        self._table_sel_sync = True
        try:
            self.table.clearSelection()
            for idx in self.table_frozen.selectionModel().selectedRows():
                self.table.selectRow(idx.row())
        finally:
            self._table_sel_sync = False

    def _cell_display_for_filter_main(self, rec: dict, field: str) -> str:
        if field == "created_at":
            return BillApp._created_at_filter_key(rec.get("created_at"))
        if field == "print_count":
            return str(int(rec.get("print_count", 0) or 0))
        if field == "last_printed_at":
            return BillApp._created_at_filter_key(rec.get("last_printed_at"))
        if field in ("type_code", "operator_code"):
            return self.display_val(rec, field) or str(rec.get(field, "") or "")
        if field == "allow_print_url":
            return self.display_val(rec, "allow_print_url")
        if field == "url":
            return first_url(rec.get(field, ""))
        return str(rec.get(field, "") or "")

    def _main_row_matches_header_filters_except(self, rec_idx: int, skip_field: str | None) -> bool:
        rec = self.records[rec_idx]
        for field, vals in self.header_filters.items():
            if skip_field and field == skip_field:
                continue
            if not vals:
                continue
            if field == "created_at":
                dk = BillApp._created_at_filter_key(rec.get("created_at"))
                allowed = {BillApp._created_at_filter_key(x) for x in vals}
                if dk not in allowed:
                    return False
            elif field == "last_printed_at":
                dk = BillApp._created_at_filter_key(rec.get("last_printed_at"))
                allowed = {BillApp._created_at_filter_key(x) for x in vals}
                if dk not in allowed:
                    return False
            elif field == "url":
                cell_lines = set(url_lines_for_filter(rec.get("url", "")))
                if not (cell_lines & set(vals)):
                    return False
            elif self._cell_display_for_filter_main(rec, field) not in vals:
                return False
        return True

    def _main_row_matches_header_filters(self, rec_idx: int) -> bool:
        return self._main_row_matches_header_filters_except(rec_idx, None)

    def _unique_display_values_main(self, field: str) -> list[str]:
        query = self.search.text().strip().lower()
        out: list[str] = []
        seen: set[str] = set()
        for i, r in enumerate(self.records):
            if query and query not in str(r.get("task_name", "")).lower():
                continue
            if not self._main_row_matches_header_filters_except(i, field):
                continue
            if field == "url":
                for dv in url_lines_for_filter(r.get("url", "")):
                    if dv not in seen:
                        seen.add(dv)
                        out.append(dv)
                continue
            dv = self._cell_display_for_filter_main(r, field)
            if dv not in seen:
                seen.add(dv)
                out.append(dv)
        if field == "created_at":
            out = BillApp._unique_created_at_filter_options(out)
        elif field == "last_printed_at":
            out = BillApp._unique_created_at_filter_options(out)
        else:
            out.sort(key=lambda s: (s == "", s.casefold()))
        return out

    @staticmethod
    def _unique_created_at_filter_options(values: list[str]) -> list[str]:
        """筛选项：仅 YYYY-MM-DD、按日期去重；空日期保留一条；日期新在前。"""
        seen: set[str] = set()
        norm: list[str] = []
        for x in values:
            k = BillApp._created_at_filter_key(x)
            if k in seen:
                continue
            seen.add(k)
            norm.append(k)

        def sort_key(d: str):
            if not d:
                return (2, "")
            try:
                return (0, -datetime.strptime(d, "%Y-%m-%d").timestamp())
            except Exception:
                return (1, d)

        norm.sort(key=sort_key)
        return norm

    def _field_for_header_section(self, mode: str, section: int) -> str | None:
        if mode in ("main_frozen", "hist_frozen"):
            return None if section == 0 else "task_name"
        if mode == "print_rec":
            if section == 0 or section > len(PRINT_LOG_DATA_FIELDS):
                return None
            return PRINT_LOG_DATA_FIELDS[section - 1]
        if mode == "main_scroll":
            return self._main_scroll_field_for_section(section)
        if mode == "hist_scroll":
            return self._history_scroll_field_for_section(section)
        return None

    def _header_show_filter_btn(self, mode: str, section: int) -> bool:
        return self._field_for_header_section(mode, section) is not None

    def _close_column_filter_popup(self):
        w = self._filter_popup
        if w is None:
            return
        w.close()
        w.deleteLater()
        self._filter_popup = None

    def _on_column_filter_popup_closed(self, dlg: "ColumnPickFilterPopup"):
        if self._filter_popup is dlg:
            self._filter_popup = None

    def _filter_button_global_bottom_right(self, mode: str, section: int) -> QPoint | None:
        if mode == "main_frozen":
            h = self.table_frozen.horizontalHeader()
        elif mode == "main_scroll":
            h = self.table.horizontalHeader()
        elif mode == "hist_frozen":
            h = self.history_table_frozen.horizontalHeader()
        elif mode == "hist_scroll":
            h = self.history_table.horizontalHeader()
        elif mode == "print_rec":
            h = self.print_log_table.horizontalHeader()
        else:
            return None
        if not isinstance(h, HoverFilterHeaderView):
            return None
        r = h._filter_btn_rect(section)
        return h.viewport().mapToGlobal(r.bottomRight())

    def _open_header_filter_from_header(self, mode: str, section: int):
        field = self._field_for_header_section(mode, section)
        if not field:
            return
        if mode.startswith("main"):
            opts = self._unique_display_values_main(field)
            cur = set(self.header_filters.get(field, ()))
        elif mode.startswith("print"):
            opts = self._unique_display_values_print(field)
            cur = set(self.print_header_filters.get(field, ()))
        else:
            opts = self._unique_display_values_hist(field)
            cur = set(self.history_header_filters.get(field, ()))
        if field in ("created_at", "last_printed_at", "printed_at"):
            cur = {BillApp._created_at_filter_key(x) for x in cur}
        title = HEADERS.get(field, field)
        self._close_column_filter_popup()
        anchor = self._filter_button_global_bottom_right(mode, section)
        if anchor is None:
            anchor = self.mapToGlobal(self.rect().topRight())
        dlg = ColumnPickFilterPopup(
            self,
            mode,
            field,
            f"筛选：{title}",
            opts,
            cur,
            anchor,
            parent=self,
        )
        dlg.destroyed.connect(lambda *a, d=dlg: self._on_column_filter_popup_closed(d))
        self._filter_popup = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _toggle_main_sort_for_field(self, field: str):
        if self.header_sort_field == field:
            if self.header_sort_order == Qt.SortOrder.AscendingOrder:
                self.header_sort_order = Qt.SortOrder.DescendingOrder
            else:
                self.header_sort_field = None
        else:
            self.header_sort_field = field
            self.header_sort_order = Qt.SortOrder.AscendingOrder
        self.refresh_table()

    def _on_main_frozen_header_section_clicked(self, section: int):
        if section == 0:
            # 由 HoverFilterHeaderView.mouseReleaseEvent 统一处理，避免与 sectionClicked 重复触发
            return
        self._toggle_main_sort_for_field("task_name")

    def _main_scroll_field_for_section(self, section: int) -> str | None:
        if 0 <= section < len(FIELDS) - 1:
            return FIELDS[section + 1]
        if section == len(FIELDS) - 1:
            return "created_at"
        if section == len(FIELDS):
            return "print_count"
        if section == len(FIELDS) + 1:
            return "last_printed_at"
        return None

    def _on_main_scroll_header_clicked(self, section: int):
        field = self._main_scroll_field_for_section(section)
        if not field:
            return
        self._toggle_main_sort_for_field(field)

    def _sort_key_for_main_field(self, rec_idx: int, field: str):
        rec = self.records[rec_idx]
        if field == "created_at":
            return self._created_at_sort_key(rec)
        if field == "print_count":
            try:
                return int(rec.get("print_count", 0) or 0)
            except (TypeError, ValueError):
                return -1
        if field == "last_printed_at":
            return str(rec.get("last_printed_at", "") or "").strip()
        if field in ("quantity", "age_max", "age_min", "pv"):
            raw = str(rec.get(field, "") or "").strip()
            if not raw:
                return float("-inf")
            try:
                return float(raw)
            except ValueError:
                return raw.lower()
        return self._cell_display_for_filter_main(rec, field).lower()

    def _apply_main_sort_to_filtered(self):
        if not self.header_sort_field:
            self.filtered_indices.sort(key=self._main_table_sort_key, reverse=True)
        else:
            field = self.header_sort_field
            rev = self.header_sort_order == Qt.SortOrder.DescendingOrder
            self.filtered_indices.sort(key=lambda i, f=field: self._sort_key_for_main_field(i, f), reverse=rev)

    def _update_main_header_sort_indicator(self):
        hf = self.table_frozen.horizontalHeader()
        hs = self.table.horizontalHeader()
        hf.setSortIndicatorShown(False)
        hs.setSortIndicatorShown(False)
        if not self.header_sort_field:
            return
        f = self.header_sort_field
        if f == "task_name":
            hf.setSortIndicatorShown(True)
            hf.setSortIndicator(1, self.header_sort_order)
            return
        if f in FIELDS:
            si = FIELDS.index(f) - 1
            if si >= 0:
                hs.setSortIndicatorShown(True)
                hs.setSortIndicator(si, self.header_sort_order)
            return
        if f == "created_at":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS) - 1, self.header_sort_order)
            return
        if f == "print_count":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS), self.header_sort_order)
            return
        if f == "last_printed_at":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS) + 1, self.header_sort_order)
            return

    def _update_main_header_tooltips(self):
        for s in range(self.table_frozen.columnCount()):
            field = self._field_for_header_section("main_frozen", s)
            if not field:
                continue
            tip = HEADERS.get(field, field) + "\n单击：排序；悬停右侧 ▼ 筛选"
            vals = self.header_filters.get(field)
            if vals:
                tip += f"\n已选 {len(vals)} 项"
            it = self.table_frozen.horizontalHeaderItem(s)
            if it:
                it.setToolTip(tip)
        for s in range(self.table.columnCount()):
            field = self._main_scroll_field_for_section(s)
            if not field:
                continue
            tip = HEADERS.get(field, field) + "\n单击：排序；悬停右侧 ▼ 筛选"
            vals = self.header_filters.get(field)
            if vals:
                tip += f"\n已选 {len(vals)} 项"
            it = self.table.horizontalHeaderItem(s)
            if it:
                it.setToolTip(tip)

    def _sync_hist_v_scroll(self, value: int, source: str):
        if self._hist_v_sync:
            return
        self._hist_v_sync = True
        try:
            if source == "scroll":
                self.history_table_frozen.verticalScrollBar().setValue(value)
            else:
                self.history_table.verticalScrollBar().setValue(value)
        finally:
            self._hist_v_sync = False

    def _sync_selection_hist_to_frozen(self):
        if self._hist_sel_sync or self.history_table.selectionModel() is None:
            return
        self._hist_sel_sync = True
        try:
            self.history_table_frozen.clearSelection()
            for idx in self.history_table.selectionModel().selectedRows():
                self.history_table_frozen.selectRow(idx.row())
        finally:
            self._hist_sel_sync = False

    def _sync_selection_hist_frozen_to_scroll(self):
        if self._hist_sel_sync or self.history_table_frozen.selectionModel() is None:
            return
        self._hist_sel_sync = True
        try:
            self.history_table.clearSelection()
            for idx in self.history_table_frozen.selectionModel().selectedRows():
                self.history_table.selectRow(idx.row())
        finally:
            self._hist_sel_sync = False

    def _cell_display_for_filter_hist(self, rec: dict, field: str) -> str:
        if field == "deleted_at":
            return str(rec.get("deleted_at", "") or "")
        if field == "print_count":
            return str(int(rec.get("print_count", 0) or 0))
        if field == "last_printed_at":
            return BillApp._created_at_filter_key(rec.get("last_printed_at"))
        if field == "created_at":
            return BillApp._created_at_filter_key(rec.get("created_at"))
        if field in ("type_code", "operator_code"):
            return self.display_val(rec, field) or str(rec.get(field, "") or "")
        if field == "allow_print_url":
            return self.display_val(rec, "allow_print_url")
        if field == "url":
            return first_url(rec.get(field, ""))
        return str(rec.get(field, "") or "")

    def _history_row_matches_header_filters_except(self, rec: dict, skip_field: str | None) -> bool:
        for field, vals in self.history_header_filters.items():
            if skip_field and field == skip_field:
                continue
            if not vals:
                continue
            if field == "created_at":
                dk = BillApp._created_at_filter_key(rec.get("created_at"))
                allowed = {BillApp._created_at_filter_key(x) for x in vals}
                if dk not in allowed:
                    return False
            elif field == "last_printed_at":
                dk = BillApp._created_at_filter_key(rec.get("last_printed_at"))
                allowed = {BillApp._created_at_filter_key(x) for x in vals}
                if dk not in allowed:
                    return False
            elif field == "url":
                cell_lines = set(url_lines_for_filter(rec.get("url", "")))
                if not (cell_lines & set(vals)):
                    return False
            elif self._cell_display_for_filter_hist(rec, field) not in vals:
                return False
        return True

    def _history_row_matches_header_filters(self, rec: dict) -> bool:
        return self._history_row_matches_header_filters_except(rec, None)

    def _unique_display_values_hist(self, field: str) -> list[str]:
        query = self.history_search.text().strip().lower()
        out: list[str] = []
        seen: set[str] = set()
        for r in self.history_records:
            if query and query not in str(r.get("task_name", "")).lower():
                continue
            if not self._history_row_matches_header_filters_except(r, field):
                continue
            if field == "url":
                for dv in url_lines_for_filter(r.get("url", "")):
                    if dv not in seen:
                        seen.add(dv)
                        out.append(dv)
                continue
            dv = self._cell_display_for_filter_hist(r, field)
            if dv not in seen:
                seen.add(dv)
                out.append(dv)
        if field == "created_at":
            out = BillApp._unique_created_at_filter_options(out)
        elif field == "last_printed_at":
            out = BillApp._unique_created_at_filter_options(out)
        else:
            out.sort(key=lambda s: (s == "", s.casefold()))
        return out

    def _history_scroll_field_for_section(self, section: int) -> str | None:
        if 0 <= section < HISTORY_SCROLL_COLUMNS:
            return HISTORY_SCROLL_FIELDS[section]
        return None

    def _sort_key_history_field(self, rec: dict, field: str | None):
        if not field:
            return self._created_at_sort_key(rec)
        if field == "deleted_at":
            return str(rec.get("deleted_at", "") or "")
        if field == "created_at":
            return self._created_at_sort_key(rec)
        if field == "print_count":
            try:
                return int(rec.get("print_count", 0) or 0)
            except (TypeError, ValueError):
                return -1
        if field == "last_printed_at":
            return str(rec.get("last_printed_at", "") or "").strip()
        if field in ("quantity", "age_max", "age_min", "pv"):
            raw = str(rec.get(field, "") or "").strip()
            if not raw:
                return float("-inf")
            try:
                return float(raw)
            except ValueError:
                return raw.lower()
        return self._cell_display_for_filter_hist(rec, field).lower()

    def _toggle_history_sort_for_field(self, field: str):
        if self.history_sort_field == field:
            if self.history_sort_order == Qt.SortOrder.AscendingOrder:
                self.history_sort_order = Qt.SortOrder.DescendingOrder
            else:
                self.history_sort_field = None
        else:
            self.history_sort_field = field
            self.history_sort_order = Qt.SortOrder.AscendingOrder
        self.refresh_history_table()

    def _on_hist_frozen_header_section_clicked(self, section: int):
        if section == 0:
            return
        self._toggle_history_sort_for_field("task_name")

    def _on_history_scroll_header_clicked(self, section: int):
        field = self._history_scroll_field_for_section(section)
        if not field:
            return
        self._toggle_history_sort_for_field(field)

    def _update_history_header_sort_indicator(self):
        hf = self.history_table_frozen.horizontalHeader()
        hs = self.history_table.horizontalHeader()
        hf.setSortIndicatorShown(False)
        hs.setSortIndicatorShown(False)
        if not self.history_sort_field:
            return
        f = self.history_sort_field
        if f == "task_name":
            hf.setSortIndicatorShown(True)
            hf.setSortIndicator(1, self.history_sort_order)
            return
        if f in FIELDS:
            si = FIELDS.index(f) - 1
            if si >= 0:
                hs.setSortIndicatorShown(True)
                hs.setSortIndicator(si, self.history_sort_order)
            return
        if f == "created_at":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS) - 1, self.history_sort_order)
            return
        if f == "print_count":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS), self.history_sort_order)
            return
        if f == "last_printed_at":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS) + 1, self.history_sort_order)
            return
        if f == "deleted_at":
            hs.setSortIndicatorShown(True)
            hs.setSortIndicator(len(FIELDS) + 2, self.history_sort_order)

    def _update_history_header_tooltips(self):
        for s in range(self.history_table_frozen.columnCount()):
            field = self._field_for_header_section("hist_frozen", s)
            if not field:
                continue
            tip = HEADERS.get(field, field) + "\n单击：排序；悬停右侧 ▼ 筛选"
            vals = self.history_header_filters.get(field)
            if vals:
                tip += f"\n已选 {len(vals)} 项"
            it = self.history_table_frozen.horizontalHeaderItem(s)
            if it:
                it.setToolTip(tip)
        for s in range(self.history_table.columnCount()):
            field = self._history_scroll_field_for_section(s)
            if not field:
                continue
            title = HEADERS.get(field, field) if field in HEADERS else field
            tip = title + "\n单击：排序；悬停右侧 ▼ 筛选"
            vals = self.history_header_filters.get(field)
            if vals:
                tip += f"\n已选 {len(vals)} 项"
            it = self.history_table.horizontalHeaderItem(s)
            if it:
                it.setToolTip(tip)

    def _url_cell_editor_display_height(self) -> int:
        """URL 编辑框固定单行高度；不换行、无滚动条，用左右方向键移动光标查看超出部分。"""
        fm = QFontMetrics(self.table.font())
        inner = fm.height() + 8
        return max(26, min(self._data_row_height - 4, inner))

    def _url_plain_editor_at(self, row: int, url_col: int) -> QPlainTextEdit | None:
        """URL 列可能为外层垂直居中容器 + UrlCellEditor。"""
        w = self.table.cellWidget(row, url_col)
        if w is None:
            return None
        if isinstance(w, QPlainTextEdit) and w.objectName() == "urlCellEditor":
            return w
        ed = w.findChild(QPlainTextEdit, "urlCellEditor")
        return ed if isinstance(ed, QPlainTextEdit) else None

    def _on_main_scroll_column_resized_for_url(self, logical_index: int, _old_size: int, _new_size: int):
        url_sc = FIELDS.index("url") - 1
        if logical_index != url_sc:
            return
        self._reflow_all_url_row_heights()

    def _reflow_all_url_row_heights(self):
        """URL 列宽变化后，统一为单行框高（不随内容增高）。"""
        if self.updating_table:
            return
        url_sc = FIELDS.index("url") - 1
        nh = self._url_cell_editor_display_height()
        rh = self._data_row_height
        for row in range(self.table.rowCount()):
            wgt = self._url_plain_editor_at(row, url_sc)
            if wgt is None:
                continue
            wgt.blockSignals(True)
            wgt.setFixedHeight(nh)
            wgt.blockSignals(False)
            self.table.setRowHeight(row, rh)
            self.table_frozen.setRowHeight(row, rh)

    @staticmethod
    def _url_editor_stylesheet(has_error: bool) -> str:
        border = "#e05263" if has_error else "rgba(148,163,184,0.35)"
        return (
            f"QPlainTextEdit#urlCellEditor {{ border: 1px solid {border}; border-radius: 4px; padding: 2px; }}"
        )

    def _on_url_cell_text_changed(self, rec_idx: int, ed: QPlainTextEdit):
        if self.updating_table:
            return
        if rec_idx < 0 or rec_idx >= len(self.records):
            return
        txt = ed.toPlainText()
        if self.records[rec_idx].get("url", "") == txt:
            return
        self.records[rec_idx]["url"] = txt
        ok, msg = self.validate_record(self.records[rec_idx])
        self.clear_error(rec_idx, "url")
        if not ok and "URL" in msg:
            self.mark_error(rec_idx, "url", msg)
            ed.setStyleSheet(self._url_editor_stylesheet(True))
        else:
            ed.setStyleSheet(self._url_editor_stylesheet(False))
        self.save_data()

    def _on_url_cell_focus_out(self, rec_idx: int, ed: QPlainTextEdit):
        """URL 失焦时若不符合规则则弹窗提示（输入过程中仅用红框，不弹窗）。"""
        if self.updating_table:
            return
        if rec_idx < 0 or rec_idx >= len(self.records):
            return
        if str(self.records[rec_idx].get("url", "")) != ed.toPlainText():
            return
        ok, msg = self.validate_urls(ed.toPlainText())
        if ok:
            return
        QMessageBox.warning(self, "URL 校验失败", msg or "URL 格式不正确")

    def refresh_table(self):
        self.updating_table = True
        query = self.search.text().strip().lower()
        self.filtered_indices = [
            i
            for i, r in enumerate(self.records)
            if (not query or query in str(r.get("task_name", "")).lower()) and self._main_row_matches_header_filters(i)
        ]
        self._apply_main_sort_to_filtered()
        n = len(self.filtered_indices)
        self.table.setRowCount(n)
        self.table_frozen.setRowCount(n)
        per_row_heights: list[int] = []
        for row, rec_idx in enumerate(self.filtered_indices):
            rec = self.records[rec_idx]
            row_max_h = self._data_row_height
            box_wrap = QWidget()
            box_layout = QHBoxLayout(box_wrap)
            box_layout.setContentsMargins(0, 0, 0, 0)
            box_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox()
            chk.setChecked(bool(rec.get("checked")))
            chk.stateChanged.connect(lambda _=0, x=rec_idx, c=chk: self.on_row_checkbox_changed(x, c.isChecked()))
            box_layout.addWidget(chk)
            self.table_frozen.setCellWidget(row, 0, box_wrap)
            ph0 = QTableWidgetItem("")
            ph0.setFlags(Qt.ItemIsEnabled)
            self.table_frozen.setItem(row, 0, ph0)
            tn_item = QTableWidgetItem(self.display_val(rec, "task_name"))
            tn_item.setFlags(
                Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsEditable
            )
            tn_item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            if not str(rec.get("task_name", "")).strip():
                tn_item.setToolTip("示例：客户全国YDDBDK-产品")
            if self.field_errors.get((rec_idx, "task_name")):
                tn_item.setBackground(QColor("#e05263"))
                tn_item.setToolTip(self.field_error_msgs.get((rec_idx, "task_name"), "字段校验失败"))
            self.table_frozen.setItem(row, 1, tn_item)
            for sc, f in enumerate(FIELDS[1:]):
                if f == "allow_print_url":
                    yes = coerce_allow_print_url(rec.get("allow_print_url"))
                    aw_item = QTableWidgetItem("是" if yes else "否")
                    aw_item.setFont(QFont(self.table_frozen.font()))
                    aw_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    aw_item.setToolTip(
                        "允许打印URL：是（单击切换为否）" if yes else "允许打印URL：否（单击切换为是）"
                    )
                    aw_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                    self.table.setItem(row, sc, aw_item)
                    continue
                if f == "url":
                    raw_u = str(rec.get("url", ""))
                    pte = UrlCellEditor(self, rec_idx)
                    pte.setFrameShape(QFrame.Shape.NoFrame)
                    pte.setTabChangesFocus(True)
                    pte.setFont(QFont(self.table.font()))
                    pte.blockSignals(True)
                    pte.setPlainText(raw_u)
                    pte.blockSignals(False)
                    pte.setFixedHeight(self._url_cell_editor_display_height())
                    pte.textChanged.connect(lambda ri=rec_idx, ed=pte: self._on_url_cell_text_changed(ri, ed))
                    pte.setStyleSheet(self._url_editor_stylesheet(bool(self.field_errors.get((rec_idx, "url")))))
                    ph_url = QTableWidgetItem()
                    ph_url.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                    self.table.setItem(row, sc, ph_url)
                    url_wrap = QWidget()
                    url_wrap.setObjectName("urlCellWrap")
                    url_lo = QVBoxLayout(url_wrap)
                    url_lo.setContentsMargins(0, 0, 0, 0)
                    url_lo.setSpacing(0)
                    url_lo.addStretch(1)
                    url_lo.addWidget(pte)
                    url_lo.addStretch(1)
                    self.table.setCellWidget(row, sc, url_wrap)
                    continue
                item = QTableWidgetItem(self.display_val(rec, f))
                if f in ("province", "exclude_province", "city", "exclude_city"):
                    item.setFlags(
                        Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable
                    )
                item.setTextAlignment(Qt.AlignCenter)
                if self.field_errors.get((rec_idx, f)):
                    item.setBackground(QColor("#e05263"))
                    item.setToolTip(self.field_error_msgs.get((rec_idx, f), "字段校验失败"))
                self.table.setItem(row, sc, item)
                if f in ("type_code", "operator_code", "duration"):
                    combo = QComboBox()
                    combo.addItem("-- 请选择 --", "")
                    if f == "type_code":
                        for code, name in TYPE_MAP.items():
                            combo.addItem(name, code)
                    elif f == "operator_code":
                        for code, name in OP_MAP.items():
                            combo.addItem(name, code)
                    else:
                        for val in DURATIONS:
                            combo.addItem(val, val)
                    combo.setCurrentIndex(max(0, combo.findData(rec.get(f, ""))))
                    combo.currentIndexChanged.connect(
                        lambda _=0, x=rec_idx, ff=f, cb=combo: self.on_inline_combo_changed(x, ff, cb.currentData() or "")
                    )
                    style_combo_centered(combo)
                    self.table.setCellWidget(row, sc, combo)
            ca_item = QTableWidgetItem(str(rec.get("created_at", "")))
            ca_item.setTextAlignment(Qt.AlignCenter)
            ca_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(row, len(FIELDS) - 1, ca_item)
            pc_item = QTableWidgetItem(str(int(rec.get("print_count", 0) or 0)))
            pc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            pc_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(row, len(FIELDS), pc_item)
            pt_item = QTableWidgetItem(str(rec.get("last_printed_at", "") or ""))
            pt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            pt_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(row, len(FIELDS) + 1, pt_item)
            btn = QPushButton("🗑 删除")
            btn.setObjectName("btnDanger")
            btn.clicked.connect(lambda _=False, x=rec_idx: self.delete_row(x))
            self.table.setCellWidget(row, MAIN_SCROLL_COLUMNS - 1, btn)
            per_row_heights.append(row_max_h)
        for r, rh in enumerate(per_row_heights):
            self.table.setRowHeight(r, rh)
            self.table_frozen.setRowHeight(r, rh)
        self.lbl_count.setText(f"提单数: {len(self.filtered_indices)}")
        self.update_header_check()
        self._update_main_header_sort_indicator()
        self._update_main_header_tooltips()
        self.updating_table = False

    def _select_all_header_paint_info(self, mode: str) -> dict[str, Any]:
        """冻结区/打印记录首列表头：对号 / X 芯片（绿对号全选、琥珀部分、红 X 取消、灰对号无行）。"""
        if mode == "main_frozen":
            indices = self.filtered_indices
            records = self.records
        elif mode == "hist_frozen":
            indices = self.history_filtered_indices
            records = self.history_records
        elif mode == "print_rec":
            indices = self.print_log_filtered_indices
            records = self.print_records
        else:
            return {"glyph": "check_muted", "interactive": False}
        if not indices:
            return {"glyph": "check_muted", "interactive": False}
        n = len(indices)
        checked_count = sum(1 for i in indices if records[i].get("checked"))
        if checked_count == n:
            return {"glyph": "x", "interactive": True}
        if checked_count == 0:
            return {"glyph": "check", "interactive": True}
        return {"glyph": "check_partial", "interactive": True}

    def update_header_check(self):
        it = self.table_frozen.horizontalHeaderItem(0)
        if not it:
            return
        info = self._select_all_header_paint_info("main_frozen")
        if not self.filtered_indices:
            it.setToolTip("当前筛选下列表为空")
        elif info["glyph"] == "x":
            it.setToolTip("单击：取消全选（清除当前列表中所有勾选）")
        else:
            it.setToolTip("单击：全选（将当前列表全部行设为勾选；部分勾选时则全选）")
        self.table_frozen.horizontalHeader().viewport().update()

    def on_header_clicked(self, section):
        if section != 0:
            return
        if not self.filtered_indices:
            return
        all_checked = all(self.records[i].get("checked") for i in self.filtered_indices)
        for i in self.filtered_indices:
            self.records[i]["checked"] = not all_checked
        self.save_data()
        self.refresh_table()

    def _on_frozen_header_col0_clicked(self, mode: str):
        if mode == "main_frozen":
            self.on_header_clicked(0)
            return
        if mode == "hist_frozen":
            if not self.history_filtered_indices:
                return
            all_checked = all(self.history_records[i].get("checked") for i in self.history_filtered_indices)
            for i in self.history_filtered_indices:
                self.history_records[i]["checked"] = not all_checked
            self.save_history_data()
            self.refresh_history_table()

    def on_cell_clicked(self, row, col):
        if self.updating_table:
            return
        if row < 0 or row >= len(self.filtered_indices):
            return
        if col == 0:
            rec_idx = self.filtered_indices[row]
            new_state = not self.records[rec_idx].get("checked", False)
            self.on_row_checkbox_changed(rec_idx, new_state)
            return
        if col == ALLOW_PRINT_URL_DISPLAY_COL:
            rec_idx = self.filtered_indices[row]
            cur = coerce_allow_print_url(self.records[rec_idx].get("allow_print_url"))
            self.records[rec_idx]["allow_print_url"] = not cur
            self.save_data()
            self.refresh_table()

    def on_row_checkbox_changed(self, rec_idx: int, checked: bool):
        if self.updating_table:
            return
        if rec_idx < 0 or rec_idx >= len(self.records):
            return
        self.records[rec_idx]["checked"] = bool(checked)
        self.save_data()
        self.refresh_table()

    def on_cell_double_click(self, row, col):
        if col <= 0 or col >= len(DISPLAY_FIELDS) - 1:
            return
        name = DISPLAY_FIELDS[col]
        if name in ("print_count", "last_printed_at", "created_at", "action"):
            return
        rec_idx = self.filtered_indices[row]
        field = name
        rec = self.records[rec_idx]
        if field == "allow_print_url":
            cur = coerce_allow_print_url(rec.get("allow_print_url"))
            rec["allow_print_url"] = not cur
            self.save_data()
            self.refresh_table()
            return
        if field == "url":
            sc = FIELDS.index("url") - 1
            ed = self._url_plain_editor_at(row, sc)
            if ed is not None:
                ed.setFocus()
            return
        if field in ("province", "exclude_province", "city", "exclude_city"):
            if field == "city" and rec.get("province", "").strip():
                QMessageBox.warning(self, "提示", "已选择省份后，地市不可再选择；请先清空省份")
                return
            if field == "province" and rec.get("city", "").strip():
                QMessageBox.warning(self, "提示", "已选择地市后，省份不可再选择；请先清空地市")
                return
            if field == "province" and split_multi(rec.get("exclude_province", "")):
                QMessageBox.warning(self, "提示", "已填写排除省份，省份不可再编辑；请先清空排除省份")
                return
            if field == "city" and (
                split_multi(rec.get("exclude_province", "")) or split_multi(rec.get("exclude_city", ""))
            ):
                QMessageBox.warning(self, "提示", "已填写排除省份或排除地市，地市不可再编辑；请先清空对应排除项")
                return
            if field == "exclude_province" and (
                rec.get("province", "").strip() or rec.get("city", "").strip()
            ):
                QMessageBox.warning(self, "提示", "已选择省份或地市后，排除省份不可再选择；请先清空省份与地市")
                return
            if field == "exclude_city" and rec.get("city", "").strip():
                QMessageBox.warning(self, "提示", "已选择地市后，排除地市不可再选择；请先清空地市")
                return
            cur_parts = split_multi(rec.get(field, ""))
            if self._geo_tokens_invalid(field, cur_parts, rec):
                QMessageBox.warning(self, "提示", "输入错误，请重新输入或者双击选择")
                return
            items = PROVINCES if "province" in field else self._city_picker_source(field, rec)
            if field == "exclude_city" and split_multi(rec.get("province", "")) and not items:
                QMessageBox.warning(self, "提示", "所选省份暂无地市级数据，无法选择排除地市。")
                return
            recent = self.picker_recent.get(field, [])
            ordered = sorted(
                items,
                key=lambda x: (
                    0 if x in split_multi(rec.get(field, "")) else 1,
                    0 if x in recent else 1,
                    recent.index(x) if x in recent else 999,
                    x,
                ),
            )
            dlg = MultiSelectDialog("选择", ordered, split_multi(rec.get(field, "")), self)
            if dlg.exec() == QDialog.Accepted:
                values = dlg.values()
                self.update_picker_recent(field, values)
                self.apply_field_update(rec_idx, field, "|".join(values))
            return

    def on_current_cell_changed(self, current_row, current_col, _previous_row, _previous_col):
        if self.updating_table:
            return
        if current_row < 0 or current_row >= len(self.filtered_indices):
            return
        if current_col <= 0 or current_col >= len(DISPLAY_FIELDS) - 1:
            return
        name = DISPLAY_FIELDS[current_col]
        if name in ("print_count", "last_printed_at", "created_at", "action"):
            return
        field = name
        if field == "task_name":
            self.statusBar().showMessage("任务名格式：客户名+省份/地市/全国/几省/几市+运营商码值+类型码值+行业编码-产品名", 5000)

    def on_table_context_menu(self, pos):
        idx = self.table.indexAt(pos)
        if not idx.isValid():
            return
        row = idx.row()
        col = idx.column() + FROZEN_COLUMNS
        if row < 0 or row >= len(self.filtered_indices):
            return
        if col <= 0 or col >= len(DISPLAY_FIELDS) - 1:
            return
        name = DISPLAY_FIELDS[col]
        if name in ("print_count", "last_printed_at", "created_at", "action"):
            return
        field = name
        if field not in ("task_name", "operator_code", "type_code", "duration"):
            return
        rec_idx = self.filtered_indices[row]
        rec = self.records[rec_idx]
        menu = self.table.createStandardContextMenu()
        action_map: dict[object, tuple[str, str]] = {}
        if field == "task_name":
            sample = f"客户{self.region_part(rec)}{rec.get('operator_code') or 'YD'}{rec.get('type_code') or 'DB'}{rec.get('industry_code') or 'DK'}-产品"
            action_fill = menu.addAction("套用任务名示例")
            action_map[action_fill] = ("task_name", sample)
        elif field == "operator_code":
            menu.addSeparator()
            for code, name in OP_MAP.items():
                act = menu.addAction(f"设置运营商：{name}（{code}）")
                action_map[act] = ("operator_code", code)
        elif field == "type_code":
            menu.addSeparator()
            for code, name in TYPE_MAP.items():
                act = menu.addAction(f"设置类型：{name}（{code}）")
                action_map[act] = ("type_code", code)
        elif field == "duration":
            menu.addSeparator()
            for val in DURATIONS:
                act = menu.addAction(f"设置时长：{val}")
                action_map[act] = ("duration", val)
        chosen = menu.exec(self.table.viewport().mapToGlobal(pos))
        if chosen in action_map:
            target_field, target_value = action_map[chosen]
            self.apply_field_update(rec_idx, target_field, target_value)

    def bulk_fill_selected(self):
        selected_idx = [idx for idx, rec in enumerate(self.records) if rec.get("checked")]
        if not selected_idx:
            QMessageBox.information(self, "提示", "请先勾选要批量填充的记录")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("批量填充")
        scope_combo = QComboBox(dlg)
        scope_combo.addItem("全部已勾选记录", "all_checked")
        scope_combo.addItem("仅当前筛选结果中已勾选", "filtered_checked")
        field_combo = QComboBox(dlg)
        field_combo.addItem("运营商", "operator_code")
        field_combo.addItem("类型", "type_code")
        field_combo.addItem("时长", "duration")
        field_combo.addItem("行业编码", "industry_code")
        field_combo.addItem("数量", "quantity")
        field_combo.addItem("年龄上限", "age_max")
        field_combo.addItem("年龄下限", "age_min")
        field_combo.addItem("pv", "pv")
        mode_combo = QComboBox(dlg)
        mode_combo.addItem("批量填充", "fill")
        mode_combo.addItem("批量清空", "clear")
        value_combo = QComboBox(dlg)
        value_combo.setMinimumWidth(260)
        value_edit = QLineEdit(dlg)
        value_edit.setMinimumWidth(260)

        def render_values():
            field = field_combo.currentData()
            value_combo.clear()
            value_edit.clear()
            value_edit.hide()
            value_combo.show()
            if mode_combo.currentData() == "clear":
                value_combo.setEnabled(False)
                value_edit.setEnabled(False)
                value_combo.addItem("（将清空所选字段）", "")
                return
            value_combo.setEnabled(True)
            value_edit.setEnabled(True)
            if field == "operator_code":
                for code, name in OP_MAP.items():
                    value_combo.addItem(f"{name}（{code}）", code)
            elif field == "type_code":
                for code, name in TYPE_MAP.items():
                    value_combo.addItem(f"{name}（{code}）", code)
            elif field == "duration":
                for val in DURATIONS:
                    value_combo.addItem(val, val)
            elif field == "industry_code":
                value_combo.hide()
                value_edit.show()
                value_edit.setPlaceholderText("请输入行业编码（字母或数字）")
            else:
                value_combo.hide()
                value_edit.show()
                value_edit.setPlaceholderText("请输入数字")

        field_combo.currentIndexChanged.connect(render_values)
        mode_combo.currentIndexChanged.connect(render_values)
        render_values()
        btns = QDialogButtonBox(QDialogButtonBox.Cancel | QDialogButtonBox.Ok)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel("作用范围"))
        lay.addWidget(scope_combo)
        lay.addWidget(QLabel("执行方式"))
        lay.addWidget(mode_combo)
        lay.addWidget(QLabel("填充字段"))
        lay.addWidget(field_combo)
        lay.addWidget(QLabel("填充值"))
        lay.addWidget(value_combo)
        lay.addWidget(value_edit)
        lay.addWidget(btns)
        if dlg.exec() != QDialog.Accepted:
            return
        mode = mode_combo.currentData()
        field = field_combo.currentData()
        scope = scope_combo.currentData()
        if scope == "filtered_checked":
            visible_set = set(self.filtered_indices)
            selected_idx = [idx for idx in selected_idx if idx in visible_set]
            if not selected_idx:
                QMessageBox.information(self, "提示", "当前筛选结果中没有已勾选记录")
                return
        if mode == "clear":
            value = ""
        elif field in ("industry_code", "quantity", "age_max", "age_min", "pv"):
            value = value_edit.text().strip()
        else:
            value = value_combo.currentData() or ""
        if mode == "fill" and field in ("quantity", "age_max", "age_min", "pv"):
            if value and not value.isdigit():
                QMessageBox.warning(self, "批量操作", "数量、年龄上下限、pv 仅支持数字")
                return
            if field == "quantity" and value == "0":
                QMessageBox.warning(self, "批量操作", "数量必须大于0")
                return
        field_label = field_combo.currentText()
        value_label = "清空" if mode == "clear" else (value_combo.currentText() if value_combo.isVisible() else value)
        confirm = QMessageBox.question(
            self,
            "确认批量操作",
            f"即将对 {len(selected_idx)} 条记录执行：{field_label} -> {value_label}\n确认继续吗？",
        )
        if confirm != QMessageBox.Yes:
            return
        failed: list[int] = []
        fail_details: list[tuple[int, str, str]] = []
        for idx in selected_idx:
            rec = self.records[idx]
            old_rec = dict(rec)
            rec[field] = value
            self.clear_error(idx, field)
            if field in ("age_min", "age_max"):
                self.clear_error(idx, "age_min")
                self.clear_error(idx, "age_max")
            self.after_field_change(rec, field)
            ok, msg = self.validate_record(rec)
            if not ok:
                failed.append(idx)
                fail_details.append((idx, field, msg))
                self.records[idx] = old_rec
                if "年龄上限必须大于等于年龄下限" in msg:
                    self.mark_error(idx, "age_min", msg)
                    self.mark_error(idx, "age_max", msg)
                else:
                    self.mark_error(idx, field, msg)
        self.save_data()
        self.refresh_table()
        if failed:
            self.focus_record(failed[0], field)
            detail_lines = []
            for rec_idx, failed_field, msg in fail_details[:8]:
                task_name = str(self.records[rec_idx].get("task_name", "")).strip() or f"第{rec_idx + 1}行"
                field_name = HEADERS.get(failed_field, failed_field)
                detail_lines.append(f"- {task_name} | {field_name} | {msg}")
            more = ""
            if len(fail_details) > 8:
                more = f"\n- ... 其余 {len(fail_details) - 8} 条请查看表格红框"
            QMessageBox.warning(
                self,
                "批量操作",
                "已更新 "
                f"{len(selected_idx) - len(failed)} 条，失败 {len(failed)} 条。\n\n失败明细（最多显示8条）：\n"
                + "\n".join(detail_lines)
                + more,
            )
            return
        QMessageBox.information(self, "批量操作", f"已更新 {len(selected_idx)} 条记录")

    def focus_record(self, rec_idx: int, field: str):
        if rec_idx not in self.filtered_indices:
            self.search.setText("")
            self.refresh_table()
        if rec_idx not in self.filtered_indices:
            return
        row = self.filtered_indices.index(rec_idx)
        try:
            col = DISPLAY_FIELDS.index(field)
        except ValueError:
            col = 1
        if col <= 1:
            self.table_frozen.setCurrentCell(row, col)
            it0 = self.table_frozen.item(row, col)
            if it0:
                self.table_frozen.scrollToItem(it0)
            return
        sc = col - FROZEN_COLUMNS
        self.table.setCurrentCell(row, sc)
        w = self._url_plain_editor_at(row, sc)
        if isinstance(w, QPlainTextEdit):
            self.table.scrollTo(self.table.model().index(row, sc))
            w.setFocus()
            return
        it = self.table.item(row, sc)
        if it:
            self.table.scrollToItem(it)

    def on_item_changed(self, item: QTableWidgetItem):
        if self.updating_table:
            return
        tw = item.tableWidget()
        base = 0 if tw is self.table_frozen else FROZEN_COLUMNS
        row = item.row()
        col = item.column() + base
        if row < 0 or row >= len(self.filtered_indices):
            return
        rec_idx = self.filtered_indices[row]
        rec = self.records[rec_idx]
        if col == 0:
            rec["checked"] = item.checkState() == Qt.Checked
            self.save_data()
            self.update_header_check()
            return
        if col >= len(DISPLAY_FIELDS) - 1:
            return
        name = DISPLAY_FIELDS[col]
        if name in ("print_count", "last_printed_at", "created_at", "checked"):
            return
        field = name
        if field in ("type_code", "operator_code", "duration", "url", "allow_print_url"):
            return
        if field in ("province", "exclude_province", "city", "exclude_city"):
            raw = item.text()
            norm = self._normalize_geo_field_value(field, raw, rec)
            if norm is None:
                QMessageBox.warning(self, "提示", "输入错误，请重新输入或者双击选择")
                self.updating_table = True
                item.setText(self.display_val(rec, field))
                self.updating_table = False
                return
            self.apply_field_update(rec_idx, field, norm)
            return
        val = item.text().replace("\n", " ").strip()
        self.apply_field_update(rec_idx, field, val)

    def on_inline_combo_changed(self, rec_idx: int, field: str, value: str):
        if self.updating_table:
            return
        if rec_idx < 0 or rec_idx >= len(self.records):
            return
        if str(self.records[rec_idx].get(field, "")) == str(value or ""):
            return
        self.apply_field_update(rec_idx, field, value or "")

    def task_name_hint(self, text: str) -> str:
        value = str(text or "").strip()
        if not value:
            return ""
        if "-" not in value:
            return "任务名建议包含 '-' 分隔产品名"
        left = value.split("-", 1)[0]
        parsed = self.parse_left(left)
        if not parsed:
            return "任务名前半段未识别到完整码值段（运营商+类型+行业编码）"
        if not parsed.get("customer", "").strip():
            return "任务名建议包含客户名"
        if not parsed.get("region", "").strip():
            return "任务名建议包含地域（省/市/全国/几省/几市）"
        return ""

    def mark_error(self, rec_idx: int, field: str, msg: str = ""):
        self.field_errors[(rec_idx, field)] = True
        self.field_error_msgs[(rec_idx, field)] = msg or "字段校验失败"

    def clear_error(self, rec_idx: int, field: str):
        self.field_errors.pop((rec_idx, field), None)
        self.field_error_msgs.pop((rec_idx, field), None)

    def _clear_field_errors_for_row(self, rec_idx: int):
        """移除该行所有单元格校验标红（用于导出前重算、避免残留背景色）。"""
        for k in list(self.field_errors.keys()):
            if k[0] == rec_idx:
                self.field_errors.pop(k, None)
        for k in list(self.field_error_msgs.keys()):
            if k[0] == rec_idx:
                self.field_error_msgs.pop(k, None)

    def _shift_field_errors_after_insert(self, insert_at: int):
        """在 insert_at 插入空行后，原索引 >= insert_at 的校验错误键整体 +1。"""
        new_err: dict[tuple[int, str], bool] = {}
        new_msg: dict[tuple[int, str], str] = {}
        for (ri, f), v in self.field_errors.items():
            new_err[(ri + 1 if ri >= insert_at else ri, f)] = v
        for (ri, f), v in self.field_error_msgs.items():
            new_msg[(ri + 1 if ri >= insert_at else ri, f)] = v
        self.field_errors = new_err
        self.field_error_msgs = new_msg

    def _shift_field_errors_after_delete(self, deleted_idx: int):
        """删除 deleted_idx 行后：去掉该行错误；原索引 > deleted_idx 的键整体 -1。"""
        new_err: dict[tuple[int, str], bool] = {}
        new_msg: dict[tuple[int, str], str] = {}
        for (ri, f), v in self.field_errors.items():
            if ri == deleted_idx:
                continue
            new_err[(ri - 1 if ri > deleted_idx else ri, f)] = v
        for (ri, f), v in self.field_error_msgs.items():
            if ri == deleted_idx:
                continue
            new_msg[(ri - 1 if ri > deleted_idx else ri, f)] = v
        self.field_errors = new_err
        self.field_error_msgs = new_msg

    def apply_field_update(self, rec_idx: int, field: str, value: str):
        if field in ("created_at", "print_count", "last_printed_at"):
            return
        rec = self.records[rec_idx]
        old = rec.get(field, "")
        cleaned = value.replace("\n", " ").replace("\r", " ").strip() if field == "task_name" else value
        if field == "task_name":
            hint = self.task_name_hint(cleaned)
            if hint:
                self.statusBar().showMessage(hint, 5000)
        rec[field] = cleaned
        self.clear_error(rec_idx, field)
        if field in ("age_min", "age_max"):
            self.clear_error(rec_idx, "age_min")
            self.clear_error(rec_idx, "age_max")
        ok, msg = self.validate_record(rec)
        if not ok:
            QMessageBox.warning(self, "校验失败", msg)
            rec[field] = old
            if "年龄上限必须大于等于年龄下限" in msg:
                self.mark_error(rec_idx, "age_min", msg)
                self.mark_error(rec_idx, "age_max", msg)
            else:
                self.mark_error(rec_idx, field, msg)
        else:
            changed_ok = self.after_field_change(rec, field)
            if field == "task_name":
                if changed_ok:
                    self.clear_error(rec_idx, "task_name")
                else:
                    if getattr(self, "_task_name_parse_user_warned", False):
                        self._task_name_parse_user_warned = False
                        self.mark_error(rec_idx, "task_name", "任务名码段与列表不匹配")
                    else:
                        parse_msg = "任务名未按规范完全解析：已保留原文本，仅同步可识别片段"
                        self.mark_error(rec_idx, "task_name", parse_msg)
                        self.statusBar().showMessage(parse_msg, 5000)
        self.save_data()
        self.refresh_table()

    def validate_record(self, rec):
        for k, title in [("quantity", "数量"), ("age_max", "年龄上限"), ("age_min", "年龄下限"), ("pv", "pv")]:
            v = str(rec.get(k, "")).strip()
            if v and not v.isdigit():
                return False, f"{title} 必须是数字"
        ok_url, url_msg = self.validate_urls(str(rec.get("url", "")))
        if not ok_url:
            return False, url_msg
        if rec.get("quantity", "").strip() == "0":
            return False, "数量必须大于0"
        amin, amax = rec.get("age_min", "").strip(), rec.get("age_max", "").strip()
        if amin and amax and int(amax) < int(amin):
            return False, "年龄上限必须大于等于年龄下限"
        return True, ""

    def validate_urls(self, raw: str):
        text = str(raw or "")
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        for idx, line in enumerate(lines, start=1):
            if " " in line:
                return False, f"URL 第{idx}行包含空格，请删除空格"
            parsed = urlparse(line)
            if parsed.scheme.lower() not in ("http", "https"):
                return False, f"URL 第{idx}行必须以 http:// 或 https:// 开头"
            if not parsed.netloc:
                return False, f"URL 第{idx}行缺少域名"
            if len(line) > 2048:
                return False, f"URL 第{idx}行长度超过 2048"
        return True, ""

    def validate_export_record(self, rec):
        required_fields = [
            ("task_name", "任务名"),
            ("type_code", "类型"),
            ("operator_code", "运营商"),
            ("industry_code", "行业编码"),
            ("quantity", "数量"),
            ("duration", "时长"),
        ]
        for field, title in required_fields:
            if not str(rec.get(field, "")).strip():
                return False, f"{title} 为必填项"
        tn = str(rec.get("task_name", "")).strip()
        if "-" not in tn:
            return False, "任务名格式错误：缺少 '-' 分隔"
        parsed = self.parse_left(tn.split("-", 1)[0])
        if not parsed:
            return False, "任务名格式错误：无法解析运营商/类型/行业编码"
        if parsed["op"] != rec.get("operator_code", ""):
            return False, "任务名中的运营商码值与运营商字段不匹配"
        if parsed["tp"] != rec.get("type_code", ""):
            return False, "任务名中的类型码值与类型字段不匹配"
        if parsed["ind"] != rec.get("industry_code", ""):
            return False, "任务名中的行业编码与行业编码字段不匹配"
        region_from_task = parsed.get("region", "") or "全国"
        region_from_fields = self.region_part(rec)
        if region_from_task != region_from_fields:
            return False, f"任务名中的地域“{region_from_task}”与字段地域“{region_from_fields}”不匹配"
        return True, ""

    def region_part(self, rec):
        p, c = split_multi(rec.get("province", "")), split_multi(rec.get("city", ""))
        if p and not c:
            return p[0] if len(p) == 1 else f"{int_to_cn(len(p))}省"
        if c and not p:
            return c[0] if len(c) == 1 else f"{int_to_cn(len(c))}市"
        return "全国"

    def _try_parse_left_region_six(self, left: str) -> tuple[dict | None, list[str]]:
        """地域（最前中文地域词）+ 紧随 6 位字母数字码；失败时返回 (None, 错误文案列表)。"""
        reg = find_earliest_region_in_left(left)
        if not reg:
            return None, []
        start, end, region = reg
        tail = left[end:].lstrip()
        if len(tail) < 6 or not re.match(r"^[A-Za-z0-9]{6}$", tail[:6]):
            return None, []
        code6 = tail[:6]
        op, tp, ind = code6[0:2], code6[2:4], code6[4:6]
        errs: list[str] = []
        if op not in OP_MAP:
            errs.append(f"运营商编码「{op}」输入不正确")
        if tp not in TYPE_MAP:
            errs.append(f"类型编码「{tp}」输入不正确")
        customer = (left[:start] + tail[6:]).strip() or "客户"
        if errs:
            return None, errs
        return {"customer": customer, "region": region, "op": op, "tp": tp, "ind": ind}, []

    def parse_left_with_errors(self, left: str) -> tuple[dict | None, list[str]]:
        left = (left or "").strip()
        if not left:
            return None, []
        v2, errs = self._try_parse_left_region_six(left)
        if errs:
            return None, errs
        if v2:
            return v2, []
        m = re.search(r"(YD|LT|DX|YX)(DB|DJ|XC|DH|DY)([A-Za-z0-9]+)$", left)
        if not m:
            return None, []
        prefix = left[: m.start()]
        region, customer = "", prefix
        for token in sorted(PROVINCES + CITIES + ["全国"], key=len, reverse=True):
            if prefix.endswith(token):
                customer, region = prefix[: -len(token)], token
                break
        if not region:
            rm = re.search(r"([一二两三四五六七八九十\d]+[省市])$", prefix)
            if rm:
                region = rm.group(1)
                customer = prefix[: -len(region)]
        return {"customer": customer or "客户", "region": region, "op": m.group(1), "tp": m.group(2), "ind": m.group(3)}, []

    def parse_left(self, left):
        d, _ = self.parse_left_with_errors(left)
        return d

    def build_task_name(self, rec):
        old = rec.get("task_name", "")
        product = old.split("-", 1)[1] if "-" in old else "产品"
        parsed = self.parse_left(old.split("-", 1)[0]) if "-" in old else None
        customer = parsed["customer"] if parsed else "客户"
        return f"{customer}{self.region_part(rec)}{rec.get('operator_code') or 'YD'}{rec.get('type_code') or 'DB'}{rec.get('industry_code') or 'DK'}-{product}"

    def parse_task_name(self, rec):
        self._task_name_parse_user_warned = False
        tn = rec.get("task_name", "")
        if "-" not in tn:
            return False
        left = tn.split("-", 1)[0]
        snap = {k: rec.get(k, "") for k in ("operator_code", "type_code", "industry_code", "province", "city")}
        parsed, errs = self.parse_left_with_errors(left)
        if errs:
            for k, v in snap.items():
                rec[k] = v
            QMessageBox.warning(self, "任务名", "\n".join(errs))
            self._task_name_parse_user_warned = True
            return False
        if not parsed:
            return False
        rec["operator_code"], rec["type_code"], rec["industry_code"] = parsed["op"], parsed["tp"], parsed["ind"]
        region = parsed["region"]
        if not region or region == "全国" or re.fullmatch(r"[一二两三四五六七八九十\d]+[省市]", region):
            rec["province"], rec["city"] = "", ""
        elif region in PROVINCES:
            rec["province"], rec["city"] = region, ""
        elif region in CITIES:
            rec["city"], rec["province"] = region, ""
        return True

    def after_field_change(self, rec, field):
        if field == "province" and rec.get("province", "").strip():
            rec["city"] = ""
            rec["exclude_province"] = ""
            provs = split_multi(rec.get("province", ""))
            allow = set(cities_under_provinces(provs)) if provs else set()
            if allow:
                rec["exclude_city"] = "|".join(x for x in split_multi(rec.get("exclude_city", "")) if x in allow)
            else:
                rec["exclude_city"] = ""
        elif field == "city" and rec.get("city", "").strip():
            rec["province"] = ""
            rec["exclude_province"] = ""
            rec["exclude_city"] = ""
        if field == "exclude_province" and rec.get("exclude_province", "").strip():
            rec["province"] = ""
            rec["city"] = ""
        if field == "exclude_city" and rec.get("exclude_city", "").strip():
            rec["city"] = ""
        if field == "task_name":
            return self.parse_task_name(rec)
        elif field in ("province", "city", "type_code", "operator_code", "industry_code"):
            rec["task_name"] = self.build_task_name(rec)
        return True

    def export_customer_name(self, rec):
        task_name = str(rec.get("task_name", "")).strip()
        left = task_name.split("-", 1)[0] if "-" in task_name else task_name
        parsed = self.parse_left(left) if left else None
        if parsed and str(parsed.get("customer", "")).strip():
            return sanitize_filename(str(parsed["customer"])[:12])
        fallback = left[:12] if left else "客户"
        return sanitize_filename(fallback)

    def export_excel(self):
        # 与当前表格一致：按 filtered_indices 的显示顺序，仅导出当前页可见且勾选的行
        selected_with_index: list[tuple[int, dict[str, Any]]] = []
        for rec_idx in self.filtered_indices:
            rec = self.records[rec_idx]
            if rec.get("checked"):
                selected_with_index.append((rec_idx, rec))
        if not selected_with_index:
            QMessageBox.information(self, "提示", "请先勾选至少一条记录")
            return
        include_url = self.chk_print_url.isChecked()
        selected = [x[1] for x in selected_with_index]
        # 导出前做一次全量校验，避免写出不合规数据。
        has_invalid = False
        for rec_idx, rec in selected_with_index:
            self._clear_field_errors_for_row(rec_idx)
            ok, msg = self.validate_record(rec)
            if ok:
                ok, msg = self.validate_export_record(rec)
            if ok:
                continue
            has_invalid = True
            if "年龄上限必须大于等于年龄下限" in msg:
                self.mark_error(rec_idx, "age_min", msg)
                self.mark_error(rec_idx, "age_max", msg)
            elif "运营商" in msg:
                self.mark_error(rec_idx, "operator_code", msg)
                self.mark_error(rec_idx, "task_name", msg)
            elif "类型" in msg:
                self.mark_error(rec_idx, "type_code", msg)
                self.mark_error(rec_idx, "task_name", msg)
            elif "行业编码" in msg:
                self.mark_error(rec_idx, "industry_code", msg)
                self.mark_error(rec_idx, "task_name", msg)
            elif "任务名中的地域" in msg:
                self.mark_error(rec_idx, "task_name", msg)
                self.mark_error(rec_idx, "province", msg)
                self.mark_error(rec_idx, "city", msg)
            elif "数量必须大于0" in msg or "数量" in msg:
                self.mark_error(rec_idx, "quantity", msg)
            elif "年龄上限" in msg:
                self.mark_error(rec_idx, "age_max", msg)
            elif "年龄下限" in msg:
                self.mark_error(rec_idx, "age_min", msg)
            elif "pv" in msg:
                self.mark_error(rec_idx, "pv", msg)
            elif "URL" in msg:
                self.mark_error(rec_idx, "url", msg)
            elif "时长" in msg:
                self.mark_error(rec_idx, "duration", msg)
            elif "任务名" in msg:
                self.mark_error(rec_idx, "task_name", msg)
        if has_invalid:
            self.refresh_table()
            QMessageBox.warning(self, "导出失败", "选中记录存在必填缺失或任务名关联不匹配（已标红），请修正后再导出")
            return
        self.refresh_table()
        if not TEMPLATE_FILE.exists():
            QMessageBox.warning(self, "导出失败", f"模板不存在: {TEMPLATE_FILE.name}")
            return
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb[wb.sheetnames[0]]
        col_map = {"task_name": "A", "type_code": "B", "operator_code": "C", "industry_code": "D", "url": "E", "quantity": "F", "duration": "G", "age_max": "H", "age_min": "I", "pv": "J", "province": "K", "exclude_province": "L", "city": "M", "exclude_city": "N"}
        start_row = 2
        for r in range(start_row, ws.max_row + 1):
            for f, c in col_map.items():
                ws[f"{c}{r}"].value = None
        for idx, rec in enumerate(selected):
            rr = start_row + idx
            for f, c in col_map.items():
                if f == "url":
                    raw_url = str(rec.get("url", ""))
                    row_allow = coerce_allow_print_url(rec.get("allow_print_url"))
                    if include_url and row_allow:
                        lines = [line for line in raw_url.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
                        ws[f"{c}{rr}"] = "\r\n".join(lines)
                    else:
                        ws[f"{c}{rr}"] = ""
                elif f == "type_code":
                    ws[f"{c}{rr}"] = TYPE_MAP.get(rec.get(f, ""), "")
                elif f == "operator_code":
                    ws[f"{c}{rr}"] = OP_MAP.get(rec.get(f, ""), "")
                else:
                    ws[f"{c}{rr}"] = rec.get(f, "")
        customer = self.export_customer_name(selected[0])
        filename = f"{customer}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        out_file, _ = QFileDialog.getSaveFileName(self, "保存导出文件", str(APP_DIR / filename), "Excel (*.xlsx)")
        if not out_file:
            return
        wb.save(out_file)
        for rec_idx, _ in selected_with_index:
            self._clear_field_errors_for_row(rec_idx)
        stamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        for rec_idx, _r in selected_with_index:
            rr = self.records[rec_idx]
            try:
                c = int(rr.get("print_count", 0) or 0)
            except (TypeError, ValueError):
                c = 0
            rr["print_count"] = c + 1
            rr["last_printed_at"] = stamp
        self.save_data()
        self.refresh_table()
        self.print_records.insert(
            0,
            {
                "id": str(uuid.uuid4()),
                "path": out_file,
                "filename": os.path.basename(out_file),
                "printed_at": stamp,
                "row_count": len(selected),
                "include_print_url": include_url,
            },
        )
        self.save_print_records()
        if self.content_stack.currentIndex() == 2:
            self.refresh_print_records_table()
        QMessageBox.information(self, "导出成功", f"已导出: {out_file}")


def _show_license_expired_dialog(parent: QWidget | None = None) -> None:
    """无控制台（pythonw）时也必须能看见提示；运行中过期时 parent 为主窗口。"""
    box = QMessageBox(parent)
    box.setIcon(QMessageBox.Icon.Critical)
    box.setWindowTitle("授权已过期")
    box.setText(LICENSE_EXPIRED_MSG)
    box.setStandardButtons(QMessageBox.StandardButton.Ok)
    box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    box.exec()


def main():
    app = QApplication(sys.argv)
    if not is_license_valid():
        _show_license_expired_dialog()
        sys.exit(0)
    ui_font = QFont()
    ui_font.setFamilies(["Segoe UI Variable", "Segoe UI", "Microsoft YaHei UI", "PingFang SC"])
    ui_font.setPointSize(10)
    ui_font.setWeight(QFont.Weight.Medium)
    app.setFont(ui_font)
    try:
        win = BillApp()
        win.show()
    except Exception:
        import traceback

        err_path = Path(os.environ.get("TEMP", os.environ.get("TMP", "."))) / "TidanMgr_startup_error.log"
        try:
            err_path.write_text(traceback.format_exc(), encoding="utf-8")
        except OSError:
            pass
        QMessageBox.critical(
            None,
            "启动失败",
            f"程序启动异常，详情已尝试写入：\n{err_path}\n\n{traceback.format_exc()[-800:]}",
        )
        sys.exit(1)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
