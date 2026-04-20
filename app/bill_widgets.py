"""侧栏图标、表头筛选、URL 单元格、多选对话框等可复用 Qt 组件（与主窗口逻辑解耦）。"""
from __future__ import annotations

import sys

from PySide6.QtCore import QEvent, QModelIndex, QObject, QPoint, QPointF, QRect, QRectF, Qt, QTimer
from PySide6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QFontMetricsF,
    QGuiApplication,
    QPainter,
    QPen,
    QPixmap,
    QPalette,
    QPolygon,
    QPainterPath,
    QPolygonF,
)
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QSizePolicy,
    QStyle,
    QStyleOptionViewItem,
    QStyledItemDelegate,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from bill_constants import PRINT_LOG_DATA_FIELDS


def make_sidebar_logo_pixmap(*, dark: bool, size: int = 44) -> QPixmap:
    """侧栏小图标：高效办公（看板+勾选+闪电）风格。"""
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    m = 3.5
    body = QRectF(m, m, size - 2 * m, size - 2 * m)
    if dark:
        panel_fill = QColor(23, 30, 46)
        panel_edge = QColor(104, 138, 220)
        tile_fill = QColor(38, 48, 72)
        tile_edge = QColor(82, 108, 166)
        text_c = QColor(178, 196, 238, 220)
        ok_fill = QColor(44, 166, 116)
        ok_edge = QColor(105, 214, 170)
        bolt_fill = QColor(255, 206, 84)
        bolt_edge = QColor(255, 227, 152)
    else:
        panel_fill = QColor(246, 250, 255)
        panel_edge = QColor(37, 99, 235)
        tile_fill = QColor(231, 240, 255)
        tile_edge = QColor(156, 191, 242)
        text_c = QColor(55, 84, 152, 220)
        ok_fill = QColor(30, 174, 114)
        ok_edge = QColor(127, 219, 180)
        bolt_fill = QColor(251, 191, 36)
        bolt_edge = QColor(245, 158, 11)

    # 外层面板
    panel_path = QPainterPath()
    panel_path.addRoundedRect(body, 7.0, 7.0)
    painter.fillPath(panel_path, QBrush(panel_fill))
    painter.strokePath(panel_path, QPen(panel_edge, 1.2))

    # 左侧任务块
    tile = QRectF(body.left() + 4.5, body.top() + 5.0, body.width() * 0.60, body.height() - 10.0)
    tile_path = QPainterPath()
    tile_path.addRoundedRect(tile, 5.0, 5.0)
    painter.fillPath(tile_path, QBrush(tile_fill))
    painter.strokePath(tile_path, QPen(tile_edge, 1.0))

    # 勾选圆标（表示“已完成”）
    d = min(11.0, tile.height() * 0.35)
    disc = QRectF(tile.left() + 4.0, tile.top() + 4.0, d, d)
    painter.setBrush(QBrush(ok_fill))
    painter.setPen(QPen(ok_edge, 1.0))
    painter.drawEllipse(disc)
    painter.setPen(QPen(QColor(255, 255, 255), 1.4))
    p1 = QPointF(disc.left() + d * 0.25, disc.top() + d * 0.56)
    p2 = QPointF(disc.left() + d * 0.45, disc.top() + d * 0.74)
    p3 = QPointF(disc.left() + d * 0.78, disc.top() + d * 0.34)
    painter.drawLine(p1, p2)
    painter.drawLine(p2, p3)

    # 任务行
    painter.setPen(QPen(text_c, 1.0))
    lx0 = disc.right() + 3.0
    lx1 = tile.right() - 4.0
    y0 = disc.top() + d * 0.35
    for i in range(3):
        y = y0 + i * 5.0
        painter.drawLine(QPointF(lx0, y), QPointF(lx1 - i * 3.5, y))

    # 右侧闪电（表示效率）
    bx = body.right() - 11.5
    by = body.center().y()
    bolt = QPolygonF(
        [
            QPointF(bx - 2.0, by - 8.5),
            QPointF(bx + 3.0, by - 8.5),
            QPointF(bx + 0.6, by - 2.0),
            QPointF(bx + 6.0, by - 2.0),
            QPointF(bx - 3.2, by + 8.5),
            QPointF(bx - 0.4, by + 1.5),
            QPointF(bx - 5.4, by + 1.5),
        ]
    )
    painter.setBrush(QBrush(bolt_fill))
    painter.setPen(QPen(bolt_edge, 1.0))
    painter.drawPolygon(bolt)
    painter.end()
    return pm


class _ComboLineEditPopupFilter(QObject):
    """可编辑且 lineEdit 只读时，点文字区域默认不弹列表；左键按下时打开下拉。"""

    def __init__(self, combo: QComboBox):
        super().__init__(combo)
        self._combo = combo

    def eventFilter(self, watched, event):
        if (
            event.type() == QEvent.Type.MouseButtonPress
            and event.button() == Qt.LeftButton
            and watched is self._combo.lineEdit()
        ):
            self._combo.showPopup()
        return False


class UrlCellEditor(QPlainTextEdit):
    """表内 URL：失焦时对 URL 规则弹窗提示（避免每键入一字都弹窗）。"""

    def __init__(self, bill: "BillApp", rec_idx: int):
        super().__init__()
        self.setObjectName("urlCellEditor")
        self._bill = bill
        self._rec_idx = rec_idx
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self._bill._on_url_cell_focus_out(self._rec_idx, self)


def style_combo_centered(combo: QComboBox):
    """表格内下拉：Windows 等用只读 lineEdit 以便文字居中点击展开。
    macOS：可编辑 + 在 MouseButtonPress 里同步 showPopup 会与同一次点击冲突，弹层一闪即关，故用原生非可编辑下拉。"""
    if sys.platform == "darwin":
        combo.setEditable(False)
        return
    combo.setEditable(True)
    le = combo.lineEdit()
    if le:
        le.setReadOnly(True)
        le.setAlignment(Qt.AlignCenter)
        le.installEventFilter(_ComboLineEditPopupFilter(combo))


class ColumnPickFilterPopup(QDialog):
    """列筛选：无确定/取消；勾选即写回并刷新；Qt.Popup 点击表格外关闭。勾选状态用 _selected 维护以便搜索时保留。"""

    def __init__(
        self,
        bill: "BillApp",
        mode: str,
        field: str,
        title: str,
        options: list[str],
        selected: set[str],
        anchor_bottom_right: QPoint,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowFlags(Qt.Dialog | Qt.Popup)
        self.setWindowTitle(title)
        self.resize(360, 440)
        self._bill = bill
        self._mode = mode
        self._field = field
        self._options = list(options)
        self._anchor_br = QPoint(anchor_bottom_right)
        self._suppress_list = False
        opt_set = set(options)
        self._selected = set(selected) & opt_set if selected else set()
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索选项...")
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.NoSelection)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        layout.addWidget(self.search)
        layout.addWidget(self.list_widget, 1)
        self.search.textChanged.connect(self._render)
        self.list_widget.itemChanged.connect(self._on_item_changed)
        self.list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)
        self._render()

    def _on_item_double_clicked(self, item: QListWidgetItem):
        if item.checkState() == Qt.CheckState.Checked:
            item.setCheckState(Qt.CheckState.Unchecked)
        else:
            item.setCheckState(Qt.CheckState.Checked)

    def _position_near_anchor(self):
        """弹窗左上角与筛选按钮右下角对齐，并限制在可用屏幕内。"""
        self.adjustSize()
        fg = self.frameGeometry()
        w, h = fg.width(), fg.height()
        x = self._anchor_br.x()
        y = self._anchor_br.y()
        scr = QGuiApplication.screenAt(self._anchor_br)
        if scr is None:
            scr = QApplication.primaryScreen()
        ag = scr.availableGeometry() if scr else QRect()
        if ag.width() > 0:
            x = max(ag.left(), min(x, ag.right() - w + 1))
            y = max(ag.top(), min(y, ag.bottom() - h + 1))
        self.move(x, y)

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._position_near_anchor)

    def _on_item_changed(self, item: QListWidgetItem):
        if self._suppress_list:
            return
        t = item.text()
        if item.checkState() == Qt.CheckState.Checked:
            self._selected.add(t)
        else:
            self._selected.discard(t)
        self._apply_to_bill()

    def _apply_to_bill(self):
        from bill_app import BillApp as _BillApp  # 运行期导入，避免与 bill_app 循环依赖

        sel = set(self._selected)
        if self._field in ("created_at", "last_printed_at", "printed_at"):
            sel = {_BillApp._created_at_filter_key(x) for x in sel}
        full = (
            {_BillApp._created_at_filter_key(x) for x in self._options}
            if self._field in ("created_at", "last_printed_at", "printed_at")
            else set(self._options)
        )
        target = (
            self._bill.header_filters
            if self._mode.startswith("main")
            else (
                self._bill.print_header_filters
                if self._mode.startswith("print")
                else (
                    self._bill._merge_header_filters
                    if self._mode.startswith("merge")
                    else self._bill.history_header_filters
                )
            )
        )
        if not sel or sel == full:
            target.pop(self._field, None)
        else:
            target[self._field] = sel
        if self._mode.startswith("main"):
            self._bill.refresh_table()
        elif self._mode.startswith("print"):
            self._bill.refresh_print_records_table()
        elif self._mode.startswith("merge"):
            cb = getattr(self._bill, "_merge_render_cb", None)
            if callable(cb):
                cb()
        else:
            self._bill.refresh_history_table()

    def _render(self):
        self._suppress_list = True
        self.list_widget.blockSignals(True)
        try:
            kw = self.search.text().strip().lower()
            self.list_widget.clear()
            for x in self._options:
                if kw and kw not in x.lower():
                    continue
                it = QListWidgetItem(x)
                it.setFlags(
                    Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsUserCheckable
                )
                it.setCheckState(Qt.CheckState.Checked if x in self._selected else Qt.CheckState.Unchecked)
                self.list_widget.addItem(it)
        finally:
            self.list_widget.blockSignals(False)
            self._suppress_list = False


class HoverFilterHeaderView(QHeaderView):
    """表头悬停时在右侧显示筛选三角；点击三角打开多选筛选。"""

    BTN_W = 20

    def __init__(self, parent_table: QTableWidget, bill: "BillApp", mode: str):
        super().__init__(Qt.Orientation.Horizontal, parent_table)
        self._table = parent_table
        self._bill = bill
        self._mode = mode
        self._hover_section = -1
        self.setMouseTracking(True)
        self.setDefaultAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)

    def leaveEvent(self, event):
        self._hover_section = -1
        self.viewport().update()
        super().leaveEvent(event)

    def mouseMoveEvent(self, event):
        idx = self.logicalIndexAt(event.position().toPoint())
        if idx != self._hover_section:
            self._hover_section = idx
            self.viewport().update()
        super().mouseMoveEvent(event)

    def _filter_btn_rect(self, logical_index: int) -> QRect:
        pos = self.sectionViewportPosition(logical_index)
        w = self.sectionSize(logical_index)
        h = self.height()
        return QRect(pos + w - self.BTN_W, 0, self.BTN_W, h)

    def _hit_filter_btn(self, pos: QPoint) -> int:
        idx = self.logicalIndexAt(pos)
        if idx < 0 or not self._bill._header_show_filter_btn(self._mode, idx):
            return -1
        r = self._filter_btn_rect(idx)
        if r.contains(pos):
            return idx
        return -1

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.position().toPoint()
            fi = self._hit_filter_btn(pos)
            if fi >= 0:
                self._bill._open_header_filter_from_header(self._mode, fi)
                event.accept()
                return
            idx = self.logicalIndexAt(pos)
            # 冻结首列勾选表头：自定义 HeaderView 下 sectionClicked 可能不触发，在此显式处理全选/取消全选
            if idx == 0 and self._mode == "print_rec":
                self._bill._on_print_log_header_col0_clicked()
                event.accept()
                return
            if idx == 0 and self._mode in ("main_frozen", "hist_frozen"):
                self._bill._on_frozen_header_col0_clicked(self._mode)
                event.accept()
                return
            if idx == 0 and self._mode == "merge_excel":
                self._bill._on_merge_header_section_clicked(idx)
                event.accept()
                return
            # 单击表头排序：与 sectionClicked 槽一致（避免仅 super 时不触发排序）
            if idx >= 0:
                if self._mode == "print_rec" and 1 <= idx <= len(PRINT_LOG_DATA_FIELDS):
                    self._bill._on_print_rec_header_section_clicked(idx)
                    event.accept()
                    return
                if self._mode == "main_frozen" and idx == 1:
                    self._bill._on_main_frozen_header_section_clicked(idx)
                    event.accept()
                    return
                if self._mode == "main_scroll":
                    self._bill._on_main_scroll_header_clicked(idx)
                    event.accept()
                    return
                if self._mode == "hist_frozen" and idx == 1:
                    self._bill._on_hist_frozen_header_section_clicked(idx)
                    event.accept()
                    return
                if self._mode == "hist_scroll":
                    self._bill._on_history_scroll_header_clicked(idx)
                    event.accept()
                    return
                if self._mode == "merge_excel":
                    self._bill._on_merge_header_section_clicked(idx)
                    event.accept()
                    return
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            idx = self.logicalIndexAt(event.position().toPoint())
            if idx >= 0 and self._mode == "merge_excel":
                self._bill._on_merge_header_section_clicked(idx)
                event.accept()
                return
        super().mouseDoubleClickEvent(event)

    def paintSection(self, painter: QPainter, rect: QRect, logical_index: int):
        super().paintSection(painter, rect, logical_index)
        if self._hover_section != logical_index:
            return
        if not self._bill._header_show_filter_btn(self._mode, logical_index):
            return
        tri = QRect(rect.right() - self.BTN_W + 4, rect.center().y() - 4, 10, 8)
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        c = self.palette().color(self.foregroundRole())
        painter.setPen(QPen(c, 1.2))
        painter.setBrush(QBrush(c))
        cx, top, bot = tri.center().x(), tri.top() + 1, tri.bottom() - 1
        painter.drawPolygon(
            QPolygon([QPoint(cx, bot), QPoint(tri.left() + 1, top), QPoint(tri.right() - 1, top)])
        )
        painter.restore()


class MultiSelectDialog(QDialog):
    """省份/地市等多选：仅双击行切换选中（单击不改变勾选）。"""

    def __init__(self, title: str, items: list[str], selected: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(320, 420)
        self._items = items
        self._state_all: set[str] = set(selected)
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索...")
        hint = QLabel("双击行：选中或取消")
        hint.setObjectName("hintLabel")
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SingleSelection)
        self.list_widget.itemDoubleClicked.connect(self._on_row_double_clicked)
        self.btns = QDialogButtonBox(QDialogButtonBox.Cancel | QDialogButtonBox.Ok)
        layout = QVBoxLayout(self)
        layout.addWidget(self.search)
        layout.addWidget(hint)
        layout.addWidget(self.list_widget)
        layout.addWidget(self.btns)
        self.search.textChanged.connect(self._render)
        self.btns.accepted.connect(self.accept)
        self.btns.rejected.connect(self.reject)
        self._render()

    @staticmethod
    def _row_label(x: str, checked: bool) -> str:
        return ("☑ " if checked else "☐ ") + x

    def _on_row_double_clicked(self, item: QListWidgetItem):
        x = str(item.data(Qt.UserRole))
        if x in self._state_all:
            self._state_all.discard(x)
        else:
            self._state_all.add(x)
        item.setText(self._row_label(x, x in self._state_all))

    def _render(self):
        keyword = self.search.text().strip().lower()
        self.list_widget.clear()
        for x in self._items:
            if keyword and keyword not in x.lower():
                continue
            checked = x in self._state_all
            item = QListWidgetItem(self._row_label(x, checked))
            item.setData(Qt.UserRole, x)
            item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.list_widget.addItem(item)
        for i in range(self.list_widget.count()):
            it = self.list_widget.item(i)
            name = str(it.data(Qt.UserRole))
            if name in self._state_all:
                self.list_widget.setCurrentRow(i)
                self.list_widget.scrollToItem(it)
                break

    def values(self):
        return [x for x in self._items if x in self._state_all]


class AllowPrintUrlCellDelegate(QStyledItemDelegate):
    """「允许」列：单元格为表格默认底色；中间绘制浅色圆形图标，「是/否」在圆内居中。"""

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = (opt.text or "").strip() or str(index.data(Qt.ItemDataRole.DisplayRole) or "").strip() or "是"
        yes = text == "是"
        rect = option.rect
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        pal = opt.palette
        if opt.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(rect, pal.brush(QPalette.ColorRole.Highlight))
        elif opt.features & QStyleOptionViewItem.ViewItemFeature.Alternate:
            painter.fillRect(rect, pal.brush(QPalette.ColorRole.AlternateBase))
        else:
            painter.fillRect(rect, pal.brush(QPalette.ColorRole.Base))

        margin = 6
        d = min(rect.width(), rect.height()) - 2 * margin
        d = max(24, min(int(d), 38))
        cx, cy = rect.center().x(), rect.center().y()
        disc = QRect(int(cx - d / 2), int(cy - d / 2), d, d)
        if yes:
            fill = QColor(200, 232, 204)
            border = QColor(165, 210, 172)
            pen_text = QColor(52, 118, 68)
        else:
            fill = QColor(222, 222, 230)
            border = QColor(198, 198, 208)
            pen_text = QColor(92, 92, 108)
        painter.setBrush(fill)
        painter.setPen(QPen(border, 1))
        painter.drawEllipse(disc)

        f = QFont(opt.font)
        f.setBold(True)
        for _ in range(4):
            fm = QFontMetricsF(f)
            if fm.horizontalAdvance(text) <= d - 8 and fm.height() <= d - 6:
                break
            f.setPointSizeF(max(7.5, f.pointSizeF() - 0.5))
        painter.setFont(f)
        painter.setPen(pen_text)
        painter.drawText(disc, Qt.AlignmentFlag.AlignCenter, text)
        painter.restore()


class BadgeCellDelegate(QStyledItemDelegate):
    """通用圆形徽标：主值走绿色样式，其余走灰色样式（视觉与「允许」列一致）。"""

    def __init__(self, primary_values: set[str], parent=None, badge_styles: dict[str, tuple[QColor, QColor, QColor]] | None = None):
        super().__init__(parent)
        self._primary_values = {str(x).strip() for x in primary_values}
        self._badge_styles = {str(k).strip(): v for k, v in (badge_styles or {}).items()}

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = (opt.text or "").strip() or str(index.data(Qt.ItemDataRole.DisplayRole) or "").strip()
        primary = text in self._primary_values
        rect = option.rect
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        pal = opt.palette
        if opt.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(rect, pal.brush(QPalette.ColorRole.Highlight))
        elif opt.features & QStyleOptionViewItem.ViewItemFeature.Alternate:
            painter.fillRect(rect, pal.brush(QPalette.ColorRole.AlternateBase))
        else:
            painter.fillRect(rect, pal.brush(QPalette.ColorRole.Base))

        margin = 6
        d = min(rect.width(), rect.height()) - 2 * margin
        d = max(24, min(int(d), 38))
        cx, cy = rect.center().x(), rect.center().y()
        disc = QRect(int(cx - d / 2), int(cy - d / 2), d, d)
        custom = self._badge_styles.get(text)
        if custom is not None:
            fill, border, pen_text = custom
        elif primary:
            fill = QColor(200, 232, 204)
            border = QColor(165, 210, 172)
            pen_text = QColor(52, 118, 68)
        else:
            fill = QColor(222, 222, 230)
            border = QColor(198, 198, 208)
            pen_text = QColor(92, 92, 108)
        painter.setBrush(fill)
        painter.setPen(QPen(border, 1))
        painter.drawEllipse(disc)

        f = QFont(opt.font)
        f.setBold(True)
        for _ in range(4):
            fm = QFontMetricsF(f)
            if fm.horizontalAdvance(text) <= d - 8 and fm.height() <= d - 6:
                break
            f.setPointSizeF(max(7.5, f.pointSizeF() - 0.5))
        painter.setFont(f)
        painter.setPen(pen_text)
        painter.drawText(disc, Qt.AlignmentFlag.AlignCenter, text)
        painter.restore()
