from __future__ import annotations

from PyQt6.QtCore import QPointF, QRectF, Qt, QTimer, QLineF
from datetime import date, timedelta
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QFontMetricsF,
    QPainter,
    QPainterPath,
    QPen,
    QPolygonF,
    QTextCharFormat,
    QTextCursor,
    QTextOption,
)
from PyQt6.QtWidgets import (
    QApplication,
    QGraphicsObject,
    QGraphicsPathItem,
    QGraphicsLineItem,
    QGraphicsEllipseItem,
    QGraphicsPolygonItem,
    QGraphicsRectItem,
    QGraphicsTextItem,
    QToolTip,
    QStyle,
    QStyleOptionGraphicsItem,
)

from .constants import (
    CONNECTOR_DEFAULT_COLOR,
    LINK_ARROW_SIZE,
    LINK_LINE_COLOR,
    LINK_LINE_WIDTH,
    RAG_AMBER_COLOR,
    RAG_GREEN_COLOR,
    TEXT_SIZE_MAX,
    TEXT_SIZE_MIN,
    TEXT_SIZE_STEP,
    TEXTBOX_ANCHOR_MARGIN,
    TEXTBOX_MIN_HEIGHT,
    TEXTBOX_MIN_WIDTH,
)
from .text_shortcuts import apply_text_action, extract_text_payload, text_shortcut_action

MONTH_NAMES = (
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)


def _iso_week_month(layout, base_year: int, week_index: int) -> tuple[int, int]:
    week_start = layout.week_index_to_date(base_year, week_index)
    anchor_day = week_start + timedelta(days=3)
    return anchor_day.year, anchor_day.month


def size_scale(size: int) -> float:
    scale = 0.5 + (0.1 * size)
    if scale < 0.6:
        return 0.6
    if scale > 1.0:
        return 1.0
    return scale


def _draw_arrowhead(
    painter: QPainter, start: QPointF, end: QPointF, color: QColor, size: float, *, outline: bool
) -> None:
    angle = end - start
    length = (angle.x() ** 2 + angle.y() ** 2) ** 0.5
    if length == 0:
        return
    ux = angle.x() / length
    uy = angle.y() / length
    left = QPointF(end.x() - ux * size - uy * (size / 2.0), end.y() - uy * size + ux * (size / 2.0))
    right = QPointF(end.x() - ux * size + uy * (size / 2.0), end.y() - uy * size - ux * (size / 2.0))
    painter.save()
    painter.setBrush(color)
    if not outline:
        painter.setPen(Qt.PenStyle.NoPen)
    painter.drawPolygon(QPolygonF([end, left, right]))
    painter.restore()


def _normalize_arrow_direction(value: object) -> str:
    direction = str(value or "none").strip().lower()
    if direction not in ("none", "left", "right"):
        return "none"
    return direction


def _arrow_tip_depth(width: float, height: float) -> float:
    depth = max(8.0, height * 0.35)
    return min(depth, max(1.0, width * 0.45))


def _textbox_anchor_point(obj, side: str | None, offset: float | None) -> QPointF:
    width = obj.width if obj.width is not None else TEXTBOX_MIN_WIDTH
    height = obj.height if obj.height is not None else TEXTBOX_MIN_HEIGHT
    x = obj.x if obj.x is not None else 0.0
    y = obj.y if obj.y is not None else 0.0
    offset_value = float(offset or 0.5)
    if offset_value < 0.0:
        offset_value = 0.0
    if offset_value > 1.0:
        offset_value = 1.0
    if side == "left":
        return QPointF(x, y + (height * offset_value))
    if side == "top":
        return QPointF(x + (width * offset_value), y)
    if side == "bottom":
        return QPointF(x + (width * offset_value), y + height)
    return QPointF(x + width, y + (height * offset_value))


def _object_center_for_link(obj, layout) -> QPointF | None:
    if obj.kind == "textbox":
        width = obj.width if obj.width is not None else TEXTBOX_MIN_WIDTH
        height = obj.height if obj.height is not None else TEXTBOX_MIN_HEIGHT
        x = obj.x if obj.x is not None else 0.0
        y = obj.y if obj.y is not None else 0.0
        return QPointF(x + (width / 2.0), y + (height / 2.0))
    if obj.kind == "milestone":
        if obj.row_id not in layout.row_map:
            return None
        row_height = layout.row_height(obj.row_id)
        size = min(layout.week_width, row_height) * size_scale(obj.size)
        center_x = layout.week_left_x(obj.start_week)
        center_y = layout.row_center_y(obj.row_id)
        return QPointF(center_x, center_y)
    if obj.kind == "circle":
        if obj.row_id not in layout.row_map:
            return None
        row_height = layout.row_height(obj.row_id)
        size = min(layout.week_width, row_height) * size_scale(obj.size)
        center_x = layout.week_center_x(obj.start_week)
        center_y = layout.row_center_y(obj.row_id)
        return QPointF(center_x, center_y)
    if obj.kind == "deadline":
        center_x = layout.week_left_x(obj.start_week)
        center_y = layout.header_height + (layout.total_height / 2.0)
        return QPointF(center_x, center_y)
    if obj.kind == "arrow":
        if obj.row_id not in layout.row_map:
            return None
        target_row = obj.target_row_id or obj.row_id
        if target_row not in layout.row_map:
            return None
        start_x = layout.week_center_x(obj.start_week)
        start_y = layout.row_center_y(obj.row_id)
        target_week = obj.target_week if obj.target_week is not None else obj.end_week
        end_x = layout.week_center_x(target_week)
        end_y = layout.row_center_y(target_row)
        return QPointF((start_x + end_x) / 2.0, (start_y + end_y) / 2.0)
    if obj.row_id not in layout.row_map:
        return None
    row_height = layout.row_height(obj.row_id)
    height = row_height * size_scale(obj.size)
    width = max(1, obj.end_week - obj.start_week + 1) * layout.week_width
    x = layout.week_left_x(obj.start_week)
    y = layout.row_top_y(obj.row_id) + ((row_height - height) / 2.0)
    return QPointF(x + (width / 2.0), y + (height / 2.0))


def _object_bounds_for_connector(obj, layout) -> QRectF | None:
    if obj.kind in ("link", "connector"):
        return None
    if obj.kind == "textbox":
        width = obj.width if obj.width is not None else TEXTBOX_MIN_WIDTH
        height = obj.height if obj.height is not None else TEXTBOX_MIN_HEIGHT
        x = obj.x if obj.x is not None else 0.0
        y = obj.y if obj.y is not None else 0.0
        return QRectF(x, y, width, height)
    if obj.kind == "milestone":
        if obj.row_id not in layout.row_map:
            return None
        row_height = layout.row_height(obj.row_id)
        size = min(layout.week_width, row_height) * size_scale(obj.size)
        half = size / 2.0
        center_x = layout.week_left_x(obj.start_week)
        center_y = layout.row_center_y(obj.row_id)
        return QRectF(center_x - half, center_y - half, size, size)
    if obj.kind == "circle":
        if obj.row_id not in layout.row_map:
            return None
        row_height = layout.row_height(obj.row_id)
        size = min(layout.week_width, row_height) * size_scale(obj.size)
        half = size / 2.0
        center_x = layout.week_center_x(obj.start_week)
        center_y = layout.row_center_y(obj.row_id)
        return QRectF(center_x - half, center_y - half, size, size)
    if obj.kind == "deadline":
        center_x = layout.week_left_x(obj.start_week)
        line_height = layout.header_height + layout.total_height
        width = max(1.0, float(obj.size))
        return QRectF(center_x - (width / 2.0), 0.0, width, line_height)
    if obj.kind == "arrow":
        if obj.row_id not in layout.row_map:
            return None
        target_row = obj.target_row_id or obj.row_id
        if target_row not in layout.row_map:
            return None
        start_x = layout.week_center_x(obj.start_week)
        start_y = layout.row_center_y(obj.row_id)
        target_week = obj.target_week if obj.target_week is not None else obj.end_week
        end_x = layout.week_center_x(target_week)
        end_y = layout.row_center_y(target_row)
        left = min(start_x, end_x)
        top = min(start_y, end_y)
        width = max(1.0, abs(end_x - start_x))
        height = max(1.0, abs(end_y - start_y))
        return QRectF(left, top, width, height)
    if obj.row_id not in layout.row_map:
        return None
    row_height = layout.row_height(obj.row_id)
    height = row_height * size_scale(obj.size)
    width = max(1, obj.end_week - obj.start_week + 1) * layout.week_width
    x = layout.week_left_x(obj.start_week)
    y = layout.row_top_y(obj.row_id) + ((row_height - height) / 2.0)
    return QRectF(x, y, width, height)


def _anchor_point_for_bounds(
    bounds: QRectF,
    side: str | None,
    offset: float | None,
    *,
    arrow_direction: str = "none",
) -> QPointF:
    offset_value = float(offset or 0.5)
    if offset_value < 0.0:
        offset_value = 0.0
    if offset_value > 1.0:
        offset_value = 1.0
    width = max(1.0, bounds.width())
    height = max(1.0, bounds.height())
    left = bounds.left()
    top = bounds.top()
    right = bounds.right()
    bottom = bounds.bottom()
    direction = _normalize_arrow_direction(arrow_direction)
    depth = _arrow_tip_depth(width, height)
    edge_factor = abs((offset_value * 2.0) - 1.0)
    center_factor = 1.0 - edge_factor
    if side == "left":
        x = left
        if direction == "left":
            x = left + (depth * edge_factor)
        elif direction == "right":
            x = left + (depth * center_factor)
        return QPointF(x, top + (height * offset_value))
    if side == "top":
        return QPointF(left + (width * offset_value), top)
    if side == "bottom":
        return QPointF(left + (width * offset_value), bottom)
    x = right
    if direction == "right":
        x = right - (depth * edge_factor)
    elif direction == "left":
        x = right - (depth * center_factor)
    return QPointF(x, top + (height * offset_value))


def _apply_text_alignment(text_item: QGraphicsTextItem, align: str) -> None:
    option = QTextOption()
    if align == "left":
        option.setAlignment(Qt.AlignmentFlag.AlignLeft)
    elif align == "right":
        option.setAlignment(Qt.AlignmentFlag.AlignRight)
    else:
        option.setAlignment(Qt.AlignmentFlag.AlignHCenter)
    text_item.document().setDefaultTextOption(option)


def _apply_text_color_override(text_item: QGraphicsTextItem, color: QColor) -> None:
    doc = text_item.document()
    if doc is None:
        return
    cursor = QTextCursor(doc)
    cursor.select(QTextCursor.SelectionType.Document)
    fmt = QTextCharFormat()
    fmt.setForeground(QBrush(color))
    cursor.mergeCharFormat(fmt)


def _set_text_content(text_item: QGraphicsTextItem, obj) -> None:
    if obj.text_html:
        text_item.setHtml(obj.text_html)
    else:
        text_item.setPlainText(obj.text)
    text_item.document().setDefaultFont(text_item.font())


def _active_create_tool(scene) -> str | None:
    if scene is None:
        return None
    views = scene.views()
    if not views:
        return None
    return getattr(views[0], "_create_tool", None)


class InlineTextItem(QGraphicsTextItem):
    def __init__(self, parent, object_id: str, allow_newlines: bool) -> None:
        super().__init__(parent)
        self._object_id = object_id
        self._allow_newlines = allow_newlines
        self._editing = False
        self._original_text = ""
        self._original_html = None
        self._original_font = None
        self._parent_movable = False
        self._cursor_kick_attempts = 0
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        self.setFlag(QGraphicsTextItem.GraphicsItemFlag.ItemIsFocusable, True)

    def start_edit(self) -> None:
        scene = self.scene()
        if scene is None or not getattr(scene, "edit_mode", True):
            return
        views = scene.views()
        if views:
            begin_edit = getattr(views[0], "begin_inline_edit", None)
            if callable(begin_edit) and begin_edit(self):
                return
        if self._editing:
            return
        self._editing = True
        self._original_font = QFont(self.font())
        self._original_text, self._original_html = extract_text_payload(
            self.document(), self._original_font
        )
        parent = self.parentItem()
        if parent is not None:
            self._parent_movable = bool(parent.flags() & parent.GraphicsItemFlag.ItemIsMovable)
            parent.setFlag(parent.GraphicsItemFlag.ItemIsMovable, False)
        self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)
        self.setAcceptedMouseButtons(Qt.MouseButton.LeftButton)
        self.setFocus(Qt.FocusReason.MouseFocusReason)
        scene.setFocusItem(self)
        views = scene.views()
        if views:
            views[0].setFocus(Qt.FocusReason.OtherFocusReason)
        if QApplication.cursorFlashTime() <= 0:
            QApplication.setCursorFlashTime(1000)
        self.setCursor(Qt.CursorShape.IBeamCursor)
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.setTextCursor(cursor)
        self.update()
        self._cursor_kick_attempts = 0
        QTimer.singleShot(0, self._ensure_cursor_visible)

    def paint(self, painter, option, widget=None) -> None:
        if option is None:
            super().paint(painter, option, widget)
            return
        opt = QStyleOptionGraphicsItem(option)
        opt.state &= ~QStyle.StateFlag.State_Selected
        super().paint(painter, opt, widget)

    def _ensure_cursor_visible(self) -> None:
        if not self._editing:
            return
        scene = self.scene()
        if scene is None:
            return
        self._cursor_kick_attempts += 1
        views = scene.views()
        if views:
            views[0].setFocus(Qt.FocusReason.OtherFocusReason)
        if scene.focusItem() is not self:
            self.setFocus(Qt.FocusReason.OtherFocusReason)
            scene.setFocusItem(self)
        self.setCursor(Qt.CursorShape.IBeamCursor)
        if self.textInteractionFlags() != Qt.TextInteractionFlag.TextEditorInteraction:
            self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)
        self.setTextCursor(self.textCursor())
        doc = self.document()
        if doc is not None:
            doc.markContentsDirty(0, max(1, doc.characterCount()))
        self.update()
        if self._cursor_kick_attempts < 4:
            QTimer.singleShot(60 * self._cursor_kick_attempts, self._ensure_cursor_visible)


    def _finish_edit(self, accept: bool) -> None:
        if not self._editing:
            return
        self._editing = False
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        self.unsetCursor()
        self.clearFocus()
        parent = self.parentItem()
        if parent is not None and self._parent_movable:
            parent.setFlag(parent.GraphicsItemFlag.ItemIsMovable, True)
        if not accept:
            if self._original_html:
                self.setHtml(self._original_html)
            else:
                self.setPlainText(self._original_text)
            if self._original_font is not None:
                self.setFont(self._original_font)
            self.document().setDefaultFont(self.font())
            return
        base_font = self._original_font or self.font()
        new_text, new_html = extract_text_payload(self.document(), base_font)
        if new_text == self._original_text and new_html == self._original_html:
            return
        scene = self.scene()
        if scene and hasattr(scene, "commit_object_change"):
            scene.commit_object_change(
                self._object_id, {"text": new_text, "text_html": new_html}, "Edit Text"
            )
        if self._original_font is not None:
            self.setFont(self._original_font)

    def focusOutEvent(self, event) -> None:
        self._finish_edit(True)
        super().focusOutEvent(event)

    def keyPressEvent(self, event) -> None:
        action = text_shortcut_action(event)
        if action:
            cursor = self.textCursor()
            if apply_text_action(
                cursor,
                action,
                self.font(),
                min_size=TEXT_SIZE_MIN,
                max_size=TEXT_SIZE_MAX,
                step=TEXT_SIZE_STEP,
            ):
                self.setTextCursor(cursor)
            event.accept()
            return
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if self._allow_newlines and not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
                super().keyPressEvent(event)
                return
            self._finish_edit(True)
            event.accept()
            return
        if event.key() == Qt.Key.Key_Escape:
            self._finish_edit(False)
            event.accept()
            return
        super().keyPressEvent(event)


class GridItem(QGraphicsObject):
    def __init__(self, scene_ref) -> None:
        super().__init__()
        self.scene_ref = scene_ref
        self.setZValue(-1000)
        self._rect = QRectF()
        self._hover_week: int | None = None
        self.setAcceptHoverEvents(True)
        self.setAcceptedMouseButtons(Qt.MouseButton.NoButton)

    def boundingRect(self) -> QRectF:
        return QRectF(self._rect)

    def set_rect(self, rect: QRectF) -> None:
        if rect == self._rect:
            return
        self.prepareGeometryChange()
        self._rect = QRectF(rect)

    def _week_for_hover_pos(self, pos: QPointF) -> int | None:
        scene = self.scene_ref
        layout = scene.layout
        week_y = (
            scene.header_year_height
            + scene.header_quarter_height
            + scene.header_month_height
        )
        if not (week_y <= pos.y() < (week_y + scene.header_week_height)):
            return None
        if pos.x() < layout.label_width:
            return None
        return layout.week_from_x(pos.x(), snap=False)

    def _clear_week_tooltip(self) -> None:
        if self._hover_week is None:
            return
        self._hover_week = None
        QToolTip.hideText()

    def hoverMoveEvent(self, event) -> None:
        week = self._week_for_hover_pos(event.pos())
        if week is None:
            self._clear_week_tooltip()
            super().hoverMoveEvent(event)
            return
        if week != self._hover_week:
            self._hover_week = week
            week_start = self.scene_ref.layout.week_index_to_date(
                self.scene_ref.model.year, week
            )
            QToolTip.showText(event.screenPos(), week_start.isoformat())
        super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event) -> None:
        self._clear_week_tooltip()
        super().hoverLeaveEvent(event)

    def paint(self, painter: QPainter, option, widget=None) -> None:
        scene = self.scene_ref
        layout = scene.layout
        model = scene.model
        rect = option.exposedRect if option else self.boundingRect()

        painter.fillRect(rect, QColor(252, 252, 252))

        header_rect = QRectF(rect.left(), 0, rect.width(), layout.header_height)
        painter.fillRect(header_rect, QColor(248, 248, 248))

        current_week_x: float | None = None
        if scene.show_current_week:
            today = date.today()
            iso_year, iso_week, _ = today.isocalendar()
            base_week = layout.week_index_for_iso_year(model.year, iso_year)
            current_week = base_week + (iso_week - 1)
            current_week_x = layout.week_left_x(current_week)
            current_week_rect = QRectF(
                current_week_x, rect.top(), layout.week_width, rect.height()
            )
            painter.fillRect(current_week_rect, QColor(210, 232, 255, 130))

        pen_grid = QPen(QColor(220, 220, 220))
        pen_grid.setWidth(1)
        painter.setPen(pen_grid)

        focused_row_id = getattr(scene, "focused_row_id", None)
        if focused_row_id and focused_row_id in layout.row_map:
            row = layout.row_map[focused_row_id]
            row_top = layout.header_height + row.y
            row_height = row.height
            focus_rect = QRectF(rect.left(), row_top, rect.width(), row_height)
            painter.fillRect(focus_rect, QColor(255, 243, 205, 120))

        # Vertical week lines
        left_week = layout.week_from_x(rect.left(), snap=False) - 1
        right_week = layout.week_from_x(rect.right(), snap=False) + 1
        for week in range(left_week, right_week + 1):
            if week == layout.origin_week:
                continue
            x = layout.week_left_x(week)
            painter.drawLine(int(x), int(rect.top()), int(x), int(rect.bottom()))

        # Horizontal row lines
        painter.drawLine(int(rect.left()), int(layout.header_height), int(rect.right()), int(layout.header_height))
        row_y_min = rect.top() - layout.header_height
        row_y_max = rect.bottom() - layout.header_height
        start_index, end_index = layout.row_index_range(row_y_min, row_y_max)
        for row in layout.rows[start_index:end_index]:
            y = layout.header_height + row.y + row.height
            painter.drawLine(int(rect.left()), int(y), int(rect.right()), int(y))

        # Header labels and bands
        painter.setPen(QPen(QColor(60, 60, 60)))
        font = painter.font()
        font.setBold(True)
        painter.setFont(font)

        left_year, _ = layout.week_index_to_year_week(model.year, left_week)
        right_year, _ = layout.week_index_to_year_week(model.year, right_week)
        quarter_y = scene.header_year_height
        year_index = 0
        for year in range(left_year, right_year + 1):
            year_weeks = layout.weeks_in_year(year)
            year_week = layout.week_index_for_iso_year(model.year, year)
            year_rect = QRectF(
                layout.week_left_x(year_week),
                0,
                layout.week_width * year_weeks,
                scene.header_year_height,
            )
            year_fill = QColor(246, 246, 246) if year_index % 2 == 0 else QColor(242, 242, 242)
            painter.fillRect(year_rect, year_fill)

            visible_year = year_rect.intersected(rect)
            if visible_year.width() > layout.week_width * 2:
                week_label = "week" if year_weeks == 1 else "weeks"
                label_rect = QRectF(
                    visible_year.left() + 4,
                    year_rect.top(),
                    visible_year.width() - 8,
                    year_rect.height(),
                )
                painter.drawText(
                    label_rect,
                    Qt.AlignmentFlag.AlignVCenter,
                    f"{year} ({year_weeks} {week_label})",
                )

            year_short = year % 100
            for q in range(4):
                quarter_offset = q * 13
                if quarter_offset >= year_weeks:
                    break
                quarter_length = min(13, year_weeks - quarter_offset)
                quarter_week = year_week + quarter_offset
                quarter_rect = QRectF(
                    layout.week_left_x(quarter_week),
                    quarter_y,
                    layout.week_width * quarter_length,
                    scene.header_quarter_height,
                )
                quarter_fill = QColor(252, 252, 252) if q % 2 == 0 else QColor(248, 248, 248)
                painter.fillRect(quarter_rect, quarter_fill)

                visible_quarter = quarter_rect.intersected(rect)
                if visible_quarter.width() > layout.week_width * 2:
                    label_rect = QRectF(
                        visible_quarter.left() + 4,
                        quarter_rect.top(),
                        visible_quarter.width() - 8,
                        quarter_rect.height(),
                    )
                    painter.drawText(
                        label_rect,
                        Qt.AlignmentFlag.AlignVCenter,
                        f"{year_short:02d}Q{q + 1}",
                    )
            year_index += 1

        font.setBold(False)
        painter.setFont(font)

        # Months
        month_y = scene.header_year_height + scene.header_quarter_height
        month_height = scene.header_month_height
        month_segments: list[tuple[int, int, int]] = []
        current_month = None
        segment_start = left_week
        for week in range(left_week, right_week + 1):
            month_key = _iso_week_month(layout, model.year, week)
            if current_month is None:
                current_month = month_key
                segment_start = week
                continue
            if month_key != current_month:
                month_segments.append((segment_start, week - 1, current_month[1]))
                segment_start = week
                current_month = month_key
        if current_month is not None:
            month_segments.append((segment_start, right_week, current_month[1]))

        text_pen = QPen(QColor(60, 60, 60))
        for index, (start_week, end_week, month_index) in enumerate(month_segments):
            month_rect = QRectF(
                layout.week_left_x(start_week),
                month_y,
                layout.week_width * (end_week - start_week + 1),
                month_height,
            )
            month_fill = QColor(252, 252, 252) if index % 2 == 0 else QColor(248, 248, 248)
            painter.fillRect(month_rect, month_fill)
            if index > 0:
                painter.setPen(pen_grid)
                boundary_x = layout.week_left_x(start_week)
                painter.drawLine(
                    int(boundary_x),
                    int(month_y),
                    int(boundary_x),
                    int(month_y + month_height),
                )
            painter.setPen(text_pen)
            visible_month = month_rect.intersected(rect)
            if visible_month.width() > layout.week_width * 0.6:
                label_rect = QRectF(
                    visible_month.left(),
                    month_rect.top(),
                    visible_month.width(),
                    month_rect.height(),
                )
                painter.drawText(
                    label_rect,
                    Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignCenter,
                    MONTH_NAMES[month_index - 1],
                )

        # Weeks
        week_y = scene.header_year_height + scene.header_quarter_height + month_height
        for week in range(left_week, right_week + 1):
            _, week_in_year = layout.week_index_to_year_week(model.year, week)
            x = layout.week_left_x(week)
            rect_wk = QRectF(x, week_y, layout.week_width, scene.header_week_height)
            painter.drawText(rect_wk, Qt.AlignmentFlag.AlignCenter, f"wk{week_in_year:02d}")

        # Emphasize quarter/year boundaries
        pen_quarter = QPen(QColor(200, 200, 200))
        pen_quarter.setWidth(1)
        pen_year = QPen(QColor(180, 180, 180))
        pen_year.setWidth(2)
        for year in range(left_year, right_year + 1):
            year_week = layout.week_index_for_iso_year(model.year, year)
            year_weeks = layout.weeks_in_year(year)
            x_year = layout.week_left_x(year_week)
            painter.setPen(pen_year)
            painter.drawLine(int(x_year), int(rect.top()), int(x_year), int(rect.bottom()))
            for q in range(1, 4):
                quarter_offset = q * 13
                if quarter_offset >= year_weeks:
                    break
                quarter_week = year_week + quarter_offset
                x_quarter = layout.week_left_x(quarter_week)
                painter.setPen(pen_quarter)
                painter.drawLine(int(x_quarter), int(rect.top()), int(x_quarter), int(rect.bottom()))

        if current_week_x is not None:
            now_rect = QRectF(current_week_x, 0, layout.week_width, scene.header_year_height)
            now_label = "TODAY"
            now_font = QFont(painter.font())
            now_font.setBold(True)
            max_label_width = max(0.0, now_rect.width() - 4.0)
            if max_label_width > 0.0:
                min_point_size = 6.0
                point_size = now_font.pointSizeF()
                if point_size <= 0.0:
                    point_size = 10.0
                    now_font.setPointSizeF(point_size)
                metrics = QFontMetricsF(now_font)
                while (
                    metrics.horizontalAdvance(now_label) > max_label_width
                    and point_size > min_point_size
                ):
                    point_size = max(min_point_size, point_size - 0.5)
                    now_font.setPointSizeF(point_size)
                    metrics = QFontMetricsF(now_font)
            painter.setFont(now_font)
            painter.setPen(QPen(QColor(36, 79, 120)))
            painter.drawText(now_rect, Qt.AlignmentFlag.AlignCenter, now_label)

class BoxItem(QGraphicsRectItem):
    def __init__(self, object_id: str) -> None:
        super().__init__()
        self.object_id = object_id
        self._drag_start = None
        self._drag_start_pos = None
        self._resizing = False
        self._resize_edge = None
        self._resize_start_pos = None
        self._resize_start_rect = None
        self._resize_start_obj = None
        self._resize_offset = 0.0
        self._arrow_direction = "none"
        self.text_item = InlineTextItem(self, object_id, allow_newlines=True)
        self._risk_badge = QGraphicsEllipseItem(self)
        self._risk_badge.setVisible(False)
        self._risk_badge.setBrush(QColor(255, 193, 7))
        badge_pen = QPen(QColor(80, 80, 80))
        badge_pen.setWidthF(0.6)
        badge_pen.setCosmetic(True)
        self._risk_badge.setPen(badge_pen)
        self._risk_badge.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        self.setFlags(
            QGraphicsRectItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsRectItem.GraphicsItemFlag.ItemIsMovable
        )
        self.setAcceptHoverEvents(True)

    def sync_from_model(self, obj, layout, show_missing_scope: bool | None = None) -> None:
        self.setZValue(obj.z_index)
        self._arrow_direction = _normalize_arrow_direction(
            getattr(obj, "arrow_direction", "none")
        )
        row_height = layout.row_height(obj.row_id)
        height = row_height * size_scale(obj.size)
        width = max(1, obj.end_week - obj.start_week + 1) * layout.week_width
        x = layout.week_left_x(obj.start_week)
        y = layout.row_top_y(obj.row_id) + ((row_height - height) / 2.0)
        self.setRect(0, 0, width, height)
        self.setPos(x, y)
        self.setBrush(QColor(obj.color))
        self.setPen(QPen(QColor(30, 30, 30)))
        self.text_item.setDefaultTextColor(QColor(20, 20, 20))
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        self._update_text_layout()
        self._update_risk_badge(obj, width, height, show_missing_scope=show_missing_scope)

    def _arrow_edge_insets(self, width: float, height: float) -> tuple[float, float]:
        depth = _arrow_tip_depth(width, height)
        if self._arrow_direction in ("left", "right"):
            return depth, depth
        return 0.0, 0.0

    def _shape_polygon(self) -> QPolygonF:
        rect = self.rect()
        width = rect.width()
        height = rect.height()
        if width <= 0.0 or height <= 0.0:
            return QPolygonF()
        left = rect.left()
        top = rect.top()
        right = left + width
        bottom = top + height
        middle_y = top + (height / 2.0)
        left_inset, right_inset = self._arrow_edge_insets(width, height)
        if self._arrow_direction == "left":
            return QPolygonF(
                [
                    QPointF(left + left_inset, top),
                    QPointF(right, top),
                    QPointF(right - right_inset, middle_y),
                    QPointF(right, bottom),
                    QPointF(left + left_inset, bottom),
                    QPointF(left, middle_y),
                ]
            )
        if self._arrow_direction == "right":
            return QPolygonF(
                [
                    QPointF(left, top),
                    QPointF(right - right_inset, top),
                    QPointF(right, middle_y),
                    QPointF(right - right_inset, bottom),
                    QPointF(left, bottom),
                    QPointF(left + left_inset, middle_y),
                ]
            )
        return QPolygonF(
            [
                QPointF(left, top),
                QPointF(right, top),
                QPointF(right, bottom),
                QPointF(left, bottom),
            ]
        )

    def anchor_local_point(self, side: str, offset: float) -> QPointF:
        rect = self.rect()
        return _anchor_point_for_bounds(
            rect,
            side,
            offset,
            arrow_direction=self._arrow_direction,
        )

    def _update_text_layout(self) -> None:
        rect = self.rect()
        width = rect.width()
        height = rect.height()
        left_inset, right_inset = self._arrow_edge_insets(width, height)
        left_padding = 3.0 + left_inset
        total_padding = 6.0 + left_inset + right_inset
        self.text_item.setTextWidth(max(1.0, width - total_padding))
        self.text_item.setPos(left_padding, 2.0)

    def shape(self) -> QPainterPath:
        path = QPainterPath()
        polygon = self._shape_polygon()
        if not polygon.isEmpty():
            path.addPolygon(polygon)
        return path

    def paint(self, painter: QPainter, option, widget=None) -> None:
        polygon = self._shape_polygon()
        if polygon.isEmpty():
            return
        painter.setPen(self.pen())
        painter.setBrush(self.brush())
        painter.drawPolygon(polygon)

    def _update_risk_badge(
        self,
        obj,
        width: float,
        height: float,
        *,
        show_missing_scope: bool | None = None,
    ) -> None:
        has_scope = bool(obj.scope and obj.scope.strip())
        has_risks = bool(obj.risks and obj.risks.strip())
        if show_missing_scope is None:
            scene = self.scene()
            show_missing_scope = (
                bool(getattr(scene, "show_missing_scope", False)) if scene else False
            )
        if has_risks:
            badge_color = QColor(RAG_AMBER_COLOR)
        elif has_scope:
            badge_color = QColor(RAG_GREEN_COLOR)
        else:
            if not show_missing_scope:
                self._risk_badge.setVisible(False)
                return
            badge_color = QColor(RAG_GREEN_COLOR)
        missing_scope = not has_scope and show_missing_scope
        if missing_scope:
            self._risk_badge.setBrush(QBrush(Qt.BrushStyle.NoBrush))
            badge_pen = QPen(badge_color)
            badge_pen.setStyle(Qt.PenStyle.DashLine)
            badge_pen.setWidthF(0.8)
            badge_pen.setCosmetic(True)
        else:
            self._risk_badge.setBrush(badge_color)
            badge_pen = QPen(QColor(80, 80, 80))
            badge_pen.setWidthF(0.6)
            badge_pen.setCosmetic(True)
        self._risk_badge.setPen(badge_pen)
        left_inset, right_inset = self._arrow_edge_insets(width, height)
        badge_size = min(12.0, max(6.0, height * 0.3))
        padding = 3.0
        usable_width = max(1.0, width - left_inset - right_inset - (padding * 2))
        badge_size = min(badge_size, usable_width)
        x = max(
            left_inset + padding,
            width - right_inset - badge_size - padding,
        )
        y = padding
        self._risk_badge.setRect(x, y, badge_size, badge_size)
        self._risk_badge.setVisible(True)

    def _resize_margin(self) -> float:
        scene = self.scene()
        if scene is None:
            return 6.0
        views = scene.views()
        if not views:
            return 6.0
        scale = max(0.01, views[0].transform().m11())
        return 6.0 / scale

    def _resize_edge_at(self, pos: QPointF) -> str | None:
        rect = self.rect()
        if rect.isNull():
            return None
        if pos.y() < 0.0 or pos.y() > rect.height():
            return None
        margin = self._resize_margin()
        if pos.x() <= margin:
            return "left"
        if (rect.width() - pos.x()) <= margin:
            return "right"
        return None

    def _start_resize(self, scene, edge: str, event) -> None:
        rect = self.rect()
        pos = self.pos()
        left = pos.x()
        right = pos.x() + rect.width()
        scene_x = event.scenePos().x()
        self._resizing = True
        self._resize_edge = edge
        self._resize_start_pos = QPointF(pos)
        self._resize_start_rect = QRectF(rect)
        self._resize_start_obj = scene.model.objects.get(self.object_id)
        self._resize_offset = scene_x - (left if edge == "left" else right)
        self._drag_start = None
        self._drag_start_pos = None
        self.setCursor(Qt.CursorShape.SizeHorCursor)

    def hoverMoveEvent(self, event) -> None:
        scene = self.scene()
        if scene and _active_create_tool(scene) in ("arrow", "connector"):
            self.unsetCursor()
            super().hoverMoveEvent(event)
            return
        if (
            scene
            and scene.edit_mode
            and self.isSelected()
            and self._resize_edge_at(event.pos())
        ):
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        else:
            self.unsetCursor()
        super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event) -> None:
        self.unsetCursor()
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if (
            scene
            and scene.edit_mode
            and event.button() == Qt.MouseButton.LeftButton
            and self.isSelected()
        ):
            edge = self._resize_edge_at(event.pos())
            if edge:
                self._start_resize(scene, edge, event)
                event.accept()
                return
        if scene:
            self._drag_start = scene.model.objects.get(self.object_id)
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._resizing and self._resize_start_pos and self._resize_start_rect:
            scene = self.scene()
            if scene is None:
                return
            layout = scene.layout
            snap = scene.snap_weeks
            start_pos = self._resize_start_pos
            start_rect = self._resize_start_rect
            left = start_pos.x()
            right = start_pos.x() + start_rect.width()
            min_width = layout.week_width
            scene_x = event.scenePos().x() - self._resize_offset
            if self._resize_edge == "left":
                if snap:
                    week = layout.week_from_x(scene_x, True)
                    scene_x = layout.week_left_x(week)
                new_left = min(scene_x, right - min_width)
                width = max(min_width, right - new_left)
                new_left = right - width
                self.setPos(new_left, start_pos.y())
                self.setRect(0, 0, width, start_rect.height())
            elif self._resize_edge == "right":
                if snap:
                    week = layout.week_from_x(scene_x, True)
                    scene_x = layout.week_left_x(week) + layout.week_width
                new_right = max(scene_x, left + min_width)
                width = max(min_width, new_right - left)
                self.setPos(left, start_pos.y())
                self.setRect(0, 0, width, start_rect.height())
            self._update_text_layout()
            if self._resize_start_obj:
                self._update_risk_badge(
                    self._resize_start_obj, self.rect().width(), self.rect().height()
                )
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self._resizing:
            scene = self.scene()
            if scene and self._resize_start_obj:
                layout = scene.layout
                obj = self._resize_start_obj
                pos = self.pos()
                width = self.rect().width()
                snap_week = scene.snap_weeks
                duration = max(1, int(round(width / layout.week_width)))
                start_week = layout.week_from_x(pos.x(), snap_week)
                end_week = start_week + duration - 1
                if (start_week, end_week) != (obj.start_week, obj.end_week):
                    scene.commit_object_change(
                        self.object_id,
                        {"start_week": start_week, "end_week": end_week},
                        "Resize Box",
                    )
                else:
                    self.sync_from_model(obj, layout)
            self._resizing = False
            self._resize_edge = None
            self._resize_start_pos = None
            self._resize_start_rect = None
            self._resize_start_obj = None
            self._resize_offset = 0.0
            self.unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start:
            return
        layout = scene.layout
        obj = self._drag_start
        pos = self.pos()
        width = self.rect().width()
        height = self.rect().height()

        snap_week = scene.snap_weeks
        snap_row = scene.snap_rows

        duration = max(1, int(round(width / layout.week_width)))
        start_week = layout.week_from_x(pos.x(), snap_week)
        end_week = start_week + duration - 1

        row_id = layout.row_at_y(pos.y() + height / 2.0)
        if row_id is None:
            row_id = obj.row_id

        if (start_week, end_week, row_id) != (obj.start_week, obj.end_week, obj.row_id):
            scene.commit_object_change(
                self.object_id,
                {
                    "start_week": start_week,
                    "end_week": end_week,
                    "row_id": row_id,
                },
                "Move Box",
            )
        else:
            self.sync_from_model(obj, layout)
        self._drag_start = None
        self._drag_start_pos = None

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class TextItem(QGraphicsRectItem):
    def __init__(self, object_id: str) -> None:
        super().__init__()
        self.object_id = object_id
        self._drag_start = None
        self._drag_start_pos = None
        self._resizing = False
        self._resize_edge = None
        self._resize_start_pos = None
        self._resize_start_rect = None
        self._resize_start_obj = None
        self._resize_offset = 0.0
        self.text_item = InlineTextItem(self, object_id, allow_newlines=False)
        self.setFlags(
            QGraphicsRectItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsRectItem.GraphicsItemFlag.ItemIsMovable
        )
        self.setAcceptHoverEvents(True)

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        row_height = layout.row_height(obj.row_id)
        height = row_height * size_scale(obj.size)
        width = max(1, obj.end_week - obj.start_week + 1) * layout.week_width
        x = layout.week_left_x(obj.start_week)
        y = layout.row_top_y(obj.row_id) + ((row_height - height) / 2.0)
        self.setRect(0, 0, width, height)
        self.setPos(x, y)
        pen = QPen(QColor(60, 60, 60))
        pen.setStyle(Qt.PenStyle.DashLine)
        self.setPen(pen)
        self.setBrush(QBrush(Qt.BrushStyle.NoBrush))
        self.text_item.setDefaultTextColor(QColor(20, 20, 20))
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        self._update_text_layout(width)

    def _update_text_layout(self, width: float) -> None:
        self.text_item.setTextWidth(max(1.0, width - 6))
        self.text_item.setPos(3, 2)

    def _resize_margin(self) -> float:
        scene = self.scene()
        if scene is None:
            return 6.0
        views = scene.views()
        if not views:
            return 6.0
        scale = max(0.01, views[0].transform().m11())
        return 6.0 / scale

    def _resize_edge_at(self, pos: QPointF) -> str | None:
        rect = self.rect()
        if rect.isNull():
            return None
        if pos.y() < 0.0 or pos.y() > rect.height():
            return None
        margin = self._resize_margin()
        if pos.x() <= margin:
            return "left"
        if (rect.width() - pos.x()) <= margin:
            return "right"
        return None

    def _start_resize(self, scene, edge: str, event) -> None:
        rect = self.rect()
        pos = self.pos()
        left = pos.x()
        right = pos.x() + rect.width()
        scene_x = event.scenePos().x()
        self._resizing = True
        self._resize_edge = edge
        self._resize_start_pos = QPointF(pos)
        self._resize_start_rect = QRectF(rect)
        self._resize_start_obj = scene.model.objects.get(self.object_id)
        self._resize_offset = scene_x - (left if edge == "left" else right)
        self._drag_start = None
        self._drag_start_pos = None
        self.setCursor(Qt.CursorShape.SizeHorCursor)

    def hoverMoveEvent(self, event) -> None:
        scene = self.scene()
        if scene and _active_create_tool(scene) in ("arrow", "connector"):
            self.unsetCursor()
            super().hoverMoveEvent(event)
            return
        if (
            scene
            and scene.edit_mode
            and self.isSelected()
            and self._resize_edge_at(event.pos())
        ):
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        else:
            self.unsetCursor()
        super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event) -> None:
        self.unsetCursor()
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if (
            scene
            and scene.edit_mode
            and event.button() == Qt.MouseButton.LeftButton
            and self.isSelected()
        ):
            edge = self._resize_edge_at(event.pos())
            if edge:
                self._start_resize(scene, edge, event)
                event.accept()
                return
        if scene:
            self._drag_start = scene.model.objects.get(self.object_id)
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._resizing and self._resize_start_pos and self._resize_start_rect:
            scene = self.scene()
            if scene is None:
                return
            layout = scene.layout
            snap = scene.snap_weeks
            start_pos = self._resize_start_pos
            start_rect = self._resize_start_rect
            left = start_pos.x()
            right = start_pos.x() + start_rect.width()
            min_width = layout.week_width
            scene_x = event.scenePos().x() - self._resize_offset
            if self._resize_edge == "left":
                if snap:
                    week = layout.week_from_x(scene_x, True)
                    scene_x = layout.week_left_x(week)
                new_left = min(scene_x, right - min_width)
                width = max(min_width, right - new_left)
                new_left = right - width
                self.setPos(new_left, start_pos.y())
                self.setRect(0, 0, width, start_rect.height())
            elif self._resize_edge == "right":
                if snap:
                    week = layout.week_from_x(scene_x, True)
                    scene_x = layout.week_left_x(week) + layout.week_width
                new_right = max(scene_x, left + min_width)
                width = max(min_width, new_right - left)
                self.setPos(left, start_pos.y())
                self.setRect(0, 0, width, start_rect.height())
            self._update_text_layout(self.rect().width())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self._resizing:
            scene = self.scene()
            if scene and self._resize_start_obj:
                layout = scene.layout
                obj = self._resize_start_obj
                pos = self.pos()
                width = self.rect().width()
                snap_week = scene.snap_weeks
                duration = max(1, int(round(width / layout.week_width)))
                start_week = layout.week_from_x(pos.x(), snap_week)
                end_week = start_week + duration - 1
                if (start_week, end_week) != (obj.start_week, obj.end_week):
                    scene.commit_object_change(
                        self.object_id,
                        {"start_week": start_week, "end_week": end_week},
                        "Resize Text",
                    )
                else:
                    self.sync_from_model(obj, layout)
            self._resizing = False
            self._resize_edge = None
            self._resize_start_pos = None
            self._resize_start_rect = None
            self._resize_start_obj = None
            self._resize_offset = 0.0
            self.unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start:
            return
        layout = scene.layout
        obj = self._drag_start
        pos = self.pos()
        width = self.rect().width()
        height = self.rect().height()

        snap_week = scene.snap_weeks
        snap_row = scene.snap_rows

        duration = max(1, int(round(width / layout.week_width)))
        start_week = layout.week_from_x(pos.x(), snap_week)
        end_week = start_week + duration - 1

        row_id = layout.row_at_y(pos.y() + height / 2.0)
        if row_id is None:
            row_id = obj.row_id

        if (start_week, end_week, row_id) != (obj.start_week, obj.end_week, obj.row_id):
            scene.commit_object_change(
                self.object_id,
                {
                    "start_week": start_week,
                    "end_week": end_week,
                    "row_id": row_id,
                },
                "Move Text",
            )
        else:
            self.sync_from_model(obj, layout)
        self._drag_start = None
        self._drag_start_pos = None


class TextboxItem(QGraphicsRectItem):
    def __init__(self, object_id: str) -> None:
        super().__init__()
        self.object_id = object_id
        self._drag_start = None
        self._drag_start_pos = None
        self._resizing = False
        self._resize_start = None
        self._resize_start_rect = None
        self._resize_start_obj = None
        self._resize_handle_size = 10.0
        self._link_dragging = False
        self._link_preview = None
        self._link_start_side = None
        self._link_start_offset = None
        self._link_start_scene = None
        self.text_item = InlineTextItem(self, object_id, allow_newlines=True)
        self.setFlags(
            QGraphicsRectItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsRectItem.GraphicsItemFlag.ItemIsMovable
        )
        self.setAcceptHoverEvents(True)

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        width = obj.width if obj.width is not None else TEXTBOX_MIN_WIDTH
        height = obj.height if obj.height is not None else TEXTBOX_MIN_HEIGHT
        x = obj.x if obj.x is not None else 0.0
        y = obj.y if obj.y is not None else 0.0
        opacity = max(0.0, min(1.0, float(getattr(obj, "opacity", 1.0))))
        alpha = int(opacity * 255)
        self.setRect(0, 0, width, height)
        self.setPos(x, y)
        fill = QColor(obj.color)
        if not fill.isValid():
            fill = QColor(255, 255, 255)
        fill.setAlpha(alpha)
        border = fill.darker(130)
        border.setAlpha(alpha)
        self.setBrush(fill)
        self.setPen(QPen(border))
        self.text_item.setDefaultTextColor(QColor(20, 20, 20))
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        self._update_text_layout(width)

    def _update_text_layout(self, width: float) -> None:
        self.text_item.setTextWidth(max(1.0, width - 8))
        self.text_item.setPos(4, 4)

    def _resize_handle_rect(self) -> QRectF:
        rect = self.rect()
        size = self._resize_handle_size
        return QRectF(rect.width() - size, rect.height() - size, size, size)

    def _anchor_local_point(self, side: str, offset: float) -> QPointF:
        rect = self.rect()
        offset_value = max(0.0, min(1.0, float(offset)))
        if side == "left":
            return QPointF(0.0, rect.height() * offset_value)
        if side == "top":
            return QPointF(rect.width() * offset_value, 0.0)
        if side == "bottom":
            return QPointF(rect.width() * offset_value, rect.height())
        return QPointF(rect.width(), rect.height() * offset_value)

    def _edge_anchor_at(self, pos: QPointF) -> tuple[str, float] | None:
        rect = self.rect()
        if rect.isNull():
            return None
        margin = TEXTBOX_ANCHOR_MARGIN
        if not rect.adjusted(-margin, -margin, margin, margin).contains(pos):
            return None
        distances = {
            "left": pos.x(),
            "right": rect.width() - pos.x(),
            "top": pos.y(),
            "bottom": rect.height() - pos.y(),
        }
        side, dist = min(distances.items(), key=lambda item: item[1])
        if dist > margin:
            return None
        if side in ("left", "right"):
            offset = pos.y() / max(1.0, rect.height())
        else:
            offset = pos.x() / max(1.0, rect.width())
        offset = max(0.0, min(1.0, offset))
        return side, offset

    def _start_link_drag(self, side: str, offset: float) -> None:
        scene = self.scene()
        if scene is None:
            return
        self._link_dragging = True
        self._link_start_side = side
        self._link_start_offset = offset
        self._link_start_scene = self.mapToScene(self._anchor_local_point(side, offset))
        pen = QPen(QColor(LINK_LINE_COLOR))
        pen.setWidth(LINK_LINE_WIDTH)
        preview = QGraphicsLineItem(QLineF(self._link_start_scene, self._link_start_scene))
        preview.setPen(pen)
        preview.setZValue(self.zValue() + 1)
        preview.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        scene.addItem(preview)
        self._link_preview = preview
        self.setCursor(Qt.CursorShape.CrossCursor)

    def _finish_link_drag(self, scene_pos: QPointF) -> None:
        scene = self.scene()
        if scene is None:
            return
        if self._link_preview is not None:
            scene.removeItem(self._link_preview)
        self._link_preview = None
        self.unsetCursor()
        self._link_dragging = False
        self._link_start_scene = None
        side = self._link_start_side
        offset = self._link_start_offset
        self._link_start_side = None
        self._link_start_offset = None
        if side is None or offset is None:
            return
        target_id = None
        for item in scene.items(scene_pos):
            if item is self or item is self.text_item:
                continue
            obj_id = item.data(0)
            if not obj_id or obj_id == self.object_id:
                continue
            target_obj = scene.model.objects.get(obj_id)
            if target_obj is None or target_obj.kind in ("link", "textbox", "connector"):
                continue
            target_id = obj_id
            break
        if target_id and scene.controller:
            scene.controller.add_anchor_link(self.object_id, target_id, side, offset)

    def paint(self, painter: QPainter, option, widget=None) -> None:
        super().paint(painter, option, widget)
        if self.isSelected():
            handle = self._resize_handle_rect()
            painter.setBrush(QColor(230, 230, 230))
            painter.setPen(QPen(QColor(140, 140, 140)))
            painter.drawRect(handle)

    def hoverMoveEvent(self, event) -> None:
        scene = self.scene()
        if scene and _active_create_tool(scene) in ("arrow", "connector"):
            self.unsetCursor()
            super().hoverMoveEvent(event)
            return
        if (
            scene
            and scene.edit_mode
            and self.isSelected()
            and self._resize_handle_rect().contains(event.pos())
        ):
            self.setCursor(Qt.CursorShape.SizeFDiagCursor)
        else:
            self.unsetCursor()
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if (
            scene
            and scene.edit_mode
            and event.button() == Qt.MouseButton.LeftButton
            and self.isSelected()
            and self._resize_handle_rect().contains(event.pos())
        ):
            self._resizing = True
            self._resize_start = QPointF(event.pos())
            self._resize_start_rect = QRectF(self.rect())
            self._resize_start_obj = scene.model.objects.get(self.object_id)
            self._drag_start = None
            self._drag_start_pos = None
            self.setCursor(Qt.CursorShape.SizeFDiagCursor)
            event.accept()
            return
        if scene and scene.edit_mode and event.button() == Qt.MouseButton.LeftButton:
            edge = self._edge_anchor_at(event.pos())
            if edge:
                self.setSelected(True)
                self._drag_start = None
                self._drag_start_pos = None
                self._start_link_drag(edge[0], edge[1])
                event.accept()
                return
        if scene:
            self._drag_start = scene.model.objects.get(self.object_id)
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._link_dragging and self._link_preview and self._link_start_scene:
            scene_pos = self.mapToScene(event.pos())
            self._link_preview.setLine(QLineF(self._link_start_scene, scene_pos))
            event.accept()
            return
        if self._resizing and self._resize_start and self._resize_start_rect:
            delta = event.pos() - self._resize_start
            width = max(TEXTBOX_MIN_WIDTH, self._resize_start_rect.width() + delta.x())
            height = max(TEXTBOX_MIN_HEIGHT, self._resize_start_rect.height() + delta.y())
            self.setRect(0, 0, width, height)
            self._update_text_layout(width)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self._link_dragging:
            self._finish_link_drag(self.mapToScene(event.pos()))
            event.accept()
            return
        if self._resizing:
            scene = self.scene()
            if scene and self._resize_start_obj:
                layout = scene.layout
                obj = self._resize_start_obj
                pos = self.pos()
                width = self.rect().width()
                height = self.rect().height()
                start_week = layout.week_from_x(pos.x(), snap=False)
                end_week = layout.week_from_x(pos.x() + width, snap=False)
                updates = {
                    "x": pos.x(),
                    "y": pos.y(),
                    "width": width,
                    "height": height,
                    "start_week": start_week,
                    "end_week": end_week,
                }
                if (
                    pos.x() != (obj.x or 0.0)
                    or pos.y() != (obj.y or 0.0)
                    or width != (obj.width or width)
                    or height != (obj.height or height)
                ):
                    scene.commit_object_change(self.object_id, updates, "Resize Textbox")
            self._resizing = False
            self._resize_start = None
            self._resize_start_rect = None
            self._resize_start_obj = None
            self.unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start:
            return
        layout = scene.layout
        obj = self._drag_start
        pos = self.pos()
        width = self.rect().width()
        height = self.rect().height()

        start_week = layout.week_from_x(pos.x(), snap=False)
        end_week = layout.week_from_x(pos.x() + width, snap=False)

        updates = {
            "x": pos.x(),
            "y": pos.y(),
            "width": width,
            "height": height,
            "start_week": start_week,
            "end_week": end_week,
        }
        if (
            pos.x() != (obj.x or 0.0)
            or pos.y() != (obj.y or 0.0)
            or width != (obj.width or width)
            or height != (obj.height or height)
        ):
            scene.commit_object_change(self.object_id, updates, "Move Textbox")
        self._drag_start = None
        self._drag_start_pos = None

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class MilestoneItem(QGraphicsPolygonItem):
    def __init__(self, object_id: str) -> None:
        super().__init__()
        self.object_id = object_id
        self._drag_start = None
        self._drag_start_pos = None
        self.text_item = InlineTextItem(self, object_id, allow_newlines=False)
        self.setFlags(
            QGraphicsPolygonItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsPolygonItem.GraphicsItemFlag.ItemIsMovable
        )

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        row_height = layout.row_height(obj.row_id)
        size = min(layout.week_width, row_height) * size_scale(obj.size)
        half = size / 2.0
        center_x = layout.week_left_x(obj.start_week)
        x = center_x - half
        y = layout.row_center_y(obj.row_id) - half
        polygon = QPolygonF(
            [
                QPointF(half, 0),
                QPointF(size, half),
                QPointF(half, size),
                QPointF(0, half),
            ]
        )
        self.setPolygon(polygon)
        self.setPos(x, y)
        self.setBrush(QColor(obj.color))
        self.setPen(QPen(QColor(30, 30, 30)))
        label_width = layout.week_width * 2
        label_x = center_x - (label_width / 2.0)
        label_y = layout.row_top_y(obj.row_id) + 2
        self.text_item.setDefaultTextColor(QColor(20, 20, 20))
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        self.text_item.setTextWidth(label_width)
        self.text_item.setPos(label_x - x, label_y - y)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if scene:
            self._drag_start = scene.model.objects.get(self.object_id)
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start:
            return
        layout = scene.layout
        obj = self._drag_start
        pos = self.pos()
        size = self.polygon().boundingRect().width()

        snap_week = scene.snap_weeks
        snap_row = scene.snap_rows

        week = layout.week_from_x(pos.x() + (size / 2.0), snap_week)
        row_id = layout.row_at_y(pos.y() + (size / 2.0)) if snap_row else layout.row_at_y(pos.y())
        if row_id is None:
            row_id = obj.row_id

        if (week, row_id) != (obj.start_week, obj.row_id):
            scene.commit_object_change(
                self.object_id,
                {
                    "start_week": week,
                    "end_week": week,
                    "row_id": row_id,
                },
                "Move Milestone",
            )
        else:
            self.sync_from_model(obj, layout)
        self._drag_start = None
        self._drag_start_pos = None

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class CircleItem(QGraphicsEllipseItem):
    def __init__(self, object_id: str) -> None:
        super().__init__()
        self.object_id = object_id
        self._drag_start = None
        self._drag_start_pos = None
        self.text_item = InlineTextItem(self, object_id, allow_newlines=False)
        self.setFlags(
            QGraphicsEllipseItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsEllipseItem.GraphicsItemFlag.ItemIsMovable
        )

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        row_height = layout.row_height(obj.row_id)
        size = min(layout.week_width, row_height) * size_scale(obj.size)
        half = size / 2.0
        center_x = layout.week_center_x(obj.start_week)
        x = center_x - half
        y = layout.row_center_y(obj.row_id) - half
        self.setRect(0, 0, size, size)
        self.setPos(x, y)
        self.setBrush(QColor(obj.color))
        self.setPen(QPen(QColor(30, 30, 30)))
        label_width = layout.week_width * 2
        label_x = center_x - (label_width / 2.0)
        label_y = layout.row_top_y(obj.row_id) + 2
        self.text_item.setDefaultTextColor(QColor(20, 20, 20))
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        self.text_item.setTextWidth(label_width)
        self.text_item.setPos(label_x - x, label_y - y)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if scene:
            self._drag_start = scene.model.objects.get(self.object_id)
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start:
            return
        layout = scene.layout
        obj = self._drag_start
        pos = self.pos()
        size = self.rect().width()

        snap_week = scene.snap_weeks
        snap_row = scene.snap_rows

        week = layout.week_from_center_x(pos.x() + (size / 2.0), snap_week)
        row_id = layout.row_at_y(pos.y() + (size / 2.0)) if snap_row else layout.row_at_y(pos.y())
        if row_id is None:
            row_id = obj.row_id

        if (week, row_id) != (obj.start_week, obj.row_id):
            scene.commit_object_change(
                self.object_id,
                {
                    "start_week": week,
                    "end_week": week,
                    "row_id": row_id,
                },
                "Move Circle",
            )
        else:
            self.sync_from_model(obj, layout)
        self._drag_start = None
        self._drag_start_pos = None

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class DeadlineItem(QGraphicsLineItem):
    def __init__(self, object_id: str) -> None:
        super().__init__()
        self.object_id = object_id
        self._drag_start = None
        self._drag_start_pos = None
        self.text_item = InlineTextItem(self, object_id, allow_newlines=False)
        self.setFlags(
            QGraphicsLineItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsLineItem.GraphicsItemFlag.ItemIsMovable
        )

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        center_x = layout.week_left_x(obj.start_week)
        line_height = layout.header_height + layout.total_height
        self.setLine(0.0, 0.0, 0.0, line_height)
        self.setPos(center_x, 0.0)
        pen = QPen(QColor(obj.color))
        pen.setWidth(max(2, int(obj.size)))
        pen.setCosmetic(True)
        self.setPen(pen)
        label_width = layout.week_width * 2
        label_x = -(label_width / 2.0)
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        text_color = QColor(obj.color)
        self.text_item.setDefaultTextColor(text_color)
        _apply_text_color_override(self.text_item, text_color)
        self.text_item.setTextWidth(label_width)
        label_height = self.text_item.boundingRect().height()
        label_y = -label_height - 4
        self.text_item.setPos(label_x, label_y)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if scene:
            self._drag_start = scene.model.objects.get(self.object_id)
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start:
            return
        layout = scene.layout
        obj = self._drag_start
        pos = self.pos()
        snap_week = scene.snap_weeks
        week = layout.week_from_x(pos.x(), snap_week)
        if week != obj.start_week:
            scene.commit_object_change(
                self.object_id,
                {"start_week": week, "end_week": week},
                "Move Deadline",
            )
        else:
            self.sync_from_model(obj, layout)
        self._drag_start = None
        self._drag_start_pos = None

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class ArrowItem(QGraphicsPathItem):
    def __init__(self, object_id: str, scene_ref) -> None:
        super().__init__()
        self.object_id = object_id
        self.scene_ref = scene_ref
        self._drag_start = None
        self._drag_start_pos = None
        self._drag_start_geom = None
        self._arrow_head_start = False
        self._arrow_head_end = True
        self.text_item = InlineTextItem(self, object_id, allow_newlines=False)
        self.setFlags(
            QGraphicsPathItem.GraphicsItemFlag.ItemIsSelectable
            | QGraphicsPathItem.GraphicsItemFlag.ItemIsMovable
        )

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        self._arrow_head_start = bool(getattr(obj, "arrow_head_start", False))
        self._arrow_head_end = bool(getattr(obj, "arrow_head_end", True))
        start_point = None
        end_point = None
        model = self.scene_ref.model if self.scene_ref else None
        if model and obj.connector_source_id and obj.connector_target_id:
            source = model.objects.get(obj.connector_source_id)
            target = model.objects.get(obj.connector_target_id)
            if (
                source is not None
                and target is not None
                and source.kind not in ("link", "connector", "textbox")
                and target.kind not in ("link", "connector", "textbox")
            ):
                source_bounds = _object_bounds_for_connector(source, layout)
                target_bounds = _object_bounds_for_connector(target, layout)
                if source_bounds is not None and target_bounds is not None:
                    start_point = _anchor_point_for_bounds(
                        source_bounds,
                        obj.connector_source_side,
                        obj.connector_source_offset,
                        arrow_direction=getattr(source, "arrow_direction", "none")
                        if source.kind == "box"
                        else "none",
                    )
                    end_point = _anchor_point_for_bounds(
                        target_bounds,
                        obj.connector_target_side,
                        obj.connector_target_offset,
                        arrow_direction=getattr(target, "arrow_direction", "none")
                        if target.kind == "box"
                        else "none",
                    )

        if start_point is not None and end_point is not None:
            start_x = start_point.x()
            start_y = start_point.y()
            end_x = end_point.x()
            end_y = end_point.y()
            if obj.arrow_mid_week is not None:
                mid_x = layout.week_center_x(obj.arrow_mid_week)
            else:
                mid_x = (start_x + end_x) / 2.0
        else:
            if obj.row_id not in layout.row_map:
                self.setPath(QPainterPath())
                return
            target_row = obj.target_row_id or obj.row_id
            if target_row not in layout.row_map:
                self.setPath(QPainterPath())
                return
            start_x = layout.week_center_x(obj.start_week)
            start_y = layout.row_center_y(obj.row_id)
            target_week = obj.target_week or obj.end_week
            end_x = layout.week_center_x(target_week)
            end_y = layout.row_center_y(target_row)
            mid_week = obj.arrow_mid_week or int(round((obj.start_week + target_week) / 2.0))
            mid_x = layout.week_center_x(mid_week)

        path = QPainterPath(QPointF(start_x, start_y))
        if start_y != end_y:
            path.lineTo(mid_x, start_y)
            path.lineTo(mid_x, end_y)
        path.lineTo(end_x, end_y)

        self.setPos(0, 0)
        self.setPath(path)
        pen = QPen(QColor(obj.color))
        pen.setWidth(max(1, obj.size))
        self.setPen(pen)
        label_width = max(layout.week_width * 3, abs(end_x - start_x))
        mid_x = (start_x + end_x) / 2.0
        mid_y = (start_y + end_y) / 2.0
        if obj.row_id in layout.row_map:
            row_height = layout.row_height(obj.row_id)
        elif layout.rows:
            row_height = layout.row_height(layout.rows[0].row_id)
        else:
            row_height = 20.0
        label_x = mid_x - (label_width / 2.0)
        label_y = mid_y - (row_height * 0.6)
        self.text_item.setDefaultTextColor(QColor(20, 20, 20))
        _set_text_content(self.text_item, obj)
        _apply_text_alignment(self.text_item, obj.text_align)
        self.text_item.setTextWidth(label_width)
        self.text_item.setPos(label_x, label_y)

    def paint(self, painter: QPainter, option, widget=None) -> None:
        super().paint(painter, option, widget)
        path = self.path()
        if path.elementCount() < 2:
            return
        size = 8
        pen_color = self.pen().color()
        if self._arrow_head_end:
            last = path.elementAt(path.elementCount() - 1)
            prev = path.elementAt(path.elementCount() - 2)
            end = QPointF(last.x, last.y)
            start = QPointF(prev.x, prev.y)
            _draw_arrowhead(painter, start, end, pen_color, size, outline=True)
        if self._arrow_head_start:
            first = path.elementAt(0)
            next_point = path.elementAt(1)
            start = QPointF(first.x, first.y)
            end = QPointF(next_point.x, next_point.y)
            _draw_arrowhead(painter, end, start, pen_color, size, outline=True)

    def mousePressEvent(self, event) -> None:
        scene = self.scene()
        if scene:
            obj = scene.model.objects.get(self.object_id)
            if obj:
                layout = scene.layout
                start_pos = QPointF(layout.week_center_x(obj.start_week), layout.row_center_y(obj.row_id))
                target_week = obj.target_week or obj.end_week
                target_row = obj.target_row_id or obj.row_id
                end_pos = QPointF(layout.week_center_x(target_week), layout.row_center_y(target_row))
                self._drag_start_geom = (start_pos, end_pos)
            self._drag_start = obj
            self._drag_start_pos = QPointF(self.pos())
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        super().mouseReleaseEvent(event)
        scene = self.scene()
        if not scene or not self._drag_start or not self._drag_start_geom:
            return
        layout = scene.layout
        obj = self._drag_start
        start_pos, end_pos = self._drag_start_geom
        delta = self.pos() - (self._drag_start_pos or QPointF(0, 0))

        snap_week = scene.snap_weeks
        snap_row = scene.snap_rows

        start_week = layout.week_from_center_x(start_pos.x() + delta.x(), snap_week)
        start_row = layout.row_at_y(start_pos.y() + delta.y()) if snap_row else layout.row_at_y(start_pos.y())
        if start_row is None:
            start_row = obj.row_id

        target_week = layout.week_from_center_x(end_pos.x() + delta.x(), snap_week)
        target_row = layout.row_at_y(end_pos.y() + delta.y()) if snap_row else layout.row_at_y(end_pos.y())
        if target_row is None:
            target_row = obj.target_row_id or obj.row_id

        updates = {
            "start_week": start_week,
            "row_id": start_row,
            "end_week": target_week,
            "target_week": target_week,
            "target_row_id": target_row,
        }

        if obj.arrow_mid_week is not None:
            delta_weeks = start_week - obj.start_week
            mid_week = obj.arrow_mid_week + delta_weeks
            updates["arrow_mid_week"] = mid_week

        if (
            start_week != obj.start_week
            or target_week != (obj.target_week or obj.end_week)
            or start_row != obj.row_id
            or target_row != (obj.target_row_id or obj.row_id)
        ):
            scene.commit_object_change(self.object_id, updates, "Move Arrow")
        else:
            self.sync_from_model(obj, layout)

        self._drag_start = None
        self._drag_start_pos = None
        self._drag_start_geom = None

    def mouseDoubleClickEvent(self, event) -> None:
        scene = self.scene()
        if scene and scene.edit_mode:
            self.text_item.start_edit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class ConnectorItem(QGraphicsPathItem):
    def __init__(self, object_id: str, scene_ref) -> None:
        super().__init__()
        self.object_id = object_id
        self.scene_ref = scene_ref
        self._arrow_head_start = False
        self._arrow_head_end = True
        self.setFlags(QGraphicsPathItem.GraphicsItemFlag.ItemIsSelectable)

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        self._arrow_head_start = bool(getattr(obj, "arrow_head_start", False))
        self._arrow_head_end = bool(getattr(obj, "arrow_head_end", True))
        model = self.scene_ref.model
        source_id = obj.connector_source_id
        target_id = obj.connector_target_id
        if not source_id or not target_id:
            self.setPath(QPainterPath())
            return
        source = model.objects.get(source_id)
        target = model.objects.get(target_id)
        if (
            source is None
            or target is None
            or source.kind in ("link", "connector")
            or target.kind in ("link", "connector")
        ):
            self.setPath(QPainterPath())
            return
        source_bounds = _object_bounds_for_connector(source, layout)
        target_bounds = _object_bounds_for_connector(target, layout)
        cache = self.scene_ref.connector_cache.setdefault(obj.id, {})
        start_point = None
        end_point = None
        source_visible = source_bounds is not None
        target_visible = target_bounds is not None
        if source_bounds is not None:
            start_point = _anchor_point_for_bounds(
                source_bounds,
                obj.connector_source_side,
                obj.connector_source_offset,
                arrow_direction=getattr(source, "arrow_direction", "none")
                if source.kind == "box"
                else "none",
            )
            cache["source"] = start_point
        if target_bounds is not None:
            end_point = _anchor_point_for_bounds(
                target_bounds,
                obj.connector_target_side,
                obj.connector_target_offset,
                arrow_direction=getattr(target, "arrow_direction", "none")
                if target.kind == "box"
                else "none",
            )
            cache["target"] = end_point
        if not source_visible and not target_visible:
            self.setPath(QPainterPath())
            return
        if start_point is None:
            start_point = cache.get("source")
        if end_point is None:
            end_point = cache.get("target")
        if start_point is None or end_point is None:
            self.setPath(QPainterPath())
            return
        path = QPainterPath(start_point)
        path.lineTo(end_point)
        self.setPos(0, 0)
        self.setPath(path)
        color = QColor(obj.color)
        if not color.isValid():
            color = QColor(CONNECTOR_DEFAULT_COLOR)
        pen = QPen(color)
        pen.setWidth(max(1, obj.size))
        self.setPen(pen)

    def paint(self, painter: QPainter, option, widget=None) -> None:
        super().paint(painter, option, widget)
        path = self.path()
        if path.elementCount() < 2:
            return
        line_width = self.pen().widthF()
        if line_width <= 0.0:
            line_width = float(self.pen().width())
        if line_width <= 0.0:
            line_width = 1.0
        size = max(LINK_ARROW_SIZE, line_width * 2.5)
        pen_color = self.pen().color()
        if self._arrow_head_end:
            last = path.elementAt(path.elementCount() - 1)
            prev = path.elementAt(path.elementCount() - 2)
            end = QPointF(last.x, last.y)
            start = QPointF(prev.x, prev.y)
            _draw_arrowhead(painter, start, end, pen_color, size, outline=False)
        if self._arrow_head_start:
            first = path.elementAt(0)
            next_point = path.elementAt(1)
            start = QPointF(first.x, first.y)
            end = QPointF(next_point.x, next_point.y)
            _draw_arrowhead(painter, end, start, pen_color, size, outline=False)


class LinkItem(QGraphicsPathItem):
    def __init__(self, object_id: str, scene_ref) -> None:
        super().__init__()
        self.object_id = object_id
        self.scene_ref = scene_ref
        self.setFlags(QGraphicsPathItem.GraphicsItemFlag.ItemIsSelectable)

    def sync_from_model(self, obj, layout) -> None:
        self.setZValue(obj.z_index)
        model = self.scene_ref.model
        source_id = obj.link_source_id
        target_id = obj.link_target_id
        if not source_id or not target_id:
            self.setPath(QPainterPath())
            return
        source = model.objects.get(source_id)
        target = model.objects.get(target_id)
        if (
            source is None
            or target is None
            or source.kind != "textbox"
            or target.kind in ("link", "connector")
        ):
            self.setPath(QPainterPath())
            return
        start = _textbox_anchor_point(source, obj.link_source_side, obj.link_source_offset)
        end = _object_center_for_link(target, layout)
        if end is None:
            self.setPath(QPainterPath())
            return
        path = QPainterPath(start)
        path.lineTo(end)
        self.setPos(0, 0)
        self.setPath(path)
        color = QColor(obj.color)
        if not color.isValid():
            color = QColor(LINK_LINE_COLOR)
        pen = QPen(color)
        pen.setWidth(LINK_LINE_WIDTH)
        self.setPen(pen)

    def paint(self, painter: QPainter, option, widget=None) -> None:
        super().paint(painter, option, widget)
        path = self.path()
        if path.elementCount() < 2:
            return
        last = path.elementAt(path.elementCount() - 1)
        prev = path.elementAt(path.elementCount() - 2)
        end = QPointF(last.x, last.y)
        start = QPointF(prev.x, prev.y)
        angle = (end - start)
        length = (angle.x() ** 2 + angle.y() ** 2) ** 0.5
        if length == 0:
            return
        ux = angle.x() / length
        uy = angle.y() / length
        size = LINK_ARROW_SIZE
        left = QPointF(end.x() - ux * size - uy * (size / 2.0), end.y() - uy * size + ux * (size / 2.0))
        right = QPointF(end.x() - ux * size + uy * (size / 2.0), end.y() - uy * size - ux * (size / 2.0))
        painter.setBrush(self.pen().color())
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawPolygon(QPolygonF([end, left, right]))
