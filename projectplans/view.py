from __future__ import annotations

from dataclasses import replace

from PyQt6.QtCore import QLineF, QPoint, QPointF, QRectF, Qt
from PyQt6.QtGui import QAction, QColor, QFont, QPainter, QPen, QTextCursor
from PyQt6.QtWidgets import (
    QFrame,
    QGraphicsLineItem,
    QGraphicsTextItem,
    QGraphicsView,
    QInputDialog,
    QMenu,
    QMessageBox,
    QTextEdit,
)

from .constants import (
    CANVAS_ROW_ID,
    CONNECTOR_DEFAULT_COLOR,
    LINK_LINE_WIDTH,
    MAX_ZOOM,
    MIN_ZOOM,
    TEXT_SIZE_MAX,
    TEXT_SIZE_MIN,
    TEXT_SIZE_STEP,
    TEXTBOX_ANCHOR_MARGIN,
    TEXTBOX_MIN_HEIGHT,
    TEXTBOX_MIN_WIDTH,
)
from .text_shortcuts import apply_text_action, extract_text_payload, text_shortcut_action

CONVERTIBLE_OBJECT_TYPES = (
    ("box", "Activity"),
    ("text", "Activity Text"),
    ("milestone", "Milestone"),
    ("deadline", "Deadline"),
    ("circle", "Circle"),
)
LABEL_RESIZE_MARGIN = 6
LABEL_RESIZE_MIN_WIDTH = 80


class _InlineTextEdit(QTextEdit):
    def __init__(self, commit_cb, cancel_cb, allow_newlines: bool, parent=None) -> None:
        super().__init__(parent)
        self._commit_cb = commit_cb
        self._cancel_cb = cancel_cb
        self._allow_newlines = allow_newlines
        self._done = False

    def _commit(self) -> None:
        if self._done:
            return
        self._done = True
        self._commit_cb()

    def _cancel(self) -> None:
        if self._done:
            return
        self._done = True
        self._cancel_cb()

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
        if event.key() == Qt.Key.Key_Escape:
            self._cancel()
            event.accept()
            return
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if self._allow_newlines and not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
                super().keyPressEvent(event)
                return
            self._commit()
            event.accept()
            return
        super().keyPressEvent(event)

    def insertFromMimeData(self, source) -> None:
        text = source.text()
        if not self._allow_newlines:
            text = text.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
        self.insertPlainText(text)

    def focusOutEvent(self, event) -> None:
        self._commit()
        super().focusOutEvent(event)


class CanvasView(QGraphicsView):
    def __init__(self, scene, controller) -> None:
        super().__init__(scene)
        self.controller = controller
        self.current_zoom = 1.0
        self.last_scene_pos = None
        self._space_pan = False
        self._right_pan = False
        self._right_pan_pos = None
        self._right_click_pending = False
        self._right_click_pos = None
        self._group_drag = False
        self._group_drag_start = None
        self._group_drag_positions = {}
        self._create_tool = None
        self._create_start = None
        self._create_start_row = None
        self._create_start_week = None
        self._connector_dragging = False
        self._connector_preview = None
        self._connector_start_item = None
        self._connector_start_obj_id = None
        self._connector_start_side = None
        self._connector_start_offset = None
        self._connector_start_scene = None
        self._arrow_dragging = False
        self._arrow_preview = None
        self._arrow_start_item = None
        self._arrow_start_obj_id = None
        self._arrow_start_side = None
        self._arrow_start_offset = None
        self._arrow_start_scene = None
        self._label_resize_active = False
        self._label_resize_hover = False
        self._inline_editor = None
        self._inline_editor_item = None
        self._inline_editor_text_item = None
        self._inline_editor_obj_id = None
        self._inline_editor_original = ""
        self._inline_editor_original_html = None
        self._inline_editor_original_font = None
        self.setRenderHints(self.renderHints() | QPainter.RenderHint.Antialiasing)
        self.setDragMode(QGraphicsView.DragMode.NoDrag)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setViewportUpdateMode(QGraphicsView.ViewportUpdateMode.FullViewportUpdate)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)

    def activate_create_tool(self, kind: str | None) -> None:
        self._create_tool = kind
        self._create_start = None
        self._create_start_row = None
        self._create_start_week = None
        self._cancel_connector_drag()
        self._cancel_arrow_drag()
        if kind:
            self._apply_cursor(Qt.CursorShape.CrossCursor)
        else:
            self._apply_cursor(Qt.CursorShape.ArrowCursor)

    def active_create_tool(self) -> str | None:
        return self._create_tool

    def _cancel_connector_drag(self) -> None:
        if self._connector_preview is not None:
            scene = self.scene()
            if scene:
                scene.removeItem(self._connector_preview)
            self._connector_preview = None
        self._connector_dragging = False
        self._connector_start_item = None
        self._connector_start_obj_id = None
        self._connector_start_side = None
        self._connector_start_offset = None
        self._connector_start_scene = None

    def _cancel_arrow_drag(self) -> None:
        if self._arrow_preview is not None:
            scene = self.scene()
            if scene:
                scene.removeItem(self._arrow_preview)
            self._arrow_preview = None
        self._arrow_dragging = False
        self._arrow_start_item = None
        self._arrow_start_obj_id = None
        self._arrow_start_side = None
        self._arrow_start_offset = None
        self._arrow_start_scene = None

    def _connector_edge_margin(self) -> float:
        scale = max(0.01, self.transform().m11())
        return TEXTBOX_ANCHOR_MARGIN / scale

    def _object_item_from_graphics_item(self, item):
        current = item
        while current is not None and not current.data(0):
            current = current.parentItem()
        if current is not None and current.data(0):
            return current
        return None

    def _object_item_at_scene(self, scene_pos: QPointF, skip_id: str | None = None):
        scene = self.scene()
        if scene is None:
            return None
        for item in scene.items(scene_pos):
            obj_item = self._object_item_from_graphics_item(item)
            if obj_item is None:
                continue
            obj_id = obj_item.data(0)
            if obj_id and obj_id != skip_id:
                return obj_item
        return None

    def _edge_anchor_for_item(
        self, item, scene_pos: QPointF, *, require_edge: bool
    ) -> tuple[str, float] | None:
        pos = item.mapFromScene(scene_pos)
        bounds = item.boundingRect()
        margin = self._connector_edge_margin()
        if require_edge and not bounds.adjusted(-margin, -margin, margin, margin).contains(pos):
            return None
        if not bounds.contains(pos) and require_edge:
            return None
        left = bounds.left()
        right = bounds.right()
        top = bounds.top()
        bottom = bounds.bottom()
        distances = {
            "left": abs(pos.x() - left),
            "right": abs(right - pos.x()),
            "top": abs(pos.y() - top),
            "bottom": abs(bottom - pos.y()),
        }
        side, dist = min(distances.items(), key=lambda entry: entry[1])
        if require_edge and dist > margin:
            return None
        width = max(1.0, bounds.width())
        height = max(1.0, bounds.height())
        if side in ("left", "right"):
            offset = (pos.y() - top) / height
        else:
            offset = (pos.x() - left) / width
        offset = max(0.0, min(1.0, offset))
        return side, offset

    def _anchor_point_for_item(self, item, side: str, offset: float) -> QPointF:
        custom_anchor = getattr(item, "anchor_local_point", None)
        if callable(custom_anchor):
            return item.mapToScene(custom_anchor(side, offset))
        bounds = item.boundingRect()
        width = max(1.0, bounds.width())
        height = max(1.0, bounds.height())
        left = bounds.left()
        top = bounds.top()
        right = bounds.right()
        bottom = bounds.bottom()
        if side == "left":
            local = QPointF(left, top + (height * offset))
        elif side == "top":
            local = QPointF(left + (width * offset), top)
        elif side == "bottom":
            local = QPointF(left + (width * offset), bottom)
        else:
            local = QPointF(right, top + (height * offset))
        return item.mapToScene(local)

    def _start_connector_drag(self, scene_pos: QPointF) -> bool:
        item = self._object_item_at_scene(scene_pos)
        if item is None:
            return False
        obj_id = item.data(0)
        scene = self.scene()
        if scene is None or not obj_id:
            return False
        obj = scene.model.objects.get(obj_id)
        if obj is None or obj.kind in ("link", "connector"):
            return False
        anchor = self._edge_anchor_for_item(item, scene_pos, require_edge=True)
        if anchor is None:
            return False
        side, offset = anchor
        start_scene = self._anchor_point_for_item(item, side, offset)
        preview = QGraphicsLineItem(QLineF(start_scene, start_scene))
        default_color = (
            self.controller.connector_default_color
            if self.controller is not None
            else CONNECTOR_DEFAULT_COLOR
        )
        preview_color = QColor(default_color)
        if not preview_color.isValid():
            preview_color = QColor(CONNECTOR_DEFAULT_COLOR)
        pen = QPen(preview_color)
        pen.setWidth(LINK_LINE_WIDTH)
        preview.setPen(pen)
        preview.setZValue(item.zValue() + 1)
        preview.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        scene.addItem(preview)
        self._connector_dragging = True
        self._connector_preview = preview
        self._connector_start_item = item
        self._connector_start_obj_id = obj_id
        self._connector_start_side = side
        self._connector_start_offset = offset
        self._connector_start_scene = start_scene
        return True

    def _finish_connector_drag(self, scene_pos: QPointF) -> None:
        if not self._connector_dragging:
            return
        scene = self.scene()
        if scene is None:
            self._cancel_connector_drag()
            self._reset_cursor()
            return
        if self._connector_preview is not None:
            scene.removeItem(self._connector_preview)
        self._connector_preview = None
        source_id = self._connector_start_obj_id
        source_side = self._connector_start_side
        source_offset = self._connector_start_offset
        target_item = self._object_item_at_scene(scene_pos, skip_id=source_id)
        target_id = target_item.data(0) if target_item else None
        created = False
        if (
            source_id
            and target_id
            and source_side is not None
            and source_offset is not None
            and target_item is not None
        ):
            target_anchor = self._edge_anchor_for_item(
                target_item, scene_pos, require_edge=False
            )
            if target_anchor:
                target_side, target_offset = target_anchor
                self.controller.add_connector_arrow(
                    source_id,
                    target_id,
                    source_side,
                    source_offset,
                    target_side,
                    target_offset,
                )
                created = True
        self._connector_dragging = False
        self._connector_start_item = None
        self._connector_start_obj_id = None
        self._connector_start_side = None
        self._connector_start_offset = None
        self._connector_start_scene = None
        if created:
            self.activate_create_tool(None)
        self._reset_cursor()

    def _start_arrow_drag(self, scene_pos: QPointF) -> bool:
        item = self._object_item_at_scene(scene_pos)
        if item is None:
            return False
        obj_id = item.data(0)
        scene = self.scene()
        if scene is None or not obj_id:
            return False
        obj = scene.model.objects.get(obj_id)
        if obj is None or obj.kind in ("link", "connector", "textbox"):
            return False
        anchor = self._edge_anchor_for_item(item, scene_pos, require_edge=True)
        if anchor is None:
            return False
        side, offset = anchor
        start_scene = self._anchor_point_for_item(item, side, offset)
        preview = QGraphicsLineItem(QLineF(start_scene, start_scene))
        default_color = (
            self.controller.arrow_default_color
            if self.controller is not None
            else CONNECTOR_DEFAULT_COLOR
        )
        preview_color = QColor(default_color)
        if not preview_color.isValid():
            preview_color = QColor(CONNECTOR_DEFAULT_COLOR)
        pen = QPen(preview_color)
        pen.setWidth(LINK_LINE_WIDTH)
        preview.setPen(pen)
        preview.setZValue(item.zValue() + 1)
        preview.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        scene.addItem(preview)
        self._arrow_dragging = True
        self._arrow_preview = preview
        self._arrow_start_item = item
        self._arrow_start_obj_id = obj_id
        self._arrow_start_side = side
        self._arrow_start_offset = offset
        self._arrow_start_scene = start_scene
        return True

    def _finish_arrow_drag(self, scene_pos: QPointF) -> None:
        if not self._arrow_dragging:
            return
        scene = self.scene()
        if scene is None:
            self._cancel_arrow_drag()
            self._reset_cursor()
            return
        if self._arrow_preview is not None:
            scene.removeItem(self._arrow_preview)
        self._arrow_preview = None
        source_id = self._arrow_start_obj_id
        source_side = self._arrow_start_side
        source_offset = self._arrow_start_offset
        target_item = self._object_item_at_scene(scene_pos, skip_id=source_id)
        target_id = target_item.data(0) if target_item else None
        created = False
        if (
            source_id
            and target_id
            and source_side is not None
            and source_offset is not None
            and target_item is not None
        ):
            target_obj = scene.model.objects.get(target_id)
            if target_obj and target_obj.kind not in ("link", "connector", "textbox"):
                target_anchor = self._edge_anchor_for_item(
                    target_item, scene_pos, require_edge=False
                )
                if target_anchor:
                    target_side, target_offset = target_anchor
                    layout = scene.layout
                    start_point = self._anchor_point_for_item(
                        self._arrow_start_item, source_side, source_offset
                    )
                    end_point = self._anchor_point_for_item(
                        target_item, target_side, target_offset
                    )
                    start_week = layout.week_from_x(start_point.x(), snap=False)
                    target_week = layout.week_from_x(end_point.x(), snap=False)
                    row_id = None
                    source_obj = scene.model.objects.get(source_id)
                    if source_obj and source_obj.row_id in layout.row_map:
                        row_id = source_obj.row_id
                    elif target_obj.row_id in layout.row_map:
                        row_id = target_obj.row_id
                    if row_id is None:
                        row_id = CANVAS_ROW_ID
                    obj = self.controller.make_default_object(
                        "arrow", row_id, start_week, target_week
                    )
                    target_row = (
                        target_obj.row_id
                        if target_obj.row_id in layout.row_map
                        else None
                    )
                    obj = replace(
                        obj,
                        target_row_id=target_row or row_id,
                        target_week=target_week,
                        end_week=target_week,
                        color=self.controller.arrow_default_color,
                        size=self.controller.arrow_default_size,
                        connector_source_id=source_id,
                        connector_target_id=target_id,
                        connector_source_side=source_side,
                        connector_source_offset=source_offset,
                        connector_target_side=target_side,
                        connector_target_offset=target_offset,
                    )
                    self.controller.add_object(obj, "Add Arrow")
                    created = True
        self._arrow_dragging = False
        self._arrow_start_item = None
        self._arrow_start_obj_id = None
        self._arrow_start_side = None
        self._arrow_start_offset = None
        self._arrow_start_scene = None
        if created:
            self.activate_create_tool(None)
        self._reset_cursor()
    def set_navigation_mode(self, enabled: bool) -> None:
        if enabled:
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        else:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)

    def wheelEvent(self, event) -> None:
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            factor = 1.1 if delta > 0 else 0.9
            self.zoom_by(factor)
            return
        super().wheelEvent(event)

    def zoom_by(self, factor: float) -> None:
        new_zoom = self.current_zoom * factor
        if new_zoom < MIN_ZOOM or new_zoom > MAX_ZOOM:
            return
        self.current_zoom = new_zoom
        self.scale(factor, factor)
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()

    def set_zoom(self, zoom: float) -> None:
        zoom = max(MIN_ZOOM, min(MAX_ZOOM, zoom))
        self.resetTransform()
        self.current_zoom = zoom
        self.scale(zoom, zoom)
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()

    def zoom_to_fit(self) -> None:
        scene_rect = self.scene().sceneRect()
        if scene_rect.isNull():
            return
        self.fitInView(scene_rect, Qt.AspectRatioMode.KeepAspectRatio)
        transform = self.transform()
        self.current_zoom = transform.m11()
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()

    def zoom_to_selection(self) -> None:
        items = self.scene().selectedItems()
        if not items:
            return
        rect = None
        for item in items:
            item_rect = item.sceneBoundingRect()
            rect = item_rect if rect is None else rect.united(item_rect)
        if rect is None or rect.isNull():
            return
        self.fitInView(rect, Qt.AspectRatioMode.KeepAspectRatio)
        transform = self.transform()
        self.current_zoom = transform.m11()
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()

    def center_on_base_year(self) -> None:
        layout = self.scene().layout
        weeks = layout.weeks_in_year(self.scene().model.year)
        week = layout.origin_week + (weeks // 2)
        self.centerOn(layout.week_center_x(week), layout.header_height)

    def drawForeground(self, painter, rect) -> None:
        super().drawForeground(painter, rect)
        scene = self.scene()
        layout = scene.layout
        if not layout.rows:
            return
        scale = max(0.01, self.transform().m11())
        label_width = self._label_width_pixels()
        viewport_rect = self.viewport().rect()
        header_top = self.mapFromScene(0, 0).y()
        header_bottom = self.mapFromScene(0, layout.header_height).y()

        painter.save()
        painter.resetTransform()
        painter.setClipRect(QRectF(0, 0, label_width, viewport_rect.height()))

        painter.fillRect(QRectF(0, 0, label_width, viewport_rect.height()), QColor(245, 245, 245))
        painter.fillRect(QRectF(0, header_top, label_width, header_bottom - header_top), QColor(248, 248, 248))

        focused_row_id = getattr(scene, "focused_row_id", None)
        if focused_row_id and focused_row_id in layout.row_map:
            row = layout.row_map[focused_row_id]
            row_top = self.mapFromScene(0, layout.header_height + row.y).y()
            row_bottom = self.mapFromScene(0, layout.header_height + row.y + row.height).y()
            painter.fillRect(
                QRectF(0, row_top, label_width, row_bottom - row_top),
                QColor(255, 243, 205, 140),
            )

        selected_row_id = getattr(scene, "selected_row_id", None)
        if selected_row_id and selected_row_id in layout.row_map:
            row = layout.row_map[selected_row_id]
            row_top = self.mapFromScene(0, layout.header_height + row.y).y()
            row_bottom = self.mapFromScene(0, layout.header_height + row.y + row.height).y()
            painter.fillRect(
                QRectF(0, row_top, label_width, row_bottom - row_top),
                QColor(227, 236, 247),
            )

        pen_grid = QPen(QColor(220, 220, 220))
        pen_grid.setWidth(1)
        painter.setPen(pen_grid)
        painter.drawLine(int(label_width), 0, int(label_width), int(viewport_rect.height()))
        painter.drawLine(0, int(header_bottom), int(label_width), int(header_bottom))
        for row in layout.rows:
            row_bottom = self.mapFromScene(0, layout.header_height + row.y + row.height).y()
            if row_bottom < 0 or row_bottom > viewport_rect.height() + 1:
                continue
            painter.drawLine(0, int(row_bottom), int(label_width), int(row_bottom))

        painter.setPen(QPen(QColor(60, 60, 60)))
        font = QFont(painter.font())
        base_size = font.pointSizeF()
        if base_size > 0:
            font.setPointSizeF(max(1.0, base_size * scale))
        elif font.pixelSize() > 0:
            font.setPixelSize(max(1, int(font.pixelSize() * scale)))
        painter.setFont(font)

        topic_font = QFont(font)
        topic_font.setBold(True)
        indent_step = 12 * scale
        indicator_offset = 6 * scale
        padding = 6 * scale

        for row in layout.rows:
            row_top = self.mapFromScene(0, layout.header_height + row.y).y()
            row_bottom = self.mapFromScene(0, layout.header_height + row.y + row.height).y()
            if row_bottom < 0 or row_top > viewport_rect.height():
                continue
            row_height = row_bottom - row_top
            label_rect = QRectF(0, row_top, label_width, row_height)
            text_x = (8 * scale) + (row.indent * indent_step)

            if row.kind == "topic":
                painter.setFont(topic_font)
                topic = scene.model.get_topic(row.row_id)
                indicator = "[+]" if topic and topic.collapsed else "[-]"
                painter.drawText(
                    label_rect.adjusted(indicator_offset, 0, 0, 0),
                    Qt.AlignmentFlag.AlignVCenter,
                    indicator,
                )
                text_x += 24 * scale
            else:
                painter.setFont(font)

            text_rect = QRectF(text_x, row_top, label_width - text_x - padding, row_height)
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter, row.name)

        painter.restore()

        painter.save()
        painter.resetTransform()
        tag_text = scene.model.classification_label()
        font = QFont(painter.font())
        if font.pointSizeF() > 0:
            font.setPointSizeF(float(scene.model.classification_size))
        elif font.pixelSize() > 0:
            font.setPixelSize(max(1, int(scene.model.classification_size)))
        painter.setFont(font)
        metrics = painter.fontMetrics()
        margin = 12
        max_width = max(1, viewport_rect.width() - (margin * 2))
        tag_text = metrics.elidedText(tag_text, Qt.TextElideMode.ElideRight, int(max_width))
        painter.setPen(QColor(90, 90, 90))
        overlay_rect = QRectF(
            margin,
            margin,
            max(1.0, viewport_rect.width() - (margin * 2)),
            max(1.0, viewport_rect.height() - (margin * 2)),
        )
        painter.drawText(
            overlay_rect,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom,
            tag_text,
        )
        painter.restore()

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self._is_over_label_resize_handle(event.pos()):
            self._label_resize_active = True
            self._label_resize_hover = False
            self._apply_cursor(Qt.CursorShape.SplitHCursor)
            event.accept()
            return
        if event.button() == Qt.MouseButton.RightButton:
            self._right_click_pending = True
            self._right_pan = False
            self._right_pan_pos = None
            self._right_click_pos = event.pos()
            event.accept()
            return
        if event.button() == Qt.MouseButton.LeftButton:
            if event.pos().x() <= self._label_width_pixels():
                scene_pos = self.mapToScene(event.pos())
                row_id = self.scene().layout.row_at_y(scene_pos.y())
                self.scene().clearSelection()
                self.set_selected_row(row_id)
                event.accept()
                return
            if self.scene().edit_mode and not self._create_tool:
                selected = self.scene().selectedItems()
                if len(selected) > 1:
                    item = self.itemAt(event.pos())
                    if item and item.isSelected():
                        self._group_drag = True
                        self._group_drag_start = self.mapToScene(event.pos())
                        self._group_drag_positions = {
                            sel: QPointF(sel.pos()) for sel in selected
                        }
                        self._apply_cursor(Qt.CursorShape.SizeHorCursor)
                        event.accept()
                        return
        if (
            self._create_tool
            and event.button() == Qt.MouseButton.LeftButton
            and self.scene().edit_mode
        ):
            if self._create_tool == "connector":
                scene_pos = self.mapToScene(event.pos())
                if self._start_connector_drag(scene_pos):
                    event.accept()
                    return
                event.accept()
                return
            if self._create_tool == "arrow":
                scene_pos = self.mapToScene(event.pos())
                if self._start_arrow_drag(scene_pos):
                    event.accept()
                    return
                event.accept()
                return
            if self._create_tool != "textbox" and event.pos().x() <= self._label_width_pixels():
                super().mousePressEvent(event)
                return
            self._create_start = self.mapToScene(event.pos())
            layout = self.scene().layout
            if self._create_tool in ("textbox", "deadline"):
                self._create_start_row = None
            else:
                self._create_start_row = layout.row_at_y(self._create_start.y())
            self._create_start_week = layout.week_from_x(
                self._create_start.x(), self.scene().snap_weeks
            )
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._label_resize_active:
            self._apply_label_resize(event.pos())
            event.accept()
            return
        if (
            not self._right_click_pending
            and not self._right_pan
            and not self._group_drag
            and not self._connector_dragging
            and not self._arrow_dragging
            and not self._space_pan
        ):
            if self._is_over_label_resize_handle(event.pos()):
                if not self._label_resize_hover:
                    self._label_resize_hover = True
                    self._apply_cursor(Qt.CursorShape.SplitHCursor)
            elif self._label_resize_hover:
                self._label_resize_hover = False
                self._reset_cursor()
        if self._connector_dragging and self._connector_preview and self._connector_start_scene:
            scene_pos = self.mapToScene(event.pos())
            self._connector_preview.setLine(QLineF(self._connector_start_scene, scene_pos))
            event.accept()
            return
        if self._arrow_dragging and self._arrow_preview and self._arrow_start_scene:
            scene_pos = self.mapToScene(event.pos())
            self._arrow_preview.setLine(QLineF(self._arrow_start_scene, scene_pos))
            event.accept()
            return
        if self._group_drag and self._group_drag_start:
            scene_pos = self.mapToScene(event.pos())
            delta_x = scene_pos.x() - self._group_drag_start.x()
            for item, start_pos in self._group_drag_positions.items():
                item.setPos(QPointF(start_pos.x() + delta_x, start_pos.y()))
            event.accept()
            return
        if self._right_click_pending:
            if not self._right_pan and self._right_click_pos is not None:
                if (event.pos() - self._right_click_pos).manhattanLength() > 6:
                    self._right_pan = True
                    self._right_pan_pos = event.pos()
                    self._apply_cursor(Qt.CursorShape.ClosedHandCursor)
            if self._right_pan and self._right_pan_pos is not None:
                delta = event.pos() - self._right_pan_pos
                self._right_pan_pos = event.pos()
                hbar = self.horizontalScrollBar()
                vbar = self.verticalScrollBar()
                hbar.setValue(hbar.value() - int(delta.x()))
                vbar.setValue(vbar.value() - int(delta.y()))
                event.accept()
                return
        self.last_scene_pos = self.mapToScene(event.pos())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self._label_resize_active:
            self._label_resize_active = False
            self._reset_cursor()
            event.accept()
            return
        if event.button() == Qt.MouseButton.LeftButton and self._connector_dragging:
            self._finish_connector_drag(self.mapToScene(event.pos()))
            event.accept()
            return
        if event.button() == Qt.MouseButton.LeftButton and self._arrow_dragging:
            self._finish_arrow_drag(self.mapToScene(event.pos()))
            event.accept()
            return
        if event.button() == Qt.MouseButton.LeftButton and self._group_drag:
            scene = self.scene()
            layout = scene.layout
            scene_pos = self.mapToScene(event.pos())
            delta_x = scene_pos.x() - (self._group_drag_start.x() if self._group_drag_start else 0.0)
            start_week = layout.week_from_x(
                self._group_drag_start.x() if self._group_drag_start else 0.0, scene.snap_weeks
            )
            end_week = layout.week_from_x(
                (self._group_drag_start.x() if self._group_drag_start else 0.0) + delta_x,
                scene.snap_weeks,
            )
            delta_week = end_week - start_week
            moving_ids = set()
            moved_textbox_ids = set()
            for item in self._group_drag_positions.keys():
                obj_id = item.data(0)
                if not obj_id or obj_id not in scene.model.objects:
                    continue
                obj = scene.model.objects[obj_id]
                if obj.kind in ("link", "connector"):
                    continue
                if obj.kind == "textbox" or delta_week:
                    moving_ids.add(obj_id)
                if obj.kind == "textbox":
                    moved_textbox_ids.add(obj_id)
            for item, start_pos in self._group_drag_positions.items():
                obj_id = item.data(0)
                if not obj_id or obj_id not in scene.model.objects:
                    continue
                obj = scene.model.objects[obj_id]
                if obj.kind in ("link", "connector") or (
                    obj.kind == "arrow"
                    and obj.connector_source_id
                    and obj.connector_target_id
                ):
                    item.setPos(start_pos)
                    continue
                if obj.kind == "textbox":
                    new_x = (obj.x or start_pos.x()) + delta_x
                    width = obj.width or TEXTBOX_MIN_WIDTH
                    start_wk = layout.week_from_x(new_x, snap=False)
                    end_wk = layout.week_from_x(new_x + width, snap=False)
                    self.controller.update_object(
                        obj.id,
                        {"x": new_x, "start_week": start_wk, "end_week": end_wk},
                        "Move Textbox",
                        skip_anchor_sources=moving_ids,
                        defer_link_updates=True,
                    )
                elif delta_week:
                    self._move_object(
                        obj,
                        delta_week,
                        0,
                        skip_anchor_sources=moving_ids,
                        defer_link_updates=True,
                    )
                else:
                    item.setPos(start_pos)
            self._group_drag = False
            self._group_drag_start = None
            self._group_drag_positions = {}
            if moved_textbox_ids:
                self.controller.refresh_anchor_offsets(moved_textbox_ids)
            if self._space_pan:
                self._apply_cursor(Qt.CursorShape.OpenHandCursor)
            else:
                self._apply_cursor(Qt.CursorShape.ArrowCursor)
            event.accept()
            return
        if event.button() == Qt.MouseButton.RightButton and self._right_click_pending:
            if self._right_pan:
                self._right_pan = False
                self._right_pan_pos = None
                if self._space_pan:
                    self._apply_cursor(Qt.CursorShape.OpenHandCursor)
                else:
                    self._apply_cursor(Qt.CursorShape.ArrowCursor)
            else:
                self._show_context_menu(event.pos())
            self._right_click_pending = False
            self._right_click_pos = None
            event.accept()
            return
        if (
            self._create_tool
            and self._create_start
            and event.button() == Qt.MouseButton.LeftButton
            and self.scene().edit_mode
        ):
            self._finish_create(event.pos())
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            if event.pos().x() <= self._label_width_pixels():
                scene_pos = self.mapToScene(event.pos())
                layout = self.scene().layout
                row_id = layout.row_at_y(scene_pos.y())
                if row_id:
                    row = layout.row_map.get(row_id)
                    if row and row.kind == "topic":
                        self.scene().toggle_topic(row_id)
                        return
        super().mouseDoubleClickEvent(event)

    def keyPressEvent(self, event) -> None:
        focus_item = self.scene().focusItem()
        if (
            isinstance(focus_item, QGraphicsTextItem)
            and focus_item.textInteractionFlags()
            & Qt.TextInteractionFlag.TextEditorInteraction
        ):
            super().keyPressEvent(event)
            return
        if event.key() == Qt.Key.Key_Escape and (
            self._create_tool or self._arrow_dragging or self._connector_dragging
        ):
            self.activate_create_tool(None)
            self._reset_cursor()
            event.accept()
            return
        if event.key() == Qt.Key.Key_F2 and self._start_inline_edit():
            event.accept()
            return
        if event.key() == Qt.Key.Key_Space and not self._space_pan:
            self._space_pan = True
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            self._apply_cursor(Qt.CursorShape.OpenHandCursor)
            return

        if event.key() == Qt.Key.Key_Delete:
            self._delete_selected()
            return

        if event.key() == Qt.Key.Key_D and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.duplicate_selected()
            return

        if event.key() in (Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down):
            self._nudge_selected(event)
            return

        super().keyPressEvent(event)

    def keyReleaseEvent(self, event) -> None:
        if event.key() == Qt.Key.Key_Space and self._space_pan:
            self._space_pan = False
            self._apply_cursor(Qt.CursorShape.ArrowCursor)
            if self.scene().edit_mode:
                self.setDragMode(QGraphicsView.DragMode.NoDrag)
            else:
                self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            return
        super().keyReleaseEvent(event)

    def begin_inline_edit(self, text_item: QGraphicsTextItem) -> bool:
        if not self.scene().edit_mode:
            return False
        parent = text_item.parentItem()
        if parent is None:
            return False
        obj_id = parent.data(0)
        if not obj_id:
            return False
        obj = self.scene().model.objects.get(obj_id)
        if obj is None:
            return False
        self._finish_inline_edit(True)
        self._inline_editor_item = parent
        self._inline_editor_text_item = text_item
        self._inline_editor_obj_id = obj_id
        self._inline_editor_original = obj.text
        self._inline_editor_original_html = obj.text_html
        self._inline_editor_original_font = QFont(text_item.font())
        text_item.setVisible(False)

        allow_newlines = obj.kind in ("textbox", "text")
        editor = _InlineTextEdit(
            self._commit_inline_edit, self._cancel_inline_edit, allow_newlines, self.viewport()
        )
        editor.setAcceptRichText(True)
        editor.setFrameStyle(QFrame.Shape.NoFrame)
        editor.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        editor.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        if allow_newlines:
            editor.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        else:
            editor.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        if obj.text_html:
            editor.setHtml(obj.text_html)
        else:
            editor.setPlainText(obj.text)
        editor.document().setDefaultFont(text_item.font())
        cursor = editor.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        editor.setTextCursor(cursor)

        text_color = text_item.defaultTextColor().name()
        editor.setStyleSheet(f"color: {text_color}; background: transparent;")
        editor.setFont(text_item.font())
        if obj.text_align == "left":
            alignment = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        elif obj.text_align == "right":
            alignment = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        else:
            alignment = Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
        editor.setAlignment(alignment)

        self._inline_editor = editor
        self._update_inline_editor_geometry()
        editor.show()
        editor.setFocus(Qt.FocusReason.MouseFocusReason)
        editor.update()
        return True

    def _commit_inline_edit(self) -> None:
        self._finish_inline_edit(True)

    def _cancel_inline_edit(self) -> None:
        self._finish_inline_edit(False)

    def _finish_inline_edit(self, accept: bool) -> None:
        editor = self._inline_editor
        if editor is None:
            return
        obj_id = self._inline_editor_obj_id
        text_item = self._inline_editor_text_item
        original = self._inline_editor_original
        original_html = self._inline_editor_original_html
        base_font = self._inline_editor_original_font or editor.document().defaultFont()
        new_text = None
        new_html = None
        if isinstance(editor, QTextEdit):
            new_text, new_html = extract_text_payload(editor.document(), base_font)
        if (
            accept
            and obj_id
            and new_text is not None
            and (new_text != original or new_html != original_html)
        ):
            self.controller.update_object(
                obj_id, {"text": new_text, "text_html": new_html}, "Edit Text"
            )
        if text_item is not None:
            text_item.setVisible(True)
            text_item.update()
        editor.hide()
        editor.deleteLater()
        self._inline_editor = None
        self._inline_editor_item = None
        self._inline_editor_text_item = None
        self._inline_editor_obj_id = None
        self._inline_editor_original = ""
        self._inline_editor_original_html = None
        self._inline_editor_original_font = None
        self.setFocus(Qt.FocusReason.OtherFocusReason)

    def _inline_editor_scene_rect(self) -> QRectF | None:
        text_item = self._inline_editor_text_item
        if text_item is None:
            return None
        rect = text_item.boundingRect()
        rect = rect.adjusted(-2, -2, 2, 2)
        return text_item.mapRectToScene(rect)

    def _update_inline_editor_geometry(self) -> None:
        editor = self._inline_editor
        if editor is None:
            return
        rect_scene = self._inline_editor_scene_rect()
        if rect_scene is None:
            return
        rect_view = self.mapFromScene(rect_scene).boundingRect()
        min_height = editor.sizeHint().height()
        if rect_view.height() < min_height:
            rect_view.setHeight(min_height)
        if rect_view.width() < 40:
            rect_view.setWidth(40)
        editor.setGeometry(rect_view)

    def _start_inline_edit(self) -> bool:
        if not self.scene().edit_mode:
            return False
        selected = self.scene().selectedItems()
        if not selected:
            return False
        for item in selected:
            text_item = getattr(item, "text_item", None)
            if text_item is not None and hasattr(text_item, "start_edit"):
                text_item.start_edit()
                return True
        return False

    def scrollContentsBy(self, dx: int, dy: int) -> None:
        super().scrollContentsBy(dx, dy)
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()

    def resizeEvent(self, event) -> None:
        anchor_scene = self.mapToScene(self._viewport_anchor())
        super().resizeEvent(event)
        self._restore_view_anchor(anchor_scene, self._viewport_anchor())
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()

    def _selected_object(self):
        items = self.scene().selectedItems()
        for item in items:
            obj_id = item.data(0)
            if obj_id:
                return self.scene().model.objects.get(obj_id)
        return None

    def _delete_selected(self) -> None:
        if not self.scene().edit_mode:
            return
        obj = self._selected_object()
        if obj:
            self.controller.remove_object(obj.id)

    def duplicate_selected(self) -> None:
        if not self.scene().edit_mode:
            return
        obj = self._selected_object()
        if obj:
            cloned = self.controller.duplicate_object(obj.id)
            if cloned:
                self.scene().clearSelection()
                item = self.scene().items_by_id.get(cloned.id)
                if item:
                    item.setSelected(True)
                self.set_selected_row(cloned.row_id)

    def _nudge_selected(self, event) -> None:
        if not self.scene().edit_mode:
            return
        items = self.scene().selectedItems()
        objects = []
        for item in items:
            obj_id = item.data(0)
            if obj_id and obj_id in self.scene().model.objects:
                obj = self.scene().model.objects[obj_id]
                if obj.kind in ("link", "connector"):
                    continue
                objects.append(obj)
        if not objects:
            return
        layout = self.scene().layout
        delta_week = 0
        delta_row = 0
        if event.key() == Qt.Key.Key_Left:
            delta_week = -1
        elif event.key() == Qt.Key.Key_Right:
            delta_week = 1
        elif event.key() == Qt.Key.Key_Up:
            delta_row = -1
        elif event.key() == Qt.Key.Key_Down:
            delta_row = 1

        if len(objects) > 1:
            if delta_week == 0:
                return
            selected_ids = {obj.id for obj in objects}
            moved_textbox_ids = {obj.id for obj in objects if obj.kind == "textbox"}
            for obj in objects:
                self._move_object(
                    obj,
                    delta_week,
                    0,
                    skip_anchor_sources=selected_ids,
                    defer_link_updates=True,
                )
            if moved_textbox_ids:
                self.controller.refresh_anchor_offsets(moved_textbox_ids)
            return

        obj = objects[0]
        if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
            self._resize_object(obj, delta_week, delta_row)
        else:
            self._move_object(obj, delta_week, delta_row)

    def _move_object(
        self,
        obj,
        delta_week: int,
        delta_row: int,
        *,
        skip_anchor_sources: set[str] | None = None,
        defer_link_updates: bool = False,
    ) -> None:
        layout = self.scene().layout
        if obj.kind in ("link", "connector") or (
            obj.kind == "arrow" and obj.connector_source_id and obj.connector_target_id
        ):
            return
        if obj.kind == "textbox":
            dx = delta_week * layout.week_width
            dy = delta_row * 20
            new_x = (obj.x or 0.0) + dx
            new_y = (obj.y or 0.0) + dy
            width = obj.width or TEXTBOX_MIN_WIDTH
            start_week = layout.week_from_x(new_x, snap=False)
            end_week = layout.week_from_x(new_x + width, snap=False)
            updates = {
                "x": new_x,
                "y": new_y,
                "start_week": start_week,
                "end_week": end_week,
            }
            self.controller.update_object(
                obj.id,
                updates,
                "Move Textbox",
                skip_anchor_sources=skip_anchor_sources,
                defer_link_updates=defer_link_updates,
            )
            return
        updates = {}
        if delta_week:
            updates["start_week"] = obj.start_week + delta_week
            updates["end_week"] = obj.end_week + delta_week
            if obj.kind == "arrow":
                target_week = obj.target_week if obj.target_week is not None else obj.end_week
                updates["target_week"] = target_week + delta_week
                updates["end_week"] = target_week + delta_week
                if obj.arrow_mid_week is not None:
                    updates["arrow_mid_week"] = obj.arrow_mid_week + delta_week
        if delta_row:
            new_row = layout.adjacent_row(obj.row_id, delta_row)
            if new_row:
                updates["row_id"] = new_row
            if obj.kind == "arrow":
                target_row = obj.target_row_id or obj.row_id
                new_target = layout.adjacent_row(target_row, delta_row)
                if new_target:
                    updates["target_row_id"] = new_target
        if updates:
            self.controller.update_object(
                obj.id,
                updates,
                "Nudge Object",
                skip_anchor_sources=skip_anchor_sources,
                defer_link_updates=defer_link_updates,
            )

    def _resize_object(self, obj, delta_week: int, delta_row: int) -> None:
        if obj.kind == "textbox":
            width = obj.width or TEXTBOX_MIN_WIDTH
            height = obj.height or TEXTBOX_MIN_HEIGHT
            width = max(TEXTBOX_MIN_WIDTH, width + (delta_week * self.scene().layout.week_width))
            height = max(TEXTBOX_MIN_HEIGHT, height + (delta_row * 10))
            x = obj.x or 0.0
            start_week = self.scene().layout.week_from_x(x, snap=False)
            end_week = self.scene().layout.week_from_x(x + width, snap=False)
            self.controller.update_object(
                obj.id,
                {"width": width, "height": height, "start_week": start_week, "end_week": end_week},
                "Resize Textbox",
            )
            return
        if obj.kind in ("box", "text"):
            updates = {"end_week": obj.end_week + delta_week}
            if updates["end_week"] < obj.start_week:
                updates["end_week"] = obj.start_week
            self.controller.update_object(obj.id, updates, "Resize Object")
            return
        if (
            obj.kind == "arrow"
            and obj.connector_source_id
            and obj.connector_target_id
        ):
            return
        if obj.kind == "arrow":
            updates = {}
            if delta_week:
                target_week = obj.target_week if obj.target_week is not None else obj.end_week
                updates["target_week"] = target_week + delta_week
                updates["end_week"] = target_week + delta_week
            if delta_row:
                layout = self.scene().layout
                target_row = obj.target_row_id or obj.row_id
                new_target = layout.adjacent_row(target_row, delta_row)
                if new_target:
                    updates["target_row_id"] = new_target
            if updates:
                self.controller.update_object(obj.id, updates, "Resize Arrow")

    def _finish_create(self, pos) -> None:
        scene = self.scene()
        layout = scene.layout
        end_pos = self.mapToScene(pos)

        kind = self._create_tool
        if kind == "textbox":
            if not scene.show_textboxes:
                self.activate_create_tool(None)
                return
            start = self._create_start or end_pos
            x1 = min(start.x(), end_pos.x())
            x2 = max(start.x(), end_pos.x())
            y1 = min(start.y(), end_pos.y())
            y2 = max(start.y(), end_pos.y())
            width = max(TEXTBOX_MIN_WIDTH, x2 - x1)
            height = max(TEXTBOX_MIN_HEIGHT, y2 - y1)
            obj = self.controller.make_textbox(x1, y1, width, height)
            start_wk = layout.week_from_x(x1, snap=False)
            end_wk = layout.week_from_x(x1 + width, snap=False)
            obj = replace(obj, start_week=start_wk, end_week=end_wk)
            self.controller.add_object(obj, "Add Textbox")
        else:
            if not layout.rows and kind not in ("deadline",):
                self.activate_create_tool(None)
                return

            start_row = self._create_start_row or layout.row_at_y(end_pos.y())
            end_row = layout.row_at_y(end_pos.y()) or start_row
            if kind != "deadline":
                if not start_row:
                    start_row = layout.rows[0].row_id
                if not end_row:
                    end_row = start_row

            start_week = self._create_start_week or layout.week_from_x(
                self._create_start.x(), scene.snap_weeks
            )
            end_week = layout.week_from_x(end_pos.x(), scene.snap_weeks)

            if kind in ("box", "text"):
                if end_week < start_week:
                    start_week, end_week = end_week, start_week
                obj = self.controller.make_default_object(kind, start_row, start_week, end_week)
                self.controller.add_object(obj, f"Add {kind.title()}")
            elif kind == "milestone":
                obj = self.controller.make_default_object(kind, start_row, start_week, start_week)
                self.controller.add_object(obj, "Add Milestone")
            elif kind == "deadline":
                obj = self.controller.make_default_object(kind, CANVAS_ROW_ID, start_week, start_week)
                self.controller.add_object(obj, "Add Deadline")
            elif kind == "circle":
                obj = self.controller.make_default_object(kind, start_row, start_week, start_week)
                self.controller.add_object(obj, "Add Circle")
            elif kind == "arrow":
                if end_week == start_week and end_row == start_row:
                    end_week = start_week + 1
                obj = self.controller.make_default_object(kind, start_row, start_week, end_week)
                obj = replace(
                    obj,
                    target_row_id=end_row,
                    target_week=end_week,
                    end_week=end_week,
                    color=self.controller.arrow_default_color,
                    size=self.controller.arrow_default_size,
                )
                self.controller.add_object(obj, "Add Arrow")

        self.activate_create_tool(None)

    def _label_width_pixels(self) -> float:
        layout = self.scene().layout
        return max(1.0, layout.label_width * self.transform().m11())

    def _is_over_label_resize_handle(self, pos) -> bool:
        label_edge = self._label_width_pixels()
        return abs(pos.x() - label_edge) <= LABEL_RESIZE_MARGIN

    def _apply_label_resize(self, pos) -> None:
        scene = self.scene()
        if scene is None:
            return
        scale = max(0.01, self.transform().m11())
        width = max(LABEL_RESIZE_MIN_WIDTH, pos.x() / scale)
        old_width = scene.layout.label_width
        if abs(old_width - width) < 0.5:
            return
        delta = width - old_width
        scene.shift_textboxes(delta)
        scene.set_label_width(width)
        if abs(delta) > 0.0:
            hbar = self.horizontalScrollBar()
            hbar.setValue(hbar.value() + int(round(delta * scale)))
        self._maybe_extend_scene()
        self._update_inline_editor_geometry()
        self.viewport().update()

    def _reset_cursor(self) -> None:
        if self._space_pan:
            self._apply_cursor(Qt.CursorShape.OpenHandCursor)
        elif self._right_pan:
            self._apply_cursor(Qt.CursorShape.ClosedHandCursor)
        elif self._create_tool:
            self._apply_cursor(Qt.CursorShape.CrossCursor)
        else:
            self._apply_cursor(Qt.CursorShape.ArrowCursor)

    def _apply_cursor(self, cursor: Qt.CursorShape) -> None:
        self.setCursor(cursor)
        self.viewport().setCursor(cursor)

    def _viewport_anchor(self) -> QPoint:
        return self.viewport().rect().center()

    def _restore_view_anchor(self, anchor_scene: QPointF, anchor_view: QPoint) -> None:
        new_view = self.mapFromScene(anchor_scene)
        delta = new_view - anchor_view
        if delta.x() == 0 and delta.y() == 0:
            return
        hbar = self.horizontalScrollBar()
        vbar = self.verticalScrollBar()
        hbar.setValue(hbar.value() + delta.x())
        vbar.setValue(vbar.value() + delta.y())

    def _maybe_extend_scene(self) -> None:
        scene = self.scene()
        if scene is None:
            return
        layout = scene.layout
        left_scene = self.mapToScene(0, 0).x()
        right_scene = self.mapToScene(self.viewport().width(), 0).x()
        scene.ensure_week_range(layout.week_from_x(left_scene, snap=False))
        scene.ensure_week_range(layout.week_from_x(right_scene, snap=False))

    def set_selected_row(self, row_id: str | None) -> None:
        scene = self.scene()
        if scene is None:
            return
        scene.set_selected_row(row_id)
        self.viewport().update()

    def set_focused_row(self, row_id: str | None) -> None:
        scene = self.scene()
        if scene is None:
            return
        scene.set_focused_row(row_id)
        self.viewport().update()

    def _show_context_menu(self, pos) -> None:
        scene = self.scene()
        if scene is None:
            return
        if pos.x() <= self._label_width_pixels():
            row_id = scene.layout.row_at_y(self.mapToScene(pos).y())
            row = scene.layout.row_map.get(row_id) if row_id else None
            if row and row.kind in ("topic", "deliverable"):
                self._show_row_context_menu(pos, row.row_id, row.kind)
                return
        item = self.itemAt(pos)
        obj_item = self._object_item_from_graphics_item(item) if item else None
        if obj_item and obj_item.data(0):
            if not obj_item.isSelected():
                scene.clearSelection()
                obj_item.setSelected(True)
        selected_ids = [itm.data(0) for itm in scene.selectedItems() if itm.data(0)]

        menu = QMenu(self)
        insert_menu = menu.addMenu("Insert")
        can_insert = scene.edit_mode
        has_rows = bool(scene.layout.rows)
        insert_actions = {}
        insert_actions[insert_menu.addAction("Activity")] = "box"
        insert_actions[insert_menu.addAction("Activity Text")] = "text"
        insert_actions[insert_menu.addAction("Milestone")] = "milestone"
        insert_actions[insert_menu.addAction("Deadline")] = "deadline"
        insert_actions[insert_menu.addAction("Circle")] = "circle"
        insert_actions[insert_menu.addAction("Arrow")] = "arrow"
        insert_actions[insert_menu.addAction("Connector Arrow")] = "connector"
        insert_actions[insert_menu.addAction("Text Box")] = "textbox"
        for action, kind in insert_actions.items():
            if kind == "textbox":
                action.setEnabled(can_insert and scene.show_textboxes)
            elif kind in ("deadline", "connector"):
                action.setEnabled(can_insert)
            else:
                action.setEnabled(can_insert and has_rows)

        convert_actions: dict[QAction, str] = {}
        convert_obj_id = obj_item.data(0) if obj_item and obj_item.data(0) else None
        convert_row_id = None
        if convert_obj_id:
            obj = scene.model.objects.get(convert_obj_id)
            if obj and obj.kind in {kind for kind, _ in CONVERTIBLE_OBJECT_TYPES}:
                convert_menu = menu.addMenu("Convert To")
                convert_row_id = scene.layout.row_at_y(self.mapToScene(pos).y())
                for kind, label in CONVERTIBLE_OBJECT_TYPES:
                    if kind == obj.kind:
                        continue
                    action = convert_menu.addAction(label)
                    convert_actions[action] = kind
                    if kind == "deadline":
                        action.setEnabled(scene.edit_mode)
                    else:
                        action.setEnabled(scene.edit_mode and has_rows)

        menu.addSeparator()
        bring_front = menu.addAction("Bring to Front")
        bring_forward = menu.addAction("Bring Forward")
        send_backward = menu.addAction("Send Backward")
        send_back = menu.addAction("Send to Back")

        enabled = bool(selected_ids)
        bring_front.setEnabled(enabled)
        bring_forward.setEnabled(enabled)
        send_backward.setEnabled(enabled)
        send_back.setEnabled(enabled)

        action = menu.exec(self.viewport().mapToGlobal(pos))
        if not action:
            return
        if action in insert_actions:
            self._create_from_context(insert_actions[action], pos)
            return
        if action in convert_actions:
            self._convert_object_kind(convert_obj_id, convert_actions[action], convert_row_id)
            return
        if not selected_ids:
            return
        if action == bring_front:
            self.controller.reorder_objects(selected_ids, "front")
        elif action == bring_forward:
            self.controller.reorder_objects(selected_ids, "forward")
        elif action == send_backward:
            self.controller.reorder_objects(selected_ids, "backward")
        elif action == send_back:
            self.controller.reorder_objects(selected_ids, "back")

    def _show_row_context_menu(self, pos, row_id: str, kind: str) -> None:
        scene = self.scene()
        if scene is None:
            return
        scene.clearSelection()
        self.set_selected_row(row_id)

        menu = QMenu(self)
        if kind == "topic":
            add_deliverable_action = menu.addAction("Add Deliverable")
            focus_action = None
            menu.addSeparator()
            rename_action = menu.addAction("Rename Topic")
            remove_action = menu.addAction("Remove Topic")
        else:
            add_deliverable_action = None
            is_focused = scene.focused_row_id == row_id
            focus_label = "Unfocus" if is_focused else "Focus on this"
            focus_action = menu.addAction(focus_label)
            menu.addSeparator()
            rename_action = menu.addAction("Rename Deliverable")
            remove_action = menu.addAction("Remove Deliverable")

        can_edit = scene.edit_mode
        if add_deliverable_action is not None:
            add_deliverable_action.setEnabled(can_edit)
        rename_action.setEnabled(can_edit)
        remove_action.setEnabled(can_edit)

        action = menu.exec(self.viewport().mapToGlobal(pos))
        if not action:
            return
        if add_deliverable_action is not None and action == add_deliverable_action:
            self._add_deliverable(row_id)
            return
        if focus_action is not None and action == focus_action:
            if scene.focused_row_id == row_id:
                self.set_focused_row(None)
            else:
                self.set_focused_row(row_id)
            return
        if action == rename_action:
            if kind == "topic":
                self._rename_topic(row_id)
            else:
                self._rename_deliverable(row_id)
        elif action == remove_action:
            if kind == "topic":
                self._remove_topic(row_id)
            else:
                self._remove_deliverable(row_id)

    def _create_from_context(self, kind: str, pos) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        if kind == "textbox" and not scene.show_textboxes:
            return
        if kind not in ("textbox", "deadline", "connector") and not scene.layout.rows:
            QMessageBox.information(self, "Add Object", "Add a topic or deliverable first.")
            return
        if kind in ("connector", "arrow"):
            self.activate_create_tool(kind)
            return
        scene_pos = self.mapToScene(pos)
        self._create_tool = kind
        self._create_start = scene_pos
        layout = scene.layout
        if kind in ("textbox", "deadline"):
            self._create_start_row = None
        else:
            self._create_start_row = layout.row_at_y(scene_pos.y())
        self._create_start_week = layout.week_from_x(scene_pos.x(), scene.snap_weeks)
        self._finish_create(pos)

    def _convert_object_kind(self, obj_id: str | None, new_kind: str, row_id: str | None) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        if not obj_id:
            return
        obj = scene.model.objects.get(obj_id)
        if obj is None or obj.kind == new_kind:
            return
        allowed = {kind for kind, _ in CONVERTIBLE_OBJECT_TYPES}
        if obj.kind not in allowed or new_kind not in allowed:
            return
        changes: dict[str, object] = {"kind": new_kind}
        if new_kind == "deadline":
            changes["row_id"] = CANVAS_ROW_ID
        else:
            target_row_id = obj.row_id
            if target_row_id == CANVAS_ROW_ID or target_row_id not in scene.layout.row_map:
                target_row_id = row_id
            if target_row_id is None and scene.layout.rows:
                target_row_id = scene.layout.rows[0].row_id
            if target_row_id is None:
                return
            changes["row_id"] = target_row_id
        label_map = {kind: label for kind, label in CONVERTIBLE_OBJECT_TYPES}
        description = f"Convert to {label_map.get(new_kind, new_kind)}"
        self.controller.update_object(obj_id, changes, description)

    def _add_deliverable(self, topic_id: str) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        topic = scene.model.get_topic(topic_id)
        if topic is None:
            return
        name, ok = QInputDialog.getText(
            self, "Add Deliverable", "Deliverable Name"
        )
        if not ok:
            return
        name = name.strip()
        if not name:
            return
        self.controller.add_deliverable(topic_id, name)

    def _rename_topic(self, topic_id: str) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        topic = scene.model.get_topic(topic_id)
        if topic is None:
            return
        name, ok = QInputDialog.getText(
            self, "Rename Topic", "Topic Name", text=topic.name
        )
        if not ok:
            return
        name = name.strip()
        if not name or name == topic.name:
            return
        new_topic = type(topic)(
            id=topic.id,
            name=name,
            color=topic.color,
            collapsed=topic.collapsed,
            deliverables=topic.deliverables,
        )
        self.controller.update_topic(new_topic)

    def _rename_deliverable(self, deliverable_id: str) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        found = scene.model.find_deliverable(deliverable_id)
        if found is None:
            return
        _topic, _index, deliverable = found
        name, ok = QInputDialog.getText(
            self, "Rename Deliverable", "Deliverable Name", text=deliverable.name
        )
        if not ok:
            return
        name = name.strip()
        if not name or name == deliverable.name:
            return
        new_deliverable = type(deliverable)(id=deliverable.id, name=name)
        self.controller.update_deliverable(new_deliverable)

    def _remove_deliverable(self, deliverable_id: str) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        found = scene.model.find_deliverable(deliverable_id)
        if found is None:
            return
        topic, _index, deliverable = found
        affected_objects = self._objects_for_rows({deliverable_id})
        detail = ""
        if affected_objects:
            detail = f"\n\nThis will remove {len(affected_objects)} related object(s) on the canvas."
        message = f"Remove deliverable '{deliverable.name}' from '{topic.name}'?{detail}"
        if QMessageBox.question(
            self,
            "Confirm Remove Deliverable",
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) == QMessageBox.StandardButton.Yes:
            self.controller.remove_deliverable(deliverable_id)

    def _remove_topic(self, topic_id: str) -> None:
        scene = self.scene()
        if scene is None or not scene.edit_mode:
            return
        topic = scene.model.get_topic(topic_id)
        if topic is None:
            return
        deliverable_count = len(topic.deliverables)
        row_ids = {topic.id, *[d.id for d in topic.deliverables]}
        affected_objects = self._objects_for_rows(row_ids)
        detail_parts = []
        if deliverable_count:
            detail_parts.append(f"{deliverable_count} deliverable(s)")
        if affected_objects:
            detail_parts.append(f"{len(affected_objects)} object(s)")
        detail = ""
        if detail_parts:
            detail = "\n\nThis will remove " + " and ".join(detail_parts) + "."
        message = f"Remove topic '{topic.name}'?{detail}"
        if QMessageBox.question(
            self,
            "Confirm Remove Topic",
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        ) == QMessageBox.StandardButton.Yes:
            self.controller.remove_topic(topic_id)

    def _objects_for_rows(self, row_ids: set[str]) -> list:
        scene = self.scene()
        if scene is None:
            return []
        row_objects = [
            obj
            for obj in scene.model.objects.values()
            if obj.kind not in ("link", "connector")
            and (obj.row_id in row_ids or (obj.target_row_id in row_ids))
        ]
        row_object_ids = {obj.id for obj in row_objects}
        link_objects = [
            obj
            for obj in scene.model.objects.values()
            if obj.kind == "link"
            and (
                (obj.link_source_id in row_object_ids)
                or (obj.link_target_id in row_object_ids)
            )
        ]
        connector_objects = [
            obj
            for obj in scene.model.objects.values()
            if obj.kind == "connector"
            and (
                (obj.connector_source_id in row_object_ids)
                or (obj.connector_target_id in row_object_ids)
            )
        ]
        return row_objects + link_objects + connector_objects
