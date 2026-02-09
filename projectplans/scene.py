from __future__ import annotations

from dataclasses import replace

from PyQt6.QtCore import QRectF, pyqtSignal
from PyQt6.QtWidgets import QGraphicsScene

from .constants import (
    EXPAND_YEAR_BUFFER,
    EXPAND_YEAR_RANGE,
    HEADER_MONTH_HEIGHT,
    HEADER_QUARTER_HEIGHT,
    HEADER_WEEK_HEIGHT,
    HEADER_YEAR_HEIGHT,
    INITIAL_YEAR_RANGE,
    TEXTBOX_MIN_HEIGHT,
    TEXTBOX_MIN_WIDTH,
    WEEKS_PER_YEAR,
)
from .items import (
    ArrowItem,
    BoxItem,
    CircleItem,
    ConnectorItem,
    DeadlineItem,
    GridItem,
    LinkItem,
    MilestoneItem,
    TextItem,
    TextboxItem,
)
from .layout import Layout


class CanvasScene(QGraphicsScene):
    label_width_changed = pyqtSignal(float)

    def __init__(self, model, controller) -> None:
        super().__init__()
        self.model = model
        self.controller = controller
        self.layout = Layout(model)
        if self.controller and hasattr(self.controller, "set_layout"):
            self.controller.set_layout(self.layout)
        self.items_by_id = {}
        self._object_cache = {}
        self.selected_row_id: str | None = None
        self.focused_row_id: str | None = None
        self.snap_weeks = True
        self.snap_rows = True
        self.edit_mode = True
        self.show_current_week = True
        self.show_missing_scope = False
        self.show_textboxes = True
        self.connector_cache = {}
        self.header_year_height = HEADER_YEAR_HEIGHT
        self.header_quarter_height = HEADER_QUARTER_HEIGHT
        self.header_month_height = HEADER_MONTH_HEIGHT
        self.header_week_height = HEADER_WEEK_HEIGHT
        self.min_week = self.layout.origin_week - (WEEKS_PER_YEAR * INITIAL_YEAR_RANGE)
        self.max_week = self.layout.origin_week + (WEEKS_PER_YEAR * INITIAL_YEAR_RANGE)
        self.week_expand = WEEKS_PER_YEAR * EXPAND_YEAR_RANGE
        self.week_expand_buffer = WEEKS_PER_YEAR * EXPAND_YEAR_BUFFER

        self.grid_item = GridItem(self)
        self.addItem(self.grid_item)

        self.model.rows_changed.connect(self.rebuild_layout)
        self.model.objects_changed.connect(self.refresh_items)
        self.model.metadata_changed.connect(self.update_headers)

        self.rebuild_layout()
        self.refresh_items()

    def update_headers(self) -> None:
        self.grid_item.update()

    def update_risk_badges(self) -> None:
        for obj_id, item in self.items_by_id.items():
            if not isinstance(item, BoxItem):
                continue
            obj = self.model.objects.get(obj_id)
            if obj is None:
                continue
            rect = item.rect()
            item._update_risk_badge(
                obj,
                rect.width(),
                rect.height(),
                show_missing_scope=self.show_missing_scope,
            )

    def rebuild_layout(self) -> None:
        self.layout.rebuild(self.model)
        if self.selected_row_id and self.selected_row_id not in self.layout.row_map:
            self.selected_row_id = None
        if self.focused_row_id and self.focused_row_id not in self.layout.row_map:
            self.focused_row_id = None
        self._update_scene_rect()
        self.grid_item.update()
        self.refresh_items(force_sync=True)

    def set_label_width(self, width: float) -> None:
        width = float(width)
        if width <= 0:
            return
        if abs(self.layout.label_width - width) < 0.5:
            return
        self.layout.label_width = width
        self._update_scene_rect()
        self.grid_item.update()
        self.refresh_items(force_sync=True)
        self.label_width_changed.emit(width)

    def shift_textboxes(self, delta: float) -> None:
        if abs(delta) < 0.01:
            return
        for obj_id, obj in list(self.model.objects.items()):
            if obj.kind != "textbox":
                continue
            current_x = obj.x if obj.x is not None else 0.0
            new_x = current_x + delta
            if abs(new_x - current_x) < 0.01:
                continue
            self.model.objects[obj_id] = replace(obj, x=new_x)

    def _update_scene_rect(self) -> None:
        height = self.layout.header_height + self.layout.total_height
        min_y = 0.0
        max_y = height
        min_x = self.layout.week_left_x(self.min_week)
        max_x = self.layout.week_left_x(self.max_week) + self.layout.week_width
        for obj in self.model.objects.values():
            if not self.show_textboxes or obj.kind != "textbox" or obj.y is None:
                continue
            box_width = obj.width if obj.width is not None else TEXTBOX_MIN_WIDTH
            box_height = obj.height if obj.height is not None else TEXTBOX_MIN_HEIGHT
            if obj.x is not None:
                min_x = min(min_x, obj.x)
                max_x = max(max_x, obj.x + box_width)
            min_y = min(min_y, obj.y)
            max_y = max(max_y, obj.y + box_height)
        grid_rect = QRectF(
            self.layout.week_left_x(self.min_week),
            0,
            (self.layout.week_left_x(self.max_week) + self.layout.week_width)
            - self.layout.week_left_x(self.min_week),
            height,
        )
        self.setSceneRect(min_x, min_y, max_x - min_x, max_y - min_y)
        self.grid_item.set_rect(grid_rect)
        self.grid_item.update()

    def ensure_week_range(self, week: int) -> None:
        changed = False
        while week < self.min_week + self.week_expand_buffer:
            self.min_week -= self.week_expand
            changed = True
        while week > self.max_week - self.week_expand_buffer:
            self.max_week += self.week_expand
            changed = True
        if changed:
            self._update_scene_rect()

    def _ensure_range_for_objects(self) -> None:
        for obj in self.model.objects.values():
            if obj.kind in ("textbox", "link", "connector"):
                continue
            self.ensure_week_range(obj.start_week)
            self.ensure_week_range(obj.end_week)
            if obj.target_week is not None:
                self.ensure_week_range(obj.target_week)

    def refresh_items(self, force_sync: bool = False) -> None:
        self._ensure_range_for_objects()
        self._update_scene_rect()
        selected_ids = {item.data(0) for item in self.selectedItems() if item.data(0)}
        existing_items = self.items_by_id
        existing_cache = self._object_cache
        new_items: dict[str, object] = {}
        new_cache: dict[str, object] = {}

        def should_display(obj) -> bool:
            if obj.kind == "textbox":
                return self.show_textboxes
            if obj.kind == "link":
                return self.show_textboxes
            if obj.kind in ("deadline", "connector"):
                return True
            if obj.kind == "arrow":
                attached = bool(obj.connector_source_id and obj.connector_target_id)
                if attached:
                    return True
                if obj.row_id not in self.layout.row_map:
                    return False
                target_row = obj.target_row_id or obj.row_id
                return target_row in self.layout.row_map
            return obj.row_id in self.layout.row_map

        def item_matches_kind(item, kind: str) -> bool:
            if kind == "textbox":
                return isinstance(item, TextboxItem)
            if kind == "deadline":
                return isinstance(item, DeadlineItem)
            if kind == "link":
                return isinstance(item, LinkItem)
            if kind == "connector":
                return isinstance(item, ConnectorItem)
            if kind == "box":
                return isinstance(item, BoxItem)
            if kind == "milestone":
                return isinstance(item, MilestoneItem)
            if kind == "circle":
                return isinstance(item, CircleItem)
            if kind == "arrow":
                return isinstance(item, ArrowItem)
            if kind == "text":
                return isinstance(item, TextItem)
            return False

        def create_item(obj):
            if obj.kind == "textbox":
                return TextboxItem(obj.id)
            if obj.kind == "deadline":
                return DeadlineItem(obj.id)
            if obj.kind == "link":
                return LinkItem(obj.id, self)
            if obj.kind == "connector":
                return ConnectorItem(obj.id, self)
            if obj.kind == "box":
                return BoxItem(obj.id)
            if obj.kind == "milestone":
                return MilestoneItem(obj.id)
            if obj.kind == "circle":
                return CircleItem(obj.id)
            if obj.kind == "arrow":
                return ArrowItem(obj.id, self)
            if obj.kind == "text":
                return TextItem(obj.id)
            return None

        for obj in self.model.objects.values():
            if not should_display(obj):
                continue
            item = existing_items.get(obj.id)
            if item is not None and not item_matches_kind(item, obj.kind):
                self.removeItem(item)
                item = None
            if item is None:
                item = create_item(obj)
                if item is None:
                    continue
                self.addItem(item)
            item.setData(0, obj.id)
            if not self.edit_mode:
                item.setFlag(item.GraphicsItemFlag.ItemIsMovable, False)
            cached_obj = existing_cache.get(obj.id)
            needs_sync = force_sync or cached_obj is None or cached_obj is not obj
            if obj.kind in ("link", "connector", "arrow"):
                needs_sync = True
            if needs_sync:
                if isinstance(item, BoxItem):
                    item.sync_from_model(obj, self.layout, self.show_missing_scope)
                else:
                    item.sync_from_model(obj, self.layout)
            new_items[obj.id] = item
            new_cache[obj.id] = obj
            if obj.id in selected_ids and not item.isSelected():
                item.setSelected(True)

        for obj_id, item in existing_items.items():
            if obj_id not in new_items:
                self.removeItem(item)
        self.items_by_id = new_items
        self._object_cache = new_cache
        connector_ids = {
            obj.id for obj in self.model.objects.values() if obj.kind == "connector"
        }
        for connector_id in list(self.connector_cache.keys()):
            if connector_id not in connector_ids:
                del self.connector_cache[connector_id]

    def set_edit_mode(self, enabled: bool) -> None:
        self.edit_mode = enabled
        for item in self.items_by_id.values():
            obj_id = item.data(0)
            obj = self.model.objects.get(obj_id) if obj_id else None
            if obj and obj.kind in ("link", "connector"):
                item.setFlag(item.GraphicsItemFlag.ItemIsMovable, False)
            elif (
                obj
                and obj.kind == "arrow"
                and obj.connector_source_id
                and obj.connector_target_id
            ):
                item.setFlag(item.GraphicsItemFlag.ItemIsMovable, False)
            else:
                item.setFlag(item.GraphicsItemFlag.ItemIsMovable, enabled)

    def commit_object_change(self, obj_id: str, changes: dict, description: str) -> None:
        if not self.controller:
            return
        self.controller.update_object(obj_id, changes, description)

    def toggle_topic(self, topic_id: str) -> None:
        if not self.controller:
            return
        self.controller.toggle_topic_collapsed(topic_id)

    def set_selected_row(self, row_id: str | None) -> None:
        if row_id == self.selected_row_id:
            return
        self.selected_row_id = row_id
        self.update()

    def set_focused_row(self, row_id: str | None) -> None:
        if row_id and row_id not in self.layout.row_map:
            row_id = None
        if row_id == self.focused_row_id:
            return
        self.focused_row_id = row_id
        self.grid_item.update()
        self.update()
