from __future__ import annotations

from dataclasses import replace

from PyQt6.QtGui import QUndoStack

from .commands import (
    AddDeliverableCommand,
    AddObjectCommand,
    AddTopicCommand,
    MoveDeliverableCommand,
    MoveDeliverableAcrossTopicsCommand,
    RemoveDeliverableCommand,
    RemoveObjectCommand,
    RemoveTopicCommand,
    ToggleTopicCollapseCommand,
    UpdateClassificationCommand,
    UpdateObjectCommand,
    UpdateDeliverableCommand,
    UpdateTopicCommand,
)
from .constants import (
    CANVAS_ROW_ID,
    DEFAULT_SIZE,
    DEADLINE_DEFAULT_COLOR,
    CONNECTOR_DEFAULT_COLOR,
    LINK_LINE_COLOR,
    TEXTBOX_DEFAULT_COLOR,
    TEXTBOX_DEFAULT_OPACITY,
    TEXTBOX_MIN_HEIGHT,
    TEXTBOX_MIN_WIDTH,
    TEXTBOX_NEW_OPACITY,
)
from .model import (
    CanvasObject,
    Deliverable,
    ProjectModel,
    Topic,
    new_id,
    DEFAULT_TOPIC_COLORS,
    normalize_arrow_direction,
)


class ProjectController:
    def __init__(self, model: ProjectModel, undo_stack: QUndoStack) -> None:
        self.model = model
        self.undo_stack = undo_stack
        self.layout = None
        self.connector_default_size = 1
        self.connector_default_color = CONNECTOR_DEFAULT_COLOR
        self.arrow_default_size = 1
        self.arrow_default_color = CONNECTOR_DEFAULT_COLOR

    def set_layout(self, layout) -> None:
        self.layout = layout

    @staticmethod
    def _size_scale(size: int) -> float:
        scale = 0.5 + (0.1 * size)
        if scale < 0.6:
            return 0.6
        if scale > 1.0:
            return 1.0
        return scale

    def _object_anchor_point(self, obj: CanvasObject) -> tuple[float, float] | None:
        layout = self.layout
        if layout is None:
            return None
        if obj.kind == "textbox":
            width = obj.width if obj.width is not None else TEXTBOX_MIN_WIDTH
            height = obj.height if obj.height is not None else TEXTBOX_MIN_HEIGHT
            x = obj.x if obj.x is not None else 0.0
            y = obj.y if obj.y is not None else 0.0
            return (x + (width / 2.0), y + (height / 2.0))
        if obj.kind == "milestone":
            if obj.row_id not in layout.row_map:
                return None
            row_height = layout.row_height(obj.row_id)
            size = min(layout.week_width, row_height) * self._size_scale(obj.size)
            center_x = layout.week_left_x(obj.start_week)
            center_y = layout.row_center_y(obj.row_id)
            return (center_x, center_y)
        if obj.kind == "circle":
            if obj.row_id not in layout.row_map:
                return None
            row_height = layout.row_height(obj.row_id)
            size = min(layout.week_width, row_height) * self._size_scale(obj.size)
            center_x = layout.week_center_x(obj.start_week)
            center_y = layout.row_center_y(obj.row_id)
            return (center_x, center_y)
        if obj.kind == "deadline":
            center_x = layout.week_center_x(obj.start_week)
            center_y = layout.header_height + (layout.total_height / 2.0)
            return (center_x, center_y)
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
            return ((start_x + end_x) / 2.0, (start_y + end_y) / 2.0)
        if obj.row_id not in layout.row_map:
            return None
        row_height = layout.row_height(obj.row_id)
        height = row_height * self._size_scale(obj.size)
        width = max(1, obj.end_week - obj.start_week + 1) * layout.week_width
        x = layout.week_left_x(obj.start_week)
        y = layout.row_top_y(obj.row_id) + ((row_height - height) / 2.0)
        return (x + (width / 2.0), y + (height / 2.0))

    @staticmethod
    def _textbox_anchor_point(
        obj: CanvasObject, side: str | None, offset: float | None
    ) -> tuple[float, float]:
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
            return (x, y + (height * offset_value))
        if side == "top":
            return (x + (width * offset_value), y)
        if side == "bottom":
            return (x + (width * offset_value), y + height)
        return (x + width, y + (height * offset_value))

    @staticmethod
    def _textbox_pos_for_anchor(
        anchor_x: float,
        anchor_y: float,
        width: float,
        height: float,
        side: str | None,
        offset: float | None,
    ) -> tuple[float, float]:
        offset_value = float(offset or 0.5)
        if offset_value < 0.0:
            offset_value = 0.0
        if offset_value > 1.0:
            offset_value = 1.0
        if side == "left":
            return (anchor_x, anchor_y - (height * offset_value))
        if side == "top":
            return (anchor_x - (width * offset_value), anchor_y)
        if side == "bottom":
            return (anchor_x - (width * offset_value), anchor_y - height)
        return (anchor_x - width, anchor_y - (height * offset_value))

    def _links_from_source(self, source_id: str) -> list[CanvasObject]:
        return [
            obj
            for obj in self.model.objects.values()
            if obj.kind == "link" and obj.link_source_id == source_id
        ]

    def _links_for_target(self, target_id: str, skip_sources: set[str]) -> list[CanvasObject]:
        return [
            obj
            for obj in self.model.objects.values()
            if obj.kind == "link"
            and obj.link_target_id == target_id
            and obj.link_source_id
            and obj.link_source_id not in skip_sources
        ]

    def _clamp_week(self, week: int) -> int:
        return int(week)

    def _next_z_index(self) -> int:
        if not self.model.objects:
            return 0
        return max(obj.z_index for obj in self.model.objects.values()) + 1

    def _normalize_object(self, obj: CanvasObject) -> CanvasObject:
        start_week = self._clamp_week(obj.start_week)
        end_week = self._clamp_week(obj.end_week)
        size = max(1, min(5, obj.size))
        z_index = int(obj.z_index)
        opacity = float(obj.opacity)
        text_html = obj.text_html
        if text_html == "":
            text_html = None
        notes_html = obj.notes_html
        if notes_html == "":
            notes_html = None
        scope_html = obj.scope_html
        if scope_html == "":
            scope_html = None
        risks_html = obj.risks_html
        if risks_html == "":
            risks_html = None
        arrow_direction = normalize_arrow_direction(getattr(obj, "arrow_direction", "none"))
        if obj.kind != "box":
            arrow_direction = "none"
        obj = replace(
            obj,
            text_html=text_html,
            notes_html=notes_html,
            scope_html=scope_html,
            risks_html=risks_html,
            arrow_direction=arrow_direction,
        )
        if opacity < 0.0:
            opacity = 0.0
        if opacity > 1.0:
            opacity = 1.0
        arrow_head_start = bool(getattr(obj, "arrow_head_start", False))
        arrow_head_end = bool(getattr(obj, "arrow_head_end", True))
        if obj.kind in ("arrow", "connector") and not (arrow_head_start or arrow_head_end):
            arrow_head_end = True

        if obj.kind in ("milestone", "circle", "deadline"):
            end_week = start_week
        elif obj.kind == "arrow":
            target_week = obj.target_week if obj.target_week is not None else end_week
            target_week = self._clamp_week(target_week)
            end_week = target_week
            obj = replace(
                obj,
                target_week=target_week,
                arrow_head_start=arrow_head_start,
                arrow_head_end=arrow_head_end,
            )
        elif obj.kind == "link":
            link_offset_x = float(obj.link_offset_x or 0.0)
            link_offset_y = float(obj.link_offset_y or 0.0)
            link_source_offset = obj.link_source_offset
            if link_source_offset is not None:
                link_source_offset = float(link_source_offset)
                link_source_offset = max(0.0, min(1.0, link_source_offset))
            return replace(
                obj,
                start_week=start_week,
                end_week=end_week,
                size=size,
                z_index=z_index,
                opacity=opacity,
                link_offset_x=link_offset_x,
                link_offset_y=link_offset_y,
                link_source_offset=link_source_offset,
            )
        elif obj.kind == "connector":
            source_offset = obj.connector_source_offset
            if source_offset is not None:
                source_offset = float(source_offset)
                source_offset = max(0.0, min(1.0, source_offset))
            target_offset = obj.connector_target_offset
            if target_offset is not None:
                target_offset = float(target_offset)
                target_offset = max(0.0, min(1.0, target_offset))
            return replace(
                obj,
                start_week=start_week,
                end_week=end_week,
                size=size,
                z_index=z_index,
                opacity=opacity,
                connector_source_offset=source_offset,
                connector_target_offset=target_offset,
                arrow_head_start=arrow_head_start,
                arrow_head_end=arrow_head_end,
            )
        elif obj.kind == "textbox":
            x = float(obj.x or 0.0)
            y = float(obj.y or 0.0)
            width = float(obj.width or TEXTBOX_MIN_WIDTH)
            height = float(obj.height or TEXTBOX_MIN_HEIGHT)
            if width < TEXTBOX_MIN_WIDTH:
                width = TEXTBOX_MIN_WIDTH
            if height < TEXTBOX_MIN_HEIGHT:
                height = TEXTBOX_MIN_HEIGHT
            return replace(
                obj,
                start_week=start_week,
                end_week=end_week,
                size=size,
                z_index=z_index,
                x=x,
                y=y,
                width=width,
                height=height,
                opacity=opacity,
                arrow_head_start=arrow_head_start,
                arrow_head_end=arrow_head_end,
            )
        else:
            if end_week < start_week:
                end_week = start_week

        return replace(
            obj,
            start_week=start_week,
            end_week=end_week,
            size=size,
            z_index=z_index,
            opacity=opacity,
            arrow_head_start=arrow_head_start,
            arrow_head_end=arrow_head_end,
        )

    def _duplicate_offset(self, obj: CanvasObject) -> int:
        end_week = obj.target_week if obj.target_week is not None else obj.end_week
        span = abs(end_week - obj.start_week) + 1
        return max(1, span)

    def add_object(self, obj: CanvasObject, description: str = "Add Object") -> None:
        if obj.z_index == 0 and self.model.objects:
            obj = replace(obj, z_index=self._next_z_index())
        obj = self._normalize_object(obj)
        self.undo_stack.push(AddObjectCommand(self.model, obj, description))

    def remove_object(self, obj_id: str) -> None:
        obj = self.model.objects.get(obj_id)
        if obj is None:
            return
        related_links = [
            link
            for link in self.model.objects.values()
            if link.kind == "link"
            and (link.link_source_id == obj_id or link.link_target_id == obj_id)
        ]
        related_connectors = [
            connector
            for connector in self.model.objects.values()
            if connector.kind == "connector"
            and (
                connector.connector_source_id == obj_id
                or connector.connector_target_id == obj_id
            )
        ]
        if related_links or related_connectors:
            self.undo_stack.beginMacro("Remove Object")
            for link in related_links:
                self.undo_stack.push(RemoveObjectCommand(self.model, link))
            for connector in related_connectors:
                self.undo_stack.push(RemoveObjectCommand(self.model, connector))
            self.undo_stack.push(RemoveObjectCommand(self.model, obj))
            self.undo_stack.endMacro()
            return
        self.undo_stack.push(RemoveObjectCommand(self.model, obj))

    def update_object(
        self,
        obj_id: str,
        changes: dict,
        description: str,
        *,
        skip_anchor_sources: set[str] | None = None,
        defer_link_updates: bool = False,
    ) -> None:
        obj = self.model.objects.get(obj_id)
        if obj is None:
            return
        new_obj = replace(obj, **changes)
        new_obj = self._normalize_object(new_obj)
        if new_obj == obj:
            return
        if obj.kind == "connector" and "size" in changes:
            self.connector_default_size = new_obj.size
        if obj.kind == "arrow" and "size" in changes:
            self.arrow_default_size = new_obj.size
        if obj.kind in ("connector", "arrow") and "color" in changes:
            self.connector_default_color = new_obj.color
        if obj.kind == "arrow" and "color" in changes:
            self.arrow_default_color = new_obj.color
        skip_sources = skip_anchor_sources or set()
        target_links = self._links_for_target(obj_id, skip_sources)
        source_links = self._links_from_source(obj_id) if obj.kind == "textbox" else []

        updates: list[tuple[CanvasObject, CanvasObject, str]] = [(obj, new_obj, description)]

        if target_links:
            new_target_point = self._object_anchor_point(new_obj)
            if new_target_point is None:
                target_links = []
            else:
                for link in target_links:
                    source_id = link.link_source_id
                    if not source_id:
                        continue
                    source = self.model.objects.get(source_id)
                    if source is None or source.kind != "textbox":
                        continue
                    anchor_x = new_target_point[0] + float(link.link_offset_x or 0.0)
                    anchor_y = new_target_point[1] + float(link.link_offset_y or 0.0)
                    width = source.width if source.width is not None else TEXTBOX_MIN_WIDTH
                    height = source.height if source.height is not None else TEXTBOX_MIN_HEIGHT
                    new_x, new_y = self._textbox_pos_for_anchor(
                        anchor_x,
                        anchor_y,
                        width,
                        height,
                        link.link_source_side,
                        link.link_source_offset,
                    )
                    changes_for_source = {"x": new_x, "y": new_y}
                    if self.layout is not None:
                        start_week = self.layout.week_from_x(new_x, snap=False)
                        end_week = self.layout.week_from_x(new_x + width, snap=False)
                        changes_for_source["start_week"] = start_week
                        changes_for_source["end_week"] = end_week
                    new_source = self._normalize_object(replace(source, **changes_for_source))
                    if new_source != source:
                        updates.append((source, new_source, "Anchor Textbox"))

        link_updates: list[tuple[CanvasObject, CanvasObject]] = []
        if source_links and not defer_link_updates:
            for link in source_links:
                target_id = link.link_target_id
                if not target_id:
                    continue
                target = self.model.objects.get(target_id)
                if target is None:
                    continue
                target_point = self._object_anchor_point(target)
                if target_point is None:
                    continue
                source_anchor = self._textbox_anchor_point(
                    new_obj, link.link_source_side, link.link_source_offset
                )
                new_offset_x = source_anchor[0] - target_point[0]
                new_offset_y = source_anchor[1] - target_point[1]
                old_offset_x = float(link.link_offset_x or 0.0)
                old_offset_y = float(link.link_offset_y or 0.0)
                if abs(new_offset_x - old_offset_x) > 0.01 or abs(new_offset_y - old_offset_y) > 0.01:
                    new_link = replace(
                        link,
                        link_offset_x=new_offset_x,
                        link_offset_y=new_offset_y,
                    )
                    link_updates.append((link, self._normalize_object(new_link)))

        if len(updates) > 1 or link_updates:
            self.undo_stack.beginMacro(description)
            for old_obj, updated_obj, label in updates:
                self.undo_stack.push(UpdateObjectCommand(self.model, old_obj, updated_obj, label))
            for old_link, new_link in link_updates:
                self.undo_stack.push(UpdateObjectCommand(self.model, old_link, new_link, "Update Anchor"))
            self.undo_stack.endMacro()
            return
        self.undo_stack.push(UpdateObjectCommand(self.model, obj, new_obj, description))

    def refresh_anchor_offsets(self, source_ids: set[str], description: str = "Update Anchors") -> None:
        if not source_ids:
            return
        if self.layout is None:
            return
        link_updates: list[tuple[CanvasObject, CanvasObject]] = []
        for source_id in source_ids:
            source = self.model.objects.get(source_id)
            if source is None or source.kind != "textbox":
                continue
            for link in self._links_from_source(source_id):
                target_id = link.link_target_id
                if not target_id:
                    continue
                target = self.model.objects.get(target_id)
                if target is None:
                    continue
                target_point = self._object_anchor_point(target)
                if target_point is None:
                    continue
                source_anchor = self._textbox_anchor_point(
                    source, link.link_source_side, link.link_source_offset
                )
                new_offset_x = source_anchor[0] - target_point[0]
                new_offset_y = source_anchor[1] - target_point[1]
                old_offset_x = float(link.link_offset_x or 0.0)
                old_offset_y = float(link.link_offset_y or 0.0)
                if abs(new_offset_x - old_offset_x) > 0.01 or abs(new_offset_y - old_offset_y) > 0.01:
                    new_link = replace(
                        link,
                        link_offset_x=new_offset_x,
                        link_offset_y=new_offset_y,
                    )
                    link_updates.append((link, self._normalize_object(new_link)))
        if not link_updates:
            return
        self.undo_stack.beginMacro(description)
        for old_link, new_link in link_updates:
            self.undo_stack.push(UpdateObjectCommand(self.model, old_link, new_link, "Update Anchor"))
        self.undo_stack.endMacro()

    def duplicate_object(self, obj_id: str) -> CanvasObject | None:
        obj = self.model.objects.get(obj_id)
        if obj is None:
            return None
        if obj.kind == "textbox":
            width = obj.width or TEXTBOX_MIN_WIDTH
            padding = max(10.0, width * 0.1)
            new_x = (obj.x or 0.0) + width + padding
            cloned = replace(
                obj,
                id=new_id(),
                x=new_x,
                z_index=self._next_z_index(),
            )
            self.add_object(cloned, "Duplicate Object")
            return cloned
        delta = self._duplicate_offset(obj)
        new_start = self._clamp_week(obj.start_week + delta)
        new_end = self._clamp_week(obj.end_week + delta)
        target_week = obj.target_week + delta if obj.target_week is not None else None
        if target_week is not None:
            target_week = self._clamp_week(target_week)
        arrow_mid_week = obj.arrow_mid_week + delta if obj.arrow_mid_week is not None else None
        if arrow_mid_week is not None:
            arrow_mid_week = self._clamp_week(arrow_mid_week)
        cloned = replace(
            obj,
            id=new_id(),
            start_week=new_start,
            end_week=new_end,
            target_week=target_week,
            arrow_mid_week=arrow_mid_week,
            z_index=self._next_z_index(),
        )
        self.add_object(cloned, "Duplicate Object")
        return cloned

    def add_topic(self, name: str, color: str | None = None) -> Topic:
        if color is None:
            color = DEFAULT_TOPIC_COLORS[len(self.model.topics) % len(DEFAULT_TOPIC_COLORS)]
        topic = Topic(id=new_id(), name=name, color=color)
        self.undo_stack.push(AddTopicCommand(self.model, topic))
        return topic

    def update_classification(self, text: str, size: int | None = None) -> None:
        new_text, new_size = self.model.normalize_classification(text, size)
        old_text = self.model.classification
        old_size = self.model.classification_size
        if new_text == old_text and new_size == old_size:
            return
        self.undo_stack.push(
            UpdateClassificationCommand(self.model, old_text, old_size, new_text, new_size)
        )

    def update_topic(self, topic: Topic) -> None:
        current = self.model.get_topic(topic.id)
        if current is None or current == topic:
            return
        self.undo_stack.push(UpdateTopicCommand(self.model, current, topic))

    def update_deliverable(self, deliverable: Deliverable) -> None:
        found = self.model.find_deliverable(deliverable.id)
        if found is None:
            return
        _topic, _index, current = found
        if current == deliverable:
            return
        self.undo_stack.push(UpdateDeliverableCommand(self.model, current, deliverable))

    def add_deliverable(self, topic_id: str, name: str) -> Deliverable | None:
        topic = self.model.get_topic(topic_id)
        if topic is None:
            return None
        deliverable = Deliverable(id=new_id(), name=name)
        self.undo_stack.push(AddDeliverableCommand(self.model, topic_id, deliverable))
        return deliverable

    def toggle_topic_collapsed(self, topic_id: str) -> None:
        topic = self.model.get_topic(topic_id)
        if topic is None:
            return
        self.undo_stack.push(ToggleTopicCollapseCommand(self.model, topic_id, topic.collapsed))

    def move_deliverable(self, deliverable_id: str, direction: int) -> bool:
        found = self.model.find_deliverable(deliverable_id)
        if found is None:
            return False
        topic, index, _deliverable = found
        if not topic.deliverables:
            return False
        topic_index = None
        for idx, entry in enumerate(self.model.topics):
            if entry.id == topic.id:
                topic_index = idx
                break
        if topic_index is None:
            return False

        if direction < 0:
            if index > 0:
                new_index = index - 1
                self.undo_stack.push(
                    MoveDeliverableCommand(self.model, deliverable_id, index, new_index)
                )
                return True
            if topic_index > 0:
                target_topic = self.model.topics[topic_index - 1]
                target_index = len(target_topic.deliverables)
                self.undo_stack.push(
                    MoveDeliverableAcrossTopicsCommand(
                        self.model,
                        deliverable_id,
                        topic.id,
                        index,
                        target_topic.id,
                        target_index,
                    )
                )
                return True
            return False

        if direction > 0:
            if index < len(topic.deliverables) - 1:
                new_index = index + 1
                self.undo_stack.push(
                    MoveDeliverableCommand(self.model, deliverable_id, index, new_index)
                )
                return True
            if topic_index < len(self.model.topics) - 1:
                target_topic = self.model.topics[topic_index + 1]
                target_index = 0
                self.undo_stack.push(
                    MoveDeliverableAcrossTopicsCommand(
                        self.model,
                        deliverable_id,
                        topic.id,
                        index,
                        target_topic.id,
                        target_index,
                    )
                )
                return True
            return False

        return False

    def remove_deliverable(self, deliverable_id: str) -> bool:
        found = self.model.find_deliverable(deliverable_id)
        if found is None:
            return False
        topic, index, deliverable = found
        removed_objects = self._objects_for_rows({deliverable_id})
        self.undo_stack.push(
            RemoveDeliverableCommand(self.model, topic.id, deliverable, index, removed_objects)
        )
        return True

    def remove_topic(self, topic_id: str) -> bool:
        topic_index = None
        topic = None
        for index, entry in enumerate(self.model.topics):
            if entry.id == topic_id:
                topic_index = index
                topic = entry
                break
        if topic is None or topic_index is None:
            return False
        row_ids = {topic.id, *[d.id for d in topic.deliverables]}
        removed_objects = self._objects_for_rows(row_ids)
        self.undo_stack.push(RemoveTopicCommand(self.model, topic, topic_index, removed_objects))
        return True

    def _objects_for_rows(self, row_ids: set[str]) -> list[CanvasObject]:
        row_objects = [
            obj
            for obj in self.model.objects.values()
            if obj.kind != "link" and (obj.row_id in row_ids or (obj.target_row_id in row_ids))
        ]
        row_object_ids = {obj.id for obj in row_objects}
        link_objects = [
            obj
            for obj in self.model.objects.values()
            if obj.kind == "link"
            and (
                (obj.link_source_id in row_object_ids)
                or (obj.link_target_id in row_object_ids)
            )
        ]
        connector_objects = [
            obj
            for obj in self.model.objects.values()
            if obj.kind == "connector"
            and (
                (obj.connector_source_id in row_object_ids)
                or (obj.connector_target_id in row_object_ids)
            )
        ]
        return row_objects + link_objects + connector_objects

    def make_default_object(
        self,
        kind: str,
        row_id: str,
        start_week: int,
        end_week: int | None = None,
        color: str | None = None,
    ) -> CanvasObject:
        if end_week is None:
            end_week = start_week
        if color is None:
            if kind == "deadline":
                color = DEADLINE_DEFAULT_COLOR
            else:
                topic = self.model.topic_for_row(row_id)
                color = topic.color if topic else DEFAULT_TOPIC_COLORS[0]
        size = DEFAULT_SIZE
        if kind in ("box", "text"):
            size = 5
        obj = CanvasObject(
            id=new_id(),
            kind=kind,
            row_id=row_id,
            start_week=start_week,
            end_week=end_week,
            text="",
            color=color,
            size=size,
            z_index=self._next_z_index(),
        )
        if kind == "arrow":
            obj = replace(obj, target_row_id=row_id, target_week=end_week)
        return self._normalize_object(obj)

    def make_textbox(self, x: float, y: float, width: float, height: float) -> CanvasObject:
        obj = CanvasObject(
            id=new_id(),
            kind="textbox",
            row_id=CANVAS_ROW_ID,
            start_week=0,
            end_week=0,
            text="",
            text_align="left",
            color=TEXTBOX_DEFAULT_COLOR,
            size=DEFAULT_SIZE,
            z_index=self._next_z_index(),
            x=x,
            y=y,
            width=width,
            height=height,
            opacity=TEXTBOX_NEW_OPACITY,
        )
        return self._normalize_object(obj)

    def add_anchor_link(
        self, source_id: str, target_id: str, side: str | None, offset: float | None
    ) -> None:
        source = self.model.objects.get(source_id)
        target = self.model.objects.get(target_id)
        if (
            source is None
            or target is None
            or source.kind != "textbox"
            or target.kind in ("link", "textbox", "connector")
            or source_id == target_id
        ):
            return
        if self.layout is None:
            return
        side_value = side if side in ("left", "right", "top", "bottom") else "right"
        offset_value = float(offset or 0.5)
        if offset_value < 0.0:
            offset_value = 0.0
        if offset_value > 1.0:
            offset_value = 1.0
        target_point = self._object_anchor_point(target)
        if target_point is None:
            return
        source_anchor = self._textbox_anchor_point(source, side_value, offset_value)
        link_offset_x = source_anchor[0] - target_point[0]
        link_offset_y = source_anchor[1] - target_point[1]
        z_index = min(source.z_index, target.z_index) - 1
        if z_index == 0:
            z_index = -1
        link_obj = CanvasObject(
            id=new_id(),
            kind="link",
            row_id=CANVAS_ROW_ID,
            start_week=0,
            end_week=0,
            text="",
            color=LINK_LINE_COLOR,
            size=DEFAULT_SIZE,
            z_index=z_index,
            link_source_id=source_id,
            link_target_id=target_id,
            link_source_side=side_value,
            link_source_offset=offset_value,
            link_offset_x=link_offset_x,
            link_offset_y=link_offset_y,
        )
        existing_links = self._links_from_source(source_id)
        self.undo_stack.beginMacro("Anchor Textbox")
        for link in existing_links:
            self.undo_stack.push(RemoveObjectCommand(self.model, link))
        self.add_object(link_obj, "Add Anchor")
        self.undo_stack.endMacro()

    def add_connector_arrow(
        self,
        source_id: str,
        target_id: str,
        source_side: str | None,
        source_offset: float | None,
        target_side: str | None,
        target_offset: float | None,
    ) -> None:
        source = self.model.objects.get(source_id)
        target = self.model.objects.get(target_id)
        if (
            source is None
            or target is None
            or source_id == target_id
            or source.kind in ("link", "connector")
            or target.kind in ("link", "connector")
        ):
            return
        source_side_value = (
            source_side if source_side in ("left", "right", "top", "bottom") else "right"
        )
        target_side_value = (
            target_side if target_side in ("left", "right", "top", "bottom") else "left"
        )
        source_offset_value = float(source_offset or 0.5)
        source_offset_value = max(0.0, min(1.0, source_offset_value))
        target_offset_value = float(target_offset or 0.5)
        target_offset_value = max(0.0, min(1.0, target_offset_value))
        z_index = min(source.z_index, target.z_index) - 1
        if z_index == 0:
            z_index = -1
        connector_obj = CanvasObject(
            id=new_id(),
            kind="connector",
            row_id=CANVAS_ROW_ID,
            start_week=0,
            end_week=0,
            text="",
            color=self.connector_default_color,
            size=self.connector_default_size,
            z_index=z_index,
            connector_source_id=source_id,
            connector_target_id=target_id,
            connector_source_side=source_side_value,
            connector_source_offset=source_offset_value,
            connector_target_side=target_side_value,
            connector_target_offset=target_offset_value,
        )
        self.add_object(connector_obj, "Add Connector Arrow")

    def reorder_objects(self, obj_ids: list[str], action: str) -> None:
        selected = [obj_id for obj_id in obj_ids if obj_id in self.model.objects]
        if not selected:
            return
        ordered_ids = self._ordered_object_ids()
        selected_set = set(selected)

        if action == "front":
            new_order = [obj_id for obj_id in ordered_ids if obj_id not in selected_set] + [
                obj_id for obj_id in ordered_ids if obj_id in selected_set
            ]
            self._apply_z_order(new_order, "Bring to Front")
            return
        if action == "back":
            new_order = [obj_id for obj_id in ordered_ids if obj_id in selected_set] + [
                obj_id for obj_id in ordered_ids if obj_id not in selected_set
            ]
            self._apply_z_order(new_order, "Send to Back")
            return
        if action == "forward":
            order = list(ordered_ids)
            for index in range(len(order) - 2, -1, -1):
                if order[index] in selected_set and order[index + 1] not in selected_set:
                    order[index], order[index + 1] = order[index + 1], order[index]
            self._apply_z_order(order, "Bring Forward")
            return
        if action == "backward":
            order = list(ordered_ids)
            for index in range(1, len(order)):
                if order[index] in selected_set and order[index - 1] not in selected_set:
                    order[index], order[index - 1] = order[index - 1], order[index]
            self._apply_z_order(order, "Send Backward")

    def _ordered_object_ids(self) -> list[str]:
        if not self.model.objects:
            return []
        order_index = {obj_id: index for index, obj_id in enumerate(self.model.objects.keys())}
        ordered = sorted(
            self.model.objects.values(),
            key=lambda obj: (obj.z_index, order_index.get(obj.id, 0)),
        )
        return [obj.id for obj in ordered]

    def _apply_z_order(self, ordered_ids: list[str], description: str) -> None:
        if not ordered_ids:
            return
        self.undo_stack.beginMacro(description)
        for index, obj_id in enumerate(ordered_ids):
            obj = self.model.objects.get(obj_id)
            if obj is None or obj.z_index == index:
                continue
            self.update_object(obj_id, {"z_index": index}, "Reorder")
        self.undo_stack.endMacro()
