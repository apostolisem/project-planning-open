from __future__ import annotations

from dataclasses import dataclass
from bisect import bisect_left, bisect_right
from datetime import date, timedelta
import math

from .constants import (
    DEFAULT_WEEK_WIDTH,
    DELIVERABLE_ROW_HEIGHT,
    HEADER_MONTH_HEIGHT,
    HEADER_QUARTER_HEIGHT,
    HEADER_WEEK_HEIGHT,
    HEADER_YEAR_HEIGHT,
    LABEL_WIDTH,
    TOPIC_ROW_HEIGHT,
)
from .model import ProjectModel


@dataclass
class RowLayout:
    row_id: str
    name: str
    kind: str
    topic_id: str
    y: float
    height: float
    indent: int


class Layout:
    def __init__(self, model: ProjectModel, week_width: int = DEFAULT_WEEK_WIDTH) -> None:
        self.week_width = week_width
        self.label_width = LABEL_WIDTH
        self.header_height = (
            HEADER_YEAR_HEIGHT
            + HEADER_QUARTER_HEIGHT
            + HEADER_MONTH_HEIGHT
            + HEADER_WEEK_HEIGHT
        )
        self.origin_week = 1
        self.rows: list[RowLayout] = []
        self.row_map: dict[str, RowLayout] = {}
        self._row_start_positions: list[float] = []
        self._row_end_positions: list[float] = []
        self._row_index_map: dict[str, int] = {}
        self.total_height = 0.0
        self.rebuild(model)

    def rebuild(self, model: ProjectModel) -> None:
        self.rows = []
        self.row_map = {}
        self._row_start_positions = []
        self._row_end_positions = []
        self._row_index_map = {}
        y = 0.0
        for topic in model.topics:
            topic_row = RowLayout(
                row_id=topic.id,
                name=topic.name,
                kind="topic",
                topic_id=topic.id,
                y=y,
                height=TOPIC_ROW_HEIGHT,
                indent=0,
            )
            self.rows.append(topic_row)
            self.row_map[topic.id] = topic_row
            y += TOPIC_ROW_HEIGHT
            if topic.collapsed:
                continue
            for deliverable in topic.deliverables:
                deliverable_row = RowLayout(
                    row_id=deliverable.id,
                    name=deliverable.name,
                    kind="deliverable",
                    topic_id=topic.id,
                    y=y,
                    height=DELIVERABLE_ROW_HEIGHT,
                    indent=1,
                )
                self.rows.append(deliverable_row)
                self.row_map[deliverable.id] = deliverable_row
                y += DELIVERABLE_ROW_HEIGHT
        self.total_height = y
        for index, row in enumerate(self.rows):
            self._row_start_positions.append(row.y)
            self._row_end_positions.append(row.y + row.height)
            self._row_index_map[row.row_id] = index

    def week_left_x(self, week: int) -> float:
        return self.label_width + (week - self.origin_week) * self.week_width

    def week_center_x(self, week: int) -> float:
        return self.week_left_x(week) + (self.week_width / 2.0)

    def row_top_y(self, row_id: str) -> float:
        row = self.row_map[row_id]
        return self.header_height + row.y

    def row_center_y(self, row_id: str) -> float:
        row = self.row_map[row_id]
        return self.header_height + row.y + (row.height / 2.0)

    def row_height(self, row_id: str) -> float:
        return self.row_map[row_id].height

    def row_at_y(self, scene_y: float) -> str | None:
        y_rel = scene_y - self.header_height
        if y_rel < 0:
            return None
        if not self.rows:
            return None
        index = bisect_right(self._row_end_positions, y_rel)
        if index >= len(self.rows):
            return None
        row = self.rows[index]
        if row.y <= y_rel < row.y + row.height:
            return row.row_id
        return None

    def row_index(self, row_id: str) -> int | None:
        return self._row_index_map.get(row_id)

    def row_index_range(self, y_min: float, y_max: float) -> tuple[int, int]:
        if not self.rows:
            return 0, 0
        start = bisect_left(self._row_end_positions, y_min)
        end = bisect_right(self._row_end_positions, y_max)
        return start, end

    def adjacent_row(self, row_id: str, direction: int) -> str | None:
        index = self.row_index(row_id)
        if index is None:
            return None
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.rows):
            return None
        return self.rows[new_index].row_id

    def week_from_x(self, scene_x: float, snap: bool = True) -> int:
        x_rel = scene_x - self.label_width
        ratio = x_rel / self.week_width
        if snap:
            if ratio >= 0:
                week_offset = int(ratio + 0.5)
            else:
                week_offset = int(ratio - 0.5)
        else:
            week_offset = int(math.floor(ratio))
        return week_offset + self.origin_week

    def week_from_center_x(self, scene_x: float, snap: bool = True) -> int:
        return self.week_from_x(scene_x - (self.week_width / 2.0), snap)

    def week_index_to_year_week(self, base_year: int, week_index: int) -> tuple[int, int]:
        week_date = self.week_index_to_date(base_year, week_index)
        iso = week_date.isocalendar()
        return iso.year, iso.week

    def week_index_to_date(self, base_year: int, week_index: int) -> date:
        base_date = self.base_week_start(base_year)
        return base_date + timedelta(weeks=week_index - self.origin_week)

    def week_index_for_iso_year(self, base_year: int, year: int) -> int:
        base_date = self.base_week_start(base_year)
        year_start = self.base_week_start(year)
        delta_weeks = (year_start - base_date).days // 7
        return self.origin_week + delta_weeks

    @staticmethod
    def base_week_start(year: int) -> date:
        return date.fromisocalendar(year, 1, 1)

    @staticmethod
    def weeks_in_year(year: int) -> int:
        return date(year, 12, 28).isocalendar().week

    @staticmethod
    def quarter_for_week(week_in_year: int) -> int:
        return min(4, ((week_in_year - 1) // 13) + 1)
