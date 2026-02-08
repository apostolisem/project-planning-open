from __future__ import annotations

from PyQt6.QtCore import QSignalBlocker, pyqtSignal
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QWidget,
    QColorDialog,
    QTextEdit,
)

from .constants import TEXT_SIZE_MAX, TEXT_SIZE_MIN, TEXT_SIZE_STEP, WEEK_INDEX_MAX, WEEK_INDEX_MIN
from .text_shortcuts import apply_text_action, extract_text_payload, text_shortcut_action


class _MetadataTextEdit(QTextEdit):
    commit_requested = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self._base_font = self.font()
        self.setAcceptRichText(True)
        self.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)

    def keyPressEvent(self, event) -> None:
        action = text_shortcut_action(event)
        if action:
            cursor = self.textCursor()
            if apply_text_action(
                cursor,
                action,
                self._base_font,
                min_size=TEXT_SIZE_MIN,
                max_size=TEXT_SIZE_MAX,
                step=TEXT_SIZE_STEP,
            ):
                self.setTextCursor(cursor)
            event.accept()
            return
        super().keyPressEvent(event)

    def insertFromMimeData(self, source) -> None:
        self.insertPlainText(source.text())

    def extract_payload(self) -> tuple[str, str | None]:
        return extract_text_payload(self.document(), self._base_font)

    def focusOutEvent(self, event) -> None:
        self.commit_requested.emit()
        super().focusOutEvent(event)


class _WeekSpinBox(QSpinBox):
    def __init__(self) -> None:
        super().__init__()
        self._layout = None
        self._model = None

    def set_context(self, layout, model) -> None:
        self._layout = layout
        self._model = model
        self.setValue(self.value())

    def textFromValue(self, value: int) -> str:
        if not self._layout or not self._model:
            return super().textFromValue(value)
        year, week = self._layout.week_index_to_year_week(self._model.year, value)
        return f"{year % 100:02d}{week:02d}"

    def valueFromText(self, text: str) -> int:
        cleaned = text.strip()
        if self._layout and self._model:
            digits = "".join(ch for ch in cleaned if ch.isdigit())
            if len(digits) == 4 and digits.isdigit():
                year_two = int(digits[:2])
                week = int(digits[2:])
                base_century = (self._model.year // 100) * 100
                year = base_century + year_two
                if abs(year - self._model.year) > 50:
                    year += 100 if year < self._model.year else -100
                week = max(1, min(self._layout.weeks_in_year(year), week))
                base_week = self._layout.week_index_for_iso_year(self._model.year, year)
                return base_week + (week - 1)
        try:
            return int(cleaned)
        except ValueError:
            return self.value()


class InspectorPanel(QWidget):
    def __init__(self, controller) -> None:
        super().__init__()
        self.controller = controller
        self._current_obj_id = None
        self._row_ids = []
        self._suppress_metadata_refresh = {"scope": None, "risks": None, "notes": None}

        self.type_label = QLabel("-")
        self.text_input = QLineEdit()
        self.start_week = _WeekSpinBox()
        self.duration_weeks = QSpinBox()
        self.row_combo = QComboBox()
        self.target_week = QSpinBox()
        self.target_row_combo = QComboBox()
        self.size_spin = QSpinBox()
        self.arrowheads_combo = QComboBox()
        self.arrow_direction_combo = QComboBox()
        self.reverse_direction_button = QPushButton("Reverse")
        self.align_combo = QComboBox()
        self.color_button = QPushButton("Color")
        self.opacity_spin = QSpinBox()
        self.scope_edit = _MetadataTextEdit()
        self.risks_edit = _MetadataTextEdit()
        self.notes_edit = _MetadataTextEdit()
        self.risks_help = QLabel(
            "One per line. Optional ;probability;impact (h/m/l). "
            "Example: development of helper script;m;l"
        )
        self.risks_help.setWordWrap(True)
        help_font = self.risks_help.font()
        if help_font.pointSize() > 0:
            help_font.setPointSize(max(8, help_font.pointSize() - 2))
        self.risks_help.setFont(help_font)

        self.start_week.setRange(WEEK_INDEX_MIN, WEEK_INDEX_MAX)
        self.duration_weeks.setRange(1, WEEK_INDEX_MAX - WEEK_INDEX_MIN + 1)
        self.target_week.setRange(WEEK_INDEX_MIN, WEEK_INDEX_MAX)
        self.size_spin.setRange(1, 5)
        self.opacity_spin.setRange(0, 100)
        self.opacity_spin.setSuffix("%")

        layout = QFormLayout()
        layout.addRow("Type", self.type_label)
        layout.addRow("Text", self.text_input)
        layout.addRow("Start Week", self.start_week)
        layout.addRow("Duration (weeks)", self.duration_weeks)
        layout.addRow("Row", self.row_combo)
        layout.addRow("Target Week", self.target_week)
        layout.addRow("Target Row", self.target_row_combo)
        layout.addRow("Size", self.size_spin)
        layout.addRow("Arrowheads", self.arrowheads_combo)
        layout.addRow("Arrow Direction", self.arrow_direction_combo)
        layout.addRow("Direction", self.reverse_direction_button)
        layout.addRow("Align", self.align_combo)
        layout.addRow("Color", self.color_button)
        layout.addRow("Opacity", self.opacity_spin)
        layout.addRow("Scope", self.scope_edit)
        layout.addRow("Risks", self.risks_edit)
        layout.addRow("", self.risks_help)
        layout.addRow("Notes", self.notes_edit)
        self.setLayout(layout)

        self.text_input.editingFinished.connect(self._apply_text)
        self.start_week.valueChanged.connect(self._apply_start_week)
        self.duration_weeks.valueChanged.connect(self._apply_duration)
        self.row_combo.currentIndexChanged.connect(self._apply_row)
        self.target_week.valueChanged.connect(self._apply_target_week)
        self.target_row_combo.currentIndexChanged.connect(self._apply_target_row)
        self.size_spin.valueChanged.connect(self._apply_size)
        self.arrowheads_combo.currentIndexChanged.connect(self._apply_arrowheads)
        self.arrow_direction_combo.currentIndexChanged.connect(self._apply_arrow_direction)
        self.reverse_direction_button.clicked.connect(self._reverse_direction)
        self.align_combo.currentIndexChanged.connect(self._apply_alignment)
        self.color_button.clicked.connect(self._pick_color)
        self.opacity_spin.valueChanged.connect(self._apply_opacity)
        self.scope_edit.commit_requested.connect(self._apply_scope)
        self.risks_edit.commit_requested.connect(self._apply_risks)
        self.notes_edit.commit_requested.connect(self._apply_notes)

        self.set_enabled(False)
        self.align_combo.addItem("Left", "left")
        self.align_combo.addItem("Center", "center")
        self.align_combo.addItem("Right", "right")
        self.arrowheads_combo.addItem("End", "end")
        self.arrowheads_combo.addItem("Start", "start")
        self.arrowheads_combo.addItem("Both", "both")
        self.arrow_direction_combo.addItem("None", "none")
        self.arrow_direction_combo.addItem("Left", "left")
        self.arrow_direction_combo.addItem("Right", "right")

    def set_enabled(self, enabled: bool) -> None:
        self.text_input.setEnabled(enabled)
        self.start_week.setEnabled(enabled)
        self.duration_weeks.setEnabled(enabled)
        self.row_combo.setEnabled(enabled)
        self.target_week.setEnabled(enabled)
        self.target_row_combo.setEnabled(enabled)
        self.size_spin.setEnabled(enabled)
        self.arrowheads_combo.setEnabled(enabled)
        self.arrow_direction_combo.setEnabled(enabled)
        self.reverse_direction_button.setEnabled(enabled)
        self.align_combo.setEnabled(enabled)
        self.color_button.setEnabled(enabled)
        self.opacity_spin.setEnabled(enabled)
        self.scope_edit.setEnabled(enabled)
        self.risks_edit.setEnabled(enabled)
        self.notes_edit.setEnabled(enabled)

    def refresh_rows(self, layout, model) -> None:
        self.start_week.set_context(layout, model)
        rows = layout.rows
        self._row_ids = [row.row_id for row in rows]
        current_obj = model.objects.get(self._current_obj_id) if self._current_obj_id else None
        row_value = current_obj.row_id if current_obj else self.row_combo.currentData()
        target_value = (
            (current_obj.target_row_id or current_obj.row_id)
            if current_obj
            else self.target_row_combo.currentData()
        )
        with QSignalBlocker(self.row_combo), QSignalBlocker(self.target_row_combo):
            self.row_combo.clear()
            self.target_row_combo.clear()
            for row in rows:
                label = row.name
                if row.indent:
                    label = "  " * row.indent + label
                self.row_combo.addItem(label, row.row_id)
                self.target_row_combo.addItem(label, row.row_id)
            if row_value:
                self._set_combo_value(self.row_combo, row_value)
            if target_value:
                self._set_combo_value(self.target_row_combo, target_value)

    def set_selected_object(self, obj) -> None:
        previous_obj_id = self._current_obj_id
        self._current_obj_id = obj.id if obj else None
        if obj is None or (previous_obj_id and previous_obj_id != obj.id):
            self._clear_metadata_suppression()
        if obj is None:
            self.type_label.setText("-")
            self.text_input.setText("")
            self.scope_edit.setPlainText("")
            self.risks_edit.setPlainText("")
            self.notes_edit.setPlainText("")
            self.set_enabled(False)
            return

        self.set_enabled(True)
        self.type_label.setText(obj.kind)

        with QSignalBlocker(self.text_input):
            self.text_input.setText(obj.text)
        with QSignalBlocker(self.start_week):
            self.start_week.setValue(obj.start_week)
        duration = max(1, obj.end_week - obj.start_week + 1)
        self._sync_duration_widget(obj.start_week, duration)
        with QSignalBlocker(self.row_combo):
            self._set_combo_value(self.row_combo, obj.row_id)
        with QSignalBlocker(self.target_week):
            self.target_week.setValue(obj.target_week or obj.end_week)
        with QSignalBlocker(self.target_row_combo):
            self._set_combo_value(self.target_row_combo, obj.target_row_id or obj.row_id)
        with QSignalBlocker(self.size_spin):
            self.size_spin.setValue(obj.size)
        with QSignalBlocker(self.arrowheads_combo):
            self._set_combo_value(self.arrowheads_combo, self._arrowheads_value(obj))
        with QSignalBlocker(self.arrow_direction_combo):
            self._set_combo_value(
                self.arrow_direction_combo,
                self._arrow_direction_value(obj),
            )
        with QSignalBlocker(self.align_combo):
            self._set_combo_value(self.align_combo, obj.text_align)
        with QSignalBlocker(self.opacity_spin):
            self.opacity_spin.setValue(int(round((obj.opacity or 0.0) * 100)))
        with QSignalBlocker(self.risks_edit):
            if self._should_refresh_metadata("risks", obj.id):
                if obj.risks_html:
                    self.risks_edit.setHtml(obj.risks_html)
                else:
                    self.risks_edit.setPlainText(obj.risks or "")
                self.risks_edit.document().setDefaultFont(self.risks_edit.font())
        with QSignalBlocker(self.scope_edit):
            if self._should_refresh_metadata("scope", obj.id):
                if obj.scope_html:
                    self.scope_edit.setHtml(obj.scope_html)
                else:
                    self.scope_edit.setPlainText(obj.scope or "")
                self.scope_edit.document().setDefaultFont(self.scope_edit.font())
        with QSignalBlocker(self.notes_edit):
            if self._should_refresh_metadata("notes", obj.id):
                if obj.notes_html:
                    self.notes_edit.setHtml(obj.notes_html)
                else:
                    self.notes_edit.setPlainText(obj.notes or "")
                self.notes_edit.document().setDefaultFont(self.notes_edit.font())
        self._set_color_button(obj.color)

        self._toggle_fields_for_kind(obj.kind)

    def _toggle_fields_for_kind(self, kind: str) -> None:
        is_arrow = kind == "arrow"
        is_milestone = kind == "milestone"
        is_circle = kind == "circle"
        is_deadline = kind == "deadline"
        is_textbox = kind == "textbox"
        is_link = kind == "link"
        is_connector = kind == "connector"
        is_box = kind == "box"
        self._set_field_visible(self.text_input, not is_link and not is_connector)
        self._set_field_visible(
            self.duration_weeks,
            not is_link
            and not is_connector
            and not is_arrow
            and not is_milestone
            and not is_circle
            and not is_deadline
            and not is_textbox,
        )
        self._set_field_visible(self.target_week, False)
        self._set_field_visible(self.target_row_combo, False)
        self._set_field_visible(
            self.row_combo,
            not is_link
            and not is_connector
            and not is_textbox
            and not is_deadline
            and not is_arrow,
        )
        self._set_field_visible(
            self.start_week,
            not is_link and not is_connector and not is_textbox and not is_arrow,
        )
        self._set_field_visible(self.size_spin, not is_link and not is_textbox)
        self._set_field_visible(self.arrowheads_combo, is_arrow or is_connector)
        self._set_field_visible(self.arrow_direction_combo, is_box)
        self._set_field_visible(self.reverse_direction_button, is_arrow or is_connector)
        self._set_field_visible(self.align_combo, not is_link and not is_connector)
        self._set_field_visible(self.color_button, not is_link)
        self._set_field_visible(self.opacity_spin, not is_link and is_textbox)

    def _set_field_visible(self, widget, visible: bool) -> None:
        widget.setVisible(visible)
        label = self.layout().labelForField(widget)
        if label is not None:
            label.setVisible(visible)

    def _apply_text(self) -> None:
        if not self._current_obj_id:
            return
        self.controller.update_object(
            self._current_obj_id, {"text": self.text_input.text(), "text_html": None}, "Edit Text"
        )

    def _apply_start_week(self) -> None:
        if not self._current_obj_id:
            return
        start_week = self.start_week.value()
        if self.duration_weeks.isVisible():
            duration = self._sync_duration_widget(start_week, self.duration_weeks.value())
            end_week = start_week + duration - 1
            self.controller.update_object(
                self._current_obj_id,
                {"start_week": start_week, "end_week": end_week},
                "Edit Start Week",
            )
            return
        self.controller.update_object(
            self._current_obj_id, {"start_week": start_week}, "Edit Start Week"
        )

    def _apply_duration(self) -> None:
        if not self._current_obj_id:
            return
        start_week = self.start_week.value()
        duration = self._sync_duration_widget(start_week, self.duration_weeks.value())
        end_week = start_week + duration - 1
        self.controller.update_object(
            self._current_obj_id, {"end_week": end_week}, "Edit Duration"
        )

    def _apply_row(self) -> None:
        if not self._current_obj_id:
            return
        row_id = self.row_combo.currentData()
        if row_id:
            self.controller.update_object(
                self._current_obj_id, {"row_id": row_id}, "Edit Row"
            )

    def _apply_target_week(self) -> None:
        if not self._current_obj_id:
            return
        self.controller.update_object(
            self._current_obj_id,
            {"target_week": self.target_week.value(), "end_week": self.target_week.value()},
            "Edit Target Week",
        )

    def _apply_target_row(self) -> None:
        if not self._current_obj_id:
            return
        row_id = self.target_row_combo.currentData()
        if row_id:
            self.controller.update_object(
                self._current_obj_id, {"target_row_id": row_id}, "Edit Target Row"
            )

    def _apply_size(self) -> None:
        if not self._current_obj_id:
            return
        self.controller.update_object(
            self._current_obj_id, {"size": self.size_spin.value()}, "Edit Size"
        )

    def _apply_arrowheads(self) -> None:
        if not self._current_obj_id:
            return
        value = self.arrowheads_combo.currentData()
        if value == "both":
            start = True
            end = True
        elif value == "start":
            start = True
            end = False
        else:
            start = False
            end = True
        self.controller.update_object(
            self._current_obj_id,
            {"arrow_head_start": start, "arrow_head_end": end},
            "Edit Arrowheads",
        )

    def _apply_arrow_direction(self) -> None:
        if not self._current_obj_id:
            return
        value = self.arrow_direction_combo.currentData()
        if value is None:
            return
        self.controller.update_object(
            self._current_obj_id,
            {"arrow_direction": value},
            "Edit Arrow Direction",
        )

    def _reverse_direction(self) -> None:
        if not self._current_obj_id:
            return
        obj = self.controller.model.objects.get(self._current_obj_id)
        if obj is None or obj.kind not in ("arrow", "connector"):
            return
        changes: dict[str, object] = {}
        if obj.kind == "arrow":
            target_week = obj.target_week if obj.target_week is not None else obj.end_week
            new_row_id = obj.target_row_id or obj.row_id
            new_target_row_id = obj.row_id
            if new_row_id == new_target_row_id:
                new_target_row_id = None
            changes["start_week"] = target_week
            changes["end_week"] = obj.start_week
            changes["target_week"] = obj.start_week
            changes["row_id"] = new_row_id
            changes["target_row_id"] = new_target_row_id
        if obj.connector_source_id and obj.connector_target_id:
            changes.update(
                {
                    "connector_source_id": obj.connector_target_id,
                    "connector_target_id": obj.connector_source_id,
                    "connector_source_side": obj.connector_target_side,
                    "connector_target_side": obj.connector_source_side,
                    "connector_source_offset": obj.connector_target_offset,
                    "connector_target_offset": obj.connector_source_offset,
                }
            )
        if changes:
            self.controller.update_object(self._current_obj_id, changes, "Reverse Arrow")

    def _apply_alignment(self) -> None:
        if not self._current_obj_id:
            return
        align = self.align_combo.currentData()
        if align:
            self.controller.update_object(
                self._current_obj_id, {"text_align": align}, "Edit Alignment"
            )

    def _pick_color(self) -> None:
        if not self._current_obj_id:
            return
        color = QColorDialog.getColor(QColor(self.color_button.property("color") or "#4E79A7"), self)
        if not color.isValid():
            return
        hex_color = color.name().upper()
        self._set_color_button(hex_color)
        self.controller.update_object(
            self._current_obj_id, {"color": hex_color}, "Edit Color"
        )

    def _apply_opacity(self) -> None:
        if not self._current_obj_id:
            return
        opacity = self.opacity_spin.value() / 100.0
        self.controller.update_object(
            self._current_obj_id, {"opacity": opacity}, "Edit Opacity"
        )

    def _apply_scope(self) -> None:
        if not self._current_obj_id:
            return
        text, html = self.scope_edit.extract_payload()
        self._mark_metadata_refresh("scope")
        self.controller.update_object(
            self._current_obj_id,
            {"scope": text, "scope_html": html},
            "Edit Scope",
        )

    def _apply_risks(self) -> None:
        if not self._current_obj_id:
            return
        text, html = self.risks_edit.extract_payload()
        self._mark_metadata_refresh("risks")
        self.controller.update_object(
            self._current_obj_id,
            {"risks": text, "risks_html": html},
            "Edit Risks",
        )

    def _apply_notes(self) -> None:
        if not self._current_obj_id:
            return
        text, html = self.notes_edit.extract_payload()
        self._mark_metadata_refresh("notes")
        self.controller.update_object(
            self._current_obj_id,
            {"notes": text, "notes_html": html},
            "Edit Notes",
        )

    def _sync_duration_widget(self, start_week: int, duration: int) -> int:
        max_duration = max(1, WEEK_INDEX_MAX - start_week + 1)
        duration = max(1, min(duration, max_duration))
        with QSignalBlocker(self.duration_weeks):
            self.duration_weeks.setRange(1, max_duration)
            self.duration_weeks.setValue(duration)
        return duration

    def _set_color_button(self, color: str) -> None:
        self.color_button.setProperty("color", color)
        self.color_button.setStyleSheet(f"background-color: {color};")

    @staticmethod
    def _field_has_focus(field: QTextEdit) -> bool:
        return field.hasFocus() or field.viewport().hasFocus()

    def _mark_metadata_refresh(self, field: str) -> None:
        if self._current_obj_id:
            self._suppress_metadata_refresh[field] = self._current_obj_id

    def _should_refresh_metadata(self, field: str, obj_id: str) -> bool:
        if self._suppress_metadata_refresh.get(field) == obj_id:
            self._suppress_metadata_refresh[field] = None
            return False
        return True

    def _clear_metadata_suppression(self) -> None:
        for key in self._suppress_metadata_refresh:
            self._suppress_metadata_refresh[key] = None

    @staticmethod
    def _arrowheads_value(obj) -> str:
        start = bool(getattr(obj, "arrow_head_start", False))
        end = bool(getattr(obj, "arrow_head_end", True))
        if start and end:
            return "both"
        if start:
            return "start"
        return "end"

    @staticmethod
    def _arrow_direction_value(obj) -> str:
        value = str(getattr(obj, "arrow_direction", "none") or "none").lower()
        if value not in ("none", "left", "right"):
            return "none"
        return value

    @staticmethod
    def _set_combo_value(combo: QComboBox, value: str) -> None:
        for index in range(combo.count()):
            if combo.itemData(index) == value:
                combo.setCurrentIndex(index)
                return
