from __future__ import annotations

import csv
import os
import sys
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import QRectF, QSettings, Qt, QTimer, QStandardPaths, QUrl, QSignalBlocker
from PyQt6.QtGui import (
    QAction,
    QColor,
    QDesktopServices,
    QFont,
    QImage,
    QPainter,
    QPageSize,
    QPdfWriter,
    QPen,
    QUndoStack,
)
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QDockWidget,
)

from .constants import (
    CANVAS_ROW_ID,
    CLASSIFICATION_SIZE_DEFAULT,
    CLASSIFICATION_SIZE_MAX,
    CLASSIFICATION_SIZE_MIN,
    DEFAULT_CLASSIFICATION,
    SCHEMA_VERSION,
)
from .controller import ProjectController
from .inspector import InspectorPanel
from .model import ProjectModel
from .persistence import load_project, save_project
from .scene import CanvasScene
from .updater import UpdateManager
from .view import CanvasView

RECENT_FILES_LIMIT = 10
UNNAMED_LABEL = "Unnamed"
AUTO_EXPORT_DEFAULT_ADDITIONAL_QUARTERS = 2
AUTO_EXPORT_VIEW_KEY = "auto_export"
AUTO_EXPORT_VIEW_ENABLED_KEY = "enabled"
AUTO_EXPORT_VIEW_PATH_KEY = "path"
AUTO_EXPORT_VIEW_ADDITIONAL_QUARTERS_KEY = "additional_quarters"


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.settings = QSettings("ProjectPlans", "ProjectPlans")
        self.undo_stack = QUndoStack(self)

        self.model = ProjectModel(year=datetime.now().year)
        self.controller = ProjectController(self.model, self.undo_stack)
        self.scene = CanvasScene(self.model, self.controller)
        self.view = CanvasView(self.scene, self.controller)
        self.setCentralWidget(self.view)

        self.inspector = InspectorPanel(self.controller)
        self.inspector_dock = QDockWidget("Properties", self)
        self.inspector_dock.setObjectName("properties_dock")
        self.inspector_dock.setWidget(self.inspector)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.inspector_dock)

        self.current_path: Path | None = None
        self._nav_mode_before_presentation = False
        self._restore_maximized = False
        self._has_saved_geometry = False
        self._last_window_maximized = False
        self._is_closing = False
        self._window_handle_connected = False
        self._maximize_attempts = 0
        self._restoring_window_state = False
        self._suppress_properties_sync = False
        self._view_dirty = False
        self._suppress_view_dirty = False
        self._auto_export_enabled = False
        self._auto_export_path: Path | None = None
        self._auto_export_additional_quarters = AUTO_EXPORT_DEFAULT_ADDITIONAL_QUARTERS
        self.update_manager = UpdateManager(self, self.settings)

        self._setup_actions()
        self._connect_model_signals()
        self.undo_stack.cleanChanged.connect(self._update_title)
        self._refresh_rows()
        self._update_title()
        self._load_last_file()
        self._restore_window_settings()
        self._restore_view_settings()
        self.update_manager.schedule_auto_check()

    def _setup_actions(self) -> None:
        file_menu = self.menuBar().addMenu("File")
        export_menu = self.menuBar().addMenu("Export")
        edit_menu = self.menuBar().addMenu("Edit")
        view_menu = self.menuBar().addMenu("View")
        options_menu = self.menuBar().addMenu("Options")
        project_menu = self.menuBar().addMenu("Project")
        insert_menu = self.menuBar().addMenu("Insert")
        help_menu = self.menuBar().addMenu("Help")

        new_action = QAction("New", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_project)
        file_menu.addAction(new_action)

        open_action = QAction("Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_project)
        file_menu.addAction(open_action)

        self.recent_menu = file_menu.addMenu("Open Recent")
        self._refresh_recent_menu()

        save_action = QAction("Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_project)
        file_menu.addAction(save_action)

        save_as_action = QAction("Save As", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_project_as)
        file_menu.addAction(save_as_action)

        export_scope_action = QAction("Scope (markdown)", self)
        export_scope_action.triggered.connect(self.export_scope)
        export_menu.addAction(export_scope_action)

        export_risks_action = QAction("Risks (csv)", self)
        export_risks_action.triggered.connect(self.export_risks)
        export_menu.addAction(export_risks_action)

        export_png_action = QAction("Planning (png)", self)
        export_png_action.triggered.connect(self.export_png)
        export_menu.addAction(export_png_action)

        copy_image_clipboard_action = QAction("Copy Image to Clipboard", self)
        copy_image_clipboard_action.triggered.connect(self.copy_image_to_clipboard)
        export_menu.addAction(copy_image_clipboard_action)

        export_pdf_action = QAction("Planning (pdf)", self)
        export_pdf_action.triggered.connect(self.export_pdf)
        export_menu.addAction(export_pdf_action)

        self.auto_export_action = QAction("Planning auto-export", self)
        self.auto_export_action.setCheckable(True)
        self.auto_export_action.toggled.connect(self._toggle_auto_export)
        export_menu.addAction(self.auto_export_action)

        auto_export_preferences_action = QAction("Auto-export Preferences", self)
        auto_export_preferences_action.triggered.connect(self.show_auto_export_preferences)
        options_menu.addAction(auto_export_preferences_action)

        file_menu.addSeparator()
        quit_action = QAction("Quit", self)
        quit_action.setShortcut("Ctrl+Q")
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        check_updates_action = QAction("Check for Updates", self)
        check_updates_action.triggered.connect(self.update_manager.manual_check)
        help_menu.addAction(check_updates_action)

        self.undo_action = self.undo_stack.createUndoAction(self, "Undo")
        self.undo_action.setShortcut("Ctrl+Z")
        self.redo_action = self.undo_stack.createRedoAction(self, "Redo")
        self.redo_action.setShortcut("Ctrl+Y")
        edit_menu.addAction(self.undo_action)
        edit_menu.addAction(self.redo_action)

        self.delete_action = QAction("Delete", self)
        self.delete_action.setShortcut("Delete")
        self.delete_action.triggered.connect(self.delete_selected)
        edit_menu.addAction(self.delete_action)

        self.duplicate_action = QAction("Duplicate", self)
        self.duplicate_action.setShortcut("Ctrl+D")
        self.duplicate_action.triggered.connect(self.duplicate_selected)
        edit_menu.addAction(self.duplicate_action)

        zoom_in_action = QAction("Zoom In", self)
        zoom_in_action.setShortcut("Ctrl++")
        zoom_in_action.triggered.connect(lambda: self.view.zoom_by(1.1))
        view_menu.addAction(zoom_in_action)

        zoom_out_action = QAction("Zoom Out", self)
        zoom_out_action.setShortcut("Ctrl+-")
        zoom_out_action.triggered.connect(lambda: self.view.zoom_by(0.9))
        view_menu.addAction(zoom_out_action)

        reset_zoom_action = QAction("Reset Zoom", self)
        reset_zoom_action.setShortcut("Ctrl+0")
        reset_zoom_action.triggered.connect(lambda: self.view.set_zoom(1.0))
        view_menu.addAction(reset_zoom_action)

        zoom_selection_action = QAction("Zoom To Selection", self)
        zoom_selection_action.triggered.connect(self.view.zoom_to_selection)
        view_menu.addAction(zoom_selection_action)

        fit_action = QAction("Zoom To Fit", self)
        fit_action.triggered.connect(self.view.zoom_to_fit)
        view_menu.addAction(fit_action)

        goto_today_action = QAction("Goto Today", self)
        goto_today_action.triggered.connect(self.goto_today)
        view_menu.addAction(goto_today_action)

        view_menu.addSeparator()
        self.presentation_mode_action = QAction("Presentation Mode", self)
        self.presentation_mode_action.setCheckable(True)
        self.presentation_mode_action.triggered.connect(self._toggle_presentation_mode)
        view_menu.addAction(self.presentation_mode_action)

        self.nav_mode_action = QAction("Navigation Mode", self)
        self.nav_mode_action.setCheckable(True)
        self.nav_mode_action.triggered.connect(self._toggle_navigation_mode)
        view_menu.addAction(self.nav_mode_action)

        self.properties_pane_action = QAction("Properties Pane", self)
        self.properties_pane_action.setCheckable(True)
        self.properties_pane_action.setChecked(True)
        self.properties_pane_action.triggered.connect(self._toggle_properties_pane)
        view_menu.addAction(self.properties_pane_action)

        self.snap_grid_action = QAction("Snap to Grid", self)
        self.snap_grid_action.setCheckable(True)
        self.snap_grid_action.setChecked(True)
        self.snap_grid_action.triggered.connect(self._toggle_snap_grid)
        view_menu.addAction(self.snap_grid_action)

        self.current_week_action = QAction("Current Week Highlight", self)
        self.current_week_action.setCheckable(True)
        self.current_week_action.setChecked(True)
        self.current_week_action.triggered.connect(self._toggle_current_week_line)
        view_menu.addAction(self.current_week_action)

        self.missing_scope_action = QAction("Show Missing Scope", self)
        self.missing_scope_action.setCheckable(True)
        self.missing_scope_action.setChecked(False)
        self.missing_scope_action.triggered.connect(self._toggle_missing_scope)
        view_menu.addAction(self.missing_scope_action)

        self.text_boxes_action = QAction("Text Boxes", self)
        self.text_boxes_action.setCheckable(True)
        self.text_boxes_action.setChecked(True)
        self.text_boxes_action.triggered.connect(self._toggle_text_boxes)
        view_menu.addAction(self.text_boxes_action)

        self.add_topic_action = QAction("Add Topic", self)
        self.add_topic_action.triggered.connect(self.add_topic)

        self.add_deliverable_action = QAction("Add Deliverable", self)
        self.add_deliverable_action.triggered.connect(self.add_deliverable)

        self.edit_topic_action = QAction("Edit Topic(s)", self)
        self.edit_topic_action.triggered.connect(self.edit_topic)

        self.edit_deliverable_action = QAction("Edit Deliverable(s)", self)
        self.edit_deliverable_action.triggered.connect(self.edit_deliverable)

        self.move_deliverable_up_action = QAction("Move Deliverable Up", self)
        self.move_deliverable_up_action.setShortcut("Alt+Shift+Up")
        self.move_deliverable_up_action.triggered.connect(self.move_deliverable_up)

        self.move_deliverable_down_action = QAction("Move Deliverable Down", self)
        self.move_deliverable_down_action.setShortcut("Alt+Shift+Down")
        self.move_deliverable_down_action.triggered.connect(self.move_deliverable_down)

        self.remove_deliverable_action = QAction("Remove Deliverable", self)
        self.remove_deliverable_action.triggered.connect(self.remove_deliverable)

        self.remove_topic_action = QAction("Remove Topic", self)
        self.remove_topic_action.triggered.connect(self.remove_topic)

        project_menu.addAction(self.add_topic_action)
        project_menu.addAction(self.edit_topic_action)
        project_menu.addAction(self.remove_topic_action)
        project_menu.addSeparator()
        project_menu.addAction(self.add_deliverable_action)
        project_menu.addAction(self.edit_deliverable_action)
        project_menu.addAction(self.move_deliverable_up_action)
        project_menu.addAction(self.move_deliverable_down_action)
        project_menu.addAction(self.remove_deliverable_action)
        project_menu.addSeparator()

        self.edit_classification_action = QAction("Edit Classification Tag...", self)
        self.edit_classification_action.triggered.connect(self.edit_classification_tag)
        project_menu.addAction(self.edit_classification_action)

        self.add_box_action = QAction("Activity", self)
        self.add_box_action.setShortcut("B")
        self.add_box_action.triggered.connect(lambda: self.create_object("box"))
        insert_menu.addAction(self.add_box_action)

        self.add_text_action = QAction("Activity Text", self)
        self.add_text_action.setShortcut("T")
        self.add_text_action.triggered.connect(lambda: self.create_object("text"))
        insert_menu.addAction(self.add_text_action)

        self.add_milestone_action = QAction("Milestone", self)
        self.add_milestone_action.setShortcut("M")
        self.add_milestone_action.triggered.connect(lambda: self.create_object("milestone"))
        insert_menu.addAction(self.add_milestone_action)

        self.add_deadline_action = QAction("Deadline", self)
        self.add_deadline_action.setShortcut("D")
        self.add_deadline_action.triggered.connect(lambda: self.create_object("deadline"))
        insert_menu.addAction(self.add_deadline_action)

        self.add_circle_action = QAction("Circle", self)
        self.add_circle_action.setShortcut("C")
        self.add_circle_action.triggered.connect(lambda: self.create_object("circle"))
        insert_menu.addAction(self.add_circle_action)

        self.add_arrow_action = QAction("Arrow", self)
        self.add_arrow_action.setShortcut("A")
        self.add_arrow_action.triggered.connect(lambda: self.create_object("arrow"))
        insert_menu.addAction(self.add_arrow_action)

        self.add_connector_action = QAction("Connector Arrow", self)
        self.add_connector_action.setShortcut("F")
        self.add_connector_action.triggered.connect(lambda: self.create_object("connector"))
        insert_menu.addAction(self.add_connector_action)

        self.add_textbox_action = QAction("Text Box", self)
        self.add_textbox_action.setShortcut("X")
        self.add_textbox_action.triggered.connect(lambda: self.create_object("textbox"))
        insert_menu.addAction(self.add_textbox_action)

        self._edit_actions = [
            self.undo_action,
            self.redo_action,
            self.delete_action,
            self.duplicate_action,
            self.edit_classification_action,
            self.add_topic_action,
            self.add_deliverable_action,
            self.edit_topic_action,
            self.edit_deliverable_action,
            self.move_deliverable_up_action,
            self.move_deliverable_down_action,
            self.remove_deliverable_action,
            self.remove_topic_action,
            self.add_box_action,
            self.add_text_action,
            self.add_milestone_action,
            self.add_deadline_action,
            self.add_circle_action,
            self.add_arrow_action,
            self.add_connector_action,
            self.add_textbox_action,
        ]

        self.inspector_dock.visibilityChanged.connect(self._sync_properties_pane_action)
        self._sync_textbox_insert_action()

    def _recent_file_entries(self) -> list[str]:
        value = self.settings.value("recent_files", [])
        if isinstance(value, str):
            entries = [value]
        elif isinstance(value, list):
            entries = [str(item) for item in value]
        else:
            entries = []
        cleaned: list[str] = []
        seen = set()
        for entry in entries:
            entry = entry.strip()
            if not entry or entry in seen:
                continue
            seen.add(entry)
            cleaned.append(entry)
        return cleaned

    def _write_recent_files(self, entries: list[str]) -> None:
        self.settings.setValue("recent_files", entries)
        self.settings.sync()

    def _refresh_recent_menu(self) -> None:
        self.recent_menu.clear()
        entries = self._recent_file_entries()
        existing_entries = [entry for entry in entries if Path(entry).exists()]
        if existing_entries != entries:
            self._write_recent_files(existing_entries)
        if not existing_entries:
            empty_action = QAction("No Recent Files", self)
            empty_action.setEnabled(False)
            self.recent_menu.addAction(empty_action)
            return
        for entry in existing_entries:
            path = Path(entry)
            label = f"{path.name} ({path.parent})"
            action = QAction(label, self)
            action.setData(entry)
            action.triggered.connect(self._open_recent_action)
            self.recent_menu.addAction(action)
        self.recent_menu.addSeparator()
        clear_action = QAction("Clear Recent", self)
        clear_action.triggered.connect(self._clear_recent_files)
        self.recent_menu.addAction(clear_action)

    def _add_recent_file(self, path: Path) -> None:
        path = Path(path).expanduser()
        try:
            resolved = path.resolve()
        except Exception:
            resolved = path.absolute()
        path_str = str(resolved)
        entries = [entry for entry in self._recent_file_entries() if entry != path_str]
        entries.insert(0, path_str)
        entries = entries[:RECENT_FILES_LIMIT]
        self._write_recent_files(entries)
        self._refresh_recent_menu()

    def _remove_recent_file(self, path: Path) -> None:
        path_str = str(path)
        entries = [entry for entry in self._recent_file_entries() if entry != path_str]
        self._write_recent_files(entries)
        self._refresh_recent_menu()

    def _clear_recent_files(self) -> None:
        self._write_recent_files([])
        self._refresh_recent_menu()

    def _open_recent_action(self) -> None:
        action = self.sender()
        if not isinstance(action, QAction):
            return
        entry = action.data()
        if not entry:
            return
        path = Path(str(entry))
        if not path.exists():
            QMessageBox.warning(self, "Open Recent", "Recent file no longer exists.")
            self._remove_recent_file(path)
            return
        if not self.maybe_save():
            return
        try:
            self._load_from_path(path)
        except Exception as exc:
            QMessageBox.warning(self, "Open Failed", f"Could not open project file.\n{exc}")

    def _disconnect_model_signals(self) -> None:
        try:
            self.scene.selectionChanged.disconnect(self._update_inspector_selection)
        except TypeError:
            pass
        try:
            self.scene.label_width_changed.disconnect(self._on_label_width_changed)
        except TypeError:
            pass
        try:
            self.model.rows_changed.disconnect(self._refresh_rows)
        except TypeError:
            pass
        try:
            self.model.objects_changed.disconnect(self._update_inspector_selection)
        except TypeError:
            pass

    def _connect_model_signals(self) -> None:
        self.scene.selectionChanged.connect(self._update_inspector_selection)
        self.scene.label_width_changed.connect(self._on_label_width_changed)
        self.model.rows_changed.connect(self._refresh_rows)
        self.model.objects_changed.connect(self._update_inspector_selection)

    def _refresh_rows(self) -> None:
        self.inspector.refresh_rows(self.scene.layout, self.model)

    def _update_inspector_selection(self) -> None:
        selected = self.scene.selectedItems()
        obj = None
        for item in selected:
            obj_id = item.data(0)
            if obj_id and obj_id in self.model.objects:
                obj = self.model.objects[obj_id]
                break
        self.inspector.set_selected_object(obj)
        if obj is not None:
            self.view.set_selected_row(obj.row_id)

    def _set_view_dirty(self, dirty: bool) -> None:
        if self._view_dirty == dirty:
            return
        self._view_dirty = dirty
        self._update_title()

    def _on_label_width_changed(self, _width: float) -> None:
        if self._suppress_view_dirty:
            return
        self._set_view_dirty(True)

    def _update_title(self) -> None:
        if self.current_path:
            name = self.current_path.name
        else:
            name = "Untitled"
        if not self.undo_stack.isClean() or self._view_dirty:
            name = f"*{name}"
        self.setWindowTitle(f"ProjectPlans - {name}")

    def _load_last_file(self) -> None:
        path = self.settings.value("last_file", "")
        if path:
            path = Path(path)
            if path.exists():
                try:
                    self._load_from_path(path)
                    return
                except Exception:
                    QMessageBox.warning(self, "Load Failed", "Could not load the last project file.")
        self._update_title()

    def _set_model(self, model: ProjectModel, view_state: dict | None = None) -> None:
        self._disconnect_model_signals()
        self.model = model
        self.controller = ProjectController(self.model, self.undo_stack)
        self.scene = CanvasScene(self.model, self.controller)
        self.view.setScene(self.scene)
        self.view.controller = self.controller
        self.inspector.controller = self.controller

        self._connect_model_signals()
        self._refresh_rows()
        self.inspector.set_selected_object(None)
        if view_state is not None:
            label_width_value = view_state.get("label_width")
            if label_width_value is not None:
                try:
                    label_width = float(label_width_value)
                except (TypeError, ValueError):
                    label_width = None
                if label_width is not None:
                    self._suppress_view_dirty = True
                    try:
                        self.scene.set_label_width(label_width)
                    finally:
                        self._suppress_view_dirty = False
        self._apply_auto_export_settings(view_state)
        if self.presentation_mode_action.isChecked():
            self.scene.set_edit_mode(False)
            self.view.set_navigation_mode(True)
        else:
            enabled = self.nav_mode_action.isChecked()
            self.scene.set_edit_mode(not enabled)
            self.view.set_navigation_mode(enabled)
        self._apply_presentation_mode(self.presentation_mode_action.isChecked())
        self._apply_text_boxes_visibility()
        self.scene.show_missing_scope = self.missing_scope_action.isChecked()
        self.scene.update_risk_badges()

        if view_state:
            zoom = float(view_state.get("zoom", 1.0))
            self.view.set_zoom(zoom)
            self.view.horizontalScrollBar().setValue(int(view_state.get("scroll_x", 0)))
            self.view.verticalScrollBar().setValue(int(view_state.get("scroll_y", 0)))
        else:
            self.view.set_zoom(1.0)
            self.view.center_on_base_year()

    def _toggle_properties_pane(self) -> None:
        self._apply_properties_visibility()

    def _apply_properties_visibility(self) -> None:
        show_properties = self.properties_pane_action.isChecked()
        visible = show_properties and not self.presentation_mode_action.isChecked()
        self.inspector_dock.setVisible(visible)

    def _sync_properties_pane_action(self, visible: bool) -> None:
        if self._suppress_properties_sync:
            return
        if self.presentation_mode_action.isChecked():
            return
        if self.properties_pane_action.isChecked() == visible:
            return
        with QSignalBlocker(self.properties_pane_action):
            self.properties_pane_action.setChecked(visible)

    def _toggle_navigation_mode(self) -> None:
        enabled = self.nav_mode_action.isChecked()
        self.scene.set_edit_mode(not enabled)
        self.view.set_navigation_mode(enabled)
        if enabled:
            self.view.activate_create_tool(None)

    def _toggle_presentation_mode(self) -> None:
        enabled = self.presentation_mode_action.isChecked()
        if enabled:
            self._nav_mode_before_presentation = self.nav_mode_action.isChecked()
            self.nav_mode_action.setChecked(True)
            self._toggle_navigation_mode()
        else:
            self.nav_mode_action.setChecked(self._nav_mode_before_presentation)
            self._toggle_navigation_mode()
        self._apply_presentation_mode(enabled)

    def _apply_presentation_mode(self, enabled: bool) -> None:
        self.nav_mode_action.setEnabled(not enabled)
        self.properties_pane_action.setEnabled(not enabled)
        for action in self._edit_actions:
            action.setEnabled(not enabled)
        self.inspector.setEnabled(not enabled)
        self._apply_properties_visibility()
        if enabled:
            self.view.activate_create_tool(None)
        self._sync_textbox_insert_action()

    def _sync_textbox_insert_action(self) -> None:
        enabled = self.scene.show_textboxes and not self.presentation_mode_action.isChecked()
        self.add_textbox_action.setEnabled(enabled)

    def _apply_text_boxes_visibility(self) -> None:
        show_textboxes = self.text_boxes_action.isChecked()
        visibility_changed = self.scene.show_textboxes != show_textboxes
        self.scene.show_textboxes = show_textboxes
        if not show_textboxes and self.view.active_create_tool() == "textbox":
            self.view.activate_create_tool(None)
        if visibility_changed:
            self.scene.refresh_items(force_sync=True)
        self._sync_textbox_insert_action()

    def _toggle_snap_grid(self) -> None:
        enabled = self.snap_grid_action.isChecked()
        self.scene.snap_weeks = enabled
        self.scene.snap_rows = enabled

    def _toggle_current_week_line(self) -> None:
        self.scene.show_current_week = self.current_week_action.isChecked()
        self.scene.update_headers()

    def _toggle_missing_scope(self) -> None:
        self.scene.show_missing_scope = self.missing_scope_action.isChecked()
        self.scene.update_risk_badges()

    def _toggle_text_boxes(self) -> None:
        self._apply_text_boxes_visibility()

    def _toggle_auto_export(self) -> None:
        enabled = self.auto_export_action.isChecked()
        previous_enabled = self._auto_export_enabled
        if enabled and self._auto_export_path is None:
            if not self._configure_auto_export(require_path=True):
                with QSignalBlocker(self.auto_export_action):
                    self.auto_export_action.setChecked(False)
                enabled = False
        self._auto_export_enabled = enabled
        if enabled != previous_enabled:
            self._set_view_dirty(True)
        if enabled:
            self._warn_auto_export_constraints()

    def maybe_save(self) -> bool:
        if self.undo_stack.isClean() and not self._view_dirty:
            return True
        result = QMessageBox.question(
            self,
            "Unsaved Changes",
            "Save changes before continuing?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No
            | QMessageBox.StandardButton.Cancel,
        )
        if result == QMessageBox.StandardButton.Yes:
            return self.save_project()
        if result == QMessageBox.StandardButton.Cancel:
            return False
        return True

    def new_project(self) -> None:
        if not self.maybe_save():
            return
        year, ok = QInputDialog.getInt(self, "New Project", "Year", datetime.now().year)
        if not ok:
            return
        self.undo_stack.clear()
        self.current_path = None
        self._set_model(ProjectModel(year=year))
        self.undo_stack.setClean()
        self._set_view_dirty(False)
        self._update_title()

    def open_project(self) -> None:
        if not self.maybe_save():
            return
        filename, _ = QFileDialog.getOpenFileName(self, "Open Project", "", "Project Plans (*.json)")
        if not filename:
            return
        try:
            self._load_from_path(Path(filename))
        except Exception as exc:
            QMessageBox.warning(self, "Open Failed", f"Could not open project file.\n{exc}")

    def _load_from_path(self, path: Path) -> None:
        model, view_state = load_project(path)
        self.undo_stack.clear()
        self.current_path = path
        self._set_model(model, view_state)
        self.undo_stack.setClean()
        self._set_view_dirty(False)
        self.settings.setValue("last_file", str(path))
        self._add_recent_file(path)
        self._update_title()

    def save_project(self) -> bool:
        if not self.current_path:
            return self.save_project_as()
        view_state = self._view_state()
        save_project(self.current_path, self.model, view_state)
        self.undo_stack.setClean()
        self._set_view_dirty(False)
        self._update_title()
        self._add_recent_file(self.current_path)
        self._maybe_auto_export_png()
        return True

    def save_project_as(self) -> bool:
        filename, _ = QFileDialog.getSaveFileName(self, "Save Project As", "", "Project Plans (*.json)")
        if not filename:
            return False
        path = Path(filename)
        if path.suffix.lower() != ".json":
            path = path.with_suffix(".json")
        self.current_path = path
        result = self.save_project()
        if result:
            self.settings.setValue("last_file", str(path))
        return result

    def show_auto_export_preferences(self) -> None:
        self._configure_auto_export(require_path=False)

    def _configure_auto_export(self, require_path: bool) -> bool:
        dialog = QDialog(self)
        dialog.setWindowTitle("Auto-export Preferences")
        layout = QFormLayout(dialog)
        layout.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        dialog.setMinimumWidth(520)
        path_input = QLineEdit(dialog)
        path_input.setMinimumWidth(320)
        if self._auto_export_path is not None:
            path_input.setText(str(self._auto_export_path))
        else:
            path_input.setPlaceholderText("Select a PNG file")
        browse_button = QPushButton("Browse...", dialog)
        path_row = QWidget(dialog)
        path_layout = QHBoxLayout(path_row)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.addWidget(path_input)
        path_layout.addWidget(browse_button)

        def _browse() -> None:
            start_path = path_input.text().strip()
            if not start_path:
                start_path = self._default_auto_export_path()
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "Select Auto-export File",
                start_path,
                "PNG Files (*.png)",
            )
            if filename:
                normalized = self._normalize_auto_export_path(Path(filename))
                path_input.setText(str(normalized))

        browse_button.clicked.connect(_browse)
        layout.addRow("Destination file", path_row)

        quarters_spin = QSpinBox(dialog)
        quarters_spin.setRange(0, 12)
        quarters_spin.setValue(self._auto_export_additional_quarters)
        quarters_spin.setSuffix(" quarters")
        layout.addRow("Additional quarters", quarters_spin)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        layout.addRow(buttons)

        def _accept() -> None:
            path_value = path_input.text().strip()
            needs_path = require_path or self._auto_export_enabled
            if needs_path and not path_value:
                QMessageBox.warning(
                    dialog, "Auto-export Preferences", "Select a destination file."
                )
                return
            if path_value:
                raw_candidate = Path(path_value).expanduser()
                if raw_candidate.exists() and raw_candidate.is_dir():
                    QMessageBox.warning(
                        dialog,
                        "Auto-export Preferences",
                        "Destination must be a file.",
                    )
                    return
                candidate = self._normalize_auto_export_path(raw_candidate)
                if candidate.parent.exists() and not candidate.parent.is_dir():
                    QMessageBox.warning(
                        dialog,
                        "Auto-export Preferences",
                        "Destination folder is invalid.",
                    )
                    return
                path_input.setText(str(candidate))
            dialog.accept()

        buttons.accepted.connect(_accept)
        buttons.rejected.connect(dialog.reject)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return False

        path_value = path_input.text().strip()
        if path_value:
            path = self._normalize_auto_export_path(Path(path_value).expanduser())
        else:
            path = None
        additional_quarters = quarters_spin.value()
        previous_path = self._auto_export_path
        previous_quarters = self._auto_export_additional_quarters
        self._auto_export_path = path
        self._auto_export_additional_quarters = additional_quarters
        if path != previous_path or additional_quarters != previous_quarters:
            self._set_view_dirty(True)
        if self._auto_export_enabled:
            self._warn_auto_export_constraints()
        return True

    def _warn_auto_export_constraints(self) -> None:
        message = self._auto_export_warning_message()
        if message:
            QMessageBox.warning(self, "Auto-export", message)

    def _auto_export_warning_message(self) -> str | None:
        if not self._auto_export_enabled:
            return None
        path = self._auto_export_path
        if path is None:
            return "Auto-export is enabled but no destination file is configured."
        if path.exists() and path.is_dir():
            return f"Auto-export destination must be a file:\n{path}"
        path = self._normalize_auto_export_path(path)
        parent = path.parent
        if parent.exists() and not parent.is_dir():
            return f"Auto-export destination folder is invalid:\n{parent}"
        if not parent.exists():
            try:
                parent.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                return (
                    "Auto-export cannot create the destination folder:\n"
                    f"{parent}\n{exc}"
                )
        if path.exists():
            if not os.access(path, os.W_OK):
                return f"Auto-export destination is not writable:\n{path}"
        else:
            if not os.access(parent, os.W_OK):
                return f"Auto-export folder is not writable:\n{parent}"
        return None

    def _maybe_auto_export_png(self) -> None:
        if not self._auto_export_enabled:
            return
        if self._auto_export_path is None:
            return
        try:
            self._auto_export_png()
        except Exception as exc:
            QMessageBox.warning(self, "Auto-export", f"Could not auto-export PNG.\n{exc}")

    def _auto_export_png(self) -> None:
        source_rect = self._auto_export_range()
        if source_rect is None:
            return
        path = self._auto_export_path
        if path is None:
            return
        if path.exists() and path.is_dir():
            QMessageBox.warning(
                self, "Auto-export", f"Destination must be a file.\n{path}"
            )
            return
        path = self._normalize_auto_export_path(path)
        folder = path.parent
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            QMessageBox.warning(
                self, "Auto-export", f"Could not create export folder.\n{exc}"
            )
            return
        if not self._export_png_to_path(path, source_rect):
            QMessageBox.warning(
                self, "Auto-export", f"Could not save auto-export PNG.\n{path}"
            )

    def _auto_export_range(self) -> QRectF | None:
        additional_quarters = max(0, int(self._auto_export_additional_quarters))
        entries = self._quarter_entries(additional_quarters)
        if not entries:
            return None
        current_year, current_quarter = self._current_iso_quarter()
        end_year, end_quarter = self._add_quarters(
            current_year, current_quarter, additional_quarters
        )
        entry_index = {
            (entry["year"], entry["quarter"]): idx for idx, entry in enumerate(entries)
        }
        start_entry = entries[entry_index.get((current_year, current_quarter), 0)]
        end_entry = entries[entry_index.get((end_year, end_quarter), len(entries) - 1)]
        return self._export_rect_for_entries(start_entry, end_entry)

    def export_png(self) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export PNG",
            self._default_export_path("png"),
            "PNG Files (*.png)",
        )
        if not filename:
            return
        path = Path(filename)
        source_rect = self._select_export_range()
        if source_rect is None:
            return
        if self._export_png_to_path(path, source_rect):
            self._prompt_open_export_folder("Export PNG", path)

    def copy_image_to_clipboard(self) -> None:
        source_rect = self._select_export_range()
        if source_rect is None:
            return
        image = self._render_export_png_image(source_rect)
        if image.isNull():
            QMessageBox.warning(
                self, "Copy Image to Clipboard", "Could not render planning image."
            )
            return
        clipboard = QApplication.clipboard()
        clipboard.setImage(image)
        QMessageBox.information(
            self, "Copy Image to Clipboard", "Planning image copied to clipboard."
        )

    def export_pdf(self) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export PDF",
            self._default_export_path("pdf"),
            "PDF Files (*.pdf)",
        )
        if not filename:
            return
        path = Path(filename)
        source_rect = self._select_export_range()
        if source_rect is None:
            return
        writer = QPdfWriter(str(path))
        writer.setResolution(72)
        writer.setPageSize(QPageSize(source_rect.size(), QPageSize.Unit.Point))
        painter = QPainter(writer)
        target = QRectF(0, 0, source_rect.width(), source_rect.height())
        self._render_export(painter, source_rect, target)
        painter.end()
        if path.exists():
            self._prompt_open_export_folder("Export PDF", path)

    def _export_png_to_path(self, path: Path, source_rect: QRectF) -> bool:
        image = self._render_export_png_image(source_rect)
        return image.save(str(path))

    def _render_export_png_image(self, source_rect: QRectF) -> QImage:
        image_width = max(1, int(source_rect.width()))
        image_height = max(1, int(source_rect.height()))
        image = QImage(image_width, image_height, QImage.Format.Format_ARGB32)
        image.fill(Qt.GlobalColor.white)
        painter = QPainter(image)
        target = QRectF(0, 0, image.width(), image.height())
        self._render_export(painter, source_rect, target)
        painter.end()
        return image

    def export_risks(self) -> None:
        rows = self._collect_risk_rows()
        if not rows:
            QMessageBox.information(self, "Export Risks", "No risks to export.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Risks",
            self._default_risks_export_path(),
            "CSV Files (*.csv)",
        )
        if not filename:
            return
        path = Path(filename)
        if path.suffix.lower() != ".csv":
            path = path.with_suffix(".csv")
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle, delimiter=";")
                writer.writerow(["week", "deliverable", "risk", "probability", "impact"])
                writer.writerows(rows)
        except Exception as exc:
            QMessageBox.warning(self, "Export Risks", f"Could not export risks.\n{exc}")
            return
        count = len(rows)
        QMessageBox.information(
            self,
            "Export Risks",
            f"Exported {count} risk{'s' if count != 1 else ''} to:\n{path}",
        )
        self._prompt_open_export_folder("Export Risks", path)

    def export_scope(self) -> None:
        selected_rows = self._select_scope_rows()
        if not selected_rows:
            return
        lines, count = self._collect_scope_lines(selected_rows)
        if count == 0:
            QMessageBox.information(self, "Export Scope", "No scope entries to export.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Scope",
            self._default_scope_export_path(),
            "Markdown Files (*.md)",
        )
        if not filename:
            return
        path = Path(filename)
        if path.suffix.lower() != ".md":
            path = path.with_suffix(".md")
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            content = "\n".join(lines)
            path.write_text(content, encoding="utf-8")
        except Exception as exc:
            QMessageBox.warning(self, "Export Scope", f"Could not export scope.\n{exc}")
            return
        QMessageBox.information(
            self,
            "Export Scope",
            f"Exported {count} scope item{'s' if count != 1 else ''} to:\n{path}",
        )
        self._prompt_open_export_folder("Export Scope", path)

    def _render_export(self, painter: QPainter, source_rect: QRectF, target_rect: QRectF) -> None:
        self.scene.render(painter, target_rect, source_rect)
        self._draw_export_labels(painter, source_rect, target_rect)

    def _prompt_open_export_folder(self, title: str, path: Path) -> None:
        folder = path.parent
        result = QMessageBox.question(
            self,
            title,
            f"Open the save folder?\n{folder}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if result == QMessageBox.StandardButton.Yes:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _default_export_path(self, extension: str) -> str:
        download_dir = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DownloadLocation
        )
        if not download_dir:
            download_dir = str(Path.home() / "Downloads")
        base = self.current_path.stem if self.current_path else "untitled"
        timestamp = datetime.now().strftime("%Y_%m_%d_%H-%M-%S")
        filename = f"export_{base}_{timestamp}.{extension}"
        return str(Path(download_dir) / filename)

    def _default_auto_export_path(self) -> str:
        download_dir = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DownloadLocation
        )
        if not download_dir:
            download_dir = str(Path.home() / "Downloads")
        base = self.current_path.stem if self.current_path else "untitled"
        filename = f"auto_export_{base}.png"
        return str(Path(download_dir) / filename)

    @staticmethod
    def _normalize_auto_export_path(path: Path) -> Path:
        if path.suffix.lower() != ".png":
            return path.with_suffix(".png")
        return path

    def _default_risks_export_path(self) -> str:
        download_dir = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DownloadLocation
        )
        if not download_dir:
            download_dir = str(Path.home() / "Downloads")
        base = self.current_path.stem if self.current_path else "untitled"
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        filename = f"risks_{base}_{timestamp}.csv"
        return str(Path(download_dir) / filename)

    def _default_scope_export_path(self) -> str:
        download_dir = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DownloadLocation
        )
        if not download_dir:
            download_dir = str(Path.home() / "Downloads")
        base = self.current_path.stem if self.current_path else "untitled"
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        filename = f"scope_{base}_{timestamp}.md"
        return str(Path(download_dir) / filename)

    def _collect_risk_rows(self) -> list[list[str]]:
        layout = self.scene.layout
        include_textboxes = self.scene.show_textboxes
        rows: list[list[str]] = []
        for obj in self.model.objects.values():
            if not include_textboxes and obj.kind == "textbox":
                continue
            if not obj.risks or not obj.risks.strip():
                continue
            week_label = self._week_label(layout, obj.start_week)
            row_label = self._row_label(obj.row_id)
            for line in obj.risks.splitlines():
                risk_text, probability, impact = self._parse_risk_line(line)
                if not risk_text:
                    continue
                rows.append([week_label, row_label, risk_text, probability, impact])
        return rows

    def _collect_scope_lines(self, selected_rows: set[str]) -> tuple[list[str], int]:
        layout = self.scene.layout
        include_textboxes = self.scene.show_textboxes
        selected_rows = set(selected_rows)
        selected_rows.add(CANVAS_ROW_ID)
        objects_by_row: dict[str, list] = {}
        all_objects: list = []
        base_text_by_id: dict[str, list[str]] = {}
        for obj in self.model.objects.values():
            if obj.kind in ("link", "connector", "arrow"):
                continue
            if not include_textboxes and obj.kind == "textbox":
                continue
            if obj.row_id not in selected_rows:
                continue
            base_lines = self._normalize_lines(obj.text)
            base_text_by_id[obj.id] = base_lines
            objects_by_row.setdefault(obj.row_id, []).append(obj)
            all_objects.append(obj)
        for row_id, row_objects in objects_by_row.items():
            row_objects.sort(key=lambda obj: (obj.start_week, obj.end_week, obj.id))

        linked_text: dict[str, list[str]] = {}
        if include_textboxes:
            for link in self.model.objects.values():
                if link.kind != "link":
                    continue
                source_id = link.link_source_id
                target_id = link.link_target_id
                if not source_id or not target_id:
                    continue
                source = self.model.objects.get(source_id)
                if source is None or source.kind != "textbox":
                    continue
                text = source.text.strip()
                if not text:
                    continue
                linked_text.setdefault(target_id, []).append(text)

        text_lines_by_id: dict[str, list[str]] = {}
        for obj_id, base_lines in base_text_by_id.items():
            lines = list(base_lines)
            for entry in linked_text.get(obj_id, []):
                lines.extend(self._normalize_lines(entry))
            text_lines_by_id[obj_id] = lines

        objects_by_id = {obj.id: obj for obj in all_objects}
        name_by_id: dict[str, str] = {}
        for obj_id, obj in objects_by_id.items():
            lines = text_lines_by_id.get(obj_id, [])
            if lines:
                name_by_id[obj_id] = lines[0]
            else:
                name_by_id[obj_id] = self._scope_unnamed_label(obj)

        dependencies = self._collect_scope_dependencies(selected_rows, objects_by_row)

        included_topics = []
        for topic in self.model.topics:
            topic_objects = objects_by_row.get(topic.id, [])
            include_topic_row = topic.id in selected_rows and bool(topic_objects)
            deliverables = [
                d
                for d in topic.deliverables
                if d.id in selected_rows and objects_by_row.get(d.id)
            ]
            if include_topic_row or deliverables:
                included_topics.append((topic, include_topic_row, deliverables))
        canvas_objects = objects_by_row.get(CANVAS_ROW_ID, [])

        lines: list[str] = []
        if all_objects:
            min_week = min(obj.start_week for obj in all_objects)
            ref_label = self._week_label(layout, min_week)
            ref_date = layout.week_index_to_date(self.model.year, min_week).isoformat()
            lines.extend(
                [
                    "## Calendar",
                    "- Interpretation: continuous_weeks",
                    "- Week standard: ISO-8601",
                    "- Reference:",
                    f"  - {ref_label}: {ref_date}",
                    "",
                    "> **Note on unnamed items**  ",
                    "> Some planning objects intentionally have no explicit name.  ",
                    '> These are exported as "(Unnamed ...)" to preserve full traceability.',
                    "",
                    "## Plan Metadata",
                    f"- Year: {self.model.year}",
                    f"- Classification: {self.model.classification}",
                    f"- Schema version: {SCHEMA_VERSION}",
                    "",
                ]
            )
        count = 0
        for index, (topic, include_topic_row, deliverables) in enumerate(included_topics):
            if index:
                lines.append("---")
                lines.append("")
            topic_title = topic.name.strip() or self._scope_unnamed_item_label("topic")
            lines.append(f"# {topic_title}")
            lines.append("")
            if include_topic_row:
                topic_objects = objects_by_row.get(topic.id, [])
                week_lines, section_count = self._scope_week_groups_lines(
                    topic_objects, layout, text_lines_by_id
                )
                if week_lines:
                    lines.extend(week_lines)
                    count += section_count
                    lines.append("")
            for deliverable in deliverables:
                deliverable_title = (
                    deliverable.name.strip()
                    or self._scope_unnamed_item_label("deliverable")
                )
                section_lines, section_count = self._scope_deliverable_section(
                    deliverable_title,
                    objects_by_row.get(deliverable.id, []),
                    layout,
                    text_lines_by_id,
                )
                if section_lines:
                    lines.extend(section_lines)
                    count += section_count

        if canvas_objects:
            if included_topics:
                lines.append("---")
                lines.append("")
            lines.append("# Canvas")
            lines.append("")
            week_lines, section_count = self._scope_week_groups_lines(
                canvas_objects, layout, text_lines_by_id
            )
            if week_lines:
                lines.extend(week_lines)
                count += section_count
                lines.append("")

        dependency_section = self._scope_dependency_section(dependencies, name_by_id)
        if dependency_section:
            lines.extend(dependency_section)
            lines.append("")

        deadline_section = self._scope_deadline_section(
            list(self.model.objects.values()), layout, name_by_id
        )
        if deadline_section:
            lines.extend(deadline_section)
            lines.append("")

        reference_section = self._scope_reference_section(
            objects_by_id, text_lines_by_id, layout
        )
        if reference_section:
            lines.extend(reference_section)
            lines.append("")

        while lines and lines[-1] == "":
            lines.pop()
        return lines, count

    def _scope_deliverable_section(
        self,
        title: str,
        objects: list,
        layout,
        text_lines_by_id: dict[str, list[str]],
    ) -> tuple[list[str], int]:
        week_lines, count = self._scope_week_groups_lines(
            objects, layout, text_lines_by_id
        )
        if not week_lines:
            return [], 0
        lines: list[str] = [f"## {title}", ""]
        lines.extend(week_lines)
        lines.append("")
        return lines, count

    def _scope_week_groups_lines(
        self,
        objects: list,
        layout,
        text_lines_by_id: dict[str, list[str]],
    ) -> tuple[list[str], int]:
        lines: list[str] = []
        if not objects:
            return lines, 0
        week_groups: dict[int, list] = {}
        for obj in objects:
            week_groups.setdefault(obj.start_week, []).append(obj)
        count = 0
        for week in sorted(week_groups):
            week_label = self._week_label(layout, week)
            lines.append(f"### {week_label}")
            lines.append("")
            week_objects = week_groups[week]
            show_roles = len(week_objects) > 1
            role_groups: dict[str, list] = {}
            for obj in week_objects:
                role_groups.setdefault(self._scope_role_for_object(obj), []).append(obj)
            for role in self._scope_role_order():
                group = role_groups.get(role)
                if not group:
                    continue
                if show_roles:
                    lines.append(f"#### {role}")
                    lines.append("")
                for obj in group:
                    lines.extend(self._scope_object_lines(obj, text_lines_by_id))
                    count += 1
                lines.append("")
            if lines and lines[-1] == "":
                lines.pop()
            lines.append("")
        if lines and lines[-1] == "":
            lines.pop()
        return lines, count

    def _scope_object_lines(
        self,
        obj,
        text_lines_by_id: dict[str, list[str]],
    ) -> list[str]:
        lines: list[str] = []
        duration = self._scope_object_duration(obj)
        if duration is not None:
            suffix = "week" if duration == 1 else "weeks"
            duration_line = f"Duration: {duration} {suffix}"
        else:
            duration_line = None

        text_lines = text_lines_by_id.get(obj.id, [])
        if text_lines:
            lines.append(f"- {text_lines[0]}")
        else:
            lines.append(f"- {self._scope_unnamed_label(obj)}")

        if len(text_lines) > 1:
            lines.append("  - Text (cont.):")
            for line in text_lines[1:]:
                lines.append(f"    - {line}")

        type_label = self._scope_object_type(obj)
        lines.append(f"  - Type: {type_label}")
        if duration_line:
            lines.append(f"  - {duration_line}")

        lines.append("  - Timing:")
        lines.append(f"    - Start offset: {obj.start_week}")
        lines.append(f"    - End offset: {obj.end_week}")

        scope_lines = self._normalize_lines(
            obj.scope, drop_labels={"scope", "scopes"}
        )
        if scope_lines:
            lines.append("  - Scope:")
            for line in scope_lines:
                lines.append(f"    - {line}")

        risk_lines = self._normalize_lines(
            obj.risks, drop_labels={"risk", "risks"}
        )
        if risk_lines:
            lines.append("  - Risks:")
            for line in risk_lines:
                lines.append(f"    - {line}")
        return lines

    def _normalize_lines(
        self, text: str, drop_labels: set[str] | None = None
    ) -> list[str]:
        lines: list[str] = []
        for raw_line in text.splitlines():
            cleaned = raw_line.strip()
            if not cleaned:
                continue
            cleaned = self._strip_bullet_prefix(cleaned)
            if drop_labels:
                cleaned = self._strip_label_prefix(cleaned, drop_labels)
                if not cleaned:
                    continue
                cleaned = self._strip_bullet_prefix(cleaned)
            if cleaned:
                lines.append(cleaned)
        return lines

    @staticmethod
    def _strip_bullet_prefix(line: str) -> str:
        cleaned = line.lstrip()
        bullet_prefixes = ("- ", "* ", "+ ")
        while True:
            original = cleaned
            for prefix in bullet_prefixes:
                if cleaned.startswith(prefix):
                    cleaned = cleaned[len(prefix) :].lstrip()
                    break
            else:
                idx = 0
                while idx < len(cleaned) and cleaned[idx].isdigit():
                    idx += 1
                if (
                    idx
                    and idx < len(cleaned)
                    and cleaned[idx] in (".", ")")
                    and idx + 1 < len(cleaned)
                    and cleaned[idx + 1].isspace()
                ):
                    cleaned = cleaned[idx + 1 :].lstrip()
                else:
                    break
            if cleaned == original:
                break
        return cleaned

    @staticmethod
    def _strip_label_prefix(line: str, labels: set[str]) -> str:
        cleaned = line.strip()
        lower = cleaned.lower()
        for label in labels:
            if lower == label or lower == f"{label}:":
                return ""
        for label in labels:
            for separator in (":", " -"):
                prefix = f"{label}{separator}"
                if lower.startswith(prefix):
                    return cleaned[len(prefix) :].strip()
        return cleaned

    def _match_dependency_object(self, objects: list, week: int):
        best = None
        for obj in objects:
            if obj.kind == "arrow":
                continue
            if not (obj.start_week <= week <= obj.end_week):
                continue
            if week == obj.end_week:
                score = 0
                rel = "finish"
            elif week == obj.start_week:
                score = 1
                rel = "start"
            else:
                dist_start = abs(week - obj.start_week)
                dist_end = abs(obj.end_week - week)
                if dist_end <= dist_start:
                    score = 2 + dist_end
                    rel = "finish"
                else:
                    score = 2 + dist_start
                    rel = "start"
            if best is None or score < best[0]:
                best = (score, obj, rel)
        if best is None:
            return None, None
        return best[1], best[2]

    def _collect_scope_dependencies(
        self,
        selected_rows: set[str],
        objects_by_row: dict[str, list],
    ) -> list[tuple[str, str, str]]:
        dependencies: list[tuple[str, str, str]] = []
        exportable_ids = {
            obj.id for row_objects in objects_by_row.values() for obj in row_objects
        }

        def _add_dep(source, target, dep_type: str) -> None:
            if source.id not in exportable_ids or target.id not in exportable_ids:
                return
            dependencies.append((source.id, target.id, dep_type))

        for connector in self.model.objects.values():
            if connector.kind != "connector":
                continue
            source_id = connector.connector_source_id
            target_id = connector.connector_target_id
            if not source_id or not target_id:
                continue
            source = self.model.objects.get(source_id)
            target = self.model.objects.get(target_id)
            if source is None or target is None:
                continue
            if source.row_id not in selected_rows or target.row_id not in selected_rows:
                continue
            source_rel = "start" if connector.connector_source_side == "left" else "finish"
            target_rel = "start" if connector.connector_target_side == "left" else "finish"
            dep_type = "SS" if source_rel == "start" and target_rel == "start" else "FS"
            _add_dep(source, target, dep_type)

        for arrow in self.model.objects.values():
            if arrow.kind != "arrow":
                continue
            source_row = arrow.row_id
            target_row = arrow.target_row_id or arrow.row_id
            if source_row not in selected_rows or target_row not in selected_rows:
                continue
            source_obj, source_rel = self._match_dependency_object(
                objects_by_row.get(source_row, []), arrow.start_week
            )
            target_week = arrow.target_week if arrow.target_week is not None else arrow.end_week
            target_obj, target_rel = self._match_dependency_object(
                objects_by_row.get(target_row, []), target_week
            )
            if (
                source_obj is None
                or target_obj is None
                or source_obj.id == target_obj.id
                or source_rel is None
                or target_rel is None
            ):
                continue
            dep_type = "SS" if source_rel == "start" and target_rel == "start" else "FS"
            _add_dep(source_obj, target_obj, dep_type)

        return dependencies

    @staticmethod
    def _scope_object_type(obj) -> str:
        kind = obj.kind
        if kind == "box":
            return "Activity"
        if kind == "text":
            return "Activity text"
        if kind == "milestone":
            return "Milestone"
        if kind == "circle":
            return "Circle"
        if kind == "deadline":
            return "Deadline"
        if kind == "textbox":
            return "Textbox"
        if kind == "arrow":
            return "Arrow"
        return kind.replace("_", " ").title()

    @staticmethod
    def _scope_unnamed_item_label(item: str) -> str:
        return f"({UNNAMED_LABEL} {item})"

    def _scope_unnamed_label(self, obj) -> str:
        type_label = self._scope_object_type(obj)
        return self._scope_unnamed_item_label(type_label.lower())

    def _scope_unnamed_reference_label(self, obj, layout) -> str:
        type_label = self._scope_object_type(obj)
        week_label = self._week_label(layout, obj.start_week)
        return self._scope_unnamed_item_label(f"{type_label.lower()} @ {week_label}")

    @staticmethod
    def _scope_object_duration(obj) -> int | None:
        if obj.kind in ("box", "text"):
            return max(1, obj.end_week - obj.start_week + 1)
        return None

    @staticmethod
    def _scope_role_for_object(obj) -> str:
        kind = obj.kind
        if kind in ("milestone", "circle"):
            return "Milestones"
        if kind in ("box", "text"):
            return "Activities starting"
        if kind == "deadline":
            return "Deadlines"
        if kind == "textbox":
            return "Annotations"
        return "Other"

    @staticmethod
    def _scope_role_order() -> list[str]:
        return ["Milestones", "Activities starting", "Deadlines", "Annotations", "Other"]

    def _scope_deadline_section(
        self,
        objects: list,
        layout,
        name_by_id: dict[str, str],
    ) -> list[str]:
        deadlines: list[tuple[int, int, str]] = []
        for obj in objects:
            if obj.kind != "deadline":
                continue
            name = name_by_id.get(obj.id, self._scope_unnamed_label(obj))
            deadlines.append((obj.start_week, obj.end_week, name))
        if not deadlines:
            return []
        lines = ["## Deadlines", ""]
        dash = "\u2014"
        for start_week, end_week, name in sorted(
            deadlines, key=lambda item: (item[0], item[2].lower())
        ):
            week_label = self._week_label(layout, start_week)
            lines.append(f"- {name} {dash} {week_label}")
            lines.append("  - Timing:")
            lines.append(f"    - Start offset: {start_week}")
            lines.append(f"    - End offset: {end_week}")
        return lines

    @staticmethod
    def _scope_dependency_section(
        dependencies: list[tuple[str, str, str]],
        name_by_id: dict[str, str],
    ) -> list[str]:
        entries: list[tuple[str, str, str]] = []
        seen: set[tuple[str, str, str]] = set()
        for source_id, target_id, dep_type in dependencies:
            source_name = name_by_id.get(source_id)
            target_name = name_by_id.get(target_id)
            if not source_name or not target_name:
                continue
            key = (source_id, target_id, dep_type)
            if key in seen:
                continue
            seen.add(key)
            entries.append((source_name, target_name, dep_type))
        if not entries:
            return []
        entries.sort(key=lambda item: (item[0].lower(), item[1].lower(), item[2]))
        lines = ["## Dependencies", ""]
        arrow = "\u2192"
        for source_name, target_name, dep_type in entries:
            lines.append(f"- {source_name} {arrow} {target_name} ({dep_type})")
        return lines

    def _scope_reference_section(
        self,
        objects_by_id: dict[str, object],
        text_lines_by_id: dict[str, list[str]],
        layout,
    ) -> list[str]:
        if not objects_by_id:
            return []
        entries: list[tuple[str, str]] = []
        for obj_id, obj in objects_by_id.items():
            lines = text_lines_by_id.get(obj_id, [])
            if lines:
                label = lines[0]
            else:
                label = self._scope_unnamed_reference_label(obj, layout)
            entries.append((label, obj_id))
        entries.sort(key=lambda item: (item[0].lower(), item[1]))
        lines = ["## Object Reference (Appendix)", ""]
        for label, obj_id in entries:
            lines.append(f"- {label}: {obj_id}")
        return lines

    def _week_label(self, layout, week_index: int) -> str:
        year, week_in_year = layout.week_index_to_year_week(self.model.year, week_index)
        return f"wk{year % 100:02d}{week_in_year:02d}"

    def _row_label(self, row_id: str) -> str:
        if row_id == CANVAS_ROW_ID:
            return "Canvas"
        result = self.model.find_row(row_id)
        if result is None:
            return ""
        _kind, topic, deliverable = result
        if deliverable is not None:
            return deliverable.name
        return topic.name

    @staticmethod
    def _normalize_risk_level(value: str) -> str:
        cleaned = value.strip().lower()
        if cleaned in ("h", "high"):
            return "high"
        if cleaned in ("l", "low"):
            return "low"
        if cleaned in ("m", "med", "medium"):
            return "medium"
        return "medium"

    def _parse_risk_line(self, line: str) -> tuple[str, str, str]:
        parts = [part.strip() for part in line.split(";")]
        if not parts:
            return ("", "medium", "medium")
        risk_text = parts[0].strip()
        probability = self._normalize_risk_level(parts[1] if len(parts) > 1 else "")
        impact = self._normalize_risk_level(parts[2] if len(parts) > 2 else "")
        return (risk_text, probability, impact)

    def _select_scope_rows(self) -> set[str] | None:
        if not self.model.topics:
            QMessageBox.information(self, "Export Scope", "No topics or deliverables available.")
            return None
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Scope Rows")
        dialog.resize(420, 320)
        layout = QVBoxLayout(dialog)
        intro = QLabel("Select the topics and deliverables to include in the export.")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        tree = QTreeWidget(dialog)
        tree.setHeaderHidden(True)
        layout.addWidget(tree)

        items: list[tuple[str, QTreeWidgetItem]] = []
        topic_items: list[QTreeWidgetItem] = []
        for topic in self.model.topics:
            topic_item = QTreeWidgetItem(tree, [topic.name])
            topic_item.setFlags(topic_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            topic_item.setCheckState(0, Qt.CheckState.Unchecked)
            items.append((topic.id, topic_item))
            topic_items.append(topic_item)
            for deliverable in topic.deliverables:
                deliverable_item = QTreeWidgetItem(topic_item, [deliverable.name])
                deliverable_item.setFlags(
                    deliverable_item.flags() | Qt.ItemFlag.ItemIsUserCheckable
                )
                deliverable_item.setCheckState(0, Qt.CheckState.Unchecked)
                items.append((deliverable.id, deliverable_item))

        tree.expandAll()

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        layout.addWidget(button_box)

        updating = False

        def _apply_topic_state(topic_item: QTreeWidgetItem, checked: bool) -> None:
            for index in range(topic_item.childCount()):
                child = topic_item.child(index)
                child.setDisabled(checked)
                child.setCheckState(
                    0, Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
                )

        def _on_item_changed(item: QTreeWidgetItem, _column: int) -> None:
            nonlocal updating
            if updating:
                return
            if item.parent() is not None:
                return
            updating = True
            _apply_topic_state(item, item.checkState(0) == Qt.CheckState.Checked)
            updating = False

        tree.itemChanged.connect(_on_item_changed)

        updating = True
        for topic_item in topic_items:
            topic_item.setCheckState(0, Qt.CheckState.Checked)
            _apply_topic_state(topic_item, True)
        updating = False

        def _selected_row_ids() -> set[str]:
            return {
                row_id
                for row_id, item in items
                if item.checkState(0) == Qt.CheckState.Checked
            }

        def _accept() -> None:
            if not _selected_row_ids():
                QMessageBox.warning(
                    dialog,
                    "Export Scope",
                    "Select at least one topic or deliverable.",
                )
                return
            dialog.accept()

        button_box.accepted.connect(_accept)
        button_box.rejected.connect(dialog.reject)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        return _selected_row_ids()

    def _draw_export_labels(
        self, painter: QPainter, source_rect: QRectF, target_rect: QRectF
    ) -> None:
        layout = self.scene.layout
        model = self.model
        scale_x = target_rect.width() / max(1.0, source_rect.width())
        scale_y = target_rect.height() / max(1.0, source_rect.height())
        scale = min(scale_x, scale_y)
        label_width = layout.label_width * scale_x
        viewport_height = target_rect.height()
        header_top = (0.0 - source_rect.top()) * scale_y
        header_bottom = (layout.header_height - source_rect.top()) * scale_y

        painter.save()
        painter.translate(target_rect.left(), target_rect.top())

        painter.fillRect(QRectF(0, 0, label_width, viewport_height), QColor(245, 245, 245))
        painter.fillRect(
            QRectF(0, header_top, label_width, header_bottom - header_top),
            QColor(248, 248, 248),
        )

        selected_row_id = getattr(self.scene, "selected_row_id", None)
        if selected_row_id and selected_row_id in layout.row_map:
            row = layout.row_map[selected_row_id]
            row_top = (layout.header_height + row.y - source_rect.top()) * scale_y
            row_bottom = row_top + (row.height * scale_y)
            painter.fillRect(
                QRectF(0, row_top, label_width, row_bottom - row_top),
                QColor(227, 236, 247),
            )

        pen_grid = QPen(QColor(220, 220, 220))
        pen_grid.setWidth(1)
        painter.setPen(pen_grid)
        painter.drawLine(int(label_width), 0, int(label_width), int(viewport_height))
        painter.drawLine(0, int(header_bottom), int(label_width), int(header_bottom))
        for row in layout.rows:
            row_bottom = (layout.header_height + row.y + row.height - source_rect.top()) * scale_y
            if row_bottom < 0 or row_bottom > viewport_height + 1:
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
            row_top = (layout.header_height + row.y - source_rect.top()) * scale_y
            row_bottom = row_top + (row.height * scale_y)
            if row_bottom < 0 or row_top > viewport_height:
                continue
            row_height = row_bottom - row_top
            label_rect = QRectF(0, row_top, label_width, row_height)
            text_x = (8 * scale) + (row.indent * indent_step)

            if row.kind == "topic":
                painter.setFont(topic_font)
                topic = model.get_topic(row.row_id)
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

    def _select_export_range(self) -> QRectF | None:
        entries = self._quarter_entries()
        if not entries:
            QMessageBox.information(self, "Export", "No quarters available to export.")
            return None
        current_year, current_quarter = self._current_iso_quarter()
        end_year, end_quarter = self._add_quarters(current_year, current_quarter, 1)
        entry_index = {(entry["year"], entry["quarter"]): idx for idx, entry in enumerate(entries)}
        start_index = entry_index.get((current_year, current_quarter), 0)
        end_index = entry_index.get((end_year, end_quarter), min(start_index + 1, len(entries) - 1))

        dialog = QDialog(self)
        dialog.setWindowTitle("Export Quarters")
        layout = QFormLayout(dialog)
        start_combo = QComboBox(dialog)
        end_combo = QComboBox(dialog)
        for entry in entries:
            start_combo.addItem(entry["label"])
            end_combo.addItem(entry["label"])
        start_combo.setCurrentIndex(start_index)
        end_combo.setCurrentIndex(end_index)

        layout.addRow("From", start_combo)
        layout.addRow("To", end_combo)
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        layout.addRow(button_box)

        def _accept() -> None:
            if start_combo.currentIndex() > end_combo.currentIndex():
                QMessageBox.warning(
                    dialog,
                    "Invalid Range",
                    "The end quarter must be the same as or after the start quarter.",
                )
                return
            dialog.accept()

        button_box.accepted.connect(_accept)
        button_box.rejected.connect(dialog.reject)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        start_entry = entries[start_combo.currentIndex()]
        end_entry = entries[end_combo.currentIndex()]
        return self._export_rect_for_entries(start_entry, end_entry)

    def _export_rect_for_entries(self, start_entry: dict, end_entry: dict) -> QRectF:
        layout_obj = self.scene.layout
        scene_rect = self.scene.sceneRect()
        start_week = start_entry["start_week"] - 1
        end_week = end_entry["end_week"] + 1
        self.scene.ensure_week_range(start_week)
        self.scene.ensure_week_range(end_week)
        start_x = layout_obj.week_left_x(start_week)
        end_x = layout_obj.week_left_x(end_week) + layout_obj.week_width
        x1 = start_x - layout_obj.label_width
        width = (end_x - start_x) + layout_obj.label_width
        return QRectF(x1, scene_rect.top(), width, scene_rect.height())

    def _quarter_entries(self, future_quarters: int = 3) -> list[dict]:
        layout = self.scene.layout
        min_year, _ = layout.week_index_to_year_week(self.model.year, self.scene.min_week)
        max_year, _ = layout.week_index_to_year_week(self.model.year, self.scene.max_week)
        current_year, current_quarter = self._current_iso_quarter()
        end_year, _ = self._add_quarters(
            current_year, current_quarter, max(0, int(future_quarters))
        )
        min_year = min(min_year, current_year)
        max_year = max(max_year, end_year)

        entries = []
        for year in range(min_year, max_year + 1):
            year_weeks = layout.weeks_in_year(year)
            base_week = layout.week_index_for_iso_year(self.model.year, year)
            for quarter in range(1, 5):
                quarter_offset = (quarter - 1) * 13
                if quarter_offset >= year_weeks:
                    break
                quarter_length = min(13, year_weeks - quarter_offset)
                start_week = base_week + quarter_offset
                end_week = start_week + quarter_length - 1
                entries.append(
                    {
                        "year": year,
                        "quarter": quarter,
                        "label": f"{year} Q{quarter}",
                        "start_week": start_week,
                        "end_week": end_week,
                    }
                )
        return entries

    def _current_iso_quarter(self) -> tuple[int, int]:
        iso = datetime.now().isocalendar()
        quarter = min(4, ((iso.week - 1) // 13) + 1)
        return iso.year, quarter

    @staticmethod
    def _add_quarters(year: int, quarter: int, offset: int) -> tuple[int, int]:
        total = (year * 4) + (quarter - 1) + offset
        new_year = total // 4
        new_quarter = (total % 4) + 1
        return new_year, new_quarter

    def _view_state(self) -> dict:
        return {
            "zoom": self.view.current_zoom,
            "scroll_x": self.view.horizontalScrollBar().value(),
            "scroll_y": self.view.verticalScrollBar().value(),
            "label_width": self.scene.layout.label_width,
            AUTO_EXPORT_VIEW_KEY: self._auto_export_state(),
        }

    def _auto_export_state(self) -> dict:
        return {
            AUTO_EXPORT_VIEW_ENABLED_KEY: self._auto_export_enabled,
            AUTO_EXPORT_VIEW_PATH_KEY: str(self._auto_export_path)
            if self._auto_export_path is not None
            else "",
            AUTO_EXPORT_VIEW_ADDITIONAL_QUARTERS_KEY: int(self._auto_export_additional_quarters),
        }

    def goto_today(self) -> None:
        iso = datetime.now().isocalendar()
        layout = self.scene.layout
        base_week = layout.week_index_for_iso_year(self.model.year, iso.year)
        week_index = base_week + (iso.week - 1)
        self.scene.ensure_week_range(week_index)
        week_center_x = layout.week_center_x(week_index)
        viewport_width = self.view.viewport().width()
        scale = max(0.01, self.view.transform().m11())
        if viewport_width <= 0:
            self.view.centerOn(week_center_x, layout.header_height)
            return
        label_width_view = layout.label_width * scale
        grid_width_view = max(0.0, viewport_width - label_width_view)
        target_view_x = label_width_view + (grid_width_view / 6.0)
        target_view_x = max(0.0, min(float(viewport_width), target_view_x))
        delta_view = (viewport_width / 2.0) - target_view_x
        delta_scene = delta_view / scale
        self.view.centerOn(week_center_x + delta_scene, layout.header_height)

    def delete_selected(self) -> None:
        items = self.scene.selectedItems()
        for item in items:
            obj_id = item.data(0)
            if obj_id:
                self.controller.remove_object(obj_id)
                break

    def duplicate_selected(self) -> None:
        if not self.scene.edit_mode:
            return
        self.view.duplicate_selected()

    def edit_classification_tag(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Classification Tag")
        layout = QFormLayout(dialog)
        text_input = QLineEdit(dialog)
        text_input.setText(self.model.classification_label())
        text_input.setPlaceholderText(DEFAULT_CLASSIFICATION)
        size_spin = QSpinBox(dialog)
        size_spin.setRange(CLASSIFICATION_SIZE_MIN, CLASSIFICATION_SIZE_MAX)
        size_spin.setValue(
            self.model.classification_size or CLASSIFICATION_SIZE_DEFAULT
        )
        size_spin.setSuffix(" pt")

        layout.addRow("Label", text_input)
        layout.addRow("Size", size_spin)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        layout.addRow(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.controller.update_classification(text_input.text(), size_spin.value())

    def add_topic(self) -> None:
        name, ok = QInputDialog.getText(self, "Add Topic", "Topic Name")
        if not ok or not name:
            return
        self.controller.add_topic(name)

    def edit_topic(self) -> None:
        if not self.model.topics:
            QMessageBox.information(self, "Edit Topic(s)", "No topics available.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Topic(s)")
        layout = QFormLayout(dialog)
        topic_combo = QComboBox(dialog)
        for topic in self.model.topics:
            topic_combo.addItem(topic.name, topic.id)
        name_input = QLineEdit(dialog)
        color_input = QPushButton("Pick Color", dialog)
        color_display = QLabel(dialog)
        color_display.setMinimumWidth(80)

        def _sync_fields(index: int) -> None:
            topic_id = topic_combo.itemData(index)
            topic = self.model.get_topic(topic_id)
            if topic is None:
                return
            name_input.setText(topic.name)
            color_display.setStyleSheet(f"background-color: {topic.color};")
            color_display.setProperty("color", topic.color)

        def _pick_color() -> None:
            topic_id = topic_combo.currentData()
            topic = self.model.get_topic(topic_id)
            if topic is None:
                return
            from PyQt6.QtWidgets import QColorDialog

            color = QColorDialog.getColor()
            if not color.isValid():
                return
            color_display.setStyleSheet(f"background-color: {color.name()};")
            color_display.setProperty("color", color.name().upper())

        topic_combo.currentIndexChanged.connect(_sync_fields)
        color_input.clicked.connect(_pick_color)

        layout.addRow("Topic", topic_combo)
        layout.addRow("Name", name_input)
        layout.addRow("Color", color_input)
        layout.addRow("Preview", color_display)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addRow(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        _sync_fields(0)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            topic_id = topic_combo.currentData()
            topic = self.model.get_topic(topic_id)
            if topic is None:
                return
            new_color = color_display.property("color") or topic.color
            new_topic = type(topic)(
                id=topic.id,
                name=name_input.text() or topic.name,
                color=new_color,
                collapsed=topic.collapsed,
                deliverables=topic.deliverables,
            )
            self.controller.update_topic(new_topic)

    def edit_deliverable(self) -> None:
        entries = []
        for topic in self.model.topics:
            for deliverable in topic.deliverables:
                entries.append((topic, deliverable))
        if not entries:
            QMessageBox.information(self, "Edit Deliverable(s)", "No deliverables available.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Deliverable(s)")
        layout = QFormLayout(dialog)
        deliverable_combo = QComboBox(dialog)
        for topic, deliverable in entries:
            label = f"{topic.name} - {deliverable.name}"
            deliverable_combo.addItem(label, deliverable.id)
        name_input = QLineEdit(dialog)

        def _sync_fields(index: int) -> None:
            deliverable_id = deliverable_combo.itemData(index)
            found = self.model.find_deliverable(deliverable_id)
            if found is None:
                return
            _topic, _index, deliverable = found
            name_input.setText(deliverable.name)

        deliverable_combo.currentIndexChanged.connect(_sync_fields)

        layout.addRow("Deliverable", deliverable_combo)
        layout.addRow("Name", name_input)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addRow(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        selected_id = self._selected_deliverable_id()
        if selected_id:
            index = deliverable_combo.findData(selected_id)
            if index != -1:
                deliverable_combo.setCurrentIndex(index)
        _sync_fields(deliverable_combo.currentIndex())

        if dialog.exec() == QDialog.DialogCode.Accepted:
            deliverable_id = deliverable_combo.currentData()
            found = self.model.find_deliverable(deliverable_id)
            if found is None:
                return
            _topic, _index, deliverable = found
            name = name_input.text().strip() or deliverable.name
            if name == deliverable.name:
                return
            new_deliverable = type(deliverable)(id=deliverable.id, name=name)
            self.controller.update_deliverable(new_deliverable)

    def add_deliverable(self) -> None:
        if not self.model.topics:
            QMessageBox.information(self, "Add Deliverable", "Add a topic first.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Deliverable")
        layout = QFormLayout(dialog)
        topic_combo = QComboBox(dialog)
        for topic in self.model.topics:
            topic_combo.addItem(topic.name, topic.id)
        name_input = QLineEdit(dialog)
        layout.addRow("Topic", topic_combo)
        layout.addRow("Deliverable", name_input)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addRow(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = name_input.text().strip()
            if not name:
                return
            topic_id = topic_combo.currentData()
            self.controller.add_deliverable(topic_id, name)

    def move_deliverable_up(self) -> None:
        deliverable_id = self._selected_deliverable_id()
        if not deliverable_id:
            QMessageBox.information(
                self,
                "Move Deliverable",
                "Select a deliverable row by clicking its name in the left column.",
            )
            return
        self.controller.move_deliverable(deliverable_id, -1)

    def move_deliverable_down(self) -> None:
        deliverable_id = self._selected_deliverable_id()
        if not deliverable_id:
            QMessageBox.information(
                self,
                "Move Deliverable",
                "Select a deliverable row by clicking its name in the left column.",
            )
            return
        self.controller.move_deliverable(deliverable_id, 1)

    def remove_deliverable(self) -> None:
        deliverable_id = self._selected_deliverable_id()
        if not deliverable_id:
            QMessageBox.information(
                self,
                "Remove Deliverable",
                "Select a deliverable row by clicking its name in the left column.",
            )
            return
        found = self.model.find_deliverable(deliverable_id)
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

    def remove_topic(self) -> None:
        topic_id = self._selected_topic_id()
        if not topic_id:
            QMessageBox.information(
                self,
                "Remove Topic",
                "Select a topic row by clicking its name in the left column.",
            )
            return
        topic = self.model.get_topic(topic_id)
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
        return [
            obj
            for obj in self.model.objects.values()
            if obj.row_id in row_ids or (obj.target_row_id in row_ids)
        ]

    def _selected_deliverable_id(self) -> str | None:
        row_id = self._selected_row_id()
        if not row_id:
            return None
        row = self.scene.layout.row_map.get(row_id)
        if not row or row.kind != "deliverable":
            return None
        return row_id

    def _selected_topic_id(self) -> str | None:
        row_id = self._selected_row_id()
        if not row_id:
            return None
        row = self.scene.layout.row_map.get(row_id)
        if not row or row.kind != "topic":
            return None
        return row_id

    def _selected_row_id(self) -> str | None:
        if self.scene.selected_row_id:
            return self.scene.selected_row_id
        selected = self.scene.selectedItems()
        for item in selected:
            obj_id = item.data(0)
            if obj_id and obj_id in self.model.objects:
                return self.model.objects[obj_id].row_id
        return None

    def create_object(self, kind: str) -> None:
        if kind == "textbox" and not self.scene.show_textboxes:
            return
        if not self.scene.layout.rows and kind not in ("textbox", "deadline", "connector"):
            QMessageBox.information(self, "Add Object", "Add a topic or deliverable first.")
            return
        if self.presentation_mode_action.isChecked():
            QMessageBox.information(self, "Presentation Mode", "Exit presentation mode to edit.")
            return
        if self.nav_mode_action.isChecked():
            self.nav_mode_action.setChecked(False)
            self._toggle_navigation_mode()
        self.view.activate_create_tool(kind)
        if kind == "connector":
            message = "Drag from an object edge to another to create a connector arrow."
        elif kind == "arrow":
            message = "Drag from an object edge to another to create an arrow."
        else:
            message = f"Click and drag to place a {kind}."
        self.statusBar().showMessage(f"{message} Press Esc to cancel.", 4000)

    def closeEvent(self, event) -> None:
        self._is_closing = True
        if not self.maybe_save():
            self._is_closing = False
            event.ignore()
            return
        self._save_window_settings()
        event.accept()

    def _restore_window_settings(self) -> None:
        geometry = self.settings.value("window/geometry")
        if geometry:
            self._has_saved_geometry = True
            self.restoreGeometry(geometry)
        state = self.settings.value("window/state")
        if state:
            self._suppress_properties_sync = True
            try:
                self.restoreState(state)
            finally:
                self._suppress_properties_sync = False
        maximized = self.settings.value("window/maximized", False)
        if isinstance(maximized, bool):
            self._restore_maximized = maximized
        else:
            self._restore_maximized = str(maximized).lower() in ("1", "true", "yes")
        self._last_window_maximized = self._restore_maximized
        self._apply_properties_visibility()

    def _restore_view_settings(self) -> None:
        zoom_value = self.settings.value("view/zoom")
        if zoom_value is None:
            zoom_value = None
        try:
            zoom = float(zoom_value)
        except (TypeError, ValueError):
            zoom = None
        if zoom is not None:
            self.view.set_zoom(zoom)

        current_week_value = self.settings.value("view/current_week_line", True)
        if isinstance(current_week_value, bool):
            current_week = current_week_value
        else:
            current_week = str(current_week_value).lower() in ("1", "true", "yes")
        self.current_week_action.setChecked(current_week)
        self.scene.show_current_week = current_week
        self.scene.update_headers()

        missing_scope_value = self.settings.value("view/show_missing_scope", False)
        if isinstance(missing_scope_value, bool):
            missing_scope = missing_scope_value
        else:
            missing_scope = str(missing_scope_value).lower() in ("1", "true", "yes")
        self.missing_scope_action.setChecked(missing_scope)
        self.scene.show_missing_scope = missing_scope
        self.scene.update_risk_badges()

        text_boxes_value = self.settings.value("view/show_text_boxes", True)
        if isinstance(text_boxes_value, bool):
            show_text_boxes = text_boxes_value
        else:
            show_text_boxes = str(text_boxes_value).lower() in ("1", "true", "yes")
        with QSignalBlocker(self.text_boxes_action):
            self.text_boxes_action.setChecked(show_text_boxes)
        self._apply_text_boxes_visibility()

        properties_value = self.settings.value("view/show_properties_pane", True)
        if isinstance(properties_value, bool):
            show_properties = properties_value
        else:
            show_properties = str(properties_value).lower() in ("1", "true", "yes")
        with QSignalBlocker(self.properties_pane_action):
            self.properties_pane_action.setChecked(show_properties)
        self._apply_properties_visibility()

    def _apply_auto_export_settings(self, view_state: dict | None) -> None:
        auto_export = None
        if isinstance(view_state, dict):
            auto_export = view_state.get(AUTO_EXPORT_VIEW_KEY)
        if isinstance(auto_export, dict):
            enabled_value = auto_export.get(AUTO_EXPORT_VIEW_ENABLED_KEY, False)
            if isinstance(enabled_value, bool):
                enabled = enabled_value
            else:
                enabled = str(enabled_value).lower() in ("1", "true", "yes")

            path_value = auto_export.get(AUTO_EXPORT_VIEW_PATH_KEY, "")
            if path_value:
                path = self._normalize_auto_export_path(
                    Path(str(path_value)).expanduser()
                )
            else:
                path = None

            quarters_value = auto_export.get(
                AUTO_EXPORT_VIEW_ADDITIONAL_QUARTERS_KEY,
                AUTO_EXPORT_DEFAULT_ADDITIONAL_QUARTERS,
            )
            try:
                additional_quarters = int(quarters_value)
            except (TypeError, ValueError):
                additional_quarters = AUTO_EXPORT_DEFAULT_ADDITIONAL_QUARTERS
        else:
            enabled = False
            path = None
            additional_quarters = AUTO_EXPORT_DEFAULT_ADDITIONAL_QUARTERS

        self._auto_export_path = path
        self._auto_export_additional_quarters = max(0, additional_quarters)
        self._auto_export_enabled = enabled
        with QSignalBlocker(self.auto_export_action):
            self.auto_export_action.setChecked(enabled)
        if enabled:
            self._warn_auto_export_constraints()

    def _save_window_settings(self) -> None:
        self.settings.setValue("window/geometry", self.saveGeometry())
        self.settings.setValue("window/state", self.saveState())
        self.settings.setValue("window/maximized", self._last_window_maximized)
        self.settings.setValue("view/zoom", self.view.current_zoom)
        self.settings.setValue("view/current_week_line", self.scene.show_current_week)
        self.settings.setValue("view/show_missing_scope", self.scene.show_missing_scope)
        self.settings.setValue("view/show_text_boxes", self.scene.show_textboxes)
        self.settings.setValue(
            "view/show_properties_pane", self.properties_pane_action.isChecked()
        )
        self.settings.sync()

    def show_with_restore(self) -> None:
        if not self._has_saved_geometry:
            self.resize(1200, 800)
        self._restoring_window_state = self._restore_maximized
        self._maximize_attempts = 0
        self.show()
        if not self._restore_maximized:
            self._restoring_window_state = False
        QTimer.singleShot(0, self._hook_window_state)
        if self._restore_maximized:
            QTimer.singleShot(50, self._ensure_window_restore)

    def _is_effectively_maximized(self) -> bool:
        if not self.isMaximized():
            return False
        handle = self.windowHandle()
        if handle is None or handle.screen() is None:
            return False
        available = handle.screen().availableGeometry()
        frame = self.frameGeometry()
        margin = 8
        return (
            abs(frame.width() - available.width()) <= margin
            and abs(frame.height() - available.height()) <= margin
        )

    def _ensure_window_restore(self) -> None:
        if self._is_closing or not self._restore_maximized:
            self._restoring_window_state = False
            return
        if self._is_effectively_maximized():
            self._restoring_window_state = False
            self._last_window_maximized = True
            return
        self._maximize_attempts += 1
        self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)
        self.showMaximized()
        if self._maximize_attempts < 10:
            QTimer.singleShot(200, self._ensure_window_restore)
        else:
            self._restoring_window_state = False

    def _hook_window_state(self) -> None:
        if self._window_handle_connected:
            return
        handle = self.windowHandle()
        if handle is None:
            QTimer.singleShot(0, self._hook_window_state)
            return
        handle.windowStateChanged.connect(self._on_window_state_changed)
        self._window_handle_connected = True
        self._on_window_state_changed(handle.windowState())

    def _on_window_state_changed(self, state: Qt.WindowState) -> None:
        if self._is_closing or self._restoring_window_state:
            return
        if state & Qt.WindowState.WindowMinimized:
            return
        maximized = bool(state & (Qt.WindowState.WindowMaximized | Qt.WindowState.WindowFullScreen))
        if maximized == self._last_window_maximized:
            return
        self._last_window_maximized = maximized
        self.settings.setValue("window/maximized", self._last_window_maximized)
        self.settings.sync()


def run() -> int:
    app = QApplication(sys.argv)
    if app.cursorFlashTime() <= 0:
        app.setCursorFlashTime(1000)
    window = MainWindow()
    window.show_with_restore()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(run())
