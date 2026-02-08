# ProjectPlans

Local single-user desktop planning tool for building week-based project timelines with grid-anchored objects.

## Overview and Scope

- Desktop-only, offline, single-user planning tool for Linux and Windows.
- Timeline uses ISO weeks (some years include week 53).
- Plans are organized into topics and deliverables (rows) with grid-aligned objects.
- Textboxes are free-positioned; links and connectors attach to objects.

## Requirements and Dependencies

- Python 3.10+
- PyQt6

Install dependencies:

```
python -m pip install -r requirements.txt
```

## Run

Recommended:

```
python app.pyw
```

Alternate:

```
python -m projectplans
```

On Windows, `pythonw app.pyw` runs without a console window.

## Usage

Modes

- Edit mode: default, objects are draggable.
- Navigation mode: View menu, pan/zoom/inspect without moving objects.
- Presentation mode: View menu, read-only and hides the Properties panel.

Project Structure

- Topics contain deliverables; manage them via the Project menu or row context menu.
- Double-click a topic label to collapse/expand; right-click a deliverable label to focus/unfocus.
- Alt+Shift+Up/Down: move the selected deliverable within its topic.
- Edit the classification tag (Project menu) to set the bottom-right label and size.

Core Controls

- Mouse wheel + Ctrl: zoom.
- Ctrl+Z / Ctrl+Y: undo/redo.
- Arrow keys: nudge selection by week or row; Shift+Arrow keys resize duration (boxes/text) or adjust arrow target.
- Ctrl+D: duplicate selected object; Delete: remove selected object.
- F2 or double-click: edit item text.
- Space (edit mode) or right-click drag: pan.
- Drag multiple selected objects: move them together left/right.
- Drag the label divider to resize the row label column.

Text Editing (shortcuts)

- Ctrl+B / Ctrl+I / Ctrl+U: bold/italic/underline.
- Alt+Shift+S: toggle strikethrough.
- Ctrl+] / Ctrl+[: increase/decrease selected text size.
- Applies to inline text and Scope/Risks/Notes fields.

Insert Tools (keyboard shortcuts)

- B: Activity (box)
- T: Activity Text
- M: Milestone
- D: Deadline
- C: Circle
- A: Arrow
- F: Connector Arrow
- X: Textbox

Notes

- Use the Insert menu or shortcuts, then click-drag to place items (Esc cancels).
- Arrows and connector arrows are created by dragging from an object edge to a target.
- Textboxes can be anchored to an object by dragging from a textbox edge to a target.
- Right-click the canvas or an object for insert/convert/z-order actions; right-click row labels for row actions.

Properties Panel (Inspector)

- Edit text, weeks, rows, size, alignment, color, opacity (textboxes), scope, risks, and notes.
- Arrows and connector arrows: configure arrowheads (start/end/both) and reverse direction; arrows expose target week/row.

View Options (View menu)

- Snap to Grid
- Current Week Line
- Show Missing Scope (dashed badge on boxes without scope)
- Properties Pane

View Actions (View menu)

- Zoom In/Out and Reset Zoom.
- Zoom to Selection and Zoom to Fit.
- Goto Today (centers the current ISO week).

## Exporting

Export menu

- Scope (markdown): select rows to include; exports markdown with dependencies (arrows/connector arrows),
  deadlines, and an object reference appendix. Textboxes anchored to items contribute their text.
- Risks (csv): semicolon-delimited; optional `risk;probability;impact` per line (h/m/l).
- Planning (png) and Planning (pdf): choose quarter range in the dialog.
- Planning auto-export: auto-save a PNG on every Save.

Options menu

- Auto-export Preferences: choose the destination PNG file and the number of additional
  quarters to export (starting from the current quarter). Auto-export overwrites the
  selected file on every Save.

Default export paths are in the OS Downloads folder when available.

## Configuration and Persistence

Per plan (stored in the plan JSON):

- Model data (year, classification tag/size, topics/deliverables and collapse state, objects).
- View state (zoom, scroll position, label width).
- Auto-export settings (enabled, destination file, additional quarters).

Per user (stored in QSettings):

- Last opened file and recent files list.
- Window size, position, and dock state.
- View toggles (current week line, missing scope, properties pane) and last used zoom.

## Known Limitations

- Single-user, offline tool (no collaboration or cloud sync).
- Exporting to PNG/PDF is limited to quarter-range selections.
- Arrow and connector targets must be valid canvas objects; invalid targets are ignored.

## Best Practices

- Use Navigation mode for inspection and Presentation mode when sharing the view.
- Keep topics/deliverables structured to simplify row selection and exports.
- Use Show Missing Scope to spot boxes that need scope details.

## File Format

Projects are saved as JSON with schema versioning, year, classification tag/size, topics/deliverables,
objects, and view state (including auto-export settings).

## License

This project is licensed under the GNU General Public License v3.0 or later. See `LICENSE`.
