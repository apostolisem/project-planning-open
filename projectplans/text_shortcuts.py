from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QTextCharFormat, QTextCursor, QTextDocument


def text_shortcut_action(event) -> str | None:
    modifiers = event.modifiers()
    key = event.key()
    if modifiers & Qt.KeyboardModifier.ControlModifier:
        if key == Qt.Key.Key_B:
            return "bold"
        if key == Qt.Key.Key_I:
            return "italic"
        if key == Qt.Key.Key_U:
            return "underline"
        if key == Qt.Key.Key_BracketRight:
            return "increase"
        if key == Qt.Key.Key_BracketLeft:
            return "decrease"
    if (
        (modifiers & Qt.KeyboardModifier.AltModifier)
        and (modifiers & Qt.KeyboardModifier.ShiftModifier)
        and key == Qt.Key.Key_S
    ):
        return "strikethrough"
    return None


def font_point_size(font: QFont, fallback: int = 12) -> int:
    size = font.pointSizeF()
    if size > 0:
        return int(round(size))
    pixel = font.pixelSize()
    if pixel > 0:
        return int(pixel)
    return fallback


def apply_text_action(
    cursor: QTextCursor,
    action: str,
    base_font: QFont,
    *,
    min_size: int,
    max_size: int,
    step: int,
    require_selection: bool = True,
) -> bool:
    if require_selection and not cursor.hasSelection():
        return False
    fmt = QTextCharFormat()
    current = cursor.charFormat()
    if action == "bold":
        weight = current.fontWeight()
        is_bold = weight >= QFont.Weight.Bold
        fmt.setFontWeight(QFont.Weight.Normal if is_bold else QFont.Weight.Bold)
    elif action == "italic":
        fmt.setFontItalic(not current.fontItalic())
    elif action == "underline":
        fmt.setFontUnderline(not current.fontUnderline())
    elif action == "strikethrough":
        fmt.setFontStrikeOut(not current.fontStrikeOut())
    elif action in ("increase", "decrease"):
        size = current.fontPointSize()
        if size <= 0:
            size = float(font_point_size(base_font))
        delta = step if action == "increase" else -step
        new_size = max(min_size, min(max_size, int(round(size + delta))))
        fmt.setFontPointSize(float(new_size))
    else:
        return False
    cursor.mergeCharFormat(fmt)
    return True


def document_has_formatting(doc: QTextDocument, base_font: QFont) -> bool:
    base_size = float(font_point_size(base_font))
    block = doc.begin()
    while block.isValid():
        it = block.begin()
        while not it.atEnd():
            fragment = it.fragment()
            if fragment.isValid():
                fmt = fragment.charFormat()
                weight = fmt.fontWeight()
                if weight <= 0:
                    weight = base_font.weight()
                if weight != base_font.weight():
                    return True
                if fmt.fontItalic() != base_font.italic():
                    return True
                if fmt.fontUnderline() != base_font.underline():
                    return True
                if fmt.fontStrikeOut() != base_font.strikeOut():
                    return True
                size = fmt.fontPointSize()
                if size > 0 and abs(size - base_size) > 0.1:
                    return True
            it += 1
        block = block.next()
    return False


def extract_text_payload(doc: QTextDocument, base_font: QFont) -> tuple[str, str | None]:
    plain = doc.toPlainText()
    if document_has_formatting(doc, base_font):
        return plain, doc.toHtml()
    return plain, None
