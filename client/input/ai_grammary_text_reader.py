from __future__ import annotations

import ctypes
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from ctypes import wintypes
from pathlib import Path
from typing import Callable


BROWSER_PROCESS_NAMES = {
    "applicationframehost.exe",
    "arc.exe",
    "chrome.exe",
    "chromium.exe",
    "microsoftedge.exe",
    "msedge.exe",
    "msedge_beta.exe",
    "msedge_canary.exe",
    "msedge_dev.exe",
    "msedgewebview2.exe",
    "firefox.exe",
    "iexplore.exe",
    "vivaldi.exe",
    "whale.exe",
    "opera.exe",
    "brave.exe",
}
WORD_PROCESS_NAMES = {"winword.exe"}
HWP_PROCESS_NAMES = {"hwp.exe", "hwp64.exe", "hwpviewer.exe", "hwpw.exe"}
HWP_ACTIVE_PROGIDS = (
    "HWPFrame.HwpObject.2",
    "HWPFrame.HwpObject.1",
    "HWPFrame.HwpObject",
)
HWP_IHWP_OBJECT_IID = "{5E6A8276-CF1C-42B8-BCED-319548B02AF6}"
NOTEPAD_PROCESS_NAMES = {"notepad.exe"}
HWP_CLASS_HINTS = ("hwp", "hwpctrl", "hnc", "afx")
HWP_TITLE_HINTS = (".hwp", ".hwpx")
HWP_TEXT_CONTROL_TYPES = ("Document", "Edit", "Pane", "Text")
HWP_EXCLUDED_TEXT_HINTS = (
    "menu",
    "toolbar",
    "status",
    "navigation",
    "ribbon",
    "dialog",
    "button",
    "tab",
    "paragraph",
    "도구",
    "메뉴",
    "상태",
)

EDITOR_CLASS_HINTS = (
    "Edit",
    "RichEdit",
    "RichEdit20W",
    "RichEditD2DPT",
    "RICHEDIT50W",
)
INPUT_CONTROL_TYPES = ("Edit", "Document", "ComboBox")


def normalize_text(text: str | None) -> str:
    if not text:
        return ""
    return (
        str(text)
        .replace("\x00", "")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\v", "\n")
        .replace("\f", "\n")
    )


def has_text_content(text: str | None) -> bool:
    return bool(str(text or "").strip())


def _optional_import(module_name: str):
    try:
        return __import__(module_name)
    except Exception:
        return None


win32con = _optional_import("win32con")
win32gui = _optional_import("win32gui")
win32process = _optional_import("win32process")
win32api = _optional_import("win32api")
pythoncom = _optional_import("pythoncom")
psutil = _optional_import("psutil")

_LOG_DIR = Path(__file__).resolve().parents[2] / ".logs"
_WORD_STYLE_LOG_PATH = _LOG_DIR / "word_style.log"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_HWP_LOG_PATH = _LOG_DIR / "hwp_monitor.log"
_HWP_COM_LOG_PATH = _LOG_DIR / "hwp_com.log"


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT),
    ]


def get_foreground_hwnd() -> int:
    if win32gui is None:
        return 0
    try:
        return int(win32gui.GetForegroundWindow())
    except Exception:
        return 0


def get_window_title(hwnd: int) -> str:
    if win32gui is None or not hwnd:
        return ""
    try:
        return normalize_text(win32gui.GetWindowText(hwnd) or "").strip()
    except Exception:
        return ""


def get_class_name(hwnd: int) -> str:
    if win32gui is None or not hwnd:
        return ""
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def get_process_name(hwnd: int) -> str:
    if win32process is None or not hwnd:
        return ""
    try:
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return ""
    if not pid:
        return ""
    if psutil is not None:
        try:
            return psutil.Process(pid).name().lower()
        except Exception:
            pass
    if win32api is not None and win32process is not None:
        try:
            handle = win32api.OpenProcess(0x1000, False, pid)
            image_path = win32process.GetModuleFileNameEx(handle, 0)
            win32api.CloseHandle(handle)
            return os.path.basename(image_path).lower()
        except Exception:
            pass
    return ""


def get_process_id(hwnd: int) -> int:
    if win32process is None or not hwnd:
        return 0
    try:
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid or 0)
    except Exception:
        return 0


def looks_like_url(text: str) -> bool:
    text = normalize_text(text).strip().lower()
    if not text or " " in text:
        return False
    if text.startswith(("http://", "https://", "ftp://", "file://", "chrome://", "edge://", "about:")):
        return True
    if text.startswith("www."):
        return True
    if "." in text and "/" in text:
        return True
    return "." in text and len(text) < 256 and any(
        text.endswith(tld) for tld in (".com", ".net", ".org", ".io", ".co", ".kr", ".dev", ".app")
    )


class BasePollingReader:
    source_name = "realtime"

    def __init__(self, debug: bool = False):
        self.debug = debug
        self.last_text: str | None = None

    def _debug(self, message: str):
        if self.debug:
            print(f"[reader] {message}", file=sys.stderr)

    def read_current_text(self) -> str:
        raise NotImplementedError

    def poll(self) -> str | None:
        text = self.read_current_text()
        if text == self.last_text:
            return None
        self.last_text = text
        return text


class NotepadReader(BasePollingReader):
    def _is_notepad(self, hwnd: int) -> bool:
        title = get_window_title(hwnd).lower()
        class_name = get_class_name(hwnd).lower()
        process_name = get_process_name(hwnd)
        return (
            process_name in NOTEPAD_PROCESS_NAMES
            or "notepad" in class_name
            or "notepad" in title
            or "\uba54\ubaa8\uc7a5" in title
        )

    def _read_text(self, hwnd: int) -> str:
        if win32gui is None or win32con is None or not hwnd:
            return ""
        try:
            length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        except Exception:
            return ""
        if length <= 0:
            return ""
        buffer = ctypes.create_unicode_buffer(length + 1)
        try:
            win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
        except Exception:
            return ""
        return normalize_text(buffer.value)

    def _child_handles(self, root_hwnd: int) -> list[int]:
        if win32gui is None or not root_hwnd:
            return []
        handles: list[int] = []

        def callback(hwnd, _):
            handles.append(hwnd)
            return True

        try:
            win32gui.EnumChildWindows(root_hwnd, callback, None)
        except Exception as exc:
            self._debug(f"EnumChildWindows failed: {exc}")
        return handles

    def _focused_child_handle(self, root_hwnd: int) -> int | None:
        if win32process is None or not root_hwnd:
            return None
        try:
            thread_id, _process_id = win32process.GetWindowThreadProcessId(root_hwnd)
        except Exception:
            return None
        info = GUITHREADINFO(cbSize=ctypes.sizeof(GUITHREADINFO))
        try:
            success = ctypes.windll.user32.GetGUIThreadInfo(thread_id, ctypes.byref(info))
        except Exception:
            return None
        if not success:
            return None
        focused_hwnd = int(info.hwndFocus or 0)
        current = focused_hwnd
        while current:
            if current == root_hwnd:
                return focused_hwnd
            try:
                current = win32gui.GetParent(current)
            except Exception:
                return None
        return None

    def _best_editor_text(self, root_hwnd: int) -> str:
        best_text = ""
        focused = self._focused_child_handle(root_hwnd)
        if focused:
            focused_text = self._read_text(focused)
            if has_text_content(focused_text):
                return focused_text

        fallback = ""
        for hwnd in self._child_handles(root_hwnd):
            class_name = get_class_name(hwnd)
            text = self._read_text(hwnd)
            if any(hint.lower() in class_name.lower() for hint in EDITOR_CLASS_HINTS):
                if len(text) > len(best_text):
                    best_text = text
            elif len(text) > len(fallback):
                fallback = text
        return best_text or fallback

    def read_current_text(self) -> str:
        hwnd = get_foreground_hwnd()
        if not self._is_notepad(hwnd):
            return ""
        return self._best_editor_text(hwnd)


class BrowserReader(BasePollingReader):
    def __init__(self, debug: bool = False):
        super().__init__(debug=debug)
        self.desktop = None
        self.uia = None
        self._init_backend()

    def _init_backend(self):
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        try:
            from pywinauto import Desktop
            from pywinauto.uia_defines import IUIA

            self.desktop = Desktop(backend="uia")
            self.uia = IUIA()
        except Exception as exc:
            self._debug(f"pywinauto backend unavailable: {exc}")

    def _is_browser(self, hwnd: int) -> bool:
        if not hwnd:
            return False
        process_name = get_process_name(hwnd)
        title = get_window_title(hwnd).lower()
        return process_name in BROWSER_PROCESS_NAMES or any(
            name.replace(".exe", "") in title for name in BROWSER_PROCESS_NAMES
        )

    def _window_wrapper(self, hwnd: int):
        if self.desktop is None or not hwnd:
            return None
        try:
            return self.desktop.window(handle=hwnd).wrapper_object()
        except Exception as exc:
            self._debug(f"wrapper lookup failed: {exc}")
            return None

    def _focused_wrapper(self):
        if self.uia is None:
            return None
        try:
            from pywinauto.controls.uiawrapper import UIAWrapper
            from pywinauto.uia_element_info import UIAElementInfo

            element = self.uia.get_focused_element()
            return UIAWrapper(UIAElementInfo(element)) if element else None
        except Exception as exc:
            self._debug(f"focused wrapper lookup failed: {exc}")
            return None

    def _extract_text(self, wrapper) -> str:
        if wrapper is None:
            return ""
        readers: tuple[tuple[str, Callable[[], object]], ...] = (
            ("value", lambda: wrapper.iface_value.CurrentValue if wrapper.iface_value else ""),
            ("legacy", lambda: wrapper.legacy_properties().get("Value", "")),
            ("uia-text", lambda: wrapper.iface_text.DocumentRange.GetText(-1)
             if wrapper.iface_text and wrapper.iface_text.DocumentRange
             else ""),
            ("texts", lambda: self._join_control_texts(wrapper)),
            ("window", lambda: wrapper.window_text()),
        )
        candidates: list[tuple[str, str]] = []
        for name, reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            text = normalize_text(str(value)) if value is not None else ""
            if has_text_content(text):
                candidates.append((name, text))
        if not candidates:
            return ""

        for _name, text in candidates:
            if "\n\n" in text:
                return text
        return max((text for _name, text in candidates), key=len)

    def _join_control_texts(self, wrapper) -> str:
        try:
            values = wrapper.texts()
        except Exception:
            return ""
        lines = [normalize_text(str(value)) for value in values]
        if not lines:
            return ""
        while lines and not has_text_content(lines[0]):
            lines.pop(0)
        while lines and not has_text_content(lines[-1]):
            lines.pop()
        return "\n".join(lines)

    def _best_input_text(self, window) -> str:
        if window is None:
            return ""
        candidates = []
        focused = self._focused_wrapper()
        if focused is not None:
            candidates.append(("focused-uia", focused))
        try:
            focused = window.get_focus()
        except Exception:
            focused = None
        if focused is not None:
            candidates.append(("focused", focused))
        current = focused
        for depth in range(3):
            try:
                current = current.parent() if current else None
            except Exception:
                current = None
            if current is not None:
                candidates.append((f"focused-parent-{depth + 1}", current))

        seen = set()
        for source, wrapper in candidates:
            try:
                key = (getattr(wrapper, "handle", None), wrapper.element_info.control_type, wrapper.window_text())
                control_type = wrapper.element_info.control_type or ""
                title = wrapper.window_text() or ""
            except Exception:
                key = id(wrapper)
                control_type = ""
                title = ""
            if key in seen:
                continue
            seen.add(key)
            text = self._extract_text(wrapper)
            self._debug(f"browser candidate={source} type={control_type!r} title={title!r} len={len(text)}")
            if control_type not in INPUT_CONTROL_TYPES or not has_text_content(text):
                continue
            lowered = f"{title}\n{text}".lower()
            if looks_like_url(text) or any(hint in lowered for hint in ("address", "search google", "url", "\uc8fc\uc18c")):
                continue
            return text
        return ""

    def read_current_text(self) -> str:
        hwnd = get_foreground_hwnd()
        if not self._is_browser(hwnd):
            return ""
        return self._best_input_text(self._window_wrapper(hwnd))


class ActiveWordReader(BasePollingReader):
    MIXED_VALUES = (None, 9999999, -9999999, 9999998, -9999998)
    MAX_STYLE_CHARACTERS = 4000
    ENABLE_WORD_LIVE_STYLE_SEGMENTS = False

    def __init__(self, debug: bool = False):
        super().__init__(debug=debug)
        self._last_word_text = ""
        self._last_word_style_info: dict = {}

    def _active_document(self):
        if pythoncom is None:
            return None
        hwnd = get_foreground_hwnd()
        if get_process_name(hwnd) not in WORD_PROCESS_NAMES:
            return None
        pythoncom.CoInitialize()
        import win32com.client as win32

        word = win32.GetActiveObject("Word.Application")
        return getattr(word, "ActiveDocument", None)

    def read_current_text(self) -> str:
        try:
            document = self._active_document()
            if document is None:
                return ""
            text = self._read_paragraph_text(document)
            if text != self._last_word_text or not self._last_word_style_info:
                self._last_word_text = text
                self._last_word_style_info = self._read_style_info_from_document(document)
            return text
        except Exception as exc:
            self._debug(f"word read failed: {exc}")
            return ""

    def _read_paragraph_text(self, document) -> str:
        quick_text = self._clean_word_paragraph_text(getattr(document.Content, "Text", "") or "")
        paragraphs = getattr(document, "Paragraphs", None)
        try:
            paragraph_count = int(paragraphs.Count) if paragraphs is not None else 0
        except Exception:
            paragraph_count = 0
        if paragraph_count <= 0:
            return quick_text

        quick_lines = quick_text.split("\n") if quick_text else []
        if len(quick_lines) >= paragraph_count:
            return quick_text

        lines: list[str] = []
        for index in range(1, paragraph_count + 1):
            try:
                raw_text = getattr(paragraphs.Item(index).Range, "Text", "") or ""
            except Exception:
                continue
            lines.append(self._clean_word_paragraph_text(raw_text))
        return "\n".join(lines)

    def _clean_word_paragraph_text(self, text: str) -> str:
        normalized = normalize_text(text).replace("\x07", "")
        return normalized.rstrip("\n")

    def read_style_info(self) -> dict:
        return dict(self._last_word_style_info)

    def _read_style_info_from_document(self, document) -> dict:
        try:
            word_range = document.Content.Duplicate
            if word_range.End > word_range.Start:
                word_range.End = word_range.End - 1
            style_info = self._style_from_range(word_range)
            style_info["line_styles"] = self._read_line_styles(document)
            self._merge_word_openxml_underlines(document, style_info["line_styles"])
            self._merge_word_openxml_strikes(document, style_info["line_styles"])
            self._merge_word_openxml_vertical_aligns(document, style_info["line_styles"])
            self._merge_word_openxml_highlights(document, style_info["line_styles"])
            self._merge_word_openxml_colors(document, style_info["line_styles"])
            style_info["segments"] = self._read_word_openxml_style_segments(document)
            if not style_info["segments"] and self.ENABLE_WORD_LIVE_STYLE_SEGMENTS:
                style_info["segments"] = self._read_style_segments(word_range)
            self._log_word_style(style_info)
            return style_info
        except Exception as exc:
            self._debug(f"word style read failed: {exc}")
            return {}

    def _read_line_styles(self, document) -> list[dict]:
        try:
            content_range = document.Content.Duplicate
            if content_range.End > content_range.Start:
                content_range.End = content_range.End - 1
            characters = content_range.Characters
            count = min(int(characters.Count), self.MAX_STYLE_CHARACTERS)
        except Exception:
            return []

        line_styles: list[dict] = []
        content_line = 0
        line_index = 0
        line_chars: list[str] = []
        line_style: dict | None = None

        def append_line():
            nonlocal content_line, line_index, line_chars, line_style
            visible_text = "".join(line_chars)
            is_blank = not has_text_content(visible_text)
            clean_style = {key: value for key, value in (line_style or {}).items() if value is not None}
            item = {
                "line": line_index,
                "is_blank": is_blank,
                "style": clean_style,
            }
            if not is_blank:
                item["content_line"] = content_line
                content_line += 1
            line_styles.append(item)
            line_index += 1
            line_chars = []
            line_style = None

        for index in range(1, count + 1):
            try:
                char_range = characters.Item(index)
                normalized = normalize_text(getattr(char_range, "Text", "") or "").replace("\x07", "")
            except Exception:
                continue
            if not normalized:
                continue
            for character in normalized:
                if character == "\n":
                    append_line()
                    continue
                if line_style is None and character.strip():
                    line_style = self._style_from_range(char_range)
                line_chars.append(character)

        if line_chars or not line_styles:
            append_line()
        return line_styles

    def _first_visible_character_range(self, paragraph_range, visible_text: str):
        if not visible_text:
            return None
        try:
            characters = paragraph_range.Characters
            count = int(characters.Count)
        except Exception:
            return None
        for index in range(1, count + 1):
            try:
                char_range = characters.Item(index)
                char_text = normalize_text(getattr(char_range, "Text", "") or "").replace("\x07", "")
            except Exception:
                continue
            if char_text and char_text not in {"\n", "\r"}:
                return char_range
        return None

    def _merge_word_openxml_vertical_aligns(self, document, line_styles: list[dict]):
        if not line_styles:
            return
        align_styles = self._read_word_openxml_vertical_aligns(document)
        if not align_styles:
            return
        styles_by_line = {int(item.get("line", -1)): item.get("style") or {} for item in line_styles}
        for line_index, align_style in align_styles.items():
            target_style = styles_by_line.get(line_index)
            if target_style is not None:
                target_style.update(align_style)

    def _read_word_openxml_vertical_aligns(self, document) -> dict[int, dict]:
        try:
            root = ET.fromstring(getattr(document.Content, "WordOpenXML", "") or "")
        except Exception:
            return {}
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        result: dict[int, dict] = {}
        for line_index, runs in self._word_openxml_logical_lines(root, namespace):
            for run in runs:
                vertical_align = run.find("w:rPr/w:vertAlign", namespace)
                if vertical_align is None:
                    continue
                value = self._word_openxml_attr(vertical_align, "val")
                if value == "superscript":
                    result[line_index] = {"superscript": True, "subscript": False}
                    break
                if value == "subscript":
                    result[line_index] = {"subscript": True, "superscript": False}
                    break
        return result
    def _merge_word_openxml_highlights(self, document, line_styles: list[dict]):
        if not line_styles:
            return
        highlight_styles = self._read_word_openxml_highlights(document)
        if not highlight_styles:
            return
        styles_by_line = {int(item.get("line", -1)): item.get("style") or {} for item in line_styles}
        for line_index, highlight_style in highlight_styles.items():
            target_style = styles_by_line.get(line_index)
            if target_style is not None:
                target_style.update(highlight_style)

    def _merge_word_openxml_colors(self, document, line_styles: list[dict]):
        if not line_styles:
            return
        color_styles = self._read_word_openxml_colors(document)
        if not color_styles:
            return
        styles_by_line = {int(item.get("line", -1)): item.get("style") or {} for item in line_styles}
        for line_index, color_style in color_styles.items():
            target_style = styles_by_line.get(line_index)
            if target_style is not None:
                target_style.update(color_style)

    def _read_word_openxml_highlights(self, document) -> dict[int, dict]:
        try:
            root = ET.fromstring(getattr(document.Content, "WordOpenXML", "") or "")
        except Exception:
            return {}
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        result: dict[int, dict] = {}
        for line_index, runs in self._word_openxml_logical_lines(root, namespace):
            for run in runs:
                highlight = run.find("w:rPr/w:highlight", namespace)
                if highlight is None:
                    continue
                value = self._word_openxml_attr(highlight, "val")
                color_index = self._word_openxml_highlight_value(value)
                if color_index is not None:
                    result[line_index] = {"highlight_color_index": color_index}
                    break
        return result
    def _read_word_openxml_colors(self, document) -> dict[int, dict]:
        try:
            root = ET.fromstring(getattr(document.Content, "WordOpenXML", "") or "")
        except Exception:
            return {}
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        result: dict[int, dict] = {}
        for line_index, runs in self._word_openxml_logical_lines(root, namespace):
            for run in runs:
                color = run.find("w:rPr/w:color", namespace)
                if color is None:
                    continue
                value = self._word_openxml_attr(color, "val")
                if not value or str(value).lower() == "auto":
                    continue
                color_hex = self._word_openxml_color_hex(value)
                if color_hex:
                    result[line_index] = {"color_hex": color_hex}
                    break
        return result

    def _word_openxml_color_hex(self, value: str | None):
        if not value:
            return None
        value = str(value).strip().lstrip("#")
        if len(value) != 6:
            return None
        try:
            int(value, 16)
        except Exception:
            return None
        return f"#{value.upper()}"

    def _word_openxml_highlight_value(self, value: str | None):
        if not value or value == "none":
            return 0
        return {
            "black": 1,
            "blue": 2,
            "cyan": 3,
            "green": 4,
            "magenta": 5,
            "red": 6,
            "yellow": 7,
            "darkBlue": 9,
            "darkCyan": 10,
            "darkGreen": 11,
            "darkMagenta": 12,
            "darkRed": 13,
            "darkYellow": 14,
            "darkGray": 15,
            "lightGray": 16,
        }.get(str(value), None)

    def _merge_word_openxml_underlines(self, document, line_styles: list[dict]):
        if not line_styles:
            return
        underline_styles = self._read_word_openxml_underlines(document)
        if not underline_styles:
            return
        styles_by_line = {int(item.get("line", -1)): item.get("style") or {} for item in line_styles}
        for line_index, underline_style in underline_styles.items():
            target_style = styles_by_line.get(line_index)
            if target_style is None:
                continue
            target_style.update(underline_style)

    def _merge_word_openxml_strikes(self, document, line_styles: list[dict]):
        if not line_styles:
            return
        strike_styles = self._read_word_openxml_strikes(document)
        if not strike_styles:
            return
        styles_by_line = {int(item.get("line", -1)): item.get("style") or {} for item in line_styles}
        for line_index, strike_style in strike_styles.items():
            target_style = styles_by_line.get(line_index)
            if target_style is None:
                continue
            target_style.update(strike_style)

    def _read_word_openxml_strikes(self, document) -> dict[int, dict]:
        try:
            xml_text = getattr(document.Content, "WordOpenXML", "") or ""
            root = ET.fromstring(xml_text)
        except Exception as exc:
            self._debug(f"word openxml strike read failed: {exc}")
            return {}

        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        result: dict[int, dict] = {}
        for line_index, runs in self._word_openxml_logical_lines(root, namespace):
            strike_style: dict = {}
            for run in runs:
                run_props = run.find("w:rPr", namespace)
                if run_props is None:
                    continue
                strike = run_props.find("w:strike", namespace)
                double_strike = run_props.find("w:dstrike", namespace)
                if self._word_openxml_toggle_enabled(double_strike):
                    strike_style["double_strike_through"] = True
                    strike_style["strike_through"] = False
                    break
                if self._word_openxml_toggle_enabled(strike):
                    strike_style["strike_through"] = True
                    strike_style["double_strike_through"] = False
                    break
            if strike_style:
                result[line_index] = strike_style
        return result
    def _word_openxml_toggle_enabled(self, element) -> bool:
        if element is None:
            return False
        value = self._word_openxml_attr(element, "val")
        return value is None or str(value).lower() not in {"0", "false", "off", "none"}

    def _read_word_openxml_underlines(self, document) -> dict[int, dict]:
        try:
            xml_text = getattr(document.Content, "WordOpenXML", "") or ""
            root = ET.fromstring(xml_text)
        except Exception as exc:
            self._debug(f"word openxml underline read failed: {exc}")
            return {}

        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        result: dict[int, dict] = {}
        for line_index, runs in self._word_openxml_logical_lines(root, namespace):
            underline_style: dict = {}
            for run in runs:
                underline = run.find("w:rPr/w:u", namespace)
                if underline is None:
                    continue
                value = self._word_openxml_attr(underline, "val") or "single"
                if value == "none":
                    continue
                underline_value = self._word_openxml_underline_value(value)
                if underline_value is not None:
                    underline_style["underline"] = underline_value
                color = self._word_openxml_attr(underline, "color")
                if color and color.lower() != "auto":
                    underline_style["underline_color"] = None
                    underline_style["underline_color_hex"] = f"#{color.upper()}"
                break
            if underline_style:
                result[line_index] = underline_style
        return result

    def _word_openxml_logical_lines(self, root, namespace):
        body = root.find(".//w:body", namespace)
        paragraphs = body.findall("w:p", namespace) if body is not None else root.findall(".//w:body/w:p", namespace)
        line_index = 0
        for paragraph in paragraphs:
            current_runs = []
            for run in paragraph.findall("w:r", namespace):
                current_runs.append(run)
                breaks = run.findall("w:br", namespace)
                for _break in breaks:
                    yield line_index, list(current_runs)
                    line_index += 1
                    current_runs = []
            yield line_index, list(current_runs)
            line_index += 1
    def _word_openxml_attr(self, element, name: str):
        return element.get(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{name}")

    def _word_openxml_underline_value(self, value: str):
        return {
            "single": 1,
            "words": 2,
            "double": 3,
            "dotted": 4,
            "thick": 6,
            "dash": 7,
            "dotDash": 9,
            "dotDotDash": 10,
            "wave": 11,
            "dottedHeavy": 20,
            "dashHeavy": 23,
            "dotDashHeavy": 25,
            "dotDotDashHeavy": 26,
            "wavyHeavy": 27,
            "dashLong": 39,
            "wavyDouble": 43,
            "dashLongHeavy": 55,
        }.get(str(value), 1)

    def _read_word_openxml_style_segments(self, document) -> list[dict]:
        try:
            root = ET.fromstring(getattr(document.Content, "WordOpenXML", "") or "")
        except Exception as exc:
            self._debug(f"word openxml style segment read failed: {exc}")
            return []
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        body = root.find(".//w:body", namespace)
        paragraphs = body.findall("w:p", namespace) if body is not None else root.findall(".//w:body/w:p", namespace)
        segments: list[dict] = []
        text_index = 0
        line_chunks: list[tuple[int, int, dict]] = []

        def flush_line():
            nonlocal line_chunks
            signatures = {
                tuple(sorted((style or {}).items()))
                for start, end, style in line_chunks
                if end > start
            }
            if len(signatures) > 1:
                for start, end, style in line_chunks:
                    self._append_openxml_style_segment(segments, start, end, style)
            line_chunks = []

        for paragraph in paragraphs:
            for run in paragraph.findall("w:r", namespace):
                style = self._word_openxml_run_style(run, namespace)
                for part_text, is_break in self._word_openxml_run_text_parts(run, namespace):
                    if is_break:
                        flush_line()
                        text_index += 1
                        continue
                    if not part_text:
                        continue
                    start = text_index
                    text_index += len(part_text)
                    line_chunks.append((start, text_index, style))
                    if text_index > self.MAX_STYLE_CHARACTERS:
                        flush_line()
                        return segments
            flush_line()
            text_index += 1
            if text_index > self.MAX_STYLE_CHARACTERS:
                return segments
        return segments

    def _word_openxml_run_text_parts(self, run, namespace):
        for child in list(run):
            tag = self._word_openxml_local_name(child.tag)
            if tag == "t":
                yield child.text or "", False
            elif tag == "tab":
                yield "\t", False
            elif tag in {"br", "cr"}:
                yield "\n", True

    def _word_openxml_local_name(self, tag: str) -> str:
        return str(tag).rsplit("}", 1)[-1]

    def _word_openxml_run_style(self, run, namespace) -> dict:
        run_props = run.find("w:rPr", namespace)
        style = {
            "bold": False,
            "italic": False,
            "underline": 0,
            "strike_through": False,
            "double_strike_through": False,
            "subscript": False,
            "superscript": False,
            "highlight_color_index": 0,
            "color_hex": "#000000",
        }
        if run_props is None:
            return style

        style["bold"] = self._word_openxml_toggle_enabled(run_props.find("w:b", namespace))
        style["italic"] = self._word_openxml_toggle_enabled(run_props.find("w:i", namespace))

        underline = run_props.find("w:u", namespace)
        if underline is not None:
            underline_value = self._word_openxml_attr(underline, "val") or "single"
            style["underline"] = 0 if underline_value == "none" else self._word_openxml_underline_value(underline_value)
            underline_color = self._word_openxml_attr(underline, "color")
            underline_color_hex = self._word_openxml_color_hex(underline_color)
            if underline_color_hex:
                style["underline_color"] = None
                style["underline_color_hex"] = underline_color_hex

        double_strike = run_props.find("w:dstrike", namespace)
        strike = run_props.find("w:strike", namespace)
        if self._word_openxml_toggle_enabled(double_strike):
            style["double_strike_through"] = True
            style["strike_through"] = False
        elif self._word_openxml_toggle_enabled(strike):
            style["strike_through"] = True
            style["double_strike_through"] = False

        vertical_align = run_props.find("w:vertAlign", namespace)
        vertical_value = self._word_openxml_attr(vertical_align, "val") if vertical_align is not None else None
        if vertical_value == "subscript":
            style["subscript"] = True
            style["superscript"] = False
        elif vertical_value == "superscript":
            style["superscript"] = True
            style["subscript"] = False

        highlight = run_props.find("w:highlight", namespace)
        if highlight is not None:
            highlight_value = self._word_openxml_highlight_value(self._word_openxml_attr(highlight, "val"))
            if highlight_value is not None:
                style["highlight_color_index"] = highlight_value

        color = run_props.find("w:color", namespace)
        if color is not None:
            color_hex = self._word_openxml_color_hex(self._word_openxml_attr(color, "val"))
            if color_hex:
                style["color_hex"] = color_hex
        return style

    def _append_openxml_style_segment(self, segments: list[dict], start: int, end: int, style: dict):
        if end <= start:
            return
        clean_style = {key: value for key, value in style.items() if value is not None}
        if not clean_style:
            return
        signature = tuple(sorted(clean_style.items()))
        if segments:
            last = segments[-1]
            if last.get("end") == start and tuple(sorted((last.get("style") or {}).items())) == signature:
                last["end"] = end
                return
        segments.append({"start": start, "end": end, "style": clean_style})

    def _read_style_segments(self, word_range) -> list[dict]:
        try:
            characters = word_range.Characters
            count = int(characters.Count)
        except Exception:
            return []
        if count <= 0 or count > self.MAX_STYLE_CHARACTERS:
            return []

        segments: list[dict] = []
        current_style = None
        segment_start = 0
        text_index = 0

        for index in range(1, count + 1):
            try:
                char_range = characters.Item(index)
                raw_text = getattr(char_range, "Text", "") or ""
            except Exception:
                continue
            normalized = normalize_text(raw_text)
            if not normalized:
                continue

            style = self._style_from_range(char_range)
            style_signature = tuple(sorted(style.items()))
            char_len = len(normalized)

            if current_style is None:
                current_style = (style_signature, style)
                segment_start = text_index
            elif style_signature != current_style[0]:
                self._append_style_segment(segments, segment_start, text_index, current_style[1])
                current_style = (style_signature, style)
                segment_start = text_index
            text_index += char_len

        if current_style is not None:
            self._append_style_segment(segments, segment_start, text_index, current_style[1])
        return segments

    def _append_style_segment(self, segments: list[dict], start: int, end: int, style: dict):
        if end <= start:
            return
        clean_style = {key: value for key, value in style.items() if value is not None}
        if not clean_style:
            return
        segments.append({"start": start, "end": end, "style": clean_style})

    def _style_from_range(self, word_range) -> dict:
        font = word_range.Font
        return {
            "font_name": self._clean_mixed_value(getattr(font, "Name", None)),
            "font_size": self._clean_mixed_value(getattr(font, "Size", None)),
            "color_hex": self._word_color_to_hex(self._clean_mixed_value(getattr(font, "Color", None))),
            "bold": self._word_bool(getattr(font, "Bold", None)),
            "italic": self._word_bool(getattr(font, "Italic", None)),
            "underline": self._clean_mixed_value(getattr(font, "Underline", None)),
            "underline_color": self._clean_mixed_value(getattr(font, "UnderlineColor", None)),
            "underline_color_hex": self._word_color_to_hex(
                self._clean_mixed_value(getattr(font, "UnderlineColor", None))
            ),
            "strike_through": self._word_bool(getattr(font, "StrikeThrough", None)),
            "double_strike_through": self._word_bool(getattr(font, "DoubleStrikeThrough", None)),
            "subscript": self._word_bool(getattr(font, "Subscript", None)),
            "superscript": self._word_bool(getattr(font, "Superscript", None)),
            "highlight_color_index": self._clean_mixed_value(getattr(word_range, "HighlightColorIndex", None)),
        }

    def _log_word_style(self, style_info: dict):
        try:
            line_preview = []
            for item in (style_info.get("line_styles") or [])[:8]:
                style = item.get("style") or {}
                line_preview.append({
                    "line": item.get("line"),
                    "underline": style.get("underline"),
                    "underline_color": style.get("underline_color"),
                    "underline_color_hex": style.get("underline_color_hex"),
                    "bold": style.get("bold"),
                    "italic": style.get("italic"),
                    "strike": style.get("strike_through"),
                    "double_strike": style.get("double_strike_through"),
                    "subscript": style.get("subscript"),
                    "superscript": style.get("superscript"),
                    "highlight": style.get("highlight_color_index"),
                    "color_hex": style.get("color_hex"),
                })
            segment_preview = []
            for segment in (style_info.get("segments") or [])[:10]:
                segment_preview.append({
                    "start": segment.get("start"),
                    "end": segment.get("end"),
                    "style": segment.get("style"),
                })
            _WORD_STYLE_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with _WORD_STYLE_LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(
                    f"{time.strftime('%Y-%m-%d %H:%M:%S')} "
                    f"base_underline={style_info.get('underline')!r} "
                    f"base_underline_color={style_info.get('underline_color')!r} "
                    f"segments={len(style_info.get('segments') or [])} "
                    f"segment_preview={segment_preview!r} "
                    f"lines={line_preview!r}\n"
                )
        except Exception:
            pass

    def _clean_mixed_value(self, value):
        if value in self.MIXED_VALUES:
            return None
        return value

    def _word_bool(self, value):
        value = self._clean_mixed_value(value)
        if value is None:
            return None
        return bool(value)

    def _word_color_to_hex(self, value):
        value = self._clean_mixed_value(value)
        if value is None:
            return None
        try:
            number = int(value)
            red = number & 0xFF
            green = (number >> 8) & 0xFF
            blue = (number >> 16) & 0xFF
            return f"#{red:02X}{green:02X}{blue:02X}"
        except Exception:
            return None


class ActiveHwpReader(BasePollingReader):
    SCAN_RANGE_OPTIONS = (0x0077, 0x0000, 0x000F, 0x0001)
    MAX_SCAN_STEPS = 2000
    MAX_EMPTY_SCAN_STREAK = 30
    COM_FAILURE_COOLDOWN_SECONDS = 5.0
    REPEATED_LOG_SUPPRESS_SECONDS = 30.0
    MAX_HWP_STYLE_SEGMENT_CHARS = 500
    ENABLE_HWP_LIVE_STYLE_SEGMENTS = False

    def __init__(self, debug: bool = False):
        super().__init__(debug=debug)
        self.desktop = None
        self.uia = None
        self._last_read_method = ""
        self._last_uia_source = ""
        self._last_hwp_text = ""
        self._last_hwp_style_info: dict = {}
        self._last_com_failure_at = 0.0
        self._last_logged_window_signature: tuple[int, str, str, str] | None = None
        self._last_uia_log_signature: tuple[int, tuple[tuple[str, str, int], ...]] | None = None
        self._last_scan_log_signature: tuple[int, int, tuple[tuple[object, int], ...]] | None = None
        self._last_log_times: dict[str, float] = {}
        self._last_com_success_signature: tuple[str, str] | None = None
        self._init_uia_backend()

    def _init_uia_backend(self):
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        try:
            from pywinauto import Desktop
            from pywinauto.uia_defines import IUIA

            self.desktop = Desktop(backend="uia")
            self.uia = IUIA()
        except Exception as exc:
            self._log_hwp(f"hwp uia backend unavailable: {type(exc).__name__}: {exc}")

    def _active_hwp(self):
        if pythoncom is None:
            return None
        hwnd = get_foreground_hwnd()
        if not self._is_hwp_window(hwnd):
            return None
        if time.monotonic() - self._last_com_failure_at < self.COM_FAILURE_COOLDOWN_SECONDS:
            return None
        pythoncom.CoInitialize()

        hwp = self._get_hwp_object_from_native_om(hwnd)
        if hwp is not None:
            return hwp

        hwp = self._get_active_hwp_object()
        if hwp is not None:
            return hwp

        hwp = self._get_hwp_object_from_rot()
        if hwp is not None:
            return hwp

        self._last_com_failure_at = time.monotonic()
        return None

    def _get_active_hwp_object(self):
        import win32com.client as win32

        for progid in HWP_ACTIVE_PROGIDS:
            try:
                hwp = win32.GetActiveObject(progid)
                hwp = self._coerce_hwp_object(hwp)
                if hwp is not None:
                    self._log_hwp_com_success("GetActiveObject", progid)
                    return hwp
                self._log_hwp_com(f"GetActiveObject unusable progid={progid!r}")
            except Exception as exc:
                self._log_hwp_com(f"GetActiveObject failed progid={progid!r}: {type(exc).__name__}: {exc}")
        return None

    def _get_hwp_object_from_rot(self):
        try:
            rot = pythoncom.GetRunningObjectTable()
            enum_moniker = rot.EnumRunning()
        except Exception as exc:
            self._log_hwp_com(f"ROT open failed: {type(exc).__name__}: {exc}")
            return None

        try:
            bind_context = pythoncom.CreateBindCtx(0)
        except Exception as exc:
            self._log_hwp_com(f"ROT bind context failed: {type(exc).__name__}: {exc}")
            return None

        while True:
            monikers = enum_moniker.Next(1)
            if not monikers:
                break
            moniker = monikers[0]
            try:
                display_name = moniker.GetDisplayName(bind_context, None)
            except Exception:
                display_name = ""
            lowered = str(display_name).lower()
            if "hwp" not in lowered and "hancom" not in lowered and "hword" not in lowered:
                continue
            self._log_hwp_com(f"ROT entry={display_name!r}")
            try:
                obj = rot.GetObject(moniker)
                hwp = self._coerce_hwp_object(obj)
                if hwp is not None:
                    self._log_hwp_com_success("ROT", str(display_name))
                    return hwp
                self._log_hwp_com(f"ROT object unusable entry={display_name!r} type={type(obj)}")
            except Exception as exc:
                self._log_hwp_com(f"ROT object failed entry={display_name!r}: {type(exc).__name__}: {exc}")
        return None

    def _get_hwp_object_from_native_om(self, hwnd: int):
        if not hwnd:
            return None
        try:
            import win32com.client as win32
            from ctypes import POINTER, byref, c_long, c_void_p
            from ctypes.wintypes import HWND

            oleacc = ctypes.oledll.oleacc
            iid_buffer = ctypes.create_string_buffer(bytes(pythoncom.IID_IDispatch))
            pdisp = c_void_p()
            accessible_object_from_window = oleacc.AccessibleObjectFromWindow
            accessible_object_from_window.argtypes = [HWND, c_long, c_void_p, POINTER(c_void_p)]
            accessible_object_from_window.restype = c_long
            result = accessible_object_from_window(
                HWND(hwnd),
                c_long(-16),  # OBJID_NATIVEOM
                ctypes.cast(iid_buffer, c_void_p),
                byref(pdisp),
            )
            if result != 0 or not pdisp.value:
                self._log_hwp_com(f"NativeOM failed hwnd={hwnd} result={result} pdisp={pdisp.value}")
                return None
            obj = pythoncom.ObjectFromAddress(pdisp.value, pythoncom.IID_IDispatch)
            hwp = self._coerce_hwp_object(win32.Dispatch(obj))
            if hwp is not None:
                self._log_hwp_com_success("NativeOM", str(hwnd))
                return hwp
            self._log_hwp_com(f"NativeOM unusable hwnd={hwnd}")
            return None
        except Exception as exc:
            self._log_hwp_com(f"NativeOM exception hwnd={hwnd}: {type(exc).__name__}: {exc}")
            return None

    def _coerce_hwp_object(self, obj):
        if obj is None:
            return None
        required = ("InitScan", "GetText", "ReleaseScan", "HAction", "HParameterSet")
        for candidate in self._hwp_dispatch_candidates(obj):
            if all(hasattr(candidate, name) for name in required):
                return candidate
        return None

    def _hwp_dispatch_candidates(self, obj):
        try:
            import win32com.client as win32
        except Exception:
            return []

        candidates = [obj]
        try:
            candidates.append(win32.Dispatch(obj))
        except Exception:
            pass
        try:
            candidates.append(win32.dynamic.Dispatch(obj))
        except Exception:
            pass

        for source in (obj, getattr(obj, "_oleobj_", None)):
            if source is None:
                continue
            query = getattr(source, "QueryInterface", None)
            if not callable(query):
                continue
            for iid in self._hwp_query_interface_iids():
                try:
                    candidates.append(win32.Dispatch(query(iid)))
                except Exception:
                    pass
                try:
                    candidates.append(win32.dynamic.Dispatch(query(iid)))
                except Exception:
                    pass

        wrapped_candidates = []
        for candidate in candidates:
            wrapped_candidates.append(candidate)
            try:
                wrapped_candidates.append(win32.CastTo(candidate, "IHwpObject"))
            except Exception:
                pass
        return wrapped_candidates

    def _hwp_query_interface_iids(self):
        iids = []
        try:
            from pywintypes import IID

            iids.append(IID(HWP_IHWP_OBJECT_IID))
        except Exception:
            pass
        if pythoncom is not None:
            try:
                iids.append(pythoncom.IID_IDispatch)
            except Exception:
                pass
        return iids

    def _is_hwp_window(self, hwnd: int) -> bool:
        process_name = get_process_name(hwnd)
        if process_name in HWP_PROCESS_NAMES or process_name.startswith("hwp"):
            return True
        if process_name:
            return False
        title = get_window_title(hwnd).lower()
        class_name = get_class_name(hwnd).lower()
        return any(hint in class_name for hint in HWP_CLASS_HINTS) or any(
            hint.lower() in title for hint in HWP_TITLE_HINTS
        )

    def read_current_text(self) -> str:
        try:
            hwnd = get_foreground_hwnd()
            if not self._is_hwp_window(hwnd):
                return ""
            self._log_hwp_window(hwnd)
            hwp = self._active_hwp()
            text = self._read_hwp_selected_plain_text(hwp) if hwp is not None else ""
            if has_text_content(text):
                self._last_read_method = "com"
                self._last_uia_source = ""
                self._last_hwp_text = text
                self._last_hwp_style_info = self._read_hwp_style_info(hwp)
                return text
            if hwp is not None:
                self._last_hwp_text = ""
                self._last_hwp_style_info = {}
                self._last_read_method = ""
                return ""
            text = self._read_text_via_uia(hwnd)
            if has_text_content(text):
                self._last_read_method = "uia"
                self._last_hwp_text = text
                self._last_hwp_style_info = {}
            return text
        except Exception as exc:
            self._debug(f"hwp read failed: {exc}")
            self._log_hwp(f"hwp read failed: {type(exc).__name__}: {exc}")
            return ""

    def read_style_info(self) -> dict:
        style_info = dict(self._last_hwp_style_info)
        if (
            self.ENABLE_HWP_LIVE_STYLE_SEGMENTS
            and
            self._last_read_method == "com"
            and self._last_hwp_text
            and style_info.get("hwp_style_scope") == "mixed_or_unknown"
        ):
            hwp = self._active_hwp()
            if hwp is not None:
                style_info["segments"] = self._read_hwp_style_segments(hwp, self._last_hwp_text)
        style_info.update({
            "read_method": self._last_read_method,
            "uia_source": self._last_uia_source,
            "_source_text": self._last_hwp_text if self._last_read_method == "com" else "",
        })
        return style_info

    def _read_hwp_selected_plain_text(self, hwp) -> str:
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            return ""
        for option in ("saveblock",):
            try:
                text = normalize_text(str(getter("TEXT", option) or "")).strip()
            except Exception as exc:
                self._log_hwp(f"HWP selected TEXT read failed option={option!r}: {type(exc).__name__}: {exc}")
                continue
            if has_text_content(text):
                self._log_hwp(f"HWP selected TEXT option={option!r} length={len(text)} preview={text[:120]!r}")
                return text
        return ""

    def _scan_text(self, hwp) -> str:
        for range_option in self.SCAN_RANGE_OPTIONS:
            text = self._scan_text_once(hwp, range_option)
            if not has_text_content(text):
                self._log_hwp(f"InitScan range=0x{range_option:04X} length={len(text)} preview={text[:120]!r}")
            if has_text_content(text):
                return text
        return ""

    def _scan_text_once(self, hwp, range_option: int) -> str:
        parts: list[str] = []
        try:
            hwp.InitScan(0x000F, range_option, 0, 0, -1, -1)
        except Exception as exc:
            self._log_hwp(f"InitScan range=0x{range_option:04X} failed: {type(exc).__name__}: {exc}")
            return ""
        try:
            empty_streak = 0
            steps = 0
            state_counts: dict[object, int] = {}
            while True:
                try:
                    state, text = hwp.GetText()
                except Exception as exc:
                    self._log_hwp(f"GetText range=0x{range_option:04X} failed: {type(exc).__name__}: {exc}")
                    break
                steps += 1
                state_counts[state] = state_counts.get(state, 0) + 1
                if text:
                    parts.append(text)
                    empty_streak = 0
                else:
                    empty_streak += 1
                if state in (0, 1):
                    break
                if state not in (2, 3) and not text:
                    self._log_hwp(f"GetText range=0x{range_option:04X} stopped on state={state!r}")
                    break
                if empty_streak >= self.MAX_EMPTY_SCAN_STREAK:
                    self._log_hwp(f"GetText range=0x{range_option:04X} stopped on empty streak")
                    break
                if steps >= self.MAX_SCAN_STEPS:
                    self._log_hwp(f"GetText range=0x{range_option:04X} stopped on max steps")
                    break
        finally:
            try:
                hwp.ReleaseScan()
            except Exception:
                pass
        text = normalize_text("".join(parts))
        self._log_scan_summary(range_option, steps, state_counts, text)
        return text

    def _read_hwp_style_info(self, hwp) -> dict:
        try:
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            char_shape = hwp.HParameterSet.HCharShape
        except Exception as exc:
            self._log_hwp(f"HWP CharShape read failed: {type(exc).__name__}: {exc}")
            return {}

        style_info = {
            "font_name": self._first_hwp_value(
                char_shape,
                (
                    "FaceNameHangul",
                    "FaceNameLatin",
                    "FaceNameHanja",
                    "FaceNameJapanese",
                    "FaceNameOther",
                    "FaceNameSymbol",
                    "FaceNameUser",
                ),
            ),
            "font_size": self._hwp_height_to_points(self._hwp_attr(char_shape, "Height")),
            "color": self._hwp_attr(char_shape, "TextColor"),
            "bold": self._hwp_bool(self._hwp_attr(char_shape, "Bold")),
            "italic": self._hwp_bool(self._hwp_attr(char_shape, "Italic")),
            "underline_type": self._hwp_attr(char_shape, "UnderlineType"),
            "underline_shape": self._hwp_attr(char_shape, "UnderlineShape"),
            "underline_color": self._hwp_attr(char_shape, "UnderlineColor"),
        }
        clean_style = {key: value for key, value in style_info.items() if value is not None}
        selection_pos = self._read_hwp_selection_position(hwp)
        if selection_pos:
            clean_style["hwp_selection_pos"] = selection_pos
        clean_style["hwp_style_scope"] = "basic" if (
            clean_style.get("font_name") and clean_style.get("font_size") is not None
        ) else "mixed_or_unknown"
        self._log_hwp(f"HWP CharShape style={clean_style!r}")
        return clean_style

    def _read_hwp_selection_position(self, hwp) -> dict:
        getter = getattr(hwp, "GetSelectedPos", None)
        if not callable(getter):
            return {}
        try:
            value = getter()
        except Exception as exc:
            self._log_hwp(f"HWP GetSelectedPos failed: {type(exc).__name__}: {exc}")
            return {}
        parsed = self._parse_hwp_selected_pos(value)
        self._log_hwp(f"HWP GetSelectedPos value={value!r} parsed={parsed!r}")
        return parsed

    def _parse_hwp_selected_pos(self, value) -> dict:
        if not isinstance(value, tuple) or len(value) < 7 or not bool(value[0]):
            return {}
        try:
            return {
                "start_list": int(value[1]),
                "start_para": int(value[2]),
                "start_pos": int(value[3]),
                "end_list": int(value[4]),
                "end_para": int(value[5]),
                "end_pos": int(value[6]),
                "raw": [item for item in value],
            }
        except Exception:
            return {"raw": [item for item in value]}

    def _read_hwp_style_segments(self, hwp, text: str) -> list[dict]:
        if len(text) > self.MAX_HWP_STYLE_SEGMENT_CHARS:
            self._log_hwp(f"HWP style segments skipped length={len(text)}")
            return []
        original_pos = self._get_hwp_pos(hwp)
        segments: list[dict] = []
        current_signature = None
        current_style = None
        segment_start = 0
        text_index = 0
        try:
            for char in text:
                char_len = len(char)
                if char == "\n":
                    text_index += char_len
                    continue
                start = self._hwp_text_index_to_position(text, text_index)
                end = self._hwp_text_index_to_position(text, text_index + char_len)
                style = self._style_for_hwp_range(hwp, start, end)
                signature = tuple(sorted(style.items()))
                if current_signature is None:
                    current_signature = signature
                    current_style = style
                    segment_start = text_index
                elif signature != current_signature:
                    self._append_hwp_style_segment(segments, segment_start, text_index, current_style)
                    current_signature = signature
                    current_style = style
                    segment_start = text_index
                text_index += char_len
            if current_signature is not None:
                self._append_hwp_style_segment(segments, segment_start, text_index, current_style)
        except Exception as exc:
            self._log_hwp(f"HWP style segments failed: {type(exc).__name__}: {exc}")
            return []
        finally:
            self._restore_hwp_pos(hwp, original_pos)
        self._log_hwp(f"HWP style segments count={len(segments)}")
        return segments

    def _style_for_hwp_range(self, hwp, start: tuple[int, int], end: tuple[int, int]) -> dict:
        try:
            hwp.SelectText(start[0], start[1], end[0], end[1])
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            char_shape = hwp.HParameterSet.HCharShape
        except Exception:
            return {}
        return self._sanitize_hwp_segment_style(
            {
                "font_name": self._first_hwp_value(
                    char_shape,
                    (
                        "FaceNameHangul",
                        "FaceNameLatin",
                        "FaceNameHanja",
                        "FaceNameJapanese",
                        "FaceNameOther",
                        "FaceNameSymbol",
                        "FaceNameUser",
                    ),
                ),
                "font_size": self._hwp_height_to_points(self._hwp_attr(char_shape, "Height")),
                "color": self._hwp_attr(char_shape, "TextColor"),
                "bold": self._hwp_bool(self._hwp_attr(char_shape, "Bold")),
                "italic": self._hwp_bool(self._hwp_attr(char_shape, "Italic")),
                "underline_type": self._hwp_attr(char_shape, "UnderlineType"),
                "underline_shape": self._hwp_attr(char_shape, "UnderlineShape"),
                "underline_color": self._hwp_attr(char_shape, "UnderlineColor"),
            }
        )

    def _sanitize_hwp_segment_style(self, style: dict) -> dict:
        clean = {key: value for key, value in style.items() if value is not None}
        if clean.get("font_size") is None:
            clean.pop("font_size", None)
        return clean

    def _append_hwp_style_segment(self, segments: list[dict], start: int, end: int, style: dict | None):
        if end <= start or not style:
            return
        segments.append({"start": start, "end": end, "style": dict(style)})

    def _hwp_text_index_to_position(self, text: str, index: int) -> tuple[int, int]:
        para = 0
        pos = 0
        for char in text[:index]:
            if char == "\n":
                para += 1
                pos = 0
            else:
                pos += 1
        return para, pos

    def _get_hwp_pos(self, hwp):
        try:
            value = hwp.GetPos()
        except Exception:
            return None
        if isinstance(value, tuple):
            numbers = [item for item in value if isinstance(item, int)]
            if len(numbers) >= 3:
                return tuple(numbers[-3:])
        return None

    def _restore_hwp_pos(self, hwp, pos):
        try:
            hwp.Run("Cancel")
        except Exception:
            pass
        if not pos:
            return
        try:
            hwp.SetPos(pos[0], pos[1], pos[2])
        except Exception:
            pass

    def _first_hwp_value(self, obj, names: tuple[str, ...]):
        for name in names:
            value = self._hwp_attr(obj, name)
            if value not in (None, ""):
                return value
        return None

    def _hwp_attr(self, obj, name: str):
        try:
            return getattr(obj, name)
        except Exception:
            return None

    def _hwp_height_to_points(self, value):
        if value is None:
            return None
        try:
            points = float(value) / 100
        except Exception:
            return None
        if points < 4 or points > 200:
            return None
        return points

    def _hwp_bool(self, value):
        if value is None:
            return None
        try:
            return bool(int(value))
        except Exception:
            return bool(value)

    def _read_text_via_uia(self, hwnd: int) -> str:
        candidates: list[tuple[str, str, str, str]] = []
        for source_name, wrapper in self._uia_candidate_wrappers(hwnd):
            descriptor = self._describe_uia_wrapper(wrapper)
            text = self._extract_uia_text(wrapper)
            if not has_text_content(text):
                continue
            candidates.append((source_name, descriptor[0], descriptor[1], text))

        filtered = self._filter_uia_candidates(candidates, get_window_title(hwnd))
        self._log_uia_candidates(hwnd, filtered or candidates)
        if not filtered:
            return ""
        combined_text = self._combine_hwp_uia_text_candidates(filtered)
        if combined_text:
            self._last_uia_source = "combined_uia"
            return combined_text
        source_name, _control_type, _title, text = max(filtered, key=lambda item: len(item[3]))
        self._last_uia_source = source_name
        return text

    def _combine_hwp_uia_text_candidates(self, candidates: list[tuple[str, str, str, str]]) -> str:
        leaf_rows = [
            (source_name, text)
            for source_name, control_type, _title, text in candidates
            if source_name.startswith("descendant-") and control_type in {"Edit", "Text"}
        ]
        if len(leaf_rows) < 2:
            return ""
        lines: list[str] = []
        seen: set[str] = set()
        for _source_name, text in leaf_rows:
            normalized = normalize_text(text).strip()
            if not normalized or normalized in seen or self._is_hwp_uia_noise_line(normalized):
                continue
            seen.add(normalized)
            lines.append(normalized)
        return "\n".join(lines)

    def _is_hwp_uia_noise_line(self, text: str) -> bool:
        value = normalize_text(text).strip()
        if not value:
            return True
        if value.upper() in {"CAPS", "NUM", "SCRL"}:
            return True
        if re.fullmatch(r"바탕글(?:\s+사본\d*)?", value):
            return True
        if value in {"굴림", "돋움", "바탕", "궁서", "함초롬바탕", "함초롬돋움", "한컴바탕"}:
            return True
        return False

    def _uia_candidate_wrappers(self, hwnd: int) -> list[tuple[str, object]]:
        wrappers: list[tuple[str, object]] = []
        window = self._window_wrapper(hwnd)
        focused = self._focused_wrapper()
        if focused is not None:
            wrappers.append(("focused", focused))

        current = focused
        for depth in range(4):
            try:
                current = current.parent() if current else None
            except Exception:
                current = None
            if current is not None:
                wrappers.append((f"focused-parent-{depth + 1}", current))

        if window is not None:
            wrappers.append(("window", window))
            wrappers.extend(self._descendant_wrappers(window, max_depth=5, max_nodes=120))

        unique: list[tuple[str, object]] = []
        seen = set()
        for source_name, wrapper in wrappers:
            key = self._wrapper_identity(wrapper)
            if key in seen:
                continue
            seen.add(key)
            unique.append((source_name, wrapper))
        return unique

    def _window_wrapper(self, hwnd: int):
        if self.desktop is None or not hwnd:
            return None
        try:
            return self.desktop.window(handle=hwnd).wrapper_object()
        except Exception as exc:
            self._log_hwp(f"hwp uia window lookup failed: {type(exc).__name__}: {exc}")
            return None

    def _focused_wrapper(self):
        if self.uia is None:
            return None
        try:
            from pywinauto.controls.uiawrapper import UIAWrapper
            from pywinauto.uia_element_info import UIAElementInfo

            element = self.uia.get_focused_element()
            return UIAWrapper(UIAElementInfo(element)) if element else None
        except Exception as exc:
            self._log_hwp(f"hwp uia focused lookup failed: {type(exc).__name__}: {exc}")
            return None

    def _descendant_wrappers(self, root, max_depth: int, max_nodes: int) -> list[tuple[str, object]]:
        results: list[tuple[str, object]] = []
        queue: list[tuple[object, int]] = [(root, 0)]
        seen = set()
        visited = 0
        while queue and visited < max_nodes:
            current, depth = queue.pop(0)
            key = self._wrapper_identity(current)
            if key in seen:
                continue
            seen.add(key)
            visited += 1

            control_type, _title, class_name = self._describe_uia_wrapper(current)
            if control_type in HWP_TEXT_CONTROL_TYPES or any(
                hint in class_name.lower() for hint in HWP_CLASS_HINTS
            ):
                results.append((f"descendant-{depth}", current))

            if depth >= max_depth:
                continue
            try:
                children = current.children()
            except Exception:
                children = []
            for child in children:
                queue.append((child, depth + 1))
        return results

    def _extract_uia_text(self, wrapper) -> str:
        if wrapper is None:
            return ""
        readers: tuple[tuple[str, Callable[[], object]], ...] = (
            ("value", lambda: wrapper.iface_value.CurrentValue if wrapper.iface_value else ""),
            ("legacy", lambda: wrapper.legacy_properties().get("Value", "")),
            ("legacy-name", lambda: wrapper.legacy_properties().get("Name", "")),
            ("uia-text", lambda: wrapper.iface_text.DocumentRange.GetText(-1)
             if wrapper.iface_text and wrapper.iface_text.DocumentRange
             else ""),
            ("texts", lambda: self._join_uia_texts(wrapper)),
            ("window", lambda: wrapper.window_text()),
        )
        values: list[str] = []
        for _name, reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            text = normalize_text(str(value)) if value is not None else ""
            if has_text_content(text):
                values.append(text)
        if not values:
            return ""
        return max(values, key=len)

    def _join_uia_texts(self, wrapper) -> str:
        try:
            values = wrapper.texts()
        except Exception:
            return ""
        lines = [normalize_text(str(value)) for value in values]
        lines = [line for line in lines if has_text_content(line)]
        return "\n".join(lines)

    def _filter_uia_candidates(
        self,
        candidates: list[tuple[str, str, str, str]],
        window_title: str,
    ) -> list[tuple[str, str, str, str]]:
        cleaned_title = normalize_text(window_title).strip().lower()
        filtered: list[tuple[str, str, str, str]] = []
        for source_name, control_type, title, text in candidates:
            normalized = normalize_text(text).strip()
            lowered = f"{title}\n{normalized}".lower()
            if not normalized:
                continue
            if cleaned_title and normalized.lower() == cleaned_title:
                continue
            if cleaned_title and cleaned_title in normalized.lower() and len(normalized) <= len(cleaned_title) + 12:
                continue
            if any(hint in lowered for hint in HWP_EXCLUDED_TEXT_HINTS):
                continue
            if len(normalized) < 2:
                continue
            filtered.append((source_name, control_type, title, normalized))
        return filtered

    def _describe_uia_wrapper(self, wrapper) -> tuple[str, str, str]:
        try:
            element_info = wrapper.element_info
            control_type = element_info.control_type or ""
            class_name = element_info.class_name or ""
        except Exception:
            control_type = ""
            class_name = ""
        try:
            title = wrapper.window_text() or ""
        except Exception:
            title = ""
        return control_type, normalize_text(title), class_name

    def _wrapper_identity(self, wrapper):
        try:
            info = wrapper.element_info
            return (
                getattr(wrapper, "handle", None),
                info.control_type,
                info.automation_id,
                info.name,
                info.class_name,
            )
        except Exception:
            return id(wrapper)

    def _log_uia_candidates(self, hwnd: int, candidates: list[tuple[str, str, str, str]]):
        summary = tuple(
            (source, control_type, len(text))
            for source, control_type, _title, text in candidates[:12]
        )
        signature = (hwnd, summary)
        if signature == self._last_uia_log_signature:
            return
        self._last_uia_log_signature = signature
        self._log_hwp(f"uia candidates count={len(candidates)}")
        for source_name, control_type, title, text in candidates[:12]:
            preview = text[:160].replace("\n", "\\n")
            self._log_hwp(
                f"uia candidate source={source_name!r} type={control_type!r} "
                f"title={title[:80]!r} length={len(text)} preview={preview!r}"
            )


    def _log_hwp_window(self, hwnd: int):
        signature = (
            hwnd,
            get_process_name(hwnd),
            get_class_name(hwnd),
            get_window_title(hwnd),
        )
        if signature == self._last_logged_window_signature:
            return
        self._last_logged_window_signature = signature
        self._log_hwp(
            "foreground "
            f"hwnd={signature[0]} process={signature[1]!r} "
            f"class={signature[2]!r} title={signature[3]!r}"
        )

    def _log_scan_summary(self, range_option: int, steps: int, state_counts: dict[object, int], text: str):
        signature = (
            range_option,
            len(text),
            tuple(sorted(state_counts.items(), key=lambda item: str(item[0]))),
        )
        if signature == self._last_scan_log_signature:
            return
        self._last_scan_log_signature = signature
        self._log_hwp(
            f"GetText range=0x{range_option:04X} summary "
            f"steps={steps} states={state_counts!r} length={len(text)} preview={text[:120]!r}"
        )

    def _log_hwp_com_success(self, source: str, detail: str):
        signature = (source, detail)
        if signature == self._last_com_success_signature:
            return
        self._last_com_success_signature = signature
        self._log_hwp_com(f"{source} object usable detail={detail!r}")

    def _log_hwp(self, message: str):
        if self._suppress_repeated_log(f"hwp:{message}"):
            return
        try:
            with _HWP_LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
        except Exception:
            pass

    def _log_hwp_com(self, message: str):
        if self._suppress_repeated_log(f"com:{message}"):
            return
        try:
            with _HWP_COM_LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
        except Exception:
            pass

    def _suppress_repeated_log(self, key: str) -> bool:
        now = time.monotonic()
        last = self._last_log_times.get(key)
        if last is not None and now - last < self.REPEATED_LOG_SUPPRESS_SECONDS:
            return True
        self._last_log_times[key] = now
        return False


@dataclass
class ReaderSnapshot:
    source: str
    window_title: str
    text: str
    reader_name: str
    window_handle: int | None = None
    style_info: dict | None = None


class UniversalActiveTextReader:
    def __init__(self, debug: bool = False):
        self.debug = debug
        self.readers = {
            "word": ActiveWordReader(debug=debug),
            "hwp": ActiveHwpReader(debug=debug),
            "browser": BrowserReader(debug=debug),
            "notepad": NotepadReader(debug=debug),
        }
        self.last_signature: tuple[str, str, str] | None = None
        self.current_process_id = os.getpid()
        self.last_active_reader_name = ""
        self.last_active_window_handle: int | None = None
        self.last_active_window_title = ""

    def _closed_snapshot_if_needed(self) -> ReaderSnapshot | None:
        if self.last_active_reader_name not in {"word", "hwp", "browser", "notepad"}:
            return None
        if self._is_live_window(self.last_active_window_handle):
            return None
        reader_name = f"{self.last_active_reader_name}_closed"
        window_title = self.last_active_window_title
        self.last_signature = None
        self.last_active_reader_name = ""
        self.last_active_window_handle = None
        self.last_active_window_title = ""
        return ReaderSnapshot(
            source="realtime",
            window_title=window_title,
            text="",
            reader_name=reader_name,
            window_handle=None,
            style_info={},
        )

    def _is_live_window(self, hwnd: int | None) -> bool:
        if win32gui is None or not hwnd:
            return False
        try:
            return bool(win32gui.IsWindow(hwnd))
        except Exception:
            return False

    def _remember_active_snapshot(self, reader_name: str, hwnd: int, window_title: str):
        self.last_active_reader_name = reader_name
        self.last_active_window_handle = hwnd
        self.last_active_window_title = window_title

    def _reader_order_for_foreground(self) -> list[tuple[str, BasePollingReader]]:
        hwnd = get_foreground_hwnd()
        process_name = get_process_name(hwnd)
        if process_name in WORD_PROCESS_NAMES:
            return [("word", self.readers["word"])]
        if process_name in HWP_PROCESS_NAMES:
            return [("hwp", self.readers["hwp"])]
        if process_name in BROWSER_PROCESS_NAMES:
            return [("browser", self.readers["browser"])]
        if process_name in NOTEPAD_PROCESS_NAMES:
            return [("notepad", self.readers["notepad"])]
        return [
            ("browser", self.readers["browser"]),
            ("notepad", self.readers["notepad"]),
            ("word", self.readers["word"]),
            ("hwp", self.readers["hwp"]),
        ]

    def _read_style_info(self, reader: BasePollingReader) -> dict:
        style_reader = getattr(reader, "read_style_info", None)
        if not callable(style_reader):
            return {}
        try:
            return style_reader() or {}
        except Exception:
            return {}

    def _is_own_window(self, hwnd: int) -> bool:
        if not hwnd:
            return False
        process_id = get_process_id(hwnd)
        if process_id and process_id == self.current_process_id:
            return True
        title = get_window_title(hwnd).lower()
        class_name = get_class_name(hwnd).lower()
        return "writing assistant" in title or "writing assistant" in class_name

    def poll_snapshot(self) -> ReaderSnapshot | None:
        closed_snapshot = self._closed_snapshot_if_needed()
        if closed_snapshot is not None:
            return closed_snapshot

        hwnd = get_foreground_hwnd()
        if self._is_own_window(hwnd):
            return None

        window_title = get_window_title(hwnd)
        process_name = get_process_name(hwnd)
        for reader_name, reader in self._reader_order_for_foreground():
            text = normalize_text(reader.read_current_text())
            if not has_text_content(text):
                continue
            signature = (reader_name, window_title, text)
            if signature == self.last_signature:
                return None
            style_info = self._read_style_info(reader)
            self.last_signature = signature
            self._remember_active_snapshot(reader_name, hwnd, window_title)
            return ReaderSnapshot(
                source="realtime",
                window_title=window_title,
                text=text,
                reader_name=reader_name,
                window_handle=hwnd,
                style_info=style_info,
            )

        if process_name in BROWSER_PROCESS_NAMES:
            return None

        empty_signature = ("", window_title, "")
        if window_title and empty_signature != self.last_signature:
            self.last_signature = empty_signature
            return ReaderSnapshot(
                source="realtime",
                window_title=window_title,
                text="",
                reader_name="unavailable",
                window_handle=hwnd,
            )
        return None





