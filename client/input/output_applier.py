from __future__ import annotations

import ctypes
import html
import os
from dataclasses import dataclass
from pathlib import Path
import tempfile
import time
import zipfile
import shutil

import pyperclip

from client.input.hwpx_document import (
    create_hwpx_fragment_from_match,
    extract_hwpx_text,
    find_hwpx_text_matches,
    has_hwpx_position_fallback_match,
)
from client.input.hwpml_document import create_hwpml_fragment_from_selection, extract_hwpml_text

try:
    import pythoncom
except Exception:  # pragma: no cover - optional Windows dependency
    pythoncom = None

try:
    import win32gui
except Exception:  # pragma: no cover - optional Windows dependency
    win32gui = None

try:
    import win32process
except Exception:  # pragma: no cover - optional Windows dependency
    win32process = None

try:
    import psutil
except Exception:  # pragma: no cover - optional Windows dependency
    psutil = None

HWP_PROCESS_NAMES = {"hwp.exe", "hwp64.exe", "hwpviewer.exe", "hwpw.exe"}
HWP_ACTIVE_PROGIDS = (
    "HWPFrame.HwpObject.2",
    "HWPFrame.HwpObject.1",
    "HWPFrame.HwpObject",
)
HWP_IHWP_OBJECT_IID = "{5E6A8276-CF1C-42B8-BCED-319548B02AF6}"
HWP_TEXTFILE_FORMATS = (
    "HTML",
    "HWPML",
    "HWPML2X",
    "HWPML2X_S",
    "HWPML2X_P",
    "HWPML2X_STYLE",
)
HWP_TEXTFILE_OPTIONS = ("", "saveblock", "selection")
ENABLE_HWP_CURSOR_SEGMENT_SELECTION = False
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
_LOG_DIR = Path(__file__).resolve().parents[2] / ".logs"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_HWP_REPLACE_LOG_PATH = _LOG_DIR / "hwp_replace.log"
_WORD_REPLACE_LOG_PATH = _LOG_DIR / "word_replace.log"
_HWP_TEXTFILE_SNAPSHOT_DIR = _LOG_DIR / "hwp_textfile_snapshots"


@dataclass
class OutputTarget:
    mode: str
    window_handle: int | None = None
    window_title: str = ""
    style_info: dict | None = None


class OutputApplier:
    def inspect_replace_availability(self, target: OutputTarget | None) -> tuple[bool, str | None]:
        if target is None:
            return False, "No source window has been captured yet."
        if target.mode == "browser_extension":
            session_id = (target.style_info or {}).get("browser_session_id")
            if session_id:
                return True, None
            return False, "The browser extension has not captured an editable field yet."
        if target.mode in ("browser", "notepad"):
            if self._is_live_window(target.window_handle):
                return True, None
            return False, "The original input window is no longer available."
        if target.mode == "word":
            return True, None
        if target.mode == "hwp":
            if self._is_live_window(target.window_handle):
                return True, None
            return False, "The original HWP window is no longer available."
        return False, f"Replace mode is not supported for {target.mode}."

    def apply(self, target: OutputTarget | None, text: str):
        if not text.strip():
            raise ValueError("There is no corrected text to apply.")

        can_replace, reason = self.inspect_replace_availability(target)
        if not can_replace:
            raise RuntimeError(reason or "Replace mode is unavailable.")

        if target.mode == "browser_extension":
            self._apply_to_browser_extension(text, target.style_info)
            return

        if target.mode == "word":
            self._focus_window(target.window_handle)
            self._apply_to_active_word(text, target.style_info)
            return

        if target.mode == "hwp":
            self._focus_window(target.window_handle)
            try:
                self._apply_to_active_hwp(text, target.style_info, target.window_handle)
                self._log_hwp_replace(
                    f"applied via HWPX COM length={len(text)} "
                    f"read_method={(target.style_info or {}).get('read_method')!r} "
                    f"style_keys={sorted((target.style_info or {}).keys())!r}"
                )
                return
            except Exception as com_exc:
                self._log_hwp_replace(f"HWPX COM apply failed: {type(com_exc).__name__}: {com_exc}")
                raise RuntimeError(
                    "HWP rich-format replacement failed. "
                    "Plain-text fallback was skipped to avoid losing formatting. "
                    f"Detail: {com_exc}"
                ) from com_exc

        self._apply_via_window_handle(target.window_handle, text)

    def _apply_to_browser_extension(self, text: str, style_info: dict | None = None):
        from client.input.browser_extension_bridge import get_browser_extension_bridge

        style_info = style_info or {}
        session_id = str(style_info.get("browser_session_id") or "")
        get_browser_extension_bridge().queue_apply(session_id, text, style_info)

    def _apply_via_window_handle(self, window_handle: int | None, text: str):
        Application, send_keys = self._load_pywinauto()
        if Application is None or send_keys is None or win32gui is None:
            raise RuntimeError("pywinauto and pywin32 are required for window replacement.")
        if not self._is_live_window(window_handle):
            raise RuntimeError("The original input window is no longer available.")

        original_clipboard = self._read_clipboard_safely()
        try:
            app = Application(backend="win32").connect(handle=window_handle)
            window = app.window(handle=window_handle)
            win32gui.ShowWindow(window_handle, 5)
            win32gui.SetForegroundWindow(window_handle)
            window.set_focus()
            time.sleep(0.25)
            self._copy_clipboard_safely(text)
            send_keys("^a")
            time.sleep(0.08)
            send_keys("{DELETE}")
            time.sleep(0.08)
            send_keys("^v")
        finally:
            if original_clipboard is not None:
                time.sleep(0.05)
                self._copy_clipboard_safely(original_clipboard)

    def _apply_to_active_word(self, text: str, style_info: dict | None = None):
        if pythoncom is None:
            raise RuntimeError("pywin32 is required for Word replacement.")
        pythoncom.CoInitialize()
        import win32com.client as win32

        word = win32.GetActiveObject("Word.Application")
        document = getattr(word, "ActiveDocument", None)
        if document is None:
            raise RuntimeError("No active Word document is available.")
        word.Visible = True
        document.Activate()
        style_info = style_info or {}
        line_styles = style_info.get("line_styles") or []
        self._log_word_replace(
            f"write text_len={len(str(text or ''))} newlines={str(text or '').count(chr(10))} "
            f"line_styles={len(line_styles)} segments={len(style_info.get('segments') or [])} "
            f"sample={str(text or '')[:80]!r}"
        )
        document.Content.Text = self._word_text_for_write(text)
        if line_styles:
            self._clear_word_direct_character_styles(document)
            self._apply_word_line_styles(document, line_styles)
        else:
            self._apply_word_style(document.Content, style_info)
        self._apply_word_style_segments(document, style_info.get("segments") or [])

    def _clear_word_direct_character_styles(self, document):
        try:
            font = document.Content.Font
            font.Bold = 0
            font.Italic = 0
            font.Underline = 0
            font.StrikeThrough = 0
            font.DoubleStrikeThrough = 0
            font.Subscript = 0
            font.Superscript = 0
            document.Content.HighlightColorIndex = 0
        except Exception:
            pass

    def _apply_word_line_styles(self, document, line_styles: list[dict]):
        if not line_styles:
            return
        try:
            paragraphs = document.Paragraphs
            paragraph_count = int(paragraphs.Count)
        except Exception:
            return

        paragraphs_by_content_line = self._word_content_paragraphs(paragraphs, paragraph_count)
        for line_style in line_styles:
            if line_style.get("is_blank"):
                continue
            paragraph_range = None
            content_line = line_style.get("content_line")
            if content_line is not None:
                try:
                    paragraph_range = paragraphs_by_content_line.get(int(content_line))
                except Exception:
                    paragraph_range = None
            if paragraph_range is None:
                try:
                    line_index = int(line_style.get("line", -1))
                    paragraph_index = line_index + 1
                    if paragraph_index < 1 or paragraph_index > paragraph_count:
                        continue
                    paragraph_range = paragraphs.Item(paragraph_index).Range.Duplicate
                except Exception:
                    continue
            try:
                raw_text = getattr(paragraph_range, "Text", "") or ""
                visible_text = raw_text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
                if not visible_text.strip():
                    continue
                if paragraph_range.End > paragraph_range.Start:
                    paragraph_range.End = paragraph_range.End - 1
                style = line_style.get("style") or {}
                self._log_word_replace(
                    f"apply line={line_style.get('line')!r} content_line={line_style.get('content_line')!r} "
                    f"text={visible_text[:60]!r} bold={style.get('bold')!r} italic={style.get('italic')!r} "
                    f"underline={style.get('underline')!r} strike={style.get('strike_through')!r} "
                    f"double_strike={style.get('double_strike_through')!r} sub={style.get('subscript')!r} "
                    f"super={style.get('superscript')!r} highlight={style.get('highlight_color_index')!r} "
                    f"color={style.get('color_hex')!r}"
                )
                self._reset_word_style_flags(paragraph_range)
                self._apply_word_style(paragraph_range, style)
                self._verify_word_style(paragraph_range, line_style)
            except Exception:
                pass

    def _reset_word_style_flags(self, word_range):
        try:
            font = word_range.Font
            font.Bold = 0
            font.Italic = 0
            font.Underline = 0
            font.StrikeThrough = 0
            font.DoubleStrikeThrough = 0
            font.Subscript = 0
            font.Superscript = 0
            word_range.HighlightColorIndex = 0
        except Exception:
            pass

    def _verify_word_style(self, word_range, line_style: dict):
        try:
            font = word_range.Font
            self._log_word_replace(
                f"verify line={line_style.get('line')!r} content_line={line_style.get('content_line')!r} "
                f"bold={getattr(font, 'Bold', None)!r} italic={getattr(font, 'Italic', None)!r} "
                f"underline={getattr(font, 'Underline', None)!r} "
                f"strike={getattr(font, 'StrikeThrough', None)!r} "
                f"double_strike={getattr(font, 'DoubleStrikeThrough', None)!r} "
                f"sub={getattr(font, 'Subscript', None)!r} super={getattr(font, 'Superscript', None)!r} "
                f"highlight={getattr(word_range, 'HighlightColorIndex', None)!r}"
            )
        except Exception as exc:
            self._log_word_replace(f"verify failed: {type(exc).__name__}: {exc}")

    def _log_word_replace(self, message: str):
        try:
            _WORD_REPLACE_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with _WORD_REPLACE_LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
        except Exception:
            pass

    def _word_content_paragraphs(self, paragraphs, paragraph_count: int) -> dict[int, object]:
        result = {}
        content_index = 0
        for paragraph_index in range(1, paragraph_count + 1):
            try:
                paragraph_range = paragraphs.Item(paragraph_index).Range.Duplicate
                raw_text = getattr(paragraph_range, "Text", "") or ""
                visible_text = raw_text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
            except Exception:
                continue
            if not visible_text.strip():
                continue
            result[content_index] = paragraph_range
            content_index += 1
        return result

    def _apply_word_style_segments(self, document, segments: list[dict]):
        if not segments:
            return
        try:
            content = document.Content
            content_start = int(content.Start)
            content_end = int(content.End)
        except Exception:
            return

        max_end = max(content_start, content_end - 1)
        for segment in segments:
            try:
                start = content_start + int(segment.get("start", 0))
                end = content_start + int(segment.get("end", 0))
            except Exception:
                continue
            start = max(content_start, min(start, max_end))
            end = max(content_start, min(end, max_end))
            if end <= start:
                continue
            try:
                segment_range = document.Range(Start=start, End=end)
                self._apply_word_style(segment_range, segment.get("style") or {})
            except Exception:
                pass

    def _apply_word_style(self, word_range, style_info: dict):
        if not style_info:
            return
        try:
            font = word_range.Font
        except Exception:
            return
        assignments = {
            "font_name": "Name",
            "font_size": "Size",
            "bold": "Bold",
            "italic": "Italic",
        }
        for key, attr in assignments.items():
            value = style_info.get(key)
            if value is None:
                continue
            try:
                if key in {"bold", "italic"}:
                    value = -1 if bool(value) else 0
                setattr(font, attr, value)
            except Exception:
                pass
        underline_value = self._word_underline_value(style_info.get("underline"))
        if underline_value is not None:
            try:
                font.Underline = underline_value
            except Exception:
                pass
        strike_value = style_info.get("strike_through")
        if strike_value is not None:
            try:
                font.StrikeThrough = -1 if bool(strike_value) else 0
                if bool(strike_value):
                    font.DoubleStrikeThrough = 0
            except Exception:
                pass
        double_strike_value = style_info.get("double_strike_through")
        if double_strike_value is not None:
            try:
                font.DoubleStrikeThrough = -1 if bool(double_strike_value) else 0
                if bool(double_strike_value):
                    font.StrikeThrough = 0
            except Exception:
                pass
        subscript_value = style_info.get("subscript")
        superscript_value = style_info.get("superscript")
        if subscript_value is not None:
            try:
                font.Subscript = -1 if bool(subscript_value) else 0
                if bool(subscript_value):
                    font.Superscript = 0
            except Exception:
                pass
        if superscript_value is not None:
            try:
                font.Superscript = -1 if bool(superscript_value) else 0
                if bool(superscript_value):
                    font.Subscript = 0
            except Exception:
                pass
        highlight_value = self._word_highlight_value(style_info.get("highlight_color_index"))
        if highlight_value is not None:
            try:
                word_range.HighlightColorIndex = highlight_value
            except Exception:
                pass
        color_value = self._word_color_from_hex(style_info.get("color_hex"))
        if color_value is not None:
            try:
                font.Color = color_value
            except Exception:
                pass

        underline_color = style_info.get("underline_color")
        if underline_color is None:
            underline_color = self._word_color_from_hex(style_info.get("underline_color_hex"))
        if underline_color is not None:
            try:
                font.UnderlineColor = int(underline_color)
            except Exception:
                pass

    def _word_highlight_value(self, value):
        if value is None:
            return None
        try:
            number = int(value)
        except Exception:
            return None
        if number in (9999999, -9999999, 9999998, -9999998):
            return None
        return number

    def _word_underline_value(self, value):
        if value is None:
            return None
        try:
            number = int(value)
        except Exception:
            return None
        if number in (9999999, -9999999, 9999998, -9999998):
            return None
        return number

    def _word_color_from_hex(self, color_hex):
        if not color_hex:
            return None
        try:
            value = str(color_hex).lstrip("#")
            if len(value) != 6:
                return None
            red = int(value[0:2], 16)
            green = int(value[2:4], 16)
            blue = int(value[4:6], 16)
            return blue * 65536 + green * 256 + red
        except Exception:
            return None

    def _word_text_for_write(self, text: str) -> str:
        normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
        return normalized.replace("\n", "\r")

    def _apply_to_active_hwp(self, text: str, style_info: dict | None = None, window_handle: int | None = None):
        if pythoncom is None:
            raise RuntimeError("pywin32 is required for HWP replacement.")
        pythoncom.CoInitialize()
        hwp = self._active_hwp_object(window_handle)
        if hwp is None:
            raise RuntimeError("No active HWP COM object is available.")
        style_info = dict(style_info or {})
        self._apply_to_hwp_hwpx_selection(hwp, text, style_info)
        return
        source_hwpml2x = ""
        if style_info.get("hwp_style_scope") == "mixed_or_unknown" and not style_info.get("segments"):
            self._diagnose_hwp_textfile_formats(hwp)
            source_hwpml2x = self._get_hwp_textfile(hwp, "HWPML2X", "selection")
            style_info["segments"] = self._capture_hwp_style_segments_from_hwpml2x(hwp)
            if not style_info["segments"]:
                style_info["segments"] = self._capture_hwp_style_segments(hwp, style_info.get("_source_text") or "")
        if source_hwpml2x and self._apply_hwpml2x_replacement(hwp, source_hwpml2x, text):
            return
        hwp.MovePos(2)
        hwp.Run("SelectAll")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp_style_info = dict(style_info)
        hwp_style_info["_replacement_text"] = text
        self._apply_hwp_style(hwp, hwp_style_info)

    def _apply_to_hwp_hwpx_selection(self, hwp, replacement_text: str, style_info: dict) -> bool:
        full_source_text = str(style_info.get("_source_text") or "").strip()
        source_text = full_source_text
        live_selection = self._read_hwp_selection_plain_text(hwp)
        if live_selection:
            source_text = live_selection
        if not source_text:
            raise RuntimeError("한글에서 드래그 선택된 원문을 찾지 못했습니다.")
        replacement_text = self._hwp_replacement_for_live_selection(
            full_source_text=full_source_text,
            full_replacement_text=replacement_text,
            selected_source_text=source_text,
        )
        work_dir = _LOG_DIR / "hwp_hwpx_blocks" / time.strftime("%Y%m%d_%H%M%S")
        work_dir.mkdir(parents=True, exist_ok=True)
        document_path = work_dir / "document.hwpx"
        try:
            if self._apply_to_hwp_hwpml_selection(hwp, replacement_text, work_dir):
                return True
            live_selection_pos = self._read_hwp_selection_position(hwp)
            if live_selection_pos:
                style_info["hwp_selection_pos"] = live_selection_pos
            self._current_hwp_selection_pos = style_info.get("hwp_selection_pos") or {}
            try:
                saved = self._save_hwp_document_as_hwpx(
                    hwp,
                    document_path,
                    expected_text=source_text,
                    prefer_full_document=bool(style_info.get("hwp_selection_pos")),
                )
            finally:
                self._current_hwp_selection_pos = {}
            if not saved:
                raise RuntimeError("Could not save active HWP document as HWPX.")
            document_text = extract_hwpx_text(document_path).strip()
            dump_dir = self._dump_hwp_selection_debug(
                hwp=hwp,
                document_path=document_path,
                raw_source_text=str(style_info.get("_source_text") or ""),
                live_selection_text=live_selection,
                extracted_hwpx_text=document_text,
                replacement_text=replacement_text,
                work_dir=work_dir,
                selection_pos=style_info.get("hwp_selection_pos") or {},
            )
            source_text = self._source_for_saved_hwpx_document(source_text, document_text)
            self._log_hwp_replace(
                f"HWPX document text length={len(document_text)} "
                f"source_len={len(source_text)} preview={document_text[:120]!r} dump_dir={dump_dir!s}"
            )
            fragment_specs = self._hwp_hwpx_fragment_specs(
                style_info=style_info,
                source_text=source_text,
                replacement_text=replacement_text,
            )
            fragment_paths = self._create_hwp_hwpx_fragments(
                document_path=document_path,
                work_dir=work_dir,
                document_text=document_text,
                selection_pos=style_info.get("hwp_selection_pos") or {},
                fragment_specs=fragment_specs,
            )
            for fragment_path in fragment_paths:
                self._dump_hwp_fragment_debug(fragment_path, dump_dir)
            self._replace_hwp_selection_with_fragments(hwp, fragment_paths)
            self._log_hwp_replace(
                "HWPX fragment apply succeeded "
                f"count={len(fragment_paths)} fragments={[str(path) for path in fragment_paths]!r}"
            )
            return True
        except Exception as exc:
            self._log_hwp_replace(f"HWPX selection apply failed: {type(exc).__name__}: {exc}")
            raise

    def _apply_to_hwp_hwpml_selection(self, hwp, replacement_text: str, work_dir: Path) -> bool:
        if not self._hwp_hwpml2x_enabled():
            self._log_hwp_replace("HWPML2X apply skipped by env")
            return False
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            self._log_hwp_replace("HWPML2X apply skipped: GetTextFile unavailable")
            return False
        try:
            hwpml = str(getter("HWPML2X", "saveblock") or "")
            if not hwpml.strip():
                self._log_hwp_replace("HWPML2X apply skipped: empty saveblock")
                return False
            source_path = work_dir / "selection.hml"
            fragment_path = work_dir / "selection.modified.hml"
            source_path.write_text(hwpml, encoding="utf-8")
            result = create_hwpml_fragment_from_selection(source_path, replacement_text, fragment_path)
            fragment_text = extract_hwpml_text(fragment_path).strip()
            self._log_hwp_replace(
                "HWPML2X fragment prepared "
                f"char_count={result.char_count} source_len={len(result.original_text)} "
                f"replacement_len={len(result.replacement_text)} fragment_len={len(fragment_text)} "
                f"path={fragment_path!s}"
            )
            if not fragment_text:
                raise RuntimeError("HWPML2X fragment text is empty.")
            self._delete_hwp_selection(hwp)
            if not self._insert_hwpml_file(hwp, fragment_path):
                raise RuntimeError("HWPML2X InsertFile failed.")
            self._log_hwp_replace(f"HWPML2X fragment apply succeeded path={fragment_path!s}")
            return True
        except Exception as exc:
            self._log_hwp_replace(f"HWPML2X apply failed; falling back to HWPX SaveAs: {type(exc).__name__}: {exc}")
            return False

    def _hwp_hwpml2x_enabled(self) -> bool:
        value = str(os.environ.get("WA_HWP_USE_HWPML2X", "")).strip().lower()
        return value not in {"0", "false", "no", "off"}

    def _hwp_hwpx_fragment_specs(
        self,
        style_info: dict,
        source_text: str,
        replacement_text: str,
    ) -> list[dict[str, str]]:
        raw_items = (
            style_info.get("hwp_hwpx_fragments")
            or style_info.get("hwpx_fragments")
            or style_info.get("replacement_fragments")
            or []
        )
        specs: list[dict[str, str]] = []
        if isinstance(raw_items, list):
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                item_source = str(item.get("source_text") or item.get("source") or "").strip()
                item_replacement = str(
                    item.get("replacement_text")
                    or item.get("replacement")
                    or item.get("text")
                    or ""
                ).strip()
                if item_source and item_replacement:
                    specs.append({"source_text": item_source, "replacement_text": item_replacement})
        if specs:
            self._log_hwp_replace(f"HWPX multi-fragment specs count={len(specs)}")
            return specs
        return [{"source_text": source_text, "replacement_text": replacement_text}]

    def _create_hwp_hwpx_fragments(
        self,
        document_path: Path,
        work_dir: Path,
        document_text: str,
        selection_pos: dict,
        fragment_specs: list[dict[str, str]],
    ) -> list[Path]:
        use_rebuilder = self._hwp_experimental_rebuilder_enabled()
        self._log_hwp_replace(f"HWP experimental Word-like HWPX rebuilder enabled={use_rebuilder}")
        fragment_paths: list[Path] = []
        for index, spec in enumerate(fragment_specs):
            fragment_path = work_dir / f"fragment.{index:03d}.modified.hwpx"
            source_text = spec["source_text"]
            replacement_text = spec["replacement_text"]
            preferred_start_paragraph = self._hwp_selection_start_paragraph(selection_pos)
            preferred_start_offset = self._hwp_selection_start_offset(selection_pos)
            preferred_end_paragraph = self._hwp_selection_end_paragraph(selection_pos)
            preferred_end_offset = self._hwp_selection_end_offset(selection_pos)
            matches = self._hwp_hwpx_matches(document_path, source_text)
            use_position_fallback = (
                preferred_start_paragraph is not None
                and not matches
                and self._hwp_position_fallback_available(document_path, source_text, selection_pos)
            )
            use_positioned_match = (
                preferred_start_paragraph is not None
                and len(matches) > 1
                and self._hwp_positioned_matches_count(
                    matches,
                    preferred_start_paragraph,
                    preferred_start_offset,
                ) == 1
            )
            self._write_hwp_hwpx_match_debug(
                work_dir=work_dir,
                index=index,
                source_text=source_text,
                matches=matches,
                preferred_start_paragraph=preferred_start_paragraph,
                preferred_start_offset=preferred_start_offset,
                use_positioned_match=use_positioned_match,
            )
            if use_rebuilder and not use_positioned_match and not use_position_fallback:
                from _tmp_hwp_word_like.hwpx_rebuilder import create_rebuilt_hwpx_fragment

                create_rebuilt_hwpx_fragment(
                    document_path=document_path,
                    source_text=source_text,
                    replacement_text=replacement_text,
                    output_path=fragment_path,
                )
            else:
                create_hwpx_fragment_from_match(
                    document_path,
                    source_text,
                    replacement_text,
                    fragment_path,
                    preferred_start_paragraph=preferred_start_paragraph,
                    preferred_start_offset=preferred_start_offset,
                    preferred_end_paragraph=preferred_end_paragraph,
                    preferred_end_offset=preferred_end_offset,
                )
            fragment_text = extract_hwpx_text(fragment_path).strip()
            self._validate_hwp_fragment_scope(
                fragment_text=fragment_text,
                document_text=document_text,
                source_text=source_text,
                replacement_text=replacement_text,
            )
            fragment_paths.append(fragment_path)
        return fragment_paths

    def _replace_hwp_selection_with_fragments(self, hwp, fragment_paths: list[Path]) -> None:
        if not fragment_paths:
            raise RuntimeError("No HWPX fragments were created.")
        self._delete_hwp_selection(hwp)
        for index, fragment_path in enumerate(fragment_paths):
            if not self._insert_hwpx_file(hwp, fragment_path):
                raise RuntimeError(f"InsertFile action failed for fragment {index}.")
            self._log_hwp_replace(f"HWPX fragment inserted index={index} path={fragment_path!s}")

    def _validate_hwp_fragment_scope(
        self,
        fragment_text: str,
        document_text: str,
        source_text: str,
        replacement_text: str,
    ) -> None:
        """Prevent deleting the selection and inserting a full-document HWPX."""
        fragment_key = self._hwp_compare_key(fragment_text)
        document_key = self._hwp_compare_key(document_text)
        source_key = self._hwp_compare_key(source_text)
        replacement_key = self._hwp_compare_key(replacement_text)
        intended_len = max(len(source_key), len(replacement_key), 1)
        self._log_hwp_replace(
            "HWP fragment scope check "
            f"fragment_key_len={len(fragment_key)} document_key_len={len(document_key)} "
            f"source_key_len={len(source_key)} replacement_key_len={len(replacement_key)}"
        )
        if not fragment_key:
            raise RuntimeError("삽입할 HWPX 조각이 비어 있습니다.")
        if len(document_key) > intended_len + 80 and len(fragment_key) > intended_len * 2 + 80:
            raise RuntimeError(
                "삽입용 HWPX가 선택 영역보다 지나치게 큽니다. "
                "전체 문서가 fragment로 들어가는 것을 차단했습니다."
            )



    def _hwp_replacement_for_live_selection(
        self,
        full_source_text: str,
        full_replacement_text: str,
        selected_source_text: str,
    ) -> str:
        """Return only the corrected text for the current HWP selection.

        The UI can hold the correction for the whole detected document, while HWP
        SaveBlock targets only the currently dragged selection. Without this crop,
        InsertFile replaces the selection with the whole corrected document.
        """
        full_source = self._normalize_hwp_source_text(full_source_text or "")
        full_replacement = self._normalize_hwp_source_text(full_replacement_text or "")
        selected = self._normalize_hwp_source_text(selected_source_text or "")
        if not full_replacement or not selected:
            return full_replacement
        if self._compare_hwp_text_lenient(full_replacement, selected):
            return full_replacement
        if not full_source or self._compare_hwp_text_lenient(full_source, selected):
            return full_replacement
        if len(self._hwp_compare_key(full_replacement)) <= len(self._hwp_compare_key(selected)) + 10:
            return full_replacement
        try:
            source_start, source_end = self._find_selection_span_in_source(full_source, selected)
            if source_start is None or source_end is None:
                self._log_hwp_replace("HWP selection crop skipped: selected text span not found in full source")
                return full_replacement
            repl_start = self._map_text_index(full_source, full_replacement, source_start)
            repl_end = self._map_text_index(full_source, full_replacement, source_end)
            cropped = full_replacement[repl_start:repl_end].strip()
            if cropped:
                self._log_hwp_replace(
                    "HWP selection crop applied "
                    f"full_replacement_len={len(full_replacement)} selected_len={len(selected)} cropped_len={len(cropped)}"
                )
                return cropped
        except Exception as exc:
            self._log_hwp_replace(f"HWP selection crop failed: {type(exc).__name__}: {exc}")
        return full_replacement

    def _find_selection_span_in_source(self, full_source: str, selected: str) -> tuple[int | None, int | None]:
        full_key, full_map = self._hwp_key_with_index_map(full_source)
        selected_key, _ = self._hwp_key_with_index_map(selected)
        if not full_key or not selected_key:
            return None, None
        pos = full_key.find(selected_key)
        if pos < 0:
            return None, None
        start = full_map[pos]
        end = full_map[pos + len(selected_key) - 1] + 1
        return start, end

    def _hwp_key_with_index_map(self, text: str) -> tuple[str, list[int]]:
        key_chars: list[str] = []
        index_map: list[int] = []
        for index, char in enumerate(self._normalize_hwp_source_text(text)):
            if char.isspace():
                continue
            # HWP GetTextFile may include visual/debug backtick near soft breaks.
            if char == "`":
                continue
            key_chars.append(char)
            index_map.append(index)
        return "".join(key_chars), index_map

    def _hwp_compare_key(self, text: str) -> str:
        return self._hwp_key_with_index_map(text)[0]

    def _compare_hwp_text_lenient(self, left: str, right: str) -> bool:
        left_key = self._hwp_compare_key(left)
        right_key = self._hwp_compare_key(right)
        if not left_key or not right_key:
            return False
        return left_key == right_key

    def _map_text_index(self, source: str, replacement: str, source_index: int) -> int:
        import difflib

        matcher = difflib.SequenceMatcher(None, source, replacement, autojunk=False)
        previous_j = 0
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if source_index < i1:
                return previous_j
            if source_index == i1:
                return j1
            if i1 < source_index < i2:
                if tag == "equal":
                    return j1 + (source_index - i1)
                if tag == "replace":
                    span = max(1, i2 - i1)
                    return j1 + round((j2 - j1) * (source_index - i1) / span)
                return j1
            if source_index == i2:
                return j2
            previous_j = j2
        return len(replacement)

    def _hwp_experimental_rebuilder_enabled(self) -> bool:
        """Enable the reversible HWPX rebuilder by default when its temp module exists.

        Set WA_HWP_EXPERIMENTAL_REBUILDER=0 to force the old matcher.
        This avoids silent non-use when the user forgets to set the env var.
        """
        value = str(os.environ.get("WA_HWP_EXPERIMENTAL_REBUILDER", "")).strip().lower()
        if value in {"0", "false", "no", "off"}:
            return False
        if value in {"1", "true", "yes", "on"}:
            return True
        try:
            import importlib.util
            return importlib.util.find_spec("_tmp_hwp_word_like.hwpx_rebuilder") is not None
        except Exception:
            return False

    def _dump_hwp_fragment_debug(self, fragment_path: Path, dump_dir: Path) -> None:
        """Copy the exact HWPX fragment that will be inserted for post-failure inspection."""
        try:
            fragment_dump = dump_dir / "fragment_modified_inserted.hwpx"
            shutil.copy2(fragment_path, fragment_dump)
            latest_dir = _LOG_DIR / "selected_block_dump" / "latest"
            if latest_dir.exists():
                shutil.copy2(fragment_path, latest_dir / "fragment_modified_inserted.hwpx")
            with zipfile.ZipFile(fragment_path, "r") as archive:
                for name in archive.namelist():
                    lower = name.lower()
                    if lower.startswith("contents/section") and lower.endswith(".xml"):
                        safe_name = "fragment_" + name.replace("/", "_").replace("\\", "_")
                        data = archive.read(name)
                        (dump_dir / safe_name).write_bytes(data)
                        if latest_dir.exists():
                            (latest_dir / safe_name).write_bytes(data)
        except Exception as exc:
            self._log_hwp_replace(f"fragment debug dump failed: {type(exc).__name__}: {exc}")


    def _dump_hwp_selection_debug(
        self,
        hwp,
        document_path: Path,
        raw_source_text: str,
        live_selection_text: str,
        extracted_hwpx_text: str,
        replacement_text: str,
        work_dir: Path,
        selection_pos: dict | None = None,
    ) -> Path:
        """Write a debuggable snapshot of the selected HWPX block without changing apply logic."""
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        dump_root = _LOG_DIR / "selected_block_dump"
        dump_dir = dump_root / timestamp
        latest_dir = dump_root / "latest"
        dump_dir.mkdir(parents=True, exist_ok=True)

        def write_text(path: Path, value: str) -> None:
            path.write_text(str(value or ""), encoding="utf-8", errors="replace")

        def codepoints(value: str, limit: int = 400) -> str:
            chars = []
            for index, char in enumerate(str(value or "")[:limit]):
                display = {"\n": "\\n", "\r": "\\r", "\t": "\\t", "\v": "\\v", "\f": "\\f"}.get(char, char)
                chars.append(f"{index:04d}: U+{ord(char):04X} {display!r}")
            return "\n".join(chars)

        try:
            shutil.copy2(document_path, dump_dir / "selected_block.hwpx")
        except Exception as exc:
            self._log_hwp_replace(f"debug dump copy hwpx failed: {type(exc).__name__}: {exc}")

        raw_gettext = ""
        getter = getattr(hwp, "GetTextFile", None)
        if callable(getter):
            try:
                raw_gettext = str(getter("TEXT", "saveblock") or "")
            except Exception as exc:
                raw_gettext = f"<GetTextFile TEXT saveblock failed: {type(exc).__name__}: {exc}>"

        normalized_raw = self._normalize_hwp_source_text(raw_source_text).strip()
        normalized_live = self._normalize_hwp_source_text(live_selection_text).strip()
        normalized_gettext = self._normalize_hwp_source_text(raw_gettext).strip()
        normalized_hwpx = self._normalize_hwp_source_text(extracted_hwpx_text).strip()

        write_text(dump_dir / "selected_text_from_style_info.txt", raw_source_text)
        write_text(dump_dir / "selected_text_from_hwp_gettext_raw.txt", raw_gettext)
        write_text(dump_dir / "selected_text_from_hwp_gettext_normalized.txt", normalized_gettext)
        write_text(dump_dir / "live_selection_text_normalized.txt", normalized_live)
        write_text(dump_dir / "extracted_text_from_hwpx.txt", extracted_hwpx_text)
        write_text(dump_dir / "extracted_text_from_hwpx_normalized.txt", normalized_hwpx)
        write_text(dump_dir / "replacement_text.txt", replacement_text)
        write_text(dump_dir / "selection_position.json", self._json_dumps(selection_pos or {}))
        write_text(dump_dir / "codepoints_gettext_raw.txt", codepoints(raw_gettext))
        write_text(dump_dir / "codepoints_hwpx_text.txt", codepoints(extracted_hwpx_text))

        info = [
            f"timestamp={timestamp}",
            f"work_dir={work_dir}",
            f"document_path={document_path}",
            f"raw_source_len={len(raw_source_text or '')}",
            f"live_selection_len={len(live_selection_text or '')}",
            f"raw_gettext_len={len(raw_gettext or '')}",
            f"extracted_hwpx_len={len(extracted_hwpx_text or '')}",
            f"replacement_len={len(replacement_text or '')}",
            f"selection_pos={selection_pos or {}}",
            f"experimental_rebuilder_enabled={self._hwp_experimental_rebuilder_enabled()}",
            f"experimental_rebuilder_env={os.environ.get('WA_HWP_EXPERIMENTAL_REBUILDER', '')}",
            f"normalized_gettext_equals_hwpx={normalized_gettext == normalized_hwpx}",
            f"normalized_live_equals_hwpx={normalized_live == normalized_hwpx}",
            f"normalized_style_equals_hwpx={normalized_raw == normalized_hwpx}",
            f"gettext_in_hwpx={bool(normalized_gettext and normalized_gettext in normalized_hwpx)}",
            f"live_in_hwpx={bool(normalized_live and normalized_live in normalized_hwpx)}",
            f"style_in_hwpx={bool(normalized_raw and normalized_raw in normalized_hwpx)}",
        ]
        write_text(dump_dir / "debug_info.txt", "\n".join(info) + "\n")

        try:
            with zipfile.ZipFile(document_path, "r") as archive:
                for name in archive.namelist():
                    lower = name.lower()
                    if lower.startswith("contents/section") and lower.endswith(".xml"):
                        safe_name = name.replace("/", "_").replace("\\", "_")
                        (dump_dir / safe_name).write_bytes(archive.read(name))
        except Exception as exc:
            self._log_hwp_replace(f"debug dump extract xml failed: {type(exc).__name__}: {exc}")

        try:
            if latest_dir.exists():
                shutil.rmtree(latest_dir)
            shutil.copytree(dump_dir, latest_dir)
        except Exception as exc:
            self._log_hwp_replace(f"debug dump latest copy failed: {type(exc).__name__}: {exc}")

        self._log_hwp_replace(f"debug dump written: {dump_dir}")
        return dump_dir

    def _write_hwp_hwpx_match_debug(
        self,
        work_dir: Path,
        index: int,
        source_text: str,
        matches,
        preferred_start_paragraph: int | None,
        preferred_start_offset: int | None,
        use_positioned_match: bool,
    ) -> None:
        try:
            data = {
                "fragment_index": index,
                "source_text": source_text,
                "source_length": len(source_text or ""),
                "preferred_start_paragraph": preferred_start_paragraph,
                "preferred_start_offset": preferred_start_offset,
                "use_positioned_match": use_positioned_match,
                "match_count": len(matches),
                "matches": [
                    {
                        "section_path": getattr(match, "section_path", ""),
                        "start": getattr(match, "start", -1),
                        "end": getattr(match, "end", -1),
                        "start_paragraph": getattr(match, "start_paragraph", -1),
                        "end_paragraph": getattr(match, "end_paragraph", -1),
                        "start_paragraph_offset": getattr(match, "start_paragraph_offset", -1),
                        "end_paragraph_offset": getattr(match, "end_paragraph_offset", -1),
                    }
                    for match in matches
                ],
            }
            (work_dir / f"fragment.{index:03d}.matches.json").write_text(
                self._json_dumps(data),
                encoding="utf-8",
            )
            latest_dir = _LOG_DIR / "selected_block_dump" / "latest"
            if latest_dir.exists():
                (latest_dir / f"fragment.{index:03d}.matches.json").write_text(
                    self._json_dumps(data),
                    encoding="utf-8",
                )
        except Exception as exc:
            self._log_hwp_replace(f"HWPX match debug write failed: {type(exc).__name__}: {exc}")

    def _json_dumps(self, value) -> str:
        try:
            import json

            return json.dumps(value, ensure_ascii=False, indent=2)
        except Exception:
            return repr(value)

    def _read_hwp_selection_plain_text(self, hwp) -> str:
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            return ""
        try:
            return self._normalize_hwp_source_text(str(getter("TEXT", "saveblock") or "")).strip()
        except Exception as exc:
            self._log_hwp_replace(f"HWP selection TEXT read failed: {type(exc).__name__}: {exc}")
            return ""

    def _read_hwp_selection_position(self, hwp) -> dict:
        getter = getattr(hwp, "GetSelectedPos", None)
        if not callable(getter):
            return {}
        try:
            value = getter()
        except Exception as exc:
            self._log_hwp_replace(f"HWP GetSelectedPos failed: {type(exc).__name__}: {exc}")
            return {}
        parsed = self._parse_hwp_selected_pos(value)
        self._log_hwp_replace(f"HWP GetSelectedPos value={value!r} parsed={parsed!r}")
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

    def _source_for_saved_hwpx_document(self, source_text: str, document_text: str) -> str:
        source = self._normalize_hwp_source_text(source_text).strip()
        document = self._normalize_hwp_source_text(document_text).strip()
        if not source:
            raise RuntimeError("한글에서 드래그 선택된 원문을 찾지 못했습니다.")
        if source == document:
            self._log_hwp_replace("HWPX source equals saved document; treating saved file as selection/all-selection block")
            return source
        if source in document:
            return source

        compact_source = self._compact_hwp_match_text(source)
        compact_document = self._compact_hwp_match_text(document)
        if compact_source == compact_document:
            self._log_hwp_replace("HWPX compact source equals saved document; using saved HWPX text as selection block")
            return document

        nobreak_source = self._linebreak_insensitive_text(source)
        nobreak_document = self._linebreak_insensitive_text(document)
        if nobreak_source and nobreak_source == nobreak_document:
            self._log_hwp_replace("HWPX linebreak-insensitive source equals saved document; using saved HWPX text")
            return document

        if nobreak_source and nobreak_source in nobreak_document and len(document) <= max(len(source) * 3, len(source) + 80):
            self._log_hwp_replace("HWPX source appears to be selected block with soft line breaks; using saved HWPX text")
            return document

        self._log_hwp_replace(
            f"HWPX source requires flexible XML match source={source[:160]!r} document={document[:160]!r}"
        )
        return source

    def _normalize_hwp_source_text(self, text: str) -> str:
        return (
            html.unescape(str(text or ""))
            .replace("\r\n", "\n")
            .replace("\r", "\n")
            .replace("\v", "\n")
            .replace("\f", "\n")
            .replace("\u2028", "\n")
            .replace("\u2029", "\n")
        )

    def _compact_hwp_match_text(self, text: str) -> str:
        lines = [line.strip() for line in self._normalize_hwp_source_text(text).split("\n")]
        return "\n".join(line for line in lines if line)

    def _linebreak_insensitive_text(self, text: str) -> str:
        return "".join(self._compact_hwp_match_text(text).split("\n"))

    def _save_hwp_document_as_hwpx(
        self,
        hwp,
        path: Path,
        expected_text: str = "",
        prefer_full_document: bool = False,
    ) -> bool:
        """Save the current HWP selection as HWPX.

        HWP 2018 may export the whole document for one SaveAs option even when
        text is selected.  Try all known options and choose the smallest HWPX
        whose extracted text matches the live selected text.  This prevents
        inserting the entire document into the dragged selection.
        """
        options = ("", "saveblock", "selection") if prefer_full_document else ("selection", "saveblock", "")
        expected_key = self._hwp_compare_key(expected_text or "")
        attempts: list[tuple[int, str, Path, int, bool, int, int, bool]] = []
        for index, option in enumerate(options):
            candidate_path = path.with_name(f"{path.stem}.candidate{index}{path.suffix}")
            try:
                if candidate_path.exists():
                    candidate_path.unlink()
                result = hwp.SaveAs(str(candidate_path), "HWPX", option)
                if not candidate_path.exists() or candidate_path.stat().st_size <= 0:
                    self._log_hwp_replace(
                        f"HWPX SaveAs option={option!r} produced no file result={result!r}"
                    )
                    continue
                try:
                    text = extract_hwpx_text(candidate_path)
                except Exception as text_exc:
                    self._log_hwp_replace(
                        f"HWPX SaveAs option={option!r} text extract failed: {type(text_exc).__name__}: {text_exc}"
                    )
                    text = ""
                text_key = self._hwp_compare_key(text)
                exact_or_contained = bool(expected_key and text_key and (expected_key in text_key or text_key in expected_key))
                positioned_match_count = self._hwp_positioned_match_count(
                    candidate_path,
                    expected_text,
                    getattr(self, "_current_hwp_selection_pos", {}),
                )
                position_fallback = self._hwp_position_fallback_available(
                    candidate_path,
                    expected_text,
                    getattr(self, "_current_hwp_selection_pos", {}),
                )
                matches = bool(
                    exact_or_contained
                    or (option == "" and prefer_full_document and (positioned_match_count >= 1 or position_fallback))
                )
                match_count = self._hwp_hwpx_match_count(candidate_path, expected_text) if exact_or_contained else 0
                attempts.append((
                    len(text_key),
                    option,
                    candidate_path,
                    index,
                    matches,
                    match_count,
                    positioned_match_count,
                    position_fallback,
                ))
                self._log_hwp_replace(
                    f"HWPX SaveAs candidate={index} option={option!r} "
                    f"result={result!r} text_key_len={len(text_key)} "
                    f"matches_selection={matches} match_count={match_count} "
                    f"positioned_match_count={positioned_match_count} "
                    f"position_fallback={position_fallback}"
                )
            except Exception as exc:
                self._log_hwp_replace(f"HWPX SaveAs option={option!r} failed: {type(exc).__name__}: {exc}")

        if not attempts:
            return False

        if prefer_full_document:
            full_unique = [
                item for item in attempts
                if item[1] == "" and item[4] and (
                    item[5] == 1
                    or item[6] >= 1
                    or item[7]
                )
            ]
            if full_unique:
                chosen = full_unique[0]
            else:
                chosen = None
        else:
            chosen = None

        matching = [item for item in attempts if item[4]]
        if matching:
            # Prefer the smallest matching export unless a unique full-document
            # export was already selected above.
            chosen = chosen or min(matching, key=lambda item: item[0])
        elif chosen is None:
            # Fallback: prefer the smallest non-empty export rather than the full document.
            chosen = min(attempts, key=lambda item: item[0])

        _, option, chosen_path, index, matches, match_count, positioned_match_count, position_fallback = chosen
        shutil.copy2(chosen_path, path)
        self._log_hwp_replace(
            f"HWPX document saved via selected candidate={index} option={option!r} "
            f"matches_selection={matches} match_count={match_count} "
            f"positioned_match_count={positioned_match_count} "
            f"position_fallback={position_fallback} "
            f"prefer_full_document={prefer_full_document} path={path!s}"
        )
        return path.exists() and path.stat().st_size > 0

    def _hwp_hwpx_match_count(self, document_path: Path, expected_text: str) -> int:
        return len(self._hwp_hwpx_matches(document_path, expected_text))

    def _hwp_hwpx_matches(self, document_path: Path, expected_text: str):
        try:
            return find_hwpx_text_matches(document_path, expected_text)
        except Exception as exc:
            self._log_hwp_replace(f"HWPX match count failed: {type(exc).__name__}: {exc}")
            return []

    def _hwp_positioned_match_count(self, document_path: Path, expected_text: str, selection_pos: dict) -> int:
        start_paragraph = self._hwp_selection_start_paragraph(selection_pos)
        if start_paragraph is None:
            return 0
        start_offset = self._hwp_selection_start_offset(selection_pos)
        return self._hwp_positioned_matches_count(
            self._hwp_hwpx_matches(document_path, expected_text),
            start_paragraph,
            start_offset,
        )

    def _hwp_position_fallback_available(self, document_path: Path, expected_text: str, selection_pos: dict) -> bool:
        start_paragraph = self._hwp_selection_start_paragraph(selection_pos)
        if start_paragraph is None:
            return False
        try:
            return has_hwpx_position_fallback_match(
                document_path,
                expected_text,
                preferred_start_paragraph=start_paragraph,
                preferred_start_offset=self._hwp_selection_start_offset(selection_pos),
                preferred_end_paragraph=self._hwp_selection_end_paragraph(selection_pos),
                preferred_end_offset=self._hwp_selection_end_offset(selection_pos),
            )
        except Exception as exc:
            self._log_hwp_replace(f"HWPX position fallback check failed: {type(exc).__name__}: {exc}")
            return False

    def _hwp_matches_at_paragraph_count(self, matches, start_paragraph: int) -> int:
        return sum(1 for match in matches if getattr(match, "start_paragraph", -1) == start_paragraph)

    def _hwp_positioned_matches_count(
        self,
        matches,
        start_paragraph: int,
        start_offset: int | None,
    ) -> int:
        if start_offset is not None:
            exact = [
                match for match in matches
                if getattr(match, "start_paragraph", -1) == start_paragraph
                and getattr(match, "start_paragraph_offset", -1) == start_offset
            ]
            if exact:
                return len(exact)
        return self._hwp_matches_at_paragraph_count(matches, start_paragraph)

    def _hwp_selection_start_paragraph(self, selection_pos: dict) -> int | None:
        try:
            value = int((selection_pos or {}).get("start_para"))
        except Exception:
            return None
        return value if value >= 0 else None

    def _hwp_selection_start_offset(self, selection_pos: dict) -> int | None:
        try:
            value = int((selection_pos or {}).get("start_pos"))
        except Exception:
            return None
        return value if value >= 0 else None

    def _hwp_selection_end_paragraph(self, selection_pos: dict) -> int | None:
        try:
            value = int((selection_pos or {}).get("end_para"))
        except Exception:
            return None
        return value if value >= 0 else None

    def _hwp_selection_end_offset(self, selection_pos: dict) -> int | None:
        try:
            value = int((selection_pos or {}).get("end_pos"))
        except Exception:
            return None
        return value if value >= 0 else None

    def _delete_hwp_selection(self, hwp) -> None:
        for action in ("Delete", "DeleteBack", "Erase"):
            try:
                hwp.Run(action)
                self._log_hwp_replace(f"HWP selection deleted via Run({action!r})")
                return
            except Exception as exc:
                self._log_hwp_replace(f"HWP delete action {action!r} failed: {type(exc).__name__}: {exc}")
        raise RuntimeError("Could not delete current HWP selection.")

    def _insert_hwpx_file(self, hwp, path: Path) -> bool:
        try:
            hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
            params = hwp.HParameterSet.HInsertFile
            for name in ("FileName", "filename", "FilePath", "FullName"):
                try:
                    setattr(params, name, str(path))
                except Exception:
                    pass
            for name, value in (
                ("KeepSection", 0),
                ("KeepCharshape", 1),
                ("KeepParashape", 1),
                ("KeepStyle", 1),
            ):
                try:
                    setattr(params, name, value)
                except Exception:
                    pass
            result = hwp.HAction.Execute("InsertFile", params.HSet)
            self._log_hwp_replace(f"HWP InsertFile result={result!r} path={path!s}")
            return bool(result is None or result)
        except Exception as exc:
            self._log_hwp_replace(f"HWP InsertFile HAction failed: {type(exc).__name__}: {exc}")

        try:
            hwp.InsertFile(str(path), "HWPX", "KeepSection:0;KeepCharshape:1;KeepParashape:1;KeepStyle:1")
            self._log_hwp_replace(f"HWP InsertFile method succeeded path={path!s}")
            return True
        except Exception as exc:
            self._log_hwp_replace(f"HWP InsertFile method failed: {type(exc).__name__}: {exc}")
            return False

    def _insert_hwpml_file(self, hwp, path: Path) -> bool:
        try:
            hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
            params = hwp.HParameterSet.HInsertFile
            for name in ("FileName", "filename", "FilePath", "FullName"):
                try:
                    setattr(params, name, str(path))
                except Exception:
                    pass
            result = hwp.HAction.Execute("InsertFile", params.HSet)
            self._log_hwp_replace(f"HWP HWPML InsertFile result={result!r} path={path!s}")
            return bool(result is None or result)
        except Exception as exc:
            self._log_hwp_replace(f"HWP HWPML InsertFile HAction failed: {type(exc).__name__}: {exc}")

        try:
            hwp.InsertFile(str(path))
            self._log_hwp_replace(f"HWP HWPML InsertFile method succeeded path={path!s}")
            return True
        except Exception as exc:
            self._log_hwp_replace(f"HWP HWPML InsertFile method failed: {type(exc).__name__}: {exc}")
            return False

    def _active_hwp_object(self, window_handle: int | None = None):
        import win32com.client as win32

        hwp = self._get_hwp_object_from_native_om(window_handle)
        if hwp is not None:
            self._log_hwp_replace(f"HWP COM object resolved via NativeOM hwnd={window_handle}")
            return hwp

        for progid in HWP_ACTIVE_PROGIDS:
            try:
                hwp = self._coerce_hwp_object(win32.GetActiveObject(progid))
                if hwp is not None:
                    self._log_hwp_replace(f"HWP COM object resolved via GetActiveObject progid={progid!r}")
                    return hwp
            except Exception:
                pass

        try:
            rot = pythoncom.GetRunningObjectTable()
            enum_moniker = rot.EnumRunning()
            bind_context = pythoncom.CreateBindCtx(0)
        except Exception:
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
            self._log_hwp_replace(f"HWP ROT entry={display_name!r}")
            try:
                hwp = self._coerce_hwp_object(rot.GetObject(moniker))
                if hwp is not None:
                    self._log_hwp_replace(f"HWP COM object resolved via ROT entry={display_name!r}")
                    return hwp
                self._log_hwp_replace(f"HWP ROT entry unusable={display_name!r}")
            except Exception as exc:
                self._log_hwp_replace(f"HWP ROT entry failed={display_name!r}: {type(exc).__name__}: {exc}")
                continue
        return None

    def _get_hwp_object_from_native_om(self, hwnd: int | None):
        if pythoncom is None or not hwnd:
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
                HWND(int(hwnd)),
                c_long(-16),  # OBJID_NATIVEOM
                ctypes.cast(iid_buffer, c_void_p),
                byref(pdisp),
            )
            if result != 0 or not pdisp.value:
                self._log_hwp_replace(f"HWP NativeOM failed hwnd={hwnd} result={result} pdisp={pdisp.value}")
                return None
            obj = pythoncom.ObjectFromAddress(pdisp.value, pythoncom.IID_IDispatch)
            hwp = self._coerce_hwp_object(win32.Dispatch(obj))
            if hwp is None:
                self._log_hwp_replace(f"HWP NativeOM unusable hwnd={hwnd}")
            return hwp
        except Exception as exc:
            self._log_hwp_replace(f"HWP NativeOM exception hwnd={hwnd}: {type(exc).__name__}: {exc}")
            return None

    def _coerce_hwp_object(self, obj):
        if obj is None:
            return None
        required = ("MovePos", "Run", "HAction", "HParameterSet")
        for candidate in self._hwp_dispatch_candidates(obj):
            if all(hasattr(candidate, name) for name in required):
                return candidate
        self._log_hwp_replace(f"HWP object coerce failed type={type(obj)}")
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

    def _apply_to_hwp_via_uia(self, window_handle: int | None, text: str):
        if not self._is_live_window(window_handle):
            raise RuntimeError("The original HWP window is no longer available.")
        wrapper = self._find_hwp_edit_wrapper(window_handle)
        if wrapper is None:
            raise RuntimeError("No writable HWP text control was found.")
        if self._set_uia_value(wrapper, text):
            return
        raise RuntimeError("The HWP text control does not expose a writable UIA value pattern.")

    def _apply_to_hwp_via_keyboard_once(self, window_handle: int | None, text: str):
        if not self._is_live_window(window_handle):
            raise RuntimeError("The original HWP window is no longer available.")
        if not self._is_hwp_window(window_handle):
            raise RuntimeError("The captured window is not an HWP window.")

        Application, send_keys = self._load_pywinauto()
        if Application is None or send_keys is None or win32gui is None:
            raise RuntimeError("pywinauto and pywin32 are required for HWP fallback replacement.")

        original_clipboard = self._read_clipboard_safely()
        try:
            app = Application(backend="win32").connect(handle=window_handle)
            window = app.window(handle=window_handle)
            win32gui.ShowWindow(window_handle, 5)
            win32gui.SetForegroundWindow(window_handle)
            window.set_focus()
            time.sleep(0.25)
            if win32gui.GetForegroundWindow() != window_handle:
                raise RuntimeError("Could not focus the original HWP window.")

            self._copy_clipboard_safely(text)
            send_keys("^a")
            time.sleep(0.12)
            send_keys("{DELETE}")
            time.sleep(0.12)
            send_keys("^v")
            time.sleep(0.15)
        finally:
            if original_clipboard is not None:
                time.sleep(0.05)
                self._copy_clipboard_safely(original_clipboard)

    def _find_hwp_edit_wrapper(self, window_handle: int | None):
        if window_handle is None:
            return None
        try:
            from pywinauto import Desktop
            from pywinauto.uia_defines import IUIA
            from pywinauto.controls.uiawrapper import UIAWrapper
            from pywinauto.uia_element_info import UIAElementInfo
        except Exception:
            return None

        candidates = []
        try:
            desktop = Desktop(backend="uia")
            window = desktop.window(handle=window_handle).wrapper_object()
        except Exception:
            window = None

        try:
            focused_element = IUIA().get_focused_element()
            focused = UIAWrapper(UIAElementInfo(focused_element)) if focused_element else None
        except Exception:
            focused = None

        if focused is not None:
            candidates.append(("focused", focused))
            current = focused
            for depth in range(4):
                try:
                    current = current.parent() if current else None
                except Exception:
                    current = None
                if current is not None:
                    candidates.append((f"focused-parent-{depth + 1}", current))

        if window is not None:
            candidates.append(("window", window))
            candidates.extend(self._descendant_wrappers(window, max_depth=5, max_nodes=140))

        best_wrapper = None
        best_length = -1
        seen = set()
        for _source, wrapper in candidates:
            key = self._wrapper_identity(wrapper)
            if key in seen:
                continue
            seen.add(key)
            if self._is_excluded_hwp_wrapper(wrapper):
                continue
            if not self._has_uia_set_value(wrapper):
                continue
            current_text = self._extract_uia_text(wrapper)
            length = len(current_text)
            if length > best_length:
                best_wrapper = wrapper
                best_length = length
        return best_wrapper

    def _descendant_wrappers(self, root, max_depth: int, max_nodes: int):
        results = []
        queue = [(root, 0)]
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
            if control_type in HWP_TEXT_CONTROL_TYPES or "hwp" in class_name.lower():
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

    def _set_uia_value(self, wrapper, text: str) -> bool:
        try:
            value_iface = getattr(wrapper, "iface_value", None)
            if value_iface is not None:
                value_iface.SetValue(text)
                return True
        except Exception:
            pass
        try:
            wrapper.set_edit_text(text)
            return True
        except Exception:
            return False

    def _has_uia_set_value(self, wrapper) -> bool:
        try:
            value_iface = getattr(wrapper, "iface_value", None)
            return bool(value_iface and not value_iface.CurrentIsReadOnly)
        except Exception:
            return False

    def _extract_uia_text(self, wrapper) -> str:
        readers = (
            lambda: wrapper.iface_value.CurrentValue if wrapper.iface_value else "",
            lambda: wrapper.legacy_properties().get("Value", ""),
            lambda: wrapper.legacy_properties().get("Name", ""),
            lambda: wrapper.iface_text.DocumentRange.GetText(-1)
            if wrapper.iface_text and wrapper.iface_text.DocumentRange
            else "",
            lambda: "\n".join(str(value) for value in wrapper.texts()),
            lambda: wrapper.window_text(),
        )
        values = []
        for reader in readers:
            try:
                value = reader()
            except Exception:
                continue
            normalized = self._normalize_text(str(value)) if value is not None else ""
            if normalized.strip():
                values.append(normalized)
        return max(values, key=len) if values else ""

    def _is_excluded_hwp_wrapper(self, wrapper) -> bool:
        control_type, title, class_name = self._describe_uia_wrapper(wrapper)
        hints = f"{control_type}\n{title}\n{class_name}".lower()
        return any(hint in hints for hint in HWP_EXCLUDED_TEXT_HINTS)

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
        return control_type, self._normalize_text(title), class_name

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

    def _normalize_text(self, text: str | None) -> str:
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

    def _log_hwp_replace(self, message: str):
        try:
            with _HWP_REPLACE_LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
        except Exception:
            pass

    def _diagnose_hwp_textfile_formats(self, hwp):
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            self._log_hwp_replace("HWP GetTextFile diagnostic skipped: method unavailable")
            return

        for fmt in HWP_TEXTFILE_FORMATS:
            for option in HWP_TEXTFILE_OPTIONS:
                try:
                    data = getter(fmt, option)
                    if isinstance(data, bytes):
                        text = data.decode("utf-8", errors="replace")
                    else:
                        text = str(data) if data is not None else ""
                    lowered = text.lower()
                    hints = [
                        token
                        for token in (
                            "charshape",
                            "charpr",
                            "textcolor",
                            "underline",
                            "fontref",
                            "facename",
                            "hcharshape",
                        )
                        if token in lowered
                    ]
                    preview = text[:300].replace("\n", "\\n").replace("\r", "\\r")
                    self._log_hwp_replace(
                        "HWP GetTextFile "
                        f"format={fmt!r} option={option!r} length={len(text)} "
                        f"hints={hints!r} preview={preview!r}"
                    )
                    if text and hints and fmt in {"HTML", "HWPML2X"}:
                        self._write_hwp_textfile_snapshot(fmt, option, text)
                except Exception as exc:
                    self._log_hwp_replace(
                        "HWP GetTextFile failed "
                        f"format={fmt!r} option={option!r}: {type(exc).__name__}: {exc}"
                    )

    def _get_hwp_textfile(self, hwp, fmt: str, option: str) -> str:
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            return ""
        try:
            return str(getter(fmt, option) or "")
        except Exception as exc:
            self._log_hwp_replace(
                f"HWP GetTextFile direct failed format={fmt!r} option={option!r}: {type(exc).__name__}: {exc}"
            )
            return ""

    def _write_hwp_textfile_snapshot(self, fmt: str, option: str, text: str):
        try:
            _HWP_TEXTFILE_SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
            safe_option = option or "default"
            path = _HWP_TEXTFILE_SNAPSHOT_DIR / f"{fmt.lower()}_{safe_option}.txt"
            path.write_text(text, encoding="utf-8", errors="replace")
            self._log_hwp_replace(f"HWP GetTextFile snapshot saved path={str(path)!r}")
        except Exception as exc:
            self._log_hwp_replace(f"HWP GetTextFile snapshot failed: {type(exc).__name__}: {exc}")

    def _apply_hwpml2x_replacement(self, hwp, source_xml: str, replacement_text: str) -> bool:
        setter = getattr(hwp, "SetTextFile", None)
        if not callable(setter):
            self._log_hwp_replace("HWPML2X replacement skipped: SetTextFile unavailable")
            return False

        rich_xml = self._build_hwpml2x_replacement(source_xml, replacement_text)
        if not rich_xml:
            return False

        for option in ("insertfile", ""):
            try:
                hwp.MovePos(2)
                hwp.Run("SelectAll")
                result = setter(rich_xml, "HWPML2X", option)
                summary = self._summarize_hwpml2x_body(self._get_hwp_textfile(hwp, "HWPML2X", "selection"))
                self._log_hwp_replace(
                    f"HWPML2X SetTextFile option={option!r} result={result!r} summary={summary!r}"
                )
                if self._hwpml2x_summary_matches_text(summary, replacement_text):
                    return True
                if self._hwpml2x_summary_has_mixed_shapes(summary):
                    return True
            except Exception as exc:
                self._log_hwp_replace(
                    f"HWPML2X SetTextFile failed option={option!r}: {type(exc).__name__}: {exc}"
                )
        return False

    def _build_hwpml2x_replacement(self, source_xml: str, replacement_text: str) -> str:
        try:
            import xml.etree.ElementTree as ET

            root = ET.fromstring(source_xml.lstrip("\ufeff"))
        except Exception as exc:
            self._log_hwp_replace(f"HWPML2X replacement build parse failed: {type(exc).__name__}: {exc}")
            return ""

        text_nodes = root.findall(".//BODY//TEXT")
        if not text_nodes:
            self._log_hwp_replace("HWPML2X replacement build failed: no BODY TEXT nodes")
            return ""

        if "\n" in self._normalize_text(replacement_text) and self._assign_hwpml2x_lines(text_nodes, replacement_text):
            xml_body = ET.tostring(root, encoding="unicode", short_empty_elements=True)
            self._log_hwp_replace(
                "HWPML2X replacement built linewise "
                f"length={len(xml_body)} text_length={len(replacement_text)} "
                f"text_nodes={len(text_nodes)} lines={len(self._split_hwp_replacement_lines(replacement_text))}"
            )
            return '<?xml version="1.0" encoding="UTF-16" standalone="no" ?>' + xml_body

        original_lengths = []
        for text_node in text_nodes:
            original_lengths.append(sum(len(char_node.text or "") for char_node in text_node.findall("CHAR")))

        cursor = 0
        for index, text_node in enumerate(text_nodes):
            length = original_lengths[index]
            if index == len(text_nodes) - 1:
                chunk = replacement_text[cursor:]
            else:
                chunk = replacement_text[cursor : cursor + length]
            cursor += length
            for char_node in list(text_node.findall("CHAR")):
                text_node.remove(char_node)
            if chunk:
                char_node = ET.Element("CHAR")
                char_node.text = chunk
                text_node.append(char_node)

        xml_body = ET.tostring(root, encoding="unicode", short_empty_elements=True)
        self._log_hwp_replace(
            f"HWPML2X replacement built length={len(xml_body)} text_length={len(replacement_text)}"
        )
        return '<?xml version="1.0" encoding="UTF-16" standalone="no" ?>' + xml_body

    def _assign_hwpml2x_lines(self, text_nodes, replacement_text: str) -> bool:
        import xml.etree.ElementTree as ET

        lines = self._split_hwp_replacement_lines(replacement_text)
        if not lines or len(lines) > len(text_nodes):
            self._log_hwp_replace(
                "HWPML2X linewise skipped "
                f"lines={len(lines)} text_nodes={len(text_nodes)}"
            )
            return False

        for index, text_node in enumerate(text_nodes):
            line = lines[index] if index < len(lines) else ""
            for char_node in list(text_node.findall("CHAR")):
                text_node.remove(char_node)
            if line:
                char_node = ET.Element("CHAR")
                char_node.text = line
                text_node.append(char_node)
        return True

    def _split_hwp_replacement_lines(self, text: str) -> list[str]:
        normalized = self._normalize_text(text)
        lines = normalized.split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        return lines

    def _hwpml2x_summary_has_mixed_shapes(self, summary: list[dict]) -> bool:
        shapes = {str(item.get("shape")) for item in summary if item.get("shape") is not None}
        return len(shapes) > 1 or bool(shapes - {"0"})

    def _hwpml2x_summary_matches_text(self, summary: list[dict], replacement_text: str) -> bool:
        summary_text = "".join(str(item.get("text") or "") for item in summary)
        replacement_content = "".join(line for line in self._split_hwp_replacement_lines(replacement_text) if line)
        if not replacement_content:
            return False
        return summary_text.startswith(replacement_content[: max(1, min(len(replacement_content), 24))])

    def _apply_hwp_style(self, hwp, style_info: dict):
        if not style_info:
            return
        segments = style_info.get("segments") or []
        if segments:
            self._apply_hwp_style_segments(hwp, segments, str(style_info.get("_replacement_text") or ""))
            return
        if style_info.get("hwp_style_scope") != "basic":
            self._log_hwp_replace(f"HWP style skipped scope={style_info.get('hwp_style_scope')!r}")
            return
        safe_style = self._sanitize_hwp_style_info(style_info)
        if not safe_style:
            self._log_hwp_replace(f"HWP style skipped unsafe style={style_info!r}")
            return
        if "font_name" not in safe_style or "font_size" not in safe_style:
            self._log_hwp_replace(f"HWP style skipped incomplete style={safe_style!r}")
            return
        try:
            hwp.Run("SelectAll")
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            char_shape = hwp.HParameterSet.HCharShape
            font_name = safe_style.get("font_name")
            if font_name:
                for attr in (
                    "FaceNameHangul",
                    "FaceNameLatin",
                    "FaceNameHanja",
                    "FaceNameJapanese",
                    "FaceNameOther",
                    "FaceNameSymbol",
                    "FaceNameUser",
                ):
                    try:
                        setattr(char_shape, attr, font_name)
                    except Exception:
                        pass
            height = self._hwp_points_to_height(safe_style.get("font_size"))
            if height is not None:
                try:
                    char_shape.Height = height
                except Exception:
                    pass
            for key, attr in (
                ("color", "TextColor"),
                ("underline_type", "UnderlineType"),
                ("underline_shape", "UnderlineShape"),
                ("underline_color", "UnderlineColor"),
            ):
                value = safe_style.get(key)
                if value is None:
                    continue
                try:
                    setattr(char_shape, attr, value)
                except Exception:
                    pass
            for key, attr in (("bold", "Bold"), ("italic", "Italic")):
                value = style_info.get(key)
                if value is None:
                    continue
                try:
                    setattr(char_shape, attr, 1 if bool(value) else 0)
                except Exception:
                    pass
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            self._log_hwp_replace(f"HWP style applied style={safe_style!r}")
        except Exception as exc:
            self._log_hwp_replace(f"HWP style apply failed: {type(exc).__name__}: {exc}")
        finally:
            try:
                hwp.Run("Cancel")
            except Exception:
                pass

    def _capture_hwp_style_segments(self, hwp, source_text: str) -> list[dict]:
        if not source_text:
            return []
        if len(source_text) > 500:
            self._log_hwp_replace(f"HWP segment capture skipped length={len(source_text)}")
            return []
        segments: list[dict] = []
        current_signature = None
        current_style = None
        segment_start = 0
        text_index = 0
        try:
            for char in source_text:
                if char == "\n":
                    text_index += 1
                    continue
                start_pos = self._hwp_text_index_to_position(source_text, text_index)
                end_pos = self._hwp_text_index_to_position(source_text, text_index + 1)
                style = self._read_hwp_style_for_range(hwp, start_pos, end_pos)
                signature = tuple(sorted(style.items()))
                if current_signature is None:
                    current_signature = signature
                    current_style = style
                    segment_start = text_index
                elif signature != current_signature:
                    self._append_hwp_segment(segments, segment_start, text_index, current_style)
                    current_signature = signature
                    current_style = style
                    segment_start = text_index
                text_index += 1
            if current_signature is not None:
                self._append_hwp_segment(segments, segment_start, text_index, current_style)
        except Exception as exc:
            self._log_hwp_replace(f"HWP segment capture failed: {type(exc).__name__}: {exc}")
            segments = []
        finally:
            try:
                hwp.Run("Cancel")
            except Exception:
                pass
        if len(segments) <= 1:
            self._log_hwp_replace(f"HWP segment capture not useful count={len(segments)}")
            return []
        self._log_hwp_replace(f"HWP segment capture count={len(segments)}")
        return segments

    def _capture_hwp_style_segments_from_hwpml2x(self, hwp) -> list[dict]:
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            return []
        for option in ("selection", ""):
            try:
                data = getter("HWPML2X", option)
                xml_text = str(data) if data is not None else ""
            except Exception as exc:
                self._log_hwp_replace(
                    f"HWPML2X segment capture failed option={option!r}: {type(exc).__name__}: {exc}"
                )
                continue
            segments = self._parse_hwpml2x_style_segments(xml_text)
            if segments:
                self._log_hwp_replace(
                    f"HWPML2X segment capture count={len(segments)} option={option!r}"
                )
                return segments
        self._log_hwp_replace("HWPML2X segment capture not useful")
        return []

    def _parse_hwpml2x_style_segments(self, xml_text: str) -> list[dict]:
        if not xml_text:
            return []
        try:
            import xml.etree.ElementTree as ET

            root = ET.fromstring(xml_text.lstrip("\ufeff"))
        except Exception as exc:
            self._log_hwp_replace(f"HWPML2X parse failed: {type(exc).__name__}: {exc}")
            return []

        font_names = self._parse_hwpml2x_font_names(root)
        char_shapes = self._parse_hwpml2x_char_shapes(root, font_names)
        if not char_shapes:
            return []

        segments: list[dict] = []
        position = 0
        previous_signature = None
        for paragraph in root.findall(".//BODY//P"):
            if position > 0:
                position += 1
            for text_node in paragraph.findall("TEXT"):
                chunk = "".join(char_node.text or "" for char_node in text_node.findall("CHAR"))
                if not chunk:
                    continue
                style = char_shapes.get(text_node.get("CharShape") or "")
                start = position
                end = position + len(chunk)
                position = end
                if not style:
                    continue
                signature = tuple(sorted(style.items()))
                if segments and signature == previous_signature and segments[-1]["end"] == start:
                    segments[-1]["end"] = end
                else:
                    segments.append({"start": start, "end": end, "style": dict(style)})
                previous_signature = signature

        if len(segments) <= 1:
            return []
        return segments

    def _parse_hwpml2x_font_names(self, root) -> dict[str, str]:
        font_names: dict[str, str] = {}
        for font_face in root.findall(".//FACENAMELIST/FONTFACE"):
            if font_face.get("Lang") != "Hangul":
                continue
            for font in font_face.findall("FONT"):
                font_id = font.get("Id")
                name = font.get("Name")
                if font_id is not None and name:
                    font_names[font_id] = name
            break
        return font_names

    def _parse_hwpml2x_char_shapes(self, root, font_names: dict[str, str]) -> dict[str, dict]:
        char_shapes: dict[str, dict] = {}
        for node in root.findall(".//CHARSHAPELIST/CHARSHAPE"):
            shape_id = node.get("Id")
            if shape_id is None:
                continue
            style = self._sanitize_hwp_style_info(
                {
                    "font_name": self._hwpml2x_font_name(node, font_names),
                    "font_size": self._hwp_height_to_points_value(node.get("Height")),
                    "color": self._safe_hwp_int(node.get("TextColor")),
                    "bold": node.find("BOLD") is not None,
                    "italic": node.find("ITALIC") is not None,
                    **self._hwpml2x_underline_style(node),
                    **self._hwpml2x_strikeout_style(node),
                },
                require_base=False,
            )
            if style:
                char_shapes[shape_id] = style
        return char_shapes

    def _hwpml2x_font_name(self, char_shape_node, font_names: dict[str, str]) -> str | None:
        font_id = char_shape_node.find("FONTID")
        if font_id is None:
            return None
        return font_names.get(font_id.get("Hangul") or "")

    def _hwpml2x_underline_style(self, char_shape_node) -> dict:
        underline = char_shape_node.find("UNDERLINE")
        if underline is None:
            return {"underline_type": 0}
        return {
            "underline_type": 1,
            "underline_shape": 0,
            "underline_color": self._safe_hwp_int(underline.get("Color")),
        }

    def _hwpml2x_strikeout_style(self, char_shape_node) -> dict:
        strikeout = char_shape_node.find("STRIKEOUT")
        if strikeout is None:
            return {"strikeout_type": 0}
        return {
            "strikeout_type": 1,
            "strikeout_shape": 0,
            "strikeout_color": self._safe_hwp_int(strikeout.get("Color")),
        }

    def _read_hwp_style_for_range(self, hwp, start_pos: tuple[int, int], end_pos: tuple[int, int]) -> dict:
        try:
            hwp.SelectText(start_pos[0], start_pos[1], end_pos[0], end_pos[1])
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            char_shape = hwp.HParameterSet.HCharShape
        except Exception:
            return {}
        return self._sanitize_hwp_style_info(
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
                "font_size": self._hwp_height_to_points_value(self._hwp_attr(char_shape, "Height")),
                "color": self._hwp_attr(char_shape, "TextColor"),
                "bold": self._hwp_bool(self._hwp_attr(char_shape, "Bold")),
                "italic": self._hwp_bool(self._hwp_attr(char_shape, "Italic")),
                "underline_type": self._hwp_attr(char_shape, "UnderlineType"),
                "underline_shape": self._hwp_attr(char_shape, "UnderlineShape"),
                "underline_color": self._hwp_attr(char_shape, "UnderlineColor"),
            },
            require_base=False,
        )

    def _append_hwp_segment(self, segments: list[dict], start: int, end: int, style: dict | None):
        if end <= start or not style:
            return
        segments.append({"start": start, "end": end, "style": dict(style)})

    def _apply_hwp_style_segments(self, hwp, segments: list[dict], replacement_text: str = ""):
        applied = 0
        for segment in segments[:200]:
            try:
                start = max(0, int(segment.get("start", 0)))
                end = max(0, int(segment.get("end", 0)))
            except Exception:
                continue
            if end <= start:
                continue
            style = self._sanitize_hwp_style_info(segment.get("style") or {}, require_base=False)
            if not style:
                continue
            start_pos = self._hwp_text_index_to_position(replacement_text, start)
            end_pos = self._hwp_text_index_to_position(replacement_text, end)
            try:
                if applied < 8:
                    self._log_hwp_replace(
                        "HWP segment style try "
                        f"range=({start},{end}) pos={start_pos}->{end_pos} "
                        f"text={replacement_text[start:end]!r} style={style!r}"
                    )
                self._select_hwp_text_range(hwp, start_pos, end_pos, end - start)
                hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
                char_shape = hwp.HParameterSet.HCharShape
                self._assign_hwp_char_shape(char_shape, style)
                hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
                applied += 1
            except Exception as exc:
                self._log_hwp_replace(f"HWP segment style apply failed: {type(exc).__name__}: {exc}")
        try:
            hwp.Run("Cancel")
        except Exception:
            pass
        self._log_hwp_replace(f"HWP segment styles applied count={applied} total={len(segments)}")
        self._log_hwpml2x_body_summary(hwp, "after_segment_apply")

    def _select_hwp_text_range(self, hwp, start_pos: tuple[int, int], end_pos: tuple[int, int], length: int):
        if self._select_hwp_text_range_by_cursor(hwp, start_pos, end_pos, length):
            return
        hwp.SelectText(start_pos[0], start_pos[1], end_pos[0], end_pos[1])

    def _select_hwp_text_range_by_cursor(
        self,
        hwp,
        start_pos: tuple[int, int],
        end_pos: tuple[int, int],
        length: int,
    ) -> bool:
        if not ENABLE_HWP_CURSOR_SEGMENT_SELECTION:
            return False
        if length <= 0 or length > 500:
            return False
        if start_pos[0] != end_pos[0]:
            return False
        try:
            hwp.Run("Cancel")
        except Exception:
            pass
        try:
            hwp.SetPos(0, start_pos[0], start_pos[1])
            for _ in range(length):
                hwp.Run("MoveSelRight")
            if length <= 3:
                self._log_hwp_replace(
                    f"HWP range selected via cursor pos={start_pos}->{end_pos} length={length}"
                )
            return True
        except Exception as exc:
            self._log_hwp_replace(
                f"HWP cursor range select failed pos={start_pos}->{end_pos}: {type(exc).__name__}: {exc}"
            )
            try:
                hwp.Run("Cancel")
            except Exception:
                pass
            return False

    def _assign_hwp_char_shape(self, char_shape, style: dict):
        font_name = style.get("font_name")
        if font_name:
            for attr in (
                "FaceNameHangul",
                "FaceNameLatin",
                "FaceNameHanja",
                "FaceNameJapanese",
                "FaceNameOther",
                "FaceNameSymbol",
                "FaceNameUser",
            ):
                try:
                    setattr(char_shape, attr, font_name)
                except Exception:
                    pass
        height = self._hwp_points_to_height(style.get("font_size"))
        if height is not None:
            try:
                char_shape.Height = height
            except Exception:
                pass

    def _log_hwpml2x_body_summary(self, hwp, label: str):
        getter = getattr(hwp, "GetTextFile", None)
        if not callable(getter):
            return
        try:
            xml_text = str(getter("HWPML2X", "selection") or "")
            summary = self._summarize_hwpml2x_body(xml_text)
            self._log_hwp_replace(f"HWPML2X body summary {label}: {summary!r}")
        except Exception as exc:
            self._log_hwp_replace(f"HWPML2X body summary failed {label}: {type(exc).__name__}: {exc}")

    def _summarize_hwpml2x_body(self, xml_text: str) -> list[dict]:
        if not xml_text:
            return []
        try:
            import xml.etree.ElementTree as ET

            root = ET.fromstring(xml_text.lstrip("\ufeff"))
        except Exception:
            return []
        summary = []
        for text_node in root.findall(".//BODY//TEXT"):
            chunk = "".join(char_node.text or "" for char_node in text_node.findall("CHAR"))
            if not chunk:
                continue
            summary.append({"shape": text_node.get("CharShape"), "text": chunk[:20]})
            if len(summary) >= 30:
                break
        return summary
        for key, attr in (
            ("color", "TextColor"),
            ("underline_type", "UnderlineType"),
            ("underline_shape", "UnderlineShape"),
            ("underline_color", "UnderlineColor"),
            ("strikeout_type", "StrikeOutType"),
            ("strikeout_shape", "StrikeOutShape"),
            ("strikeout_color", "StrikeOutColor"),
        ):
            value = style.get(key)
            if value is None:
                continue
            try:
                setattr(char_shape, attr, value)
            except Exception:
                pass
        for key, attr in (("bold", "Bold"), ("italic", "Italic")):
            value = style.get(key)
            if value is None:
                continue
            try:
                setattr(char_shape, attr, 1 if bool(value) else 0)
            except Exception:
                pass

    def _hwp_points_to_height(self, value):
        if value is None:
            return None
        try:
            points = float(value)
        except Exception:
            return None
        if points < 4 or points > 200:
            return None
        return int(round(points * 100))

    def _sanitize_hwp_style_info(self, style_info: dict, require_base: bool = True) -> dict:
        safe: dict = {}
        font_name = style_info.get("font_name")
        if isinstance(font_name, str) and font_name.strip():
            safe["font_name"] = font_name.strip()

        font_size = self._safe_hwp_float(style_info.get("font_size"))
        if font_size is not None and 4 <= font_size <= 200:
            safe["font_size"] = font_size

        color = self._safe_hwp_int(style_info.get("color"))
        if color is not None and 0 <= color <= 0xFFFFFF:
            safe["color"] = color

        underline_color = self._safe_hwp_int(style_info.get("underline_color"))
        if underline_color is not None and 0 <= underline_color <= 0xFFFFFF:
            safe["underline_color"] = underline_color

        for key in ("underline_type", "underline_shape", "strikeout_type", "strikeout_shape"):
            value = self._safe_hwp_int(style_info.get(key))
            if value is not None and 0 <= value <= 20:
                safe[key] = value

        strikeout_color = self._safe_hwp_int(style_info.get("strikeout_color"))
        if strikeout_color is not None and 0 <= strikeout_color <= 0xFFFFFF:
            safe["strikeout_color"] = strikeout_color

        for key in ("bold", "italic"):
            value = style_info.get(key)
            if isinstance(value, bool):
                safe[key] = value
            elif value in (0, 1):
                safe[key] = bool(value)
        if require_base and ("font_name" not in safe or "font_size" not in safe):
            return {}
        return safe

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

    def _hwp_bool(self, value):
        if value is None:
            return None
        try:
            return bool(int(value))
        except Exception:
            return bool(value)

    def _hwp_height_to_points_value(self, value):
        if value is None:
            return None
        try:
            points = float(value) / 100
        except Exception:
            return None
        if points < 4 or points > 200:
            return None
        return points

    def _hwp_text_index_to_position(self, text: str, index: int) -> tuple[int, int]:
        para = 0
        pos = 0
        for char in (text or "")[:index]:
            if char == "\n":
                para += 1
                pos = 0
            else:
                pos += 1
        return para, pos

    def _safe_hwp_int(self, value):
        if value is None or isinstance(value, bool):
            return None
        try:
            return int(value)
        except Exception:
            return None

    def _safe_hwp_float(self, value):
        if value is None or isinstance(value, bool):
            return None
        try:
            return float(value)
        except Exception:
            return None

    def _is_live_window(self, window_handle: int | None) -> bool:
        if win32gui is None or not window_handle:
            return False
        try:
            return bool(win32gui.IsWindow(window_handle))
        except Exception:
            return False

    def _is_hwp_window(self, window_handle: int | None) -> bool:
        if not self._is_live_window(window_handle):
            return False
        process_name = self._process_name_for_window(window_handle)
        if process_name:
            return process_name in HWP_PROCESS_NAMES or process_name.startswith("hwp")
        try:
            class_name = (win32gui.GetClassName(window_handle) or "").lower()
        except Exception:
            class_name = ""
        return "hwp" in class_name or "hnc" in class_name

    def _process_name_for_window(self, window_handle: int | None) -> str:
        if win32process is None or not window_handle:
            return ""
        try:
            _thread_id, process_id = win32process.GetWindowThreadProcessId(window_handle)
        except Exception:
            return ""
        if not process_id or psutil is None:
            return ""
        try:
            return psutil.Process(process_id).name().lower()
        except Exception:
            return ""

    def _focus_window(self, window_handle: int | None):
        if win32gui is None or not self._is_live_window(window_handle):
            return
        try:
            win32gui.ShowWindow(window_handle, 5)
            win32gui.SetForegroundWindow(window_handle)
            time.sleep(0.2)
        except Exception:
            pass

    def _read_clipboard_safely(self):
        for _ in range(3):
            try:
                return pyperclip.paste()
            except Exception:
                time.sleep(0.05)
        return None

    def _copy_clipboard_safely(self, text):
        for _ in range(3):
            try:
                pyperclip.copy(text)
                return True
            except Exception:
                time.sleep(0.05)
        return False

    def _load_pywinauto(self):
        try:
            from pywinauto import Application
            from pywinauto.keyboard import send_keys

            return Application, send_keys
        except Exception:
            return None, None
