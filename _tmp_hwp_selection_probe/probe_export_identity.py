from __future__ import annotations

import argparse
import base64
import json
from pathlib import Path
import sys
import time


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CLIENT_INPUT = PROJECT_ROOT / "client" / "input"
if str(CLIENT_INPUT) not in sys.path:
    sys.path.insert(0, str(CLIENT_INPUT))

from hwpx_document import extract_hwpx_text


LOG_ROOT = PROJECT_ROOT / ".logs" / "hwp_export_identity_probe"
PROGIDS = ("HWPFrame.HwpObject.2", "HWPFrame.HwpObject.1", "HWPFrame.HwpObject")
TEXTFILE_FORMATS = ("HWPX", "HWPML2X", "HWPML2X_S", "HWPML2X_P", "HWPML", "HTML", "TEXT")
TEXTFILE_OPTIONS = ("", "saveblock", "selection")
SAVEAS_OPTIONS = ("", "saveblock", "selection")


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe HWP export methods and active-document identity changes.")
    parser.add_argument("--delay", type=float, default=0.0)
    parser.add_argument("--saveas", action="store_true", help="Also try SaveAs(..., 'HWPX', option). This may rename the active document.")
    args = parser.parse_args()

    if args.delay:
        print(f"Waiting {args.delay:g}s. Keep the target HWP document active.")
        time.sleep(args.delay)

    run_dir = LOG_ROOT / time.strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    report = {"run_dir": str(run_dir), "saveas_enabled": args.saveas, "errors": []}
    try:
        hwp = active_hwp_object()
        if hwp is None:
            raise RuntimeError("active HWP COM object not found")
        report["identity_before"] = document_identity(hwp)
        report["textfile_exports"] = probe_get_text_file(hwp, run_dir)
        report["identity_after_textfile"] = document_identity(hwp)
        if args.saveas:
            report["saveas_exports"] = probe_save_as(hwp, run_dir)
            report["identity_after_saveas"] = document_identity(hwp)
        report["status"] = "ok"
    except Exception as exc:
        report["errors"].append(f"{type(exc).__name__}: {exc}")
        report["status"] = "fail"

    path = run_dir / "report.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not report["errors"] else 1


def active_hwp_object():
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:
        raise RuntimeError(f"pywin32 is required: {exc}") from exc

    pythoncom.CoInitialize()
    for progid in PROGIDS:
        try:
            return win32.GetActiveObject(progid)
        except Exception:
            pass

    rot = pythoncom.GetRunningObjectTable()
    enum_moniker = rot.EnumRunning()
    bind_context = pythoncom.CreateBindCtx(0)
    while True:
        monikers = enum_moniker.Next(1)
        if not monikers:
            break
        moniker = monikers[0]
        try:
            name = str(moniker.GetDisplayName(bind_context, None) or "")
        except Exception:
            name = ""
        if "hwp" not in name.lower():
            continue
        try:
            obj = rot.GetObject(moniker)
            query = getattr(obj, "QueryInterface", None)
            if callable(query):
                return win32.dynamic.Dispatch(query(pythoncom.IID_IDispatch))
            return win32.dynamic.Dispatch(obj)
        except Exception:
            continue
    return None


def document_identity(hwp) -> dict:
    row = {"properties": {}, "methods": {}, "xhwp": {}}
    for name in (
        "Name",
        "FullName",
        "Path",
        "FileName",
        "CurrentFileName",
        "CurFileName",
        "DocumentPath",
        "Version",
    ):
        row["properties"][name] = safe_read(lambda n=name: getattr(hwp, n))
    for name in ("GetCurFileName", "GetFileName", "GetPath", "GetTitle"):
        attr = getattr(hwp, name, None)
        if callable(attr):
            row["methods"][name] = safe_read(attr)
    xdocs = getattr(hwp, "XHwpDocuments", None)
    row["xhwp"]["documents_count"] = safe_read(lambda: int(xdocs.Count)) if xdocs is not None else None
    active = safe_read(lambda: xdocs.Active, None) if xdocs is not None else None
    if active is not None:
        row["xhwp"]["active"] = {}
        for name in ("Name", "FullName", "Path", "Title", "Saved"):
            row["xhwp"]["active"][name] = safe_read(lambda n=name: getattr(active, n))
    return row


def probe_get_text_file(hwp, run_dir: Path) -> list[dict]:
    rows = []
    getter = getattr(hwp, "GetTextFile", None)
    if not callable(getter):
        return [{"ok": False, "error": "GetTextFile not callable"}]
    for fmt in TEXTFILE_FORMATS:
        for option in TEXTFILE_OPTIONS:
            before = document_identity(hwp)
            row = {"format": fmt, "option": option, "identity_before": before}
            try:
                value = getter(fmt, option)
                data = value_to_bytes(value)
                row.update(describe_export_bytes(data, run_dir / f"gettext_{fmt}_{option or 'default'}"))
                row["ok"] = True
            except Exception as exc:
                row.update({"ok": False, "error": f"{type(exc).__name__}: {exc}"})
            row["identity_after"] = document_identity(hwp)
            row["identity_changed"] = identity_key(row["identity_before"]) != identity_key(row["identity_after"])
            rows.append(row)
    return rows


def probe_save_as(hwp, run_dir: Path) -> list[dict]:
    rows = []
    for option in SAVEAS_OPTIONS:
        path = run_dir / f"saveas_{option or 'default'}.hwpx"
        if path.exists():
            path.unlink()
        before = document_identity(hwp)
        row = {"option": option, "path": str(path), "identity_before": before}
        try:
            result = hwp.SaveAs(str(path), "HWPX", option)
            row.update({
                "ok": True,
                "result": repr(result),
                "exists": path.exists(),
                "size": path.stat().st_size if path.exists() else 0,
            })
            if path.exists() and path.stat().st_size > 0:
                try:
                    text = extract_hwpx_text(path)
                    row["text_length"] = len(text)
                    row["text_preview"] = text[:300]
                except Exception as exc:
                    row["text_error"] = f"{type(exc).__name__}: {exc}"
        except Exception as exc:
            row.update({"ok": False, "error": f"{type(exc).__name__}: {exc}"})
        row["identity_after"] = document_identity(hwp)
        row["identity_changed"] = identity_key(row["identity_before"]) != identity_key(row["identity_after"])
        rows.append(row)
    return rows


def describe_export_bytes(data: bytes, stem: Path) -> dict:
    row = {"byte_length": len(data), "zip_like": data.startswith(b"PK")}
    if not data:
        return row
    suffix = ".hwpx" if row["zip_like"] else ".txt"
    path = stem.with_suffix(suffix)
    path.write_bytes(data)
    row["path"] = str(path)
    row["base64_head"] = base64.b64encode(data[:80]).decode("ascii")
    if row["zip_like"]:
        try:
            text = extract_hwpx_text(path)
            row["text_length"] = len(text)
            row["text_preview"] = text[:300]
        except Exception as exc:
            row["text_error"] = f"{type(exc).__name__}: {exc}"
    else:
        preview = data[:600].decode("utf-8", errors="replace")
        row["text_preview"] = preview
    return row


def value_to_bytes(value) -> bytes:
    if value is None:
        return b""
    if isinstance(value, bytes):
        return value
    if isinstance(value, bytearray):
        return bytes(value)
    text = str(value)
    if text.startswith("PK"):
        return text.encode("latin1", errors="ignore")
    return text.encode("utf-8", errors="replace")


def identity_key(identity: dict) -> tuple:
    props = identity.get("properties") or {}
    methods = identity.get("methods") or {}
    active = ((identity.get("xhwp") or {}).get("active") or {})
    keys = []
    for source in (props, methods, active):
        for name in ("FullName", "Path", "FileName", "CurrentFileName", "CurFileName", "GetCurFileName", "GetFileName", "Name", "Title"):
            value = source.get(name)
            if isinstance(value, str) and value:
                keys.append((name, value))
    return tuple(keys)


def safe_read(callback, default=""):
    try:
        value = callback()
        if value is None:
            return ""
        if isinstance(value, (str, int, float, bool)):
            return value
        return repr(value)
    except Exception as exc:
        return f"<{type(exc).__name__}: {exc}>"


if __name__ == "__main__":
    raise SystemExit(main())
