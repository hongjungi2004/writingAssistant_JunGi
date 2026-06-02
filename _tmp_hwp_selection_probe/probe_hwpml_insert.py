from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import time


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CLIENT_INPUT = PROJECT_ROOT / "client" / "input"
if str(CLIENT_INPUT) not in sys.path:
    sys.path.insert(0, str(CLIENT_INPUT))
LOG_ROOT = PROJECT_ROOT / ".logs" / "hwpml_insert_probe"
PROGIDS = ("HWPFrame.HwpObject.2", "HWPFrame.HwpObject.1", "HWPFrame.HwpObject")

from hwpml_document import create_hwpml_fragment_from_selection, extract_hwpml_text


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe whether GetTextFile(HWPML2X) output can be inserted without SaveAs.")
    parser.add_argument("--delay", type=float, default=0.0)
    parser.add_argument("--apply", action="store_true", help="Delete selection and InsertFile the captured HWPML2X file. Use only on a copied document.")
    parser.add_argument("--option", default="saveblock", choices=["", "saveblock", "selection"])
    parser.add_argument("--replacement", default="", help="Replacement text. If omitted, inserts the original captured HWPML2X.")
    args = parser.parse_args()

    if args.delay:
        print(f"Waiting {args.delay:g}s. Keep the HWP selection active.")
        time.sleep(args.delay)

    run_dir = LOG_ROOT / time.strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    report = {"run_dir": str(run_dir), "apply": args.apply, "option": args.option, "errors": []}
    try:
        hwp = active_hwp_object()
        if hwp is None:
            raise RuntimeError("active HWP COM object not found")
        report["identity_before"] = document_identity(hwp)
        hwpml = str(hwp.GetTextFile("HWPML2X", args.option) or "")
        report["hwpml_length"] = len(hwpml)
        report["hwpml_preview"] = hwpml[:500]
        if not hwpml.strip():
            raise RuntimeError("GetTextFile(HWPML2X) returned empty text")

        paths = []
        for ext in ("hml", "xml", "hwpml"):
            path = run_dir / f"selection.{ext}"
            path.write_text(hwpml, encoding="utf-8")
            paths.append(path)
        report["written_paths"] = [str(path) for path in paths]
        report["captured_text"] = extract_hwpml_text(paths[0])
        report["identity_after_gettext"] = document_identity(hwp)

        insert_path = paths[0]
        if args.replacement.strip():
            modified_path = run_dir / "selection.modified.hml"
            result = create_hwpml_fragment_from_selection(paths[0], args.replacement, modified_path)
            insert_path = modified_path
            report["modify_hwpml"] = {
                "ok": True,
                "path": str(modified_path),
                "original_text": result.original_text,
                "replacement_text": result.replacement_text,
                "modified_text": extract_hwpml_text(modified_path),
                "char_count": result.char_count,
            }

        if args.apply:
            insert_results = []
            delete_selection(hwp)
            for path in [insert_path]:
                row = {"path": str(path)}
                try:
                    insert_file(hwp, path)
                    row["ok"] = True
                    insert_results.append(row)
                    break
                except Exception as exc:
                    row.update({"ok": False, "error": f"{type(exc).__name__}: {exc}"})
                    insert_results.append(row)
            report["insert_results"] = insert_results
            report["identity_after_insert"] = document_identity(hwp)
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
    return {
        "GetCurFileName": safe_call(lambda: hwp.GetCurFileName()),
        "GetFileName": safe_call(lambda: hwp.GetFileName()),
        "Path": safe_call(lambda: getattr(hwp, "Path")),
        "FullName": safe_call(lambda: getattr(hwp, "FullName")),
    }


def delete_selection(hwp) -> None:
    hwp.Run("Delete")


def insert_file(hwp, path: Path) -> None:
    try:
        hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
        params = hwp.HParameterSet.HInsertFile
        params.FileName = str(path)
        hwp.HAction.Execute("InsertFile", params.HSet)
        return
    except Exception:
        pass
    hwp.InsertFile(str(path))


def safe_call(callback):
    try:
        value = callback()
        return "" if value is None else str(value)
    except Exception as exc:
        return f"<{type(exc).__name__}: {exc}>"


if __name__ == "__main__":
    raise SystemExit(main())
