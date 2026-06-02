from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import tempfile
import time

import pythoncom
import win32com.client as win32

from hwpx_document import create_hwpx_fragment_from_match, extract_hwpx_text


PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOG_ROOT = PROJECT_ROOT / ".logs" / "hwp_insertfile_probe"
PROGIDS = ("HWPFrame.HwpObject.2", "HWPFrame.HwpObject.1", "HWPFrame.HwpObject")


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe HWP selection HWPX save/edit/InsertFile flow.")
    parser.add_argument("--delay", type=float, default=0.0)
    parser.add_argument("--source", default="", help="Expected selected source text. If omitted, use saved selection text.")
    parser.add_argument("--replacement", default="", help="Replacement text.")
    parser.add_argument("--apply", action="store_true", help="Delete selection and InsertFile modified HWPX.")
    args = parser.parse_args()

    if args.delay:
        print(f"{args.delay:g}초 안에 한글에서 대상 범위를 드래그 선택해 주세요...")
        time.sleep(args.delay)

    run_dir = LOG_ROOT / time.strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    report = {"run_dir": str(run_dir), "apply": args.apply, "errors": []}

    try:
        hwp = active_hwp_object()
        if hwp is None:
            raise RuntimeError("active HWP COM object not found")
        report["com"] = "ok"

        save_result = save_document_as_hwpx(hwp, run_dir)
        report["save_document"] = save_result
        document_path = Path(save_result["selected_path"])
        document_text = extract_hwpx_text(document_path).strip()
        report["document_text_length"] = len(document_text)
        report["document_text_preview"] = document_text[:500]

        source = args.source.strip() or document_text
        replacement = args.replacement.strip()
        report["source_length"] = len(source)
        report["replacement_length"] = len(replacement)
        if not replacement:
            report["status"] = "dry_run_no_replacement"
            return finish(report)
        if source not in document_text:
            raise RuntimeError("source text is not contained in saved document HWPX")

        modified_path = run_dir / "fragment.modified.hwpx"
        apply_result = create_hwpx_fragment_from_match(document_path, source, replacement, modified_path)
        report["modify_hwpx"] = {
            "ok": True,
            "path": str(modified_path),
            "section_path": apply_result.section_path,
            "match_start": apply_result.match_start,
            "match_end": apply_result.match_end,
        }
        report["modified_text_preview"] = extract_hwpx_text(modified_path).strip()[:500]

        if args.apply:
            delete_selection(hwp)
            insert_file(hwp, modified_path)
            report["insert_file"] = {"ok": True}
            report["status"] = "applied"
        else:
            report["status"] = "dry_run_modified"
    except Exception as exc:
        report["errors"].append(f"{type(exc).__name__}: {exc}")
        report["status"] = "fail"

    return finish(report)


def finish(report: dict) -> int:
    path = Path(report["run_dir"]) / "report.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    report["report_path"] = str(path)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not report.get("errors") else 1


def active_hwp_object():
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
            name = moniker.GetDisplayName(bind_context, None)
        except Exception:
            name = ""
        if "hwp" not in str(name).lower():
            continue
        obj = rot.GetObject(moniker)
        query = getattr(obj, "QueryInterface", None)
        if callable(query):
            try:
                return win32.dynamic.Dispatch(query(pythoncom.IID_IDispatch))
            except Exception:
                pass
    return None


def save_document_as_hwpx(hwp, run_dir: Path) -> dict:
    attempts = []
    for option in ("saveblock", "selection", ""):
        suffix = option or "default"
        path = run_dir / f"document_{suffix}.hwpx"
        try:
            result = hwp.SaveAs(str(path), "HWPX", option)
            row = {
                "option": option,
                "path": str(path),
                "result": repr(result),
                "exists": path.exists(),
                "size": path.stat().st_size if path.exists() else 0,
            }
            if row["exists"] and row["size"] > 0:
                try:
                    text = extract_hwpx_text(path).strip()
                    row["text_length"] = len(text)
                    row["text_preview"] = text[:300]
                except Exception as exc:
                    row["text_error"] = f"{type(exc).__name__}: {exc}"
            attempts.append(row)
        except Exception as exc:
            attempts.append({"option": option, "error": f"{type(exc).__name__}: {exc}"})
    usable = [row for row in attempts if row.get("exists") and row.get("size", 0) > 0]
    if not usable:
        raise RuntimeError(f"selection SaveAs HWPX failed: {attempts}")
    selected = min(usable, key=lambda row: row.get("text_length", 10**9))
    return {"ok": True, "selected_option": selected["option"], "selected_path": selected["path"], "attempts": attempts}


def delete_selection(hwp) -> None:
    hwp.Run("Delete")


def insert_file(hwp, path: Path) -> None:
    hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
    params = hwp.HParameterSet.HInsertFile
    params.FileName = str(path)
    for name, value in (("KeepSection", 0), ("KeepCharshape", 1), ("KeepParashape", 1), ("KeepStyle", 1)):
        try:
            setattr(params, name, value)
        except Exception:
            pass
    hwp.HAction.Execute("InsertFile", params.HSet)


if __name__ == "__main__":
    raise SystemExit(main())
