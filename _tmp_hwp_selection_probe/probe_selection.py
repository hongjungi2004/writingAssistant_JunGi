from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import time
import uuid


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CLIENT_INPUT = PROJECT_ROOT / "client" / "input"
if str(CLIENT_INPUT) not in sys.path:
    sys.path.insert(0, str(CLIENT_INPUT))

from hwpx_document import create_hwpx_fragment_from_match, extract_hwpx_text, find_hwpx_text_matches


LOG_ROOT = PROJECT_ROOT / ".logs" / "hwp_selection_probe"
PROGIDS = ("HWPFrame.HwpObject.2", "HWPFrame.HwpObject.1", "HWPFrame.HwpObject")


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe HWP full-HWPX selection anchoring strategies.")
    parser.add_argument("--delay", type=float, default=0.0, help="Seconds to wait before probing.")
    parser.add_argument("--replacement", default="", help="Replacement text for positioned dry-run fragment.")
    parser.add_argument("--apply", action="store_true", help="Delete current selection and InsertFile generated fragment.")
    parser.add_argument("--marker-test", action="store_true", help="Briefly insert unique markers, save, then undo.")
    args = parser.parse_args()

    if args.delay:
        print(f"{args.delay:g}초 안에 한글에서 대상 범위를 드래그 선택해 주세요...")
        time.sleep(args.delay)

    run_dir = LOG_ROOT / time.strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    report: dict = {
        "run_dir": str(run_dir),
        "marker_test": args.marker_test,
        "apply": args.apply,
        "steps": [],
        "errors": [],
    }

    try:
        hwp = active_hwp_object()
        if hwp is None:
            raise RuntimeError("active HWP COM object not found")
        report["steps"].append({"name": "connect_hwp", "ok": True})

        selection_texts = read_selection_texts(hwp)
        report["selection_texts"] = selection_texts
        selected_text = best_selection_text(selection_texts)
        report["selected_text"] = selected_text
        if not selected_text.strip():
            raise RuntimeError("selected text is empty; drag-select text in HWP first")

        position_report = probe_positions(hwp)
        report["position_probe"] = position_report

        full_path = run_dir / "full_document.hwpx"
        save_report = save_full_hwpx(hwp, full_path)
        report["full_hwpx"] = save_report
        full_text = extract_hwpx_text(full_path)
        (run_dir / "full_document_text.txt").write_text(full_text, encoding="utf-8")
        report["full_text_length"] = len(full_text)
        report["full_text_preview"] = full_text[:500]

        matches = find_hwpx_text_matches(full_path, selected_text)
        report["full_hwpx_matches"] = [
            {
                "section_path": item.section_path,
                "start": item.start,
                "end": item.end,
                "start_paragraph": getattr(item, "start_paragraph", -1),
                "end_paragraph": getattr(item, "end_paragraph", -1),
                "start_paragraph_offset": getattr(item, "start_paragraph_offset", -1),
                "end_paragraph_offset": getattr(item, "end_paragraph_offset", -1),
            }
            for item in matches
        ]
        report["match_count"] = len(matches)
        report["match_assessment"] = assess_matches(matches, position_report)
        report["position_selected_match"] = selected_match_by_position(matches, position_report)
        report["main_app_selection_policy"] = main_app_selection_policy(matches, position_report)

        replacement = normalize_text(args.replacement).strip()
        if replacement:
            fragment_report = create_positioned_fragment(
                full_path=full_path,
                selected_text=selected_text,
                replacement_text=replacement,
                position_report=position_report,
                run_dir=run_dir,
            )
            report["positioned_fragment"] = fragment_report
            if fragment_report.get("ok"):
                report["main_app_selection_policy"] = {
                    "would_use_full_hwpx": True,
                    "reason": "position_range_fallback" if not matches else "positioned_fragment_ok",
                    "fragment_match": fragment_report.get("match"),
                }
            if args.apply:
                if not fragment_report.get("ok"):
                    raise RuntimeError("positioned fragment was not created")
                delete_selection(hwp)
                insert_file(hwp, Path(fragment_report["path"]))
                report["apply_result"] = {"ok": True, "method": "delete_selection_then_insertfile"}

        if args.marker_test:
            report["marker_probe"] = marker_probe(hwp, run_dir)

        report["status"] = "ok"
    except Exception as exc:
        report["errors"].append(f"{type(exc).__name__}: {exc}")
        report["status"] = "fail"

    return finish(report)


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


def read_selection_texts(hwp) -> list[dict]:
    rows = []
    getter = getattr(hwp, "GetTextFile", None)
    if not callable(getter):
        return [{"option": None, "ok": False, "error": "GetTextFile not callable"}]
    for option in ("saveblock", "selection", ""):
        try:
            text = normalize_text(str(getter("TEXT", option) or ""))
            rows.append({"option": option, "ok": True, "length": len(text), "text": text, "preview": text[:200]})
        except Exception as exc:
            rows.append({"option": option, "ok": False, "error": f"{type(exc).__name__}: {exc}"})
    return rows


def best_selection_text(rows: list[dict]) -> str:
    usable = [row for row in rows if row.get("ok") and str(row.get("text") or "").strip()]
    if not usable:
        return ""
    preferred = [row for row in usable if row.get("option") in {"saveblock", "selection"}]
    return str(min(preferred or usable, key=lambda row: len(str(row.get("text") or ""))).get("text") or "")


def probe_positions(hwp) -> dict:
    report = {"snapshots": [], "callables": {}}
    for name in ("GetPos", "GetPosBySet", "GetSelectedPos", "GetCaretPos", "KeyIndicator"):
        attr = getattr(hwp, name, None)
        report["callables"][name] = callable(attr)
        if not callable(attr):
            continue
        try:
            value = attr()
            row = {"method": name, "ok": True, "value": safe_repr(value)}
            if name == "GetSelectedPos":
                row["parsed"] = parse_selected_pos(value)
            report["snapshots"].append(row)
        except Exception as exc:
            report["snapshots"].append({"method": name, "ok": False, "error": f"{type(exc).__name__}: {exc}"})
    return report


def save_full_hwpx(hwp, path: Path) -> dict:
    if path.exists():
        path.unlink()
    result = hwp.SaveAs(str(path), "HWPX", "")
    return {
        "path": str(path),
        "result": repr(result),
        "exists": path.exists(),
        "size": path.stat().st_size if path.exists() else 0,
    }


def assess_matches(matches: list, position_report: dict | None = None) -> str:
    if len(matches) == 1:
        return "unique_by_text"
    if len(matches) == 0:
        return "not_found"
    if selected_match_by_position(matches, position_report):
        return "unique_by_position"
    return "ambiguous_by_text"


def main_app_selection_policy(matches: list, position_report: dict | None) -> dict:
    selected = selected_match_by_position(matches, position_report)
    if len(matches) == 1:
        return {"would_use_full_hwpx": True, "reason": "unique_text_match"}
    if selected:
        return {"would_use_full_hwpx": True, "reason": selected.get("reason"), "selected_match": selected}
    return {
        "would_use_full_hwpx": False,
        "reason": "ambiguous_without_position_match" if matches else "no_text_match",
    }


def selected_match_by_position(matches: list, position_report: dict | None) -> dict:
    selection_pos = selected_pos_from_report(position_report or {})
    if not selection_pos:
        return {}
    start_para = selection_pos.get("start_para")
    start_pos = selection_pos.get("start_pos")
    exact = [
        match for match in matches
        if getattr(match, "start_paragraph", -1) == start_para
        and getattr(match, "start_paragraph_offset", -1) == start_pos
    ]
    if len(exact) == 1:
        match = exact[0]
        return {
            "reason": "paragraph_and_offset",
            "section_path": match.section_path,
            "start": match.start,
            "end": match.end,
            "start_paragraph": getattr(match, "start_paragraph", -1),
            "start_paragraph_offset": getattr(match, "start_paragraph_offset", -1),
        }
    paragraph_only = [
        match for match in matches
        if getattr(match, "start_paragraph", -1) == start_para
    ]
    if len(paragraph_only) == 1:
        match = paragraph_only[0]
        return {
            "reason": "paragraph_only",
            "section_path": match.section_path,
            "start": match.start,
            "end": match.end,
            "start_paragraph": getattr(match, "start_paragraph", -1),
            "start_paragraph_offset": getattr(match, "start_paragraph_offset", -1),
        }
    return {}


def selected_pos_from_report(position_report: dict) -> dict:
    for row in position_report.get("snapshots") or []:
        if row.get("method") != "GetSelectedPos" or not row.get("ok"):
            continue
        if isinstance(row.get("parsed"), dict) and row["parsed"]:
            return row["parsed"]
        parsed = parse_selected_pos_repr(str(row.get("value") or ""))
        if parsed:
            return parsed
    return {}


def parse_selected_pos(value) -> dict:
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
        }
    except Exception:
        return {}


def parse_selected_pos_repr(value: str) -> dict:
    text = value.strip()
    if not text.startswith("(") or not text.endswith(")"):
        return {}
    parts = [part.strip() for part in text[1:-1].split(",")]
    if len(parts) < 7 or parts[0] != "True":
        return {}
    try:
        return {
            "start_list": int(parts[1]),
            "start_para": int(parts[2]),
            "start_pos": int(parts[3]),
            "end_list": int(parts[4]),
            "end_para": int(parts[5]),
            "end_pos": int(parts[6]),
        }
    except Exception:
        return {}


def marker_probe(hwp, run_dir: Path) -> dict:
    marker_id = uuid.uuid4().hex
    start_marker = f"__WA_START_{marker_id}__"
    end_marker = f"__WA_END_{marker_id}__"
    report = {"start_marker": start_marker, "end_marker": end_marker, "undo_attempted": False}
    try:
        insert_text(hwp, start_marker)
        insert_text(hwp, end_marker)
        marked_path = run_dir / "marked_document.hwpx"
        hwp.SaveAs(str(marked_path), "HWPX", "")
        report["marked_path"] = str(marked_path)
        report["marked_exists"] = marked_path.exists()
        if marked_path.exists():
            marked_text = extract_hwpx_text(marked_path)
            report["start_found"] = start_marker in marked_text
            report["end_found"] = end_marker in marked_text
            report["marked_text_preview"] = marked_text[:500]
            (run_dir / "marked_document_text.txt").write_text(marked_text, encoding="utf-8")
    finally:
        report["undo_attempted"] = True
        for _ in range(2):
            try:
                hwp.Run("Undo")
            except Exception:
                pass
    return report


def create_positioned_fragment(
    full_path: Path,
    selected_text: str,
    replacement_text: str,
    position_report: dict,
    run_dir: Path,
) -> dict:
    selection_pos = selected_pos_from_report(position_report)
    start_para = selection_pos.get("start_para")
    start_pos = selection_pos.get("start_pos")
    fragment_path = run_dir / "positioned_fragment.modified.hwpx"
    report = {
        "ok": False,
        "path": str(fragment_path),
        "replacement_length": len(replacement_text),
        "preferred_start_paragraph": start_para,
        "preferred_start_offset": start_pos,
    }
    try:
        result = create_hwpx_fragment_from_match(
            full_path,
            selected_text,
            replacement_text,
            fragment_path,
            preferred_start_paragraph=start_para,
            preferred_start_offset=start_pos,
            preferred_end_paragraph=selection_pos.get("end_para"),
            preferred_end_offset=selection_pos.get("end_pos"),
        )
        fragment_text = extract_hwpx_text(fragment_path)
        (run_dir / "positioned_fragment_text.txt").write_text(fragment_text, encoding="utf-8")
        report.update({
            "ok": True,
            "section_path": result.section_path,
            "match_start": result.match_start,
            "match_end": result.match_end,
            "fragment_text_length": len(fragment_text),
            "fragment_text_preview": fragment_text[:500],
        })
    except Exception as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"
    return report


def insert_text(hwp, text: str) -> None:
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


def delete_selection(hwp) -> None:
    for action in ("Delete", "DeleteBack", "Erase"):
        try:
            hwp.Run(action)
            return
        except Exception:
            pass
    raise RuntimeError("selection delete failed")


def insert_file(hwp, path: Path) -> None:
    try:
        hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
        params = hwp.HParameterSet.HInsertFile
        params.FileName = str(path)
        for name, value in (("KeepSection", 0), ("KeepCharshape", 1), ("KeepParashape", 1), ("KeepStyle", 1)):
            try:
                setattr(params, name, value)
            except Exception:
                pass
        result = hwp.HAction.Execute("InsertFile", params.HSet)
        if result is False:
            raise RuntimeError("InsertFile HAction returned False")
        return
    except Exception:
        hwp.InsertFile(str(path), "HWPX", "KeepSection:0;KeepCharshape:1;KeepParashape:1;KeepStyle:1")


def normalize_text(text: str) -> str:
    return (
        str(text or "")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\v", "\n")
        .replace("\f", "\n")
        .replace("\u2028", "\n")
        .replace("\u2029", "\n")
    )


def safe_repr(value) -> str:
    try:
        return repr(value)
    except Exception:
        return f"<unreprable {type(value).__name__}>"


def finish(report: dict) -> int:
    path = Path(report["run_dir"]) / "report.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    report["report_path"] = str(path)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not report.get("errors") else 1


if __name__ == "__main__":
    raise SystemExit(main())
