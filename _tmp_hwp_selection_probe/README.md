# HWP Selection Probe

Isolated runtime probes for selecting a precise range inside a full-document
HWPX export. Nothing here is imported by the main app.

Run while HWP is open and the target text is drag-selected:

```powershell
python _tmp_hwp_selection_probe/probe_selection.py --delay 3
```

To create a positioned replacement fragment without applying it:

```powershell
python _tmp_hwp_selection_probe/probe_selection.py --delay 3 --replacement "교정문"
```

To apply the generated fragment to the current selected range, use a copied test
document first:

```powershell
python _tmp_hwp_selection_probe/probe_selection.py --delay 3 --replacement "교정문" --apply
```

Probe order:

1. Save the full document as HWPX.
2. Read the selected text with `GetTextFile`.
3. Try non-destructive COM position snapshots.
4. Find candidate matches in the full HWPX text.
5. Optionally create a positioned fragment from the full HWPX.
6. Optionally test marker anchoring with `--marker-test`.

Marker mode briefly modifies the document and calls Undo. Use only on a copy
until the report looks safe.

To check whether an export method changes the active HWP document path:

```powershell
python _tmp_hwp_selection_probe/probe_export_identity.py --delay 3
```

This first tries `GetTextFile` formats only. To also test risky `SaveAs` export:

```powershell
python _tmp_hwp_selection_probe/probe_export_identity.py --delay 3 --saveas
```

Use `--saveas` only on a copied test document.
