from __future__ import annotations

from dataclasses import dataclass
import difflib
import html
from pathlib import Path
import re


P_NODE_RE = re.compile(
    r"(?P<open><P(?:\s[^<>]*?)?>)(?P<body>.*?)(?P<close></P>)",
    re.DOTALL | re.IGNORECASE,
)
CHAR_NODE_RE = re.compile(
    r"(?P<open><CHAR(?:\s[^<>]*?)?>)(?P<text>.*?)(?P<close></CHAR>)",
    re.DOTALL | re.IGNORECASE,
)


class HwpmlApplyError(RuntimeError):
    pass


@dataclass(frozen=True)
class HwpmlApplyResult:
    output_path: Path
    original_text: str
    replacement_text: str
    char_count: int


@dataclass(frozen=True)
class _Paragraph:
    text: str
    run_texts: list[str]


def extract_hwpml_text(path_or_text: str | Path) -> str:
    data = _read_path_or_text(path_or_text)
    paragraphs = _parse_paragraphs(data)
    if paragraphs:
        return "\n".join(paragraph.text for paragraph in paragraphs)
    return "\n".join(_decode_char_text(match.group("text")) for match in CHAR_NODE_RE.finditer(data))


def create_hwpml_fragment_from_selection(
    hwpml_path: str | Path,
    replacement_text: str,
    output_path: str | Path,
) -> HwpmlApplyResult:
    path = Path(hwpml_path)
    data = path.read_text(encoding="utf-8")
    paragraphs = _parse_paragraphs(data)
    if not paragraphs:
        raise HwpmlApplyError("HWPML2X CHAR nodes were not found.")

    original_lines = [paragraph.text for paragraph in paragraphs]
    original_text = "\n".join(original_lines)
    replacement_text = _strip_visual_prefixes_from_replacement(original_lines, replacement_text)
    replacement_lines = _project_replacement_lines(original_lines, replacement_text)
    paragraph_iter = iter(replacement_lines)
    char_count = sum(len(paragraph.run_texts) for paragraph in paragraphs)

    def replace_paragraph(match: re.Match) -> str:
        body = match.group("body")
        char_matches = list(CHAR_NODE_RE.finditer(body))
        if not char_matches:
            return match.group(0)
        paragraph_text = next(paragraph_iter)
        run_texts = [_decode_char_text(item.group("text")) for item in char_matches]
        projected_runs = _project_paragraph_runs(run_texts, paragraph_text)
        run_iter = iter(projected_runs)

        def replace_char(char_match: re.Match) -> str:
            return f"{char_match.group('open')}{_encode_char_text(next(run_iter))}{char_match.group('close')}"

        updated_body = CHAR_NODE_RE.sub(replace_char, body, count=len(char_matches))
        return f"{match.group('open')}{updated_body}{match.group('close')}"

    updated = P_NODE_RE.sub(replace_paragraph, data, count=len(paragraphs))
    output = Path(output_path)
    output.write_text(updated, encoding="utf-8")
    return HwpmlApplyResult(
        output_path=output,
        original_text=original_text,
        replacement_text="\n".join(replacement_lines),
        char_count=char_count,
    )


def _parse_paragraphs(data: str) -> list[_Paragraph]:
    paragraphs: list[_Paragraph] = []
    for paragraph_match in P_NODE_RE.finditer(data):
        run_texts = [
            _decode_char_text(char_match.group("text"))
            for char_match in CHAR_NODE_RE.finditer(paragraph_match.group("body"))
        ]
        if run_texts:
            paragraphs.append(_Paragraph("".join(run_texts), run_texts))
    return paragraphs


def _read_path_or_text(path_or_text: str | Path) -> str:
    path = Path(path_or_text)
    if path.exists():
        return path.read_text(encoding="utf-8")
    return str(path_or_text)


def _project_replacement_lines(original_lines: list[str], replacement_text: str) -> list[str]:
    replacement_lines = _normalize_text(replacement_text).split("\n")
    if len(replacement_lines) == len(original_lines):
        return replacement_lines
    if len(replacement_lines) < len(original_lines):
        return replacement_lines + [""] * (len(original_lines) - len(replacement_lines))

    projected = replacement_lines[: len(original_lines) - 1]
    projected.append("\n".join(replacement_lines[len(original_lines) - 1:]))
    return projected


def _project_paragraph_runs(original_runs: list[str], replacement_text: str) -> list[str]:
    if not original_runs:
        return []
    original_text = "".join(original_runs)
    if len(original_runs) == 1:
        return [replacement_text]
    if replacement_text == original_text:
        return list(original_runs)
    if replacement_text.startswith(original_text):
        projected = list(original_runs)
        projected[_last_non_empty_index(projected)] += replacement_text[len(original_text):]
        return projected

    projected = ["" for _ in original_runs]
    spans = _run_spans(original_runs)
    matcher = difflib.SequenceMatcher(None, original_text, replacement_text, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for run_index, run_start, run_end in spans:
                overlap_start = max(i1, run_start)
                overlap_end = min(i2, run_end)
                if overlap_start < overlap_end:
                    repl_start = j1 + (overlap_start - i1)
                    repl_end = j1 + (overlap_end - i1)
                    projected[run_index] += replacement_text[repl_start:repl_end]
        elif tag in {"insert", "replace"}:
            projected[_run_index_for_insert(spans, i1)] += replacement_text[j1:j2]
    return projected


def _run_spans(runs: list[str]) -> list[tuple[int, int, int]]:
    spans: list[tuple[int, int, int]] = []
    cursor = 0
    for index, text in enumerate(runs):
        spans.append((index, cursor, cursor + len(text)))
        cursor += len(text)
    return spans


def _run_index_for_insert(spans: list[tuple[int, int, int]], offset: int) -> int:
    if not spans:
        return 0
    for index, start, end in spans:
        if start <= offset < end:
            return index
    for index, _start, end in reversed(spans):
        if offset >= end:
            return index
    return spans[0][0]


def _last_non_empty_index(values: list[str]) -> int:
    for index in range(len(values) - 1, -1, -1):
        if values[index]:
            return index
    return len(values) - 1


def _strip_visual_prefixes_from_replacement(original_lines: list[str], replacement_text: str) -> str:
    replacement_lines = _normalize_text(replacement_text).split("\n")
    if not original_lines or not replacement_lines:
        return replacement_text
    for index, original_line in enumerate(original_lines[: len(replacement_lines)]):
        replacement_lines[index] = _strip_visual_prefix_for_line(original_line, replacement_lines[index])
    return "\n".join(replacement_lines)


def _strip_visual_prefix_for_line(original_line: str, replacement_line: str) -> str:
    original = original_line.strip()
    if not original:
        return replacement_line
    candidates = _visual_prefix_strip_candidates(replacement_line)
    original_core = _strip_all_visual_prefixes(original)
    best = replacement_line
    best_score = _visual_prefix_alignment_score(original, original_core, replacement_line)
    for candidate in candidates[1:]:
        score = _visual_prefix_alignment_score(original, original_core, candidate)
        if score > best_score:
            best = candidate
            best_score = score
    return best


def _visual_prefix_alignment_score(original: str, original_core: str, candidate: str) -> int:
    stripped = candidate.strip()
    candidate_core = _strip_all_visual_prefixes(stripped)
    if stripped.startswith(original):
        return 100
    if original_core and candidate_core.startswith(original_core):
        return 70
    if original in stripped:
        return 40
    if original_core and original_core in candidate_core:
        return 20
    return 0


def _visual_prefix_strip_candidates(text: str) -> list[str]:
    candidates = [text]
    current = text
    while True:
        stripped = _strip_one_visual_prefix(current)
        if stripped == current:
            break
        candidates.append(stripped)
        current = stripped
    return candidates


def _strip_one_visual_prefix(text: str) -> str:
    return re.sub(r"^(\s*)(?:\d+[.)]|[\uac00-\ud7a3A-Za-z][.)]|[\u00b7\u2022\u25aa\-])\s*", r"\1", text, count=1)


def _strip_all_visual_prefixes(text: str) -> str:
    current = text
    while True:
        stripped = _strip_one_visual_prefix(current).strip()
        if stripped == current.strip():
            return stripped
        current = stripped


def _decode_char_text(text: str) -> str:
    return html.unescape(text or "")


def _encode_char_text(text: str) -> str:
    return html.escape(_normalize_text(text), quote=False)


def _normalize_text(text: str) -> str:
    return (
        html.unescape(str(text or ""))
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\v", "\n")
    )
