from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import difflib
import html
import re
import shutil
import zipfile

SECTION_PATH_RE = re.compile(r"^Contents/section\d+\.xml$", re.IGNORECASE)
TEXT_NODE_RE = re.compile(
    rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?t(?:\s[^<>]*?)?>"
    rb"(?P<text>.*?)"
    rb"</(?P=prefix)?t\s*>",
    re.DOTALL,
)
SOFT_BREAK_RE = re.compile(
    rb"<(?:[A-Za-z_][\w.-]*:)?(?:lineBreak|br)(?:\s[^<>]*?)?(?:/>|>\s*</(?:[A-Za-z_][\w.-]*:)?(?:lineBreak|br)\s*>)",
    re.IGNORECASE,
)
TEXT_OR_BREAK_RE = re.compile(TEXT_NODE_RE.pattern + rb"|" + SOFT_BREAK_RE.pattern, re.DOTALL | re.IGNORECASE)
PARAGRAPH_START_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?p(?:\s|>)", re.IGNORECASE)
SECTION_CLOSE_RE = re.compile(rb"</(?:[A-Za-z_][\w.-]*:)?sec\s*>", re.IGNORECASE)
RUN_RE = re.compile(
    rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?run(?P<attrs>[^<>]*?)>"
    rb"(?P<body>.*?)"
    rb"</(?P=prefix)?run\s*>",
    re.DOTALL | re.IGNORECASE,
)
EMPTY_TEXT_NODE_RE = re.compile(
    rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?t(?:\s[^<>]*?)?>\s*</(?P=prefix)?t\s*>",
    re.DOTALL | re.IGNORECASE,
)
TAG_RE = re.compile(rb"<[^<>]+>")


@dataclass
class PartRef:
    slot_index: int
    part_index: int
    line_index: int
    text: str


@dataclass
class TextSlot:
    xml_start: int
    xml_end: int
    original_payload: bytes
    parts: list[str]
    break_tags: list[bytes]


@dataclass
class ParsedSection:
    path: str
    xml: bytes
    slots: list[TextSlot]
    lines: list[list[PartRef]]


def create_rebuilt_hwpx_fragment(
    document_path: str | Path,
    replacement_text: str,
    output_path: str | Path,
    source_text: str | None = None,
) -> Path:
    """Create a *selection fragment* using a Word-like line/slot rebuild strategy.

    Important: ``document_path`` may be a full temporary HWPX export.  InsertFile
    must never receive that full document.  This function rewrites only the
    matched selected line range, then crops the HWPX section down to that range
    before writing ``output_path``.
    """
    src = Path(document_path)
    out = Path(output_path)
    if not src.exists():
        raise FileNotFoundError(src)

    with zipfile.ZipFile(src, "r") as archive:
        section_names = _section_names(archive)
        sections = [_parse_section(name, archive.read(name)) for name in section_names]

    section = _select_section(sections, source_text)
    if section is None:
        raise RuntimeError("No editable HWPX section text was found.")

    updated_xml, start_line, end_line = _rebuild_section_xml_and_range(
        section, replacement_text, source_text=source_text
    )
    fragment_xml = _section_fragment_for_line_range(section, updated_xml, start_line, end_line)

    shutil.copy2(src, out)
    replacements: dict[str, bytes] = {}
    for parsed in sections:
        if parsed.path == section.path:
            replacements[parsed.path] = fragment_xml
        else:
            replacements[parsed.path] = _empty_section_xml(parsed.xml)
    _replace_zip_members(out, replacements)
    return out


def rebuild_section_xml(section: ParsedSection, replacement_text: str, source_text: str | None = None) -> bytes:
    updated, _start_line, _end_line = _rebuild_section_xml_and_range(
        section, replacement_text, source_text=source_text
    )
    return updated


def _rebuild_section_xml_and_range(
    section: ParsedSection, replacement_text: str, source_text: str | None = None
) -> tuple[bytes, int, int]:
    all_source_lines = [_line_text(line) for line in section.lines]
    start_line, end_line = _find_selected_line_range(all_source_lines, source_text)
    selected_lines = all_source_lines[start_line:end_line]
    replacement_lines = _project_replacement_to_lines(selected_lines, replacement_text)

    part_outputs: dict[tuple[int, int], str] = {}
    # Default: preserve every existing text part. Only the matched/selected line
    # range is rewritten. This keeps surrounding paragraphs intact before the
    # final fragment crop.
    for line in section.lines:
        for part_ref in line:
            part_outputs[(part_ref.slot_index, part_ref.part_index)] = part_ref.text

    for offset, line in enumerate(section.lines[start_line:end_line]):
        line_replacement = replacement_lines[offset] if offset < len(replacement_lines) else ""
        for part_ref, value in zip(line, _diff_replacement_by_parts(line, line_replacement)):
            part_outputs[(part_ref.slot_index, part_ref.part_index)] = value

    edits: list[tuple[int, int, bytes]] = []
    encoding = _xml_encoding(section.xml)
    for slot_index, slot in enumerate(section.slots):
        new_parts = [part_outputs.get((slot_index, idx), "") for idx in range(len(slot.parts))]
        payload = _encode_slot_payload(new_parts, slot.break_tags, encoding)
        edits.append((slot.xml_start, slot.xml_end, payload))

    updated = section.xml
    for start, end, payload in sorted(edits, key=lambda item: item[0], reverse=True):
        updated = updated[:start] + payload + updated[end:]
    return _cleanup_empty_text_runs(updated), start_line, end_line


def extract_experimental_lines(document_path: str | Path) -> list[str]:
    """Debug helper: return visual lines detected by the experimental parser."""
    with zipfile.ZipFile(document_path, "r") as archive:
        sections = [_parse_section(name, archive.read(name)) for name in _section_names(archive)]
    section = _select_section(sections, None)
    return [_line_text(line) for line in section.lines] if section else []


def _parse_section(path: str, xml: bytes) -> ParsedSection:
    slots: list[TextSlot] = []
    lines: list[list[PartRef]] = [[]]
    line_index = 0
    previous_paragraph_start = -1
    encoding = _xml_encoding(xml)

    for match in TEXT_OR_BREAK_RE.finditer(xml):
        paragraph_start = _nearest_paragraph_start(xml, match.start())
        if paragraph_start >= 0 and previous_paragraph_start >= 0 and paragraph_start != previous_paragraph_start:
            if lines[-1]:
                lines.append([])
                line_index += 1
        previous_paragraph_start = paragraph_start

        token = match.group(0)
        if SOFT_BREAK_RE.fullmatch(token):
            if lines[-1]:
                lines.append([])
                line_index += 1
            continue

        payload = match.groupdict().get("text")
        if payload is None:
            continue
        parts, break_tags = _decode_payload_parts(payload, encoding)
        slot_index = len(slots)
        slots.append(TextSlot(match.start("text"), match.end("text"), payload, parts, break_tags))
        for part_index, part in enumerate(parts):
            if part:
                lines[-1].append(PartRef(slot_index, part_index, line_index, part))
            if part_index < len(break_tags):
                if lines[-1]:
                    lines.append([])
                    line_index += 1

    lines = [line for line in lines if line]
    # Renumber after removing blank structural lines.
    for new_index, line in enumerate(lines):
        for ref in line:
            ref.line_index = new_index
    return ParsedSection(path, xml, slots, lines)


def _select_section(sections: list[ParsedSection], source_text: str | None) -> ParsedSection | None:
    sections = [section for section in sections if section.slots and section.lines]
    if not sections:
        return None
    if source_text:
        key = _compare_key(source_text)
        exact = [section for section in sections if key and key in _compare_key("\n".join(_line_text(line) for line in section.lines))]
        if len(exact) == 1:
            return exact[0]
    return max(sections, key=lambda section: sum(len(_line_text(line)) for line in section.lines))


def _decode_payload_parts(payload: bytes, encoding: str) -> tuple[list[str], list[bytes]]:
    parts: list[str] = []
    break_tags: list[bytes] = []
    cursor = 0
    for match in SOFT_BREAK_RE.finditer(payload):
        parts.append(html.unescape(payload[cursor:match.start()].decode(encoding, errors="replace")))
        break_tags.append(match.group(0))
        cursor = match.end()
    parts.append(html.unescape(payload[cursor:].decode(encoding, errors="replace")))
    return [_normalize_text(part) for part in parts], break_tags


def _encode_slot_payload(parts: list[str], break_tags: list[bytes], encoding: str) -> bytes:
    output = bytearray()
    for index, part in enumerate(parts):
        output.extend(html.escape(_normalize_text(part), quote=False).encode(encoding))
        if index < len(break_tags):
            output.extend(break_tags[index] or b"<hp:lineBreak/>")
    return bytes(output)


def _cleanup_empty_text_runs(xml: bytes) -> bytes:
    without_empty_text = EMPTY_TEXT_NODE_RE.sub(b"", xml)

    def replace_run(match: re.Match[bytes]) -> bytes:
        body = match.group("body") or b""
        if _run_body_has_non_text_payload(body):
            return match.group(0)
        return b""

    return RUN_RE.sub(replace_run, without_empty_text)


def _run_body_has_non_text_payload(body: bytes) -> bool:
    if not body.strip():
        return False
    if TAG_RE.sub(b"", body).strip():
        return True
    return bool(TAG_RE.search(body))


def _line_text(line: list[PartRef]) -> str:
    return "".join(ref.text for ref in line)



def _find_selected_line_range(source_lines: list[str], source_text: str | None) -> tuple[int, int]:
    if not source_text:
        return 0, len(source_lines)
    needle = _compare_key(source_text)
    if not needle:
        return 0, len(source_lines)

    best: tuple[int, int] | None = None
    best_score = -1
    max_span = min(len(source_lines), 30)
    for start in range(len(source_lines)):
        combined = ""
        for end in range(start + 1, min(len(source_lines), start + max_span) + 1):
            combined = _compare_key("\n".join(source_lines[start:end]))
            if not combined:
                continue
            score = -1
            if combined == needle:
                score = 10_000 - (end - start)
            elif needle in combined:
                score = 8_000 - (len(combined) - len(needle)) - (end - start)
            elif combined in needle:
                score = 6_000 - (len(needle) - len(combined)) - (end - start)
            else:
                ratio = difflib.SequenceMatcher(None, combined, needle, autojunk=False).ratio()
                if ratio >= 0.82:
                    score = int(ratio * 1000) - (end - start)
            if score > best_score:
                best_score = score
                best = (start, end)
    if best is not None and best_score >= 700:
        return best
    return 0, len(source_lines)

def _project_replacement_to_lines(source_lines: list[str], replacement_text: str) -> list[str]:
    replacement = _normalize_text(replacement_text)
    if not source_lines:
        return [replacement]
    explicit = replacement.split("\n")
    if len(explicit) == len(source_lines):
        return explicit

    # HWP GetTextFile often injects blank paragraph separators that are not real
    # editable HWPX lines. Ignore empty lines before projecting the correction.
    nonempty_explicit = [line for line in explicit if line.strip()]
    if len(nonempty_explicit) == len(source_lines):
        return nonempty_explicit
    if len(source_lines) == 1:
        return [" ".join(nonempty_explicit) if nonempty_explicit else replacement.replace("\n", "")]

    collapsed = " ".join(nonempty_explicit) if nonempty_explicit else replacement.replace("\n", " ")
    source_joined = " ".join(source_lines)
    boundaries: list[int] = []
    cursor = 0
    for line in source_lines[:-1]:
        cursor += len(line)
        boundaries.append(cursor)
        cursor += 1  # projected separator in source_joined

    split_points = [_map_source_index_to_replacement(source_joined, collapsed, boundary) for boundary in boundaries]
    result: list[str] = []
    start = 0
    previous = 0
    for point in split_points:
        point = max(previous, min(len(collapsed), point))
        result.append(collapsed[start:point].rstrip())
        start = point
        while start < len(collapsed) and collapsed[start] == " ":
            start += 1
        previous = start
    result.append(collapsed[start:])
    return result


def _diff_replacement_by_parts(line: list[PartRef], replacement_text: str) -> list[str]:
    source = "".join(ref.text for ref in line)
    replacement = _normalize_text(replacement_text)
    if not line:
        return []
    source_part_indexes: list[int] = []
    for part_index, ref in enumerate(line):
        source_part_indexes.extend([part_index] * len(ref.text))

    outputs = ["" for _ in line]
    matcher = difflib.SequenceMatcher(None, source, replacement, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for source_index, char in zip(range(i1, i2), replacement[j1:j2]):
                outputs[source_part_indexes[source_index]] += char
        elif tag == "delete":
            continue
        elif tag == "insert":
            outputs[_part_for_insertion(source_part_indexes, i1, len(line))] += replacement[j1:j2]
        elif tag == "replace":
            for part_index, chunk in _split_replacement_for_parts(source_part_indexes[i1:i2], replacement[j1:j2], len(line)):
                outputs[part_index] += chunk
    return outputs


def _part_for_insertion(source_part_indexes: list[int], source_index: int, part_count: int) -> int:
    if source_index > 0 and source_part_indexes:
        return source_part_indexes[source_index - 1]
    if source_index < len(source_part_indexes):
        return source_part_indexes[source_index]
    return max(0, part_count - 1)


def _split_replacement_for_parts(source_part_indexes: list[int], replacement: str, part_count: int) -> list[tuple[int, str]]:
    if not replacement:
        return []
    if not source_part_indexes:
        return [(_part_for_insertion([], 0, part_count), replacement)]
    ordered: list[int] = []
    counts: dict[int, int] = {}
    for idx in source_part_indexes:
        counts[idx] = counts.get(idx, 0) + 1
        if idx not in ordered:
            ordered.append(idx)
    if len(ordered) == 1:
        return [(ordered[0], replacement)]
    total = len(source_part_indexes)
    cursor = 0
    chunks: list[tuple[int, str]] = []
    for pos, idx in enumerate(ordered):
        if pos == len(ordered) - 1:
            end = len(replacement)
        else:
            end = round(len(replacement) * sum(counts[item] for item in ordered[: pos + 1]) / total)
            end = max(cursor, min(len(replacement), end))
        chunks.append((idx, replacement[cursor:end]))
        cursor = end
    return chunks


def _map_source_index_to_replacement(source: str, replacement: str, source_index: int) -> int:
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
                source_span = max(1, i2 - i1)
                return j1 + round((j2 - j1) * (source_index - i1) / source_span)
            return j1
        if source_index == i2:
            return j2
        previous_j = j2
    return len(replacement)


def _section_names(archive: zipfile.ZipFile) -> list[str]:
    return sorted(name for name in archive.namelist() if SECTION_PATH_RE.match(name))


def _nearest_paragraph_start(xml: bytes, offset: int) -> int:
    last = -1
    for match in PARAGRAPH_START_RE.finditer(xml, 0, offset):
        last = match.start()
    return last


def _section_fragment_for_line_range(
    section: ParsedSection, updated_xml: bytes, start_line: int, end_line: int
) -> bytes:
    selected_lines = section.lines[start_line:end_line]
    affected_slot_indexes = {ref.slot_index for line in selected_lines for ref in line}
    if not affected_slot_indexes:
        raise RuntimeError("Selected HWPX line range did not map to text slots.")

    paragraph_starts = _paragraph_starts(section.xml)
    affected_ordinals: list[int] = []
    for slot_index in sorted(affected_slot_indexes):
        slot = section.slots[slot_index]
        paragraph_start = _nearest_paragraph_start(section.xml, slot.xml_start)
        if paragraph_start in paragraph_starts:
            ordinal = paragraph_starts.index(paragraph_start)
            if ordinal not in affected_ordinals:
                affected_ordinals.append(ordinal)
    if not affected_ordinals:
        raise RuntimeError("Could not find paragraph XML range for selected HWPX lines.")

    updated_starts = _paragraph_starts(updated_xml)
    start_ordinal = min(affected_ordinals)
    end_ordinal = max(affected_ordinals)
    if start_ordinal >= len(updated_starts) or end_ordinal >= len(updated_starts):
        raise RuntimeError("Could not map selected paragraphs after HWPX rebuild.")

    start = updated_starts[start_ordinal]
    next_ordinal = end_ordinal + 1
    end = updated_starts[next_ordinal] if next_ordinal < len(updated_starts) else _section_close_start(updated_xml)
    body = updated_xml[start:end]

    first_paragraph = _first_paragraph_start(updated_xml)
    if first_paragraph < 0:
        raise RuntimeError("Could not find section paragraph start for fragment.")
    section_close = _section_close_start(updated_xml)
    return updated_xml[:first_paragraph] + body + updated_xml[section_close:]


def _empty_section_xml(xml: bytes) -> bytes:
    first_paragraph = _first_paragraph_start(xml)
    section_close = _section_close_start(xml)
    if first_paragraph < 0 or section_close < first_paragraph:
        return xml
    return xml[:first_paragraph] + xml[section_close:]


def _paragraph_starts(xml: bytes) -> list[int]:
    return [match.start() for match in PARAGRAPH_START_RE.finditer(xml)]


def _first_paragraph_start(xml: bytes) -> int:
    match = PARAGRAPH_START_RE.search(xml)
    return match.start() if match else -1


def _section_close_start(xml: bytes) -> int:
    match = SECTION_CLOSE_RE.search(xml)
    return match.start() if match else len(xml)


def _replace_zip_member(path: Path, member_name: str, data: bytes) -> None:
    _replace_zip_members(path, {member_name: data})


def _replace_zip_members(path: Path, replacements: dict[str, bytes]) -> None:
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(temp_path, "w") as dst:
        for item in src.infolist():
            payload = replacements.get(item.filename)
            if payload is None:
                payload = src.read(item.filename)
            dst.writestr(item, payload)
    temp_path.replace(path)


def _xml_encoding(xml: bytes) -> str:
    head = xml[:200]
    match = re.search(rb'encoding=["\']([^"\']+)["\']', head, re.IGNORECASE)
    if not match:
        return "utf-8"
    try:
        return match.group(1).decode("ascii", errors="ignore") or "utf-8"
    except Exception:
        return "utf-8"


def _normalize_text(text: str) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").replace("\f", "\n").replace("\u2028", "\n").replace("\u2029", "\n")


def _compare_key(text: str) -> str:
    return re.sub(r"\s+", "", _normalize_text(text))
