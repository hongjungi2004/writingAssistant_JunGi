from __future__ import annotations

from dataclasses import dataclass
import difflib
from pathlib import Path
import html
import re
import shutil
import tempfile
import zipfile
import unicodedata


SECTION_PATH_RE = re.compile(r"^Contents/section\d+\.xml$", re.IGNORECASE)
TEXT_NODE_RE = re.compile(
    rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?t(?:\s[^<>]*?)?>"
    rb"(?P<text>.*?)"
    rb"</(?P=prefix)?t\s*>",
    re.DOTALL,
)
PARAGRAPH_START_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?p(?:\s|>)", re.IGNORECASE)
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
SOFT_LINE_BREAK_RE = re.compile(
    rb"<(?:[A-Za-z_][\w.-]*:)?(?:lineBreak|br)(?:\s[^<>]*?)?(?:/>|>\s*</(?:[A-Za-z_][\w.-]*:)?(?:lineBreak|br)\s*>)",
    re.IGNORECASE,
)
TEXT_OR_BREAK_RE = re.compile(
    TEXT_NODE_RE.pattern + rb"|" + SOFT_LINE_BREAK_RE.pattern,
    re.DOTALL | re.IGNORECASE,
)
INNER_SOFT_LINE_BREAK_RE = re.compile(
    rb"<(?:[A-Za-z_][\w.-]*:)?(?:lineBreak|br)(?:\s[^<>]*?)?(?:/>|>\s*</(?:[A-Za-z_][\w.-]*:)?(?:lineBreak|br)\s*>)",
    re.IGNORECASE,
)


class HwpxApplyError(RuntimeError):
    pass


@dataclass(frozen=True)
class TextSlot:
    section_path: str
    xml_start: int
    xml_end: int
    text_start: int
    text_end: int
    value: str


@dataclass(frozen=True)
class HwpxTextMatch:
    section_path: str
    start: int
    end: int
    start_paragraph: int = -1
    end_paragraph: int = -1
    start_paragraph_offset: int = -1
    end_paragraph_offset: int = -1


@dataclass(frozen=True)
class HwpxApplyResult:
    document_path: Path
    backup_path: Path
    section_path: str
    match_start: int
    match_end: int


def extract_hwpx_text(document_path: str | Path) -> str:
    path = _validate_hwpx_path(document_path)
    parts: list[str] = []
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            text, _slots = _section_text_and_slots(section_path, archive.read(section_path))
            if text:
                parts.append(text)
    return "\n".join(parts)


def find_hwpx_text_matches(document_path: str | Path, source_text: str) -> list[HwpxTextMatch]:
    path = _validate_hwpx_path(document_path)
    needle = _normalize_text(source_text)
    if not needle:
        return []
    matches: list[HwpxTextMatch] = []
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            xml_bytes = archive.read(section_path)
            section_text, slots = _section_text_and_slots(section_path, xml_bytes)
            for start, end in _find_text_ranges(section_text, needle):
                match_pos = _match_paragraph_position(xml_bytes, slots, start, end)
                matches.append(HwpxTextMatch(section_path, start, end, *match_pos))
    return matches


def apply_text_to_hwpx(document_path: str | Path, source_text: str, replacement_text: str) -> HwpxApplyResult:
    path = _validate_hwpx_path(document_path)
    needle = _normalize_text(source_text)
    if not needle:
        raise HwpxApplyError("Source text is empty.")

    candidates = []
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            xml_bytes = archive.read(section_path)
            section_text, slots = _section_text_and_slots(section_path, xml_bytes)
            for start, end in _find_text_ranges(section_text, needle):
                candidates.append((section_path, xml_bytes, slots, start, end))

    if not candidates:
        fallback = _whole_block_candidate_if_compatible(path, needle)
        if fallback is None:
            raise HwpxApplyError("Source text was not found in the HWPX document.")
        candidates.append(fallback)
    candidate = _select_best_match_candidate(candidates)
    if candidate is None:
        raise HwpxApplyError(f"Source text appears {len(candidates)} times.")

    section_path, xml_bytes, slots, start, end = candidate
    updated_xml = _replace_text_in_section(xml_bytes, slots, start, end, replacement_text)
    backup_path = _create_backup(path)
    _replace_zip_member(path, section_path, updated_xml)
    return HwpxApplyResult(path, backup_path, section_path, start, end)


def create_hwpx_fragment_from_match(
    document_path: str | Path,
    source_text: str,
    replacement_text: str,
    output_path: str | Path,
    preferred_start_paragraph: int | None = None,
    preferred_start_offset: int | None = None,
    preferred_end_paragraph: int | None = None,
    preferred_end_offset: int | None = None,
) -> HwpxApplyResult:
    path = _validate_hwpx_path(document_path)
    output = Path(output_path)
    needle = _normalize_text(source_text)
    if not needle:
        raise HwpxApplyError("Source text is empty.")

    candidates = []
    used_position_fallback = False
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            xml_bytes = archive.read(section_path)
            section_text, slots = _section_text_and_slots(section_path, xml_bytes)
            for start, end in _find_text_ranges(section_text, needle):
                candidates.append((section_path, xml_bytes, slots, start, end))

    if not candidates:
        fallback = _whole_block_candidate_if_compatible(path, needle)
        if fallback is None and preferred_start_paragraph is not None:
            fallback = _position_range_candidate_if_available(
                path,
                needle,
                preferred_start_paragraph=preferred_start_paragraph,
                preferred_start_offset=preferred_start_offset,
                preferred_end_paragraph=preferred_end_paragraph,
                preferred_end_offset=preferred_end_offset,
            )
            used_position_fallback = fallback is not None
        if fallback is None:
            raise HwpxApplyError("Source text was not found in the HWPX document.")
        candidates.append(fallback)
    candidate = _select_best_match_candidate(
        candidates,
        preferred_start_paragraph=preferred_start_paragraph,
        preferred_start_offset=preferred_start_offset,
    )
    if candidate is None:
        raise HwpxApplyError(f"Source text appears {len(candidates)} times.")

    section_path, xml_bytes, slots, start, end = candidate
    if used_position_fallback:
        replacement_text = _strip_visual_prefix_from_position_replacement(
            _candidate_text_from_slots(slots, start, end),
            replacement_text,
        )
    updated_xml = _replace_text_in_section(xml_bytes, slots, start, end, replacement_text)
    fragment_xml = _section_fragment_for_match(xml_bytes, updated_xml, slots, start, end)
    shutil.copy2(path, output)
    _replace_zip_members(
        output,
        _fragment_section_replacements(output, section_path, fragment_xml),
    )
    return HwpxApplyResult(output, output, section_path, start, end)


def _strip_visual_prefix_from_position_replacement(candidate_text: str, replacement_text: str) -> str:
    candidate_lines = _normalize_text(candidate_text).split("\n")
    replacement_lines = _normalize_text(replacement_text).split("\n")
    if not candidate_lines or not replacement_lines:
        return replacement_text
    candidate_first = candidate_lines[0].strip()
    replacement_first = replacement_lines[0].strip()
    if not candidate_first or replacement_first == candidate_first:
        return replacement_text
    if replacement_first.endswith(candidate_first):
        prefix_len = len(replacement_lines[0]) - len(candidate_first)
        visual_prefix = replacement_lines[0][:prefix_len]
        if _looks_like_visual_prefix(visual_prefix):
            replacement_lines[0] = replacement_lines[0][prefix_len:]
            return "\n".join(replacement_lines)
    return replacement_text


def _looks_like_visual_prefix(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    return bool(re.fullmatch(r"(?:[\d]+[.)]\s*|[가-힣A-Za-z][.)]\s*|[·•▪\-]\s*)+", stripped))


def has_hwpx_position_fallback_match(
    document_path: str | Path,
    source_text: str,
    preferred_start_paragraph: int,
    preferred_start_offset: int | None = None,
    preferred_end_paragraph: int | None = None,
    preferred_end_offset: int | None = None,
) -> bool:
    path = _validate_hwpx_path(document_path)
    needle = _normalize_text(source_text)
    if not needle:
        return False
    return _position_range_candidate_if_available(
        path,
        needle,
        preferred_start_paragraph=preferred_start_paragraph,
        preferred_start_offset=preferred_start_offset,
        preferred_end_paragraph=preferred_end_paragraph,
        preferred_end_offset=preferred_end_offset,
    ) is not None


def _position_range_candidate_if_available(
    path: Path,
    source_text: str,
    preferred_start_paragraph: int,
    preferred_start_offset: int | None = None,
    preferred_end_paragraph: int | None = None,
    preferred_end_offset: int | None = None,
):
    candidates = []
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            xml_bytes = archive.read(section_path)
            _section_text, slots = _section_text_and_slots(section_path, xml_bytes)
            candidate = _position_range_candidate_for_section(
                section_path,
                xml_bytes,
                slots,
                source_text=source_text,
                preferred_start_paragraph=preferred_start_paragraph,
                preferred_start_offset=preferred_start_offset,
                preferred_end_paragraph=preferred_end_paragraph,
                preferred_end_offset=preferred_end_offset,
            )
            if candidate is not None:
                candidates.append(candidate)
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        scored = [
            (_position_candidate_score(candidate, source_text), candidate)
            for candidate in candidates
        ]
        scored.sort(key=lambda item: item[0], reverse=True)
        if scored[0][0] >= 0.45 and (len(scored) == 1 or scored[0][0] > scored[1][0] + 0.08):
            return scored[0][1]
        return candidates[0]
    return None


def _position_range_candidate_for_section(
    section_path: str,
    xml_bytes: bytes,
    slots: list[TextSlot],
    source_text: str,
    preferred_start_paragraph: int,
    preferred_start_offset: int | None = None,
    preferred_end_paragraph: int | None = None,
    preferred_end_offset: int | None = None,
):
    paragraphs = _paragraph_text_ranges(xml_bytes, slots)
    if not paragraphs:
        return None
    direct = _paragraph_range_from_position(
        section_path,
        xml_bytes,
        slots,
        paragraphs,
        preferred_start_paragraph,
        preferred_start_offset,
        preferred_end_paragraph,
        preferred_end_offset,
    )
    fuzzy = _best_fuzzy_paragraph_range(
        section_path,
        xml_bytes,
        slots,
        paragraphs,
        source_text,
        preferred_start_paragraph,
        preferred_start_offset,
        preferred_end_paragraph,
        preferred_end_offset,
    )
    if fuzzy is not None and fuzzy[-1] >= 0.45:
        return fuzzy[:-1]
    return direct


def _position_candidate_score(candidate, source_text: str) -> float:
    _section_path, _xml_bytes, slots, start, end = candidate
    source_key = _position_compare_key(source_text)
    candidate_key = _position_compare_key(_candidate_text_from_slots(slots, start, end))
    if not source_key or not candidate_key:
        return 0.0
    ratio = difflib.SequenceMatcher(None, candidate_key, source_key, autojunk=False).ratio()
    if candidate_key in source_key or source_key in candidate_key:
        shorter = min(len(candidate_key), len(source_key))
        longer = max(len(candidate_key), len(source_key))
        ratio = max(ratio, shorter / max(1, longer))
    return ratio


def _candidate_text_from_slots(slots: list[TextSlot], match_start: int, match_end: int) -> str:
    parts: list[str] = []
    cursor = match_start
    for slot in slots:
        if slot.text_end <= match_start or slot.text_start >= match_end:
            continue
        local_start = max(0, match_start - slot.text_start)
        local_end = min(len(slot.value), match_end - slot.text_start)
        selected_start = slot.text_start + local_start
        selected_end = slot.text_start + local_end
        if selected_end <= selected_start:
            continue
        if selected_start > cursor:
            parts.append("\n" * (selected_start - cursor))
        parts.append(slot.value[local_start:local_end])
        cursor = selected_end
    return "".join(parts)


def _paragraph_text_ranges(xml_bytes: bytes, slots: list[TextSlot]) -> list[dict]:
    paragraph_starts = _paragraph_starts(xml_bytes)
    ranges: dict[int, dict] = {}
    for slot in slots:
        paragraph_start = _nearest_paragraph_start(xml_bytes, slot.xml_start)
        if paragraph_start not in paragraph_starts:
            continue
        ordinal = paragraph_starts.index(paragraph_start)
        item = ranges.setdefault(
            ordinal,
            {"ordinal": ordinal, "text_start": slot.text_start, "text_end": slot.text_end, "text": ""},
        )
        item["text_start"] = min(item["text_start"], slot.text_start)
        item["text_end"] = max(item["text_end"], slot.text_end)
    section_text_parts = []
    last_end = 0
    for slot in slots:
        if slot.text_start > last_end:
            section_text_parts.append("\n" * (slot.text_start - last_end))
        section_text_parts.append(slot.value)
        last_end = slot.text_end
    section_text = "".join(section_text_parts)
    for item in ranges.values():
        item["text"] = section_text[item["text_start"] : item["text_end"]]
    return [ranges[key] for key in sorted(ranges)]


def _paragraph_range_from_position(
    section_path: str,
    xml_bytes: bytes,
    slots: list[TextSlot],
    paragraphs: list[dict],
    start_paragraph: int,
    start_offset: int | None,
    end_paragraph: int | None,
    end_offset: int | None,
):
    by_ordinal = {item["ordinal"]: item for item in paragraphs}
    start_item = by_ordinal.get(start_paragraph)
    if start_item is None:
        return None
    end_item = by_ordinal.get(end_paragraph) if end_paragraph is not None else start_item
    if end_item is None:
        end_item = start_item
    start = start_item["text_start"] + max(0, int(start_offset or 0))
    if end_offset is None:
        end = end_item["text_end"]
    else:
        end = end_item["text_start"] + max(0, int(end_offset))
    start = max(start_item["text_start"], min(start, start_item["text_end"]))
    end = max(end_item["text_start"], min(end, end_item["text_end"]))
    if end <= start:
        start = start_item["text_start"]
        end = end_item["text_end"]
    if end <= start:
        return None
    return (section_path, xml_bytes, slots, start, end)


def _best_fuzzy_paragraph_range(
    section_path: str,
    xml_bytes: bytes,
    slots: list[TextSlot],
    paragraphs: list[dict],
    source_text: str,
    start_paragraph: int,
    start_offset: int | None,
    end_paragraph: int | None,
    end_offset: int | None,
):
    source_key = _position_compare_key(source_text)
    if not source_key:
        return None
    span = max(1, (end_paragraph if end_paragraph is not None else start_paragraph) - start_paragraph + 1)
    ordinals = [item["ordinal"] for item in paragraphs]
    best = None
    best_score = -1.0
    for candidate_start in range(start_paragraph - 8, start_paragraph + 9):
        if candidate_start not in ordinals:
            continue
        for candidate_span in range(max(1, span - 2), span + 3):
            candidate_end = candidate_start + candidate_span - 1
            selected = [item for item in paragraphs if candidate_start <= item["ordinal"] <= candidate_end]
            if not selected:
                continue
            text = "\n".join(item["text"] for item in selected)
            key = _position_compare_key(text)
            if not key:
                continue
            ratio = difflib.SequenceMatcher(None, key, source_key, autojunk=False).ratio()
            if key in source_key or source_key in key:
                shorter = min(len(key), len(source_key))
                longer = max(len(key), len(source_key))
                ratio = max(ratio, shorter / max(1, longer))
            distance_penalty = abs(candidate_start - start_paragraph) * 0.015
            score = ratio - distance_penalty
            if score > best_score:
                best_score = score
                start_item = selected[0]
                end_item = selected[-1]
                start = start_item["text_start"] + max(0, int(start_offset or 0))
                if end_offset is None:
                    end = end_item["text_end"]
                else:
                    end = end_item["text_start"] + max(0, int(end_offset))
                start = max(start_item["text_start"], min(start, start_item["text_end"]))
                end = max(end_item["text_start"], min(end, end_item["text_end"]))
                if end <= start:
                    start = start_item["text_start"]
                    end = end_item["text_end"]
                best = (section_path, xml_bytes, slots, start, end, best_score)
    return best


def _position_compare_key(text: str) -> str:
    normalized = _selection_compare_key(text)
    normalized = re.sub(r"&#\d+;", "", normalized)
    normalized = normalized.replace("·", "").replace("•", "").replace("●", "")
    return normalized


def _select_best_match_candidate(
    candidates,
    preferred_start_paragraph: int | None = None,
    preferred_start_offset: int | None = None,
):
    if len(candidates) == 1:
        return candidates[0]

    if preferred_start_paragraph is not None and preferred_start_offset is not None:
        exact_position_candidates = [
            candidate for candidate in candidates
            if _candidate_start_paragraph(candidate) == preferred_start_paragraph
            and _candidate_start_paragraph_offset(candidate) == preferred_start_offset
        ]
        if len(exact_position_candidates) == 1:
            return exact_position_candidates[0]

    if preferred_start_paragraph is not None:
        paragraph_candidates = [
            candidate for candidate in candidates
            if _candidate_start_paragraph(candidate) == preferred_start_paragraph
        ]
        if len(paragraph_candidates) == 1:
            return paragraph_candidates[0]

    # If Shift+Enter is involved, the intended selected block is the range whose
    # affected XML text nodes contain an actual HWPX lineBreak/br tag.
    soft_break_candidates = [
        candidate for candidate in candidates
        if _candidate_contains_soft_break(candidate)
    ]
    if len(soft_break_candidates) == 1:
        return soft_break_candidates[0]

    return None


def _candidate_start_paragraph(candidate) -> int:
    _section_path, xml_bytes, slots, start, end = candidate
    start_paragraph, _end_paragraph, _start_offset, _end_offset = _match_paragraph_position(
        xml_bytes,
        slots,
        start,
        end,
    )
    return start_paragraph


def _candidate_start_paragraph_offset(candidate) -> int:
    _section_path, xml_bytes, slots, start, end = candidate
    _start_paragraph, _end_paragraph, start_offset, _end_offset = _match_paragraph_position(
        xml_bytes,
        slots,
        start,
        end,
    )
    return start_offset


def _match_paragraph_position(
    xml_bytes: bytes,
    slots: list[TextSlot],
    match_start: int,
    match_end: int,
) -> tuple[int, int, int, int]:
    affected = [slot for slot in slots if slot.text_end > match_start and slot.text_start < match_end]
    if not affected:
        return -1, -1, -1, -1
    paragraph_starts = _paragraph_starts(xml_bytes)
    ordinals = []
    paragraph_slot_starts: dict[int, int] = {}
    for slot in affected:
        paragraph_start = _nearest_paragraph_start(xml_bytes, slot.xml_start)
        if paragraph_start in paragraph_starts:
            ordinal = paragraph_starts.index(paragraph_start)
            if ordinal not in ordinals:
                ordinals.append(ordinal)
    for slot in slots:
        paragraph_start = _nearest_paragraph_start(xml_bytes, slot.xml_start)
        if paragraph_start not in paragraph_starts:
            continue
        ordinal = paragraph_starts.index(paragraph_start)
        paragraph_slot_starts[ordinal] = min(paragraph_slot_starts.get(ordinal, slot.text_start), slot.text_start)
    if not ordinals:
        return -1, -1, -1, -1
    start_paragraph = min(ordinals)
    end_paragraph = max(ordinals)
    start_offset = match_start - paragraph_slot_starts.get(start_paragraph, match_start)
    end_offset = match_end - paragraph_slot_starts.get(end_paragraph, match_end)
    return start_paragraph, end_paragraph, start_offset, end_offset


def _whole_block_candidate_if_compatible(path: Path, needle: str):
    """Fallback for HWP saveblock fragments.

    HWP 2018 may return Shift+Enter selections from GetTextFile with control
    characters/newlines that do not exactly match the selected HWPX XML.  When
    the saved HWPX is already a selected block, replacing the whole extracted
    text is safer than falling back to plain text, but only when the two texts
    are clearly the same block.
    """
    section_candidates = []
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            xml_bytes = archive.read(section_path)
            section_text, slots = _section_text_and_slots(section_path, xml_bytes)
            if not slots:
                continue
            if _is_compatible_selected_block(section_text, needle):
                section_candidates.append((
                    section_path,
                    xml_bytes,
                    slots,
                    min(slot.text_start for slot in slots),
                    max(slot.text_end for slot in slots),
                ))
    if len(section_candidates) == 1:
        return section_candidates[0]
    return None


def _is_compatible_selected_block(section_text: str, needle: str) -> bool:
    section_key = _selection_compare_key(section_text)
    needle_key = _selection_compare_key(needle)
    if not section_key or not needle_key:
        return False
    shorter = min(len(section_key), len(needle_key))
    longer = max(len(section_key), len(needle_key))
    if shorter / longer < 0.72:
        return False
    if section_key == needle_key:
        return True
    if section_key in needle_key or needle_key in section_key:
        return shorter / longer >= 0.86
    return difflib.SequenceMatcher(None, section_key, needle_key, autojunk=False).ratio() >= 0.82


def _selection_compare_key(text: str) -> str:
    normalized = _normalize_text(text)
    # Ignore only layout/control differences that HWP GetTextFile and HWPX
    # commonly disagree on. Keep ordinary punctuation/content intact.
    ignored = {"\n", "\t", " ", "\u00a0", "\u200b"}
    return "".join(char for char in normalized if char not in ignored and not unicodedata.category(char).startswith("C"))


def _candidate_contains_soft_break(candidate) -> bool:
    _section_path, xml_bytes, slots, start, end = candidate
    for slot in slots:
        if slot.text_end <= start or slot.text_start >= end:
            continue
        payload = xml_bytes[slot.xml_start:slot.xml_end]
        if INNER_SOFT_LINE_BREAK_RE.search(payload):
            return True
    return False


def _find_text_ranges(text: str, needle: str) -> list[tuple[int, int]]:
    normalized_text = _normalize_text(text)
    normalized_needle = _normalize_text(needle)
    if not normalized_needle:
        return []

    exact = _find_exact_ranges(normalized_text, normalized_needle)
    if exact:
        return exact

    comparable_text, text_map = _comparable_text_with_index_map(normalized_text)
    comparable_needle, _needle_map = _comparable_text_with_index_map(normalized_needle)
    if comparable_needle:
        ranges: list[tuple[int, int]] = []
        for start, end in _find_exact_ranges(comparable_text, comparable_needle):
            original_start = text_map[start]
            original_end = text_map[end - 1] + 1
            ranges.append((original_start, original_end))
        if ranges:
            return _dedupe_ranges(ranges)

    # HWP Shift+Enter can appear as <hp:lineBreak/> in HWPX, while
    # GetTextFile(TEXT, saveblock) may return the same selection with no newline.
    # Match once more with newlines ignored, but keep an index map so replacement
    # still targets the original XML text range.
    flat_text, flat_text_map = _linebreakless_text_with_index_map(normalized_text)
    flat_needle, _flat_needle_map = _linebreakless_text_with_index_map(normalized_needle)
    if not flat_needle:
        return []

    ranges = []
    for start, end in _find_exact_ranges(flat_text, flat_needle):
        original_start = flat_text_map[start]
        original_end = flat_text_map[end - 1] + 1
        ranges.append((original_start, original_end))
    return _dedupe_ranges(ranges)


def _find_exact_ranges(text: str, needle: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    start = text.find(needle)
    while start >= 0:
        ranges.append((start, start + len(needle)))
        start = text.find(needle, start + 1)
    return ranges


def _comparable_text_with_index_map(text: str) -> tuple[str, list[int]]:
    lines = _normalize_text(text).split("\n")
    pieces: list[str] = []
    index_map: list[int] = []
    original_cursor = 0
    for line in lines:
        raw_start = original_cursor
        raw_end = raw_start + len(line)
        stripped = line.strip()
        if stripped:
            leading = len(line) - len(line.lstrip())
            if pieces:
                pieces.append("\n")
                index_map.append(max(0, raw_start - 1))
            for index, char in enumerate(stripped):
                pieces.append(char)
                index_map.append(raw_start + leading + index)
        original_cursor = raw_end + 1
    return "".join(pieces), index_map


def _linebreakless_text_with_index_map(text: str) -> tuple[str, list[int]]:
    normalized = _normalize_text(text)
    pieces: list[str] = []
    index_map: list[int] = []
    for index, char in enumerate(normalized):
        if char == "\n":
            continue
        pieces.append(char)
        index_map.append(index)
    return "".join(pieces), index_map


def _dedupe_ranges(ranges: list[tuple[int, int]]) -> list[tuple[int, int]]:
    seen: set[tuple[int, int]] = set()
    result: list[tuple[int, int]] = []
    for item in ranges:
        if item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


def _section_text_and_slots(section_path: str, xml_bytes: bytes) -> tuple[str, list[TextSlot]]:
    chunks: list[str] = []
    slots: list[TextSlot] = []
    cursor = 0
    previous_paragraph_start = -1
    encoding = _xml_encoding(xml_bytes)
    for match in TEXT_OR_BREAK_RE.finditer(xml_bytes):
        paragraph_start = _nearest_paragraph_start(xml_bytes, match.start())
        if paragraph_start >= 0 and previous_paragraph_start >= 0 and paragraph_start != previous_paragraph_start:
            chunks.append("\n")
            cursor += 1
        previous_paragraph_start = paragraph_start

        if SOFT_LINE_BREAK_RE.fullmatch(match.group(0)):
            chunks.append("\n")
            cursor += 1
            continue

        value_bytes = match.groupdict().get("text")
        if value_bytes is None:
            continue
        value = _decode_hwpx_text_node(value_bytes, encoding)
        if not value:
            continue
        value = _normalize_text(value)
        chunks.append(value)
        slots.append(TextSlot(section_path, match.start("text"), match.end("text"), cursor, cursor + len(value), value))
        cursor += len(value)
    return "".join(chunks), slots


def _replace_text_in_section(
    xml_bytes: bytes,
    slots: list[TextSlot],
    match_start: int,
    match_end: int,
    replacement_text: str,
) -> bytes:
    affected = [slot for slot in slots if slot.text_end > match_start and slot.text_start < match_end]
    if not affected:
        raise HwpxApplyError("Matched text did not map to text nodes.")

    new_values = _paragraph_aware_replacement_values(xml_bytes, affected, match_start, match_end, replacement_text)
    edits: list[tuple[int, int, bytes]] = []
    encoding = _xml_encoding(xml_bytes)
    for slot, new_value in zip(affected, new_values):
        original_payload = xml_bytes[slot.xml_start:slot.xml_end]
        edits.append((slot.xml_start, slot.xml_end, _encode_hwpx_text_node(new_value, encoding, original_payload)))

    updated = xml_bytes
    for start, end, data in sorted(edits, key=lambda item: item[0], reverse=True):
        updated = updated[:start] + data + updated[end:]
    return _cleanup_empty_text_runs(updated)


def _decode_hwpx_text_node(value_bytes: bytes, encoding: str) -> str:
    # Some HWPX files store soft line breaks illegally/nestingly inside <hp:t>.
    # Treat those inner tags as text newlines for matching.
    value_bytes = INNER_SOFT_LINE_BREAK_RE.sub(b"\n", value_bytes)
    return html.unescape(value_bytes.decode(encoding, errors="replace"))


def _encode_hwpx_text_node(value: str, encoding: str, original_payload: bytes) -> bytes:
    normalized = _normalize_text(value)
    if "\n" not in normalized:
        return html.escape(normalized, quote=False).encode(encoding)

    break_tag = b"<hp:lineBreak/>" if b"<hp:lineBreak" in original_payload else b"<hp:lineBreak/>"
    parts = normalized.split("\n")
    encoded_parts = [html.escape(part, quote=False).encode(encoding) for part in parts]
    return break_tag.join(encoded_parts)


def _cleanup_empty_text_runs(xml_bytes: bytes) -> bytes:
    without_empty_text = EMPTY_TEXT_NODE_RE.sub(b"", xml_bytes)

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
    # Any remaining element means the run still carries non-text content such as
    # lineBreak, fields, controls, drawings, tables, or other HWPX structures.
    return bool(TAG_RE.search(body))


def _paragraph_aware_replacement_values(
    xml_bytes: bytes,
    affected: list[TextSlot],
    match_start: int,
    match_end: int,
    replacement_text: str,
) -> list[str]:
    paragraph_groups: list[list[TextSlot]] = []
    paragraph_keys: list[int] = []
    for slot in affected:
        key = _nearest_paragraph_start(xml_bytes, slot.xml_start)
        if not paragraph_keys or paragraph_keys[-1] != key:
            paragraph_keys.append(key)
            paragraph_groups.append([])
        paragraph_groups[-1].append(slot)

    replacement_lines = _split_replacement_lines(replacement_text)
    group_sources = [
        _source_text_for_group(group, match_start, match_end)
        for group in paragraph_groups
    ]
    if len(paragraph_groups) > 1 and len(replacement_lines) == 1:
        replacement_lines = _project_collapsed_replacement_to_source_lines(group_sources, replacement_lines[0])
    if len(paragraph_groups) > 1 and len(replacement_lines) == len(paragraph_groups):
        values_by_slot_id: dict[int, str] = {}
        for group, line in zip(paragraph_groups, replacement_lines):
            group_start = max(match_start, min(slot.text_start for slot in group))
            group_end = min(match_end, max(slot.text_end for slot in group))
            group_values = _diff_replacement_by_slot(group, group_start, group_end, line)
            for slot, value in zip(group, group_values):
                values_by_slot_id[id(slot)] = value
        return [values_by_slot_id[id(slot)] for slot in affected]

    if len(paragraph_groups) == 1 and len(replacement_lines) == 1 and "\n" in group_sources[0]:
        projected = _project_collapsed_replacement_to_source_lines(group_sources[0].split("\n"), replacement_lines[0])

        # Case A: HWP 2018 sometimes stores Shift+Enter *inside* one <hp:t>,
        # e.g. <hp:t>했다.`<hp:lineBreak/></hp:t>.  If we split by slot groups,
        # that slot is rewritten without a newline and the lineBreak tag is lost.
        # Keep a newline in the replacement so _encode_hwpx_text_node rewrites it
        # back as <hp:lineBreak/>.
        if any("\n" in slot.value for slot in affected):
            return _diff_replacement_by_slot(affected, match_start, match_end, "\n".join(projected))

        # Case B: Shift+Enter is a separate XML tag between text nodes.  Do not
        # insert another newline into text nodes, or the existing tag becomes a
        # duplicated visual line break.
        soft_groups = _soft_line_groups_from_slots(affected, match_start, match_end)
        if len(soft_groups) == len(projected):
            values_by_slot_id: dict[int, str] = {}
            for group, line in zip(soft_groups, projected):
                group_start = max(match_start, min(slot.text_start for slot in group))
                group_end = min(match_end, max(slot.text_end for slot in group))
                group_values = _diff_replacement_by_slot(group, group_start, group_end, line)
                for slot, value in zip(group, group_values):
                    values_by_slot_id[id(slot)] = value
            return [values_by_slot_id[id(slot)] for slot in affected]
        return _diff_replacement_by_slot(affected, match_start, match_end, "\n".join(projected))

    return _diff_replacement_by_slot(affected, match_start, match_end, _remove_synthetic_newlines(replacement_text))


def _soft_line_groups_from_slots(group: list[TextSlot], match_start: int, match_end: int) -> list[list[TextSlot]]:
    groups: list[list[TextSlot]] = []
    cursor = match_start
    for slot in group:
        local_start = max(0, match_start - slot.text_start)
        local_end = min(len(slot.value), match_end - slot.text_start)
        selected_start = slot.text_start + local_start
        selected_end = slot.text_start + local_end
        if selected_end <= selected_start:
            continue
        if selected_start > cursor or not groups:
            groups.append([])
        groups[-1].append(slot)
        cursor = selected_end
    return groups


def _source_text_for_group(group: list[TextSlot], match_start: int, match_end: int) -> str:
    """Return selected source text for a paragraph group, including soft breaks.

    TextSlot only represents <hp:t> payloads.  A Shift+Enter stored as a
    separate <hp:lineBreak/> advances text positions but has no TextSlot, so a
    plain slot-value join collapses the line boundary.  Reconstruct the selected
    text using each slot's text_start/text_end gap so collapsed correction
    results can be projected back onto the original line structure.
    """
    parts: list[str] = []
    cursor = match_start
    for slot in group:
        local_start = max(0, match_start - slot.text_start)
        local_end = min(len(slot.value), match_end - slot.text_start)
        selected_start = slot.text_start + local_start
        selected_end = slot.text_start + local_end
        if selected_end <= selected_start:
            continue
        if selected_start > cursor:
            parts.append("\n" * (selected_start - cursor))
        parts.append(slot.value[local_start:local_end])
        cursor = selected_end
    return "".join(parts)


def _project_collapsed_replacement_to_source_lines(source_lines: list[str], replacement_line: str) -> list[str]:
    if len(source_lines) <= 1:
        return [replacement_line]
    source = "".join(source_lines)
    replacement = _normalize_text(replacement_line).replace("\n", "")
    if not source or not replacement:
        return [replacement_line]

    boundaries: list[int] = []
    cursor = 0
    for line in source_lines[:-1]:
        cursor += len(line)
        boundaries.append(cursor)

    mapped = [_map_source_index_to_replacement(source, replacement, boundary) for boundary in boundaries]
    split_points: list[int] = []
    previous = 0
    for point in mapped:
        point = max(previous, min(len(replacement), point))
        split_points.append(point)
        previous = point

    chunks: list[str] = []
    start = 0
    for point in split_points:
        chunks.append(replacement[start:point])
        start = point
    chunks.append(replacement[start:])
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
                replacement_span = j2 - j1
                return j1 + round(replacement_span * (source_index - i1) / source_span)
            return j1
        if source_index == i2:
            return j2
        previous_j = j2
    return len(replacement)


def _split_replacement_lines(text: str) -> list[str]:
    normalized = _normalize_text(text)
    return [line for line in normalized.split("\n")]


def _remove_synthetic_newlines(text: str) -> str:
    return _normalize_text(text).replace("\n", "")


def _diff_replacement_by_slot(
    affected: list[TextSlot],
    match_start: int,
    match_end: int,
    replacement_text: str,
) -> list[str]:
    source_chars: list[str] = []
    source_slot_indexes: list[int] = []
    prefixes: list[str] = []
    suffixes: list[str] = []

    for slot_index, slot in enumerate(affected):
        local_start = max(0, match_start - slot.text_start)
        local_end = min(len(slot.value), match_end - slot.text_start)
        prefixes.append(slot.value[:local_start])
        suffixes.append(slot.value[local_end:])
        for char in slot.value[local_start:local_end]:
            source_chars.append(char)
            source_slot_indexes.append(slot_index)

    source = "".join(source_chars)
    replacement = _normalize_text(replacement_text)
    middles = ["" for _slot in affected]
    matcher = difflib.SequenceMatcher(None, source, replacement, autojunk=False)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for source_index, char in zip(range(i1, i2), replacement[j1:j2]):
                middles[source_slot_indexes[source_index]] += char
        elif tag == "delete":
            continue
        elif tag == "insert":
            slot_index = _slot_for_insertion(source_slot_indexes, i1, len(affected))
            middles[slot_index] += replacement[j1:j2]
        elif tag == "replace":
            chunks = _split_replacement_for_source_slots(source_slot_indexes[i1:i2], replacement[j1:j2], len(affected))
            for slot_index, chunk in chunks:
                middles[slot_index] += chunk

    return [prefix + middle + suffix for prefix, middle, suffix in zip(prefixes, middles, suffixes)]


def _slot_for_insertion(source_slot_indexes: list[int], source_index: int, slot_count: int) -> int:
    if source_index > 0:
        return source_slot_indexes[source_index - 1]
    if source_index < len(source_slot_indexes):
        return source_slot_indexes[source_index]
    return max(0, slot_count - 1)


def _split_replacement_for_source_slots(
    source_slot_indexes: list[int],
    replacement: str,
    slot_count: int,
) -> list[tuple[int, str]]:
    if not replacement:
        return []
    if not source_slot_indexes:
        return [(_slot_for_insertion([], 0, slot_count), replacement)]

    ordered_slots: list[int] = []
    counts: dict[int, int] = {}
    for slot_index in source_slot_indexes:
        counts[slot_index] = counts.get(slot_index, 0) + 1
        if slot_index not in ordered_slots:
            ordered_slots.append(slot_index)

    if len(ordered_slots) == 1:
        return [(ordered_slots[0], replacement)]

    total = len(source_slot_indexes)
    cursor = 0
    chunks: list[tuple[int, str]] = []
    for index, slot_index in enumerate(ordered_slots):
        if index == len(ordered_slots) - 1:
            take_end = len(replacement)
        else:
            target = round(len(replacement) * sum(counts[item] for item in ordered_slots[: index + 1]) / total)
            take_end = max(cursor, min(len(replacement), target))
        chunks.append((slot_index, replacement[cursor:take_end]))
        cursor = take_end
    return chunks


def _section_fragment_for_match(
    original_xml: bytes,
    updated_xml: bytes,
    slots: list[TextSlot],
    match_start: int,
    match_end: int,
) -> bytes:
    affected = [slot for slot in slots if slot.text_end > match_start and slot.text_start < match_end]
    if not affected:
        raise HwpxApplyError("Matched text did not map to text nodes.")
    original_starts = _paragraph_starts(original_xml)
    affected_ordinals = []
    for slot in affected:
        paragraph_start = _nearest_paragraph_start(original_xml, slot.xml_start)
        if paragraph_start in original_starts:
            ordinal = original_starts.index(paragraph_start)
            if ordinal not in affected_ordinals:
                affected_ordinals.append(ordinal)
    if not affected_ordinals:
        raise HwpxApplyError("Could not find paragraph XML range for match.")

    updated_starts = _paragraph_starts(updated_xml)
    if max(affected_ordinals) >= len(updated_starts):
        raise HwpxApplyError("Could not map matched paragraphs after replacement.")
    start_ordinal = min(affected_ordinals)
    end_ordinal = max(affected_ordinals)
    start = updated_starts[start_ordinal]
    next_ordinal = max(affected_ordinals) + 1
    end = updated_starts[next_ordinal] if next_ordinal < len(updated_starts) else _section_close_start(updated_xml)
    body = updated_xml[start:end]
    if end_ordinal > start_ordinal:
        body += _empty_paragraph_after_fragment(updated_xml, end_ordinal)

    first_paragraph = _first_paragraph_start(updated_xml)
    if first_paragraph < 0:
        raise HwpxApplyError("Could not find section paragraph start.")
    section_close = _section_close_start(updated_xml)
    return updated_xml[:first_paragraph] + body + updated_xml[section_close:]


def _empty_paragraph_after_fragment(xml_bytes: bytes, paragraph_ordinal: int) -> bytes:
    starts = _paragraph_starts(xml_bytes)
    if paragraph_ordinal < 0 or paragraph_ordinal >= len(starts):
        return b""
    start = starts[paragraph_ordinal]
    end = starts[paragraph_ordinal + 1] if paragraph_ordinal + 1 < len(starts) else _section_close_start(xml_bytes)
    paragraph = xml_bytes[start:end]
    textless = re.sub(
        rb"<(?P<prefix>[A-Za-z_][\w.-]*:)?t(?:\s[^<>]*?)?>.*?</(?P=prefix)?t\s*>",
        b"",
        paragraph,
        flags=re.DOTALL,
    )
    if b"<" not in textless:
        return b""
    return textless


def _fragment_section_replacements(path: Path, target_section_path: str, target_xml: bytes) -> dict[str, bytes]:
    replacements: dict[str, bytes] = {}
    with zipfile.ZipFile(path, "r") as archive:
        for section_path in _section_names(archive):
            if section_path == target_section_path:
                replacements[section_path] = target_xml
            else:
                replacements[section_path] = _empty_section_xml(archive.read(section_path))
    return replacements


def _empty_section_xml(xml_bytes: bytes) -> bytes:
    section_open = re.search(rb"<(?:[A-Za-z_][\w.-]*:)?sec(?:\s[^<>]*?)?>", xml_bytes, re.IGNORECASE)
    section_close = _section_close_start(xml_bytes)
    if section_open is None or section_close < section_open.end():
        return xml_bytes
    return xml_bytes[:section_open.end()] + xml_bytes[section_close:]


def _paragraph_xml_range(xml_bytes: bytes, offset: int) -> tuple[int, int]:
    starts = _paragraph_starts(xml_bytes)
    if not starts:
        return -1, -1
    start = -1
    for value in starts:
        if value <= offset:
            start = value
        else:
            break
    if start < 0:
        return -1, -1
    next_starts = [value for value in starts if value > start]
    end = next_starts[0] if next_starts else _section_close_start(xml_bytes)
    return start, end


def _paragraph_starts(xml_bytes: bytes) -> list[int]:
    return [match.start() for match in PARAGRAPH_START_RE.finditer(xml_bytes)]


def _section_close_start(xml_bytes: bytes) -> int:
    match = re.search(rb"</(?:[A-Za-z_][\w.-]*:)?sec\s*>", xml_bytes, re.IGNORECASE)
    return match.start() if match else len(xml_bytes)


def _first_paragraph_start(xml_bytes: bytes) -> int:
    match = PARAGRAPH_START_RE.search(xml_bytes)
    return match.start() if match else -1


def _section_names(archive: zipfile.ZipFile) -> list[str]:
    return sorted(name for name in archive.namelist() if SECTION_PATH_RE.match(name))


def _nearest_paragraph_start(xml_bytes: bytes, offset: int) -> int:
    last = -1
    for match in PARAGRAPH_START_RE.finditer(xml_bytes, 0, offset):
        last = match.start()
    return last


def _create_backup(path: Path) -> Path:
    backup_dir = path.parent / ".writing-assistant-backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stem = path.stem
    index = 1
    while True:
        backup_path = backup_dir / f"{stem}.before_{index:03d}{path.suffix}"
        if not backup_path.exists():
            shutil.copy2(path, backup_path)
            return backup_path
        index += 1


def _replace_zip_member(path: Path, member_name: str, data: bytes) -> None:
    _replace_zip_members(path, {member_name: data})


def _replace_zip_members(path: Path, replacements: dict[str, bytes]) -> None:
    temp_handle = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx", dir=str(path.parent))
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    try:
        with zipfile.ZipFile(path, "r") as source_zip, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
            for info in source_zip.infolist():
                payload = replacements.get(info.filename)
                if payload is None:
                    payload = source_zip.read(info.filename)
                target_zip.writestr(info, payload)
        temp_path.replace(path)
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


def _xml_encoding(xml_bytes: bytes) -> str:
    match = re.search(br'encoding=["\']([^"\']+)["\']', xml_bytes[:200])
    return match.group(1).decode("ascii", errors="ignore") if match else "utf-8"


def _normalize_text(text: str) -> str:
    return (
        str(text or "")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\v", "\n")
        .replace("\f", "\n")
        .replace("\u2028", "\n")
        .replace("\u2029", "\n")
    )


def _validate_hwpx_path(document_path: str | Path) -> Path:
    path = Path(document_path)
    if path.suffix.lower() != ".hwpx":
        raise HwpxApplyError(f"Not an HWPX file: {path}")
    if not path.exists():
        raise HwpxApplyError(f"HWPX file does not exist: {path}")
    return path
