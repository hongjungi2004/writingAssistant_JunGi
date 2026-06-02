from __future__ import annotations

from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import zipfile


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "client" / "input"))

from hwpx_document import (
    HwpxApplyError,
    apply_text_to_hwpx,
    create_hwpx_fragment_from_match,
    extract_hwpx_text,
    find_hwpx_text_matches,
)
from _tmp_hwp_word_like.hwpx_rebuilder import create_rebuilt_hwpx_fragment
from client.input.output_applier import OutputApplier


def create_hwpx(path: Path, section_xml: str) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", "application/hwp+zip")
        archive.writestr("Contents/section0.xml", section_xml.encode("utf-8"))


def read_section(path: Path) -> str:
    with zipfile.ZipFile(path, "r") as archive:
        return archive.read("Contents/section0.xml").decode("utf-8")


class HwpxDocumentTest(unittest.TestCase):
    def test_replaces_text_without_changing_run_shape_refs(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="11"><hp:t>school애서</hp:t></hp:run><hp:run charPrIDRef="10"><hp:t> 와 sul레를</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            result = apply_text_to_hwpx(path, "school애서 와 sul레를", "school에서 와 sul래를")

            updated = read_section(path)
            self.assertTrue(result.backup_path.exists())
            self.assertIn('charPrIDRef="11"><hp:t>school에서</hp:t>', updated)
            self.assertIn('charPrIDRef="10"><hp:t> 와 sul래를</hp:t>', updated)
            self.assertEqual(updated.count("charPrIDRef="), xml.count("charPrIDRef="))

    def test_duplicate_source_is_blocked(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>same</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>same</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "duplicate.hwpx"
            create_hwpx(path, xml)

            with self.assertRaises(HwpxApplyError):
                apply_text_to_hwpx(path, "same", "changed")

    def test_duplicate_source_can_be_disambiguated_by_paragraph(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>same</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>same</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "duplicate.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            matches = find_hwpx_text_matches(path, "same")
            self.assertEqual([(m.start_paragraph, m.end_paragraph) for m in matches], [(0, 0), (1, 1)])

            create_hwpx_fragment_from_match(
                path,
                "same",
                "changed",
                fragment,
                preferred_start_paragraph=1,
            )

            updated = read_section(fragment)
            self.assertNotIn('charPrIDRef="1"', updated)
            self.assertIn('charPrIDRef="2"><hp:t>changed</hp:t>', updated)

    def test_duplicate_source_in_same_paragraph_can_be_disambiguated_by_offset(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>same </hp:t></hp:run><hp:run charPrIDRef="2"><hp:t>same</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "duplicate.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            matches = find_hwpx_text_matches(path, "same")
            self.assertEqual(
                [(m.start_paragraph, m.start_paragraph_offset) for m in matches],
                [(0, 0), (0, 5)],
            )

            create_hwpx_fragment_from_match(
                path,
                "same",
                "changed",
                fragment,
                preferred_start_paragraph=0,
                preferred_start_offset=5,
            )

            updated = read_section(fragment)
            self.assertIn('charPrIDRef="1"><hp:t>same </hp:t>', updated)
            self.assertIn('charPrIDRef="2"><hp:t>changed</hp:t>', updated)

    def test_position_fallback_handles_numbered_selection_text_mismatch(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>이전 문단</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>다른 문단</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"><hp:t>2. 문단 정렬과 간격</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="4"><hp:t>왼쪽 정렬 문단입니다.</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="5"><hp:t>가운데 정렬 문단입니다.</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="6"><hp:t>이후 문단</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "numbered.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(
                path,
                "1. 2. 문단 정렬과 간격\n왼쪽 정렬 문단입니다.\n가운데 정렬 문단입니다.",
                "교정문",
                fragment,
                preferred_start_paragraph=0,
                preferred_start_offset=0,
                preferred_end_paragraph=2,
                preferred_end_offset=20,
            )

            updated = read_section(fragment)
            self.assertEqual(extract_hwpx_text(fragment), "교정문")
            self.assertNotIn("이전 문단", updated)
            self.assertNotIn("다른 문단", updated)
            self.assertNotIn("이후 문단", updated)

    def test_position_fallback_uses_direct_range_when_text_is_unmatched(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"><hp:t>CCC</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "position.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(
                path,
                "unmatched selection text",
                "교정문",
                fragment,
                preferred_start_paragraph=1,
                preferred_start_offset=0,
                preferred_end_paragraph=1,
                preferred_end_offset=3,
            )

            updated = read_section(fragment)
            self.assertEqual(extract_hwpx_text(fragment), "교정문")
            self.assertNotIn("AAA", updated)
            self.assertNotIn("CCC", updated)

    def test_position_fallback_strips_visual_auto_number_prefix_from_replacement(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>2. TITLE</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BODY</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "numbered.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(
                path,
                "1. 2. TITLE\nBODY",
                "1. 2. TITLE\nBODY fixed",
                fragment,
                preferred_start_paragraph=0,
                preferred_start_offset=0,
                preferred_end_paragraph=1,
                preferred_end_offset=4,
            )

            self.assertEqual(extract_hwpx_text(fragment), "2. TITLE\nBODY fixed")
            updated = read_section(fragment)
            self.assertNotIn(">1. 2. TITLE<", updated)

    def test_fragment_keeps_only_matched_paragraph(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(path, "BBB", "CCC", fragment)

            self.assertEqual(extract_hwpx_text(fragment), "CCC")
            updated = read_section(fragment)
            self.assertNotIn("AAA", updated)
            self.assertIn('charPrIDRef="2"', updated)

    def test_create_rebuilt_hwpx_fragment_preserves_style_and_fragment_scope(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="11"><hp:t>school애서</hp:t></hp:run><hp:run charPrIDRef="10"><hp:t> 와 sul레를</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="20"><hp:t>다른문장</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_rebuilt_hwpx_fragment(path, "school에서 와 sul래를", fragment, source_text="school애서 와 sul레를")

            self.assertEqual(extract_hwpx_text(fragment), "school에서 와 sul래를")
            updated = read_section(fragment)
            self.assertIn('charPrIDRef="11"><hp:t>school에서</hp:t>', updated)
            self.assertIn('charPrIDRef="10"><hp:t> 와 sul래를</hp:t>', updated)
            self.assertNotIn("다른문장", updated)
            self.assertEqual(updated.count("charPrIDRef="), 2)

    def test_flexible_match_tolerates_blank_lines_without_falling_back_to_one_line(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"><hp:t>CCC</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(path, "AAA\n\nBBB\nCCC", "AXAA\nBXBB\nCCX", fragment)

            self.assertEqual(extract_hwpx_text(fragment), "AXAA\nBXBB\nCCX")
            updated = read_section(fragment)
            self.assertIn('charPrIDRef="1"><hp:t>AXAA</hp:t>', updated)
            self.assertIn('charPrIDRef="2"><hp:t>BXBB</hp:t>', updated)
            self.assertIn('charPrIDRef="3"><hp:t>CCX</hp:t>', updated)

    def test_multiline_mismatch_does_not_modify_only_longest_line(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            with self.assertRaises(HwpxApplyError):
                create_hwpx_fragment_from_match(path, "ZZZ\nBBB", "changed", fragment)

    def test_fragment_preserves_empty_paragraph_inside_selection(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"/></hp:p>
  <hp:p><hp:run charPrIDRef="4"><hp:t>CCC</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="5"><hp:t>DDD</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(path, "BBB\n\nCCC", "BXBB\n\nCCX", fragment)

            updated = read_section(fragment)
            self.assertNotIn("AAA", updated)
            self.assertNotIn("DDD", updated)
            self.assertIn('charPrIDRef="2"><hp:t>BXBB</hp:t>', updated)
            self.assertIn('charPrIDRef="4"><hp:t>CCX</hp:t>', updated)
            self.assertIn('charPrIDRef="3"', updated)

    def test_multiline_replacement_does_not_write_synthetic_newlines_into_text_nodes(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            apply_text_to_hwpx(path, "AAA\nBBB", "AXAA\nBXBB")

            updated = read_section(path)
            self.assertIn('charPrIDRef="1"><hp:t>AXAA</hp:t>', updated)
            self.assertIn('charPrIDRef="2"><hp:t>BXBB</hp:t>', updated)
            self.assertNotIn("<hp:t>AXAA\n", updated)

    def test_fragment_range_survives_length_changes_across_three_paragraphs(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="0"><hp:t>BEFORE</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"><hp:t>CCC</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="4"><hp:t>AFTER</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(path, "AAA\nBBB\nCCC", "AAAAA\nB\nCCCCCC", fragment)

            self.assertEqual(extract_hwpx_text(fragment), "AAAAA\nB\nCCCCCC")
            updated = read_section(fragment)
            self.assertNotIn("BEFORE", updated)
            self.assertNotIn("AFTER", updated)
            self.assertIn('charPrIDRef="1"><hp:t>AAAAA</hp:t>', updated)
            self.assertIn('charPrIDRef="2"><hp:t>B</hp:t>', updated)
            self.assertIn('charPrIDRef="3"><hp:t>CCCCCC</hp:t>', updated)
            self.assertGreaterEqual(updated.count("<hp:p"), 4)

    def test_collapsed_multiline_replacement_is_projected_back_to_paragraphs(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>BBB</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"><hp:t>CCC</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_hwpx_fragment_from_match(path, "AAA\nBBB\nCCC", "AXAABXBBCCX [done]", fragment)

            self.assertEqual(extract_hwpx_text(fragment), "AXAA\nBXBB\nCCX [done]")
            updated = read_section(fragment)
            self.assertIn('charPrIDRef="1"><hp:t>AXAA</hp:t>', updated)
            self.assertIn('charPrIDRef="2"><hp:t>BXBB</hp:t>', updated)
            self.assertIn('charPrIDRef="3"><hp:t>CCX [done]</hp:t>', updated)

    def test_diff_keeps_inserted_text_near_original_style_slot(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="3"><hp:t>강강</hp:t></hp:run><hp:run charPrIDRef="4"><hp:t>술레를</hp:t></hp:run><hp:run charPrIDRef="5"><hp:t>했다.</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            apply_text_to_hwpx(path, "강강술레를했다.", "강강술래를 했다.")

            updated = read_section(path)
            self.assertIn('charPrIDRef="3"><hp:t>강강</hp:t>', updated)
            self.assertIn('charPrIDRef="4"><hp:t>술래를 </hp:t>', updated)
            self.assertIn('charPrIDRef="5"><hp:t>했다.</hp:t>', updated)

    def test_deleted_sentence_removes_empty_text_run_debris(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>안녕하세요.</hp:t></hp:run><hp:run charPrIDRef="2"><hp:t> 반갑습니다.</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            apply_text_to_hwpx(path, "안녕하세요. 반갑습니다.", "안녕하세요.")

            updated = read_section(path)
            self.assertEqual(extract_hwpx_text(path), "안녕하세요.")
            self.assertIn('charPrIDRef="1"><hp:t>안녕하세요.</hp:t>', updated)
            self.assertNotIn('charPrIDRef="2"', updated)
            self.assertNotIn("<hp:t></hp:t>", updated)

    def test_deleted_sentence_cleanup_preserves_non_text_run_content(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t></hp:run><hp:run charPrIDRef="2"><hp:t>BBB</hp:t><hp:lineBreak/></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            apply_text_to_hwpx(path, "AAABBB", "AAA")

            updated = read_section(path)
            self.assertIn('charPrIDRef="1"><hp:t>AAA</hp:t>', updated)
            self.assertIn('charPrIDRef="2">', updated)
            self.assertIn("<hp:lineBreak/>", updated)
            self.assertNotIn("<hp:t></hp:t>", updated)

    def test_rebuilt_fragment_deletion_removes_empty_text_run_debris(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>안녕하세요.</hp:t></hp:run><hp:run charPrIDRef="2"><hp:t> 반갑습니다.</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            fragment = Path(temp_dir) / "fragment.hwpx"
            create_hwpx(path, xml)

            create_rebuilt_hwpx_fragment(path, "안녕하세요.", fragment, source_text="안녕하세요. 반갑습니다.")

            updated = read_section(fragment)
            self.assertEqual(extract_hwpx_text(fragment), "안녕하세요.")
            self.assertIn('charPrIDRef="1"><hp:t>안녕하세요.</hp:t>', updated)
            self.assertNotIn('charPrIDRef="2"', updated)
            self.assertNotIn("<hp:t></hp:t>", updated)


    def test_shift_enter_soft_line_break_matches_source_newline(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t><hp:lineBreak/><hp:t>BBB</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            self.assertEqual(extract_hwpx_text(path), "AAA\nBBB")
            apply_text_to_hwpx(path, "AAA\nBBB", "AXAA\nBXBB")

            updated = read_section(path)
            self.assertIn("<hp:lineBreak/>", updated)
            self.assertIn("<hp:t>AXAA</hp:t>", updated)
            self.assertIn("<hp:t>BXBB</hp:t>", updated)


    def test_shift_enter_closed_tag_and_vertical_tab_source_match(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>AAA</hp:t><hp:lineBreak></hp:lineBreak><hp:t>BBB</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            self.assertEqual(extract_hwpx_text(path), "AAA\nBBB")
            apply_text_to_hwpx(path, "AAA\vBBB", "AXAA\nBXBB")

            updated = read_section(path)
            self.assertIn("<hp:lineBreak></hp:lineBreak>", updated)
            self.assertIn("<hp:t>AXAA</hp:t>", updated)
            self.assertIn("<hp:t>BXBB</hp:t>", updated)


    def test_collapsed_replacement_preserves_soft_line_break_between_text_nodes(self) -> None:
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>학교에서 </hp:t></hp:run><hp:run charPrIDRef="2"><hp:t>친구들과</hp:t><hp:lineBreak/><hp:t> 강강술레를 했다.</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.hwpx"
            create_hwpx(path, xml)

            apply_text_to_hwpx(path, "학교에서 친구들과 강강술레를 했다.", "학교에서 친구들과 강강술래를 했다. [맞춤법 기능 사용 됨]")

            self.assertEqual(extract_hwpx_text(path), "학교에서 친구들과\n 강강술래를 했다. [맞춤법 기능 사용 됨]")
            updated = read_section(path)
            self.assertIn("<hp:lineBreak/>", updated)
            self.assertIn("친구들과", updated)
            self.assertIn("강강술래를 했다.", updated)

    def test_real_s_hwpx_sample_when_available(self) -> None:
        source_path = find_real_sample()
        if source_path is None:
            self.skipTest("real s.hwpx sample not found")
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "s.hwpx"
            shutil.copy2(source_path, path)
            text = extract_hwpx_text(path)
            self.assertTrue(text.strip())

            source = first_existing(text, ["school애서", "학교애서", "강강술레를했다.", "강강술래를 했다!"])
            self.assertTrue(source)
            matches = find_hwpx_text_matches(path, source)
            self.assertGreaterEqual(len(matches), 1)

    def test_output_applier_accepts_explicit_hwpx_fragment_specs(self) -> None:
        applier = OutputApplier()

        specs = applier._hwp_hwpx_fragment_specs(
            style_info={
                "hwp_hwpx_fragments": [
                    {"source_text": "AAA", "replacement_text": "AXA"},
                    {"source": "BBB", "text": "BXB"},
                    {"source_text": "", "replacement_text": "ignored"},
                ]
            },
            source_text="AAA\nBBB",
            replacement_text="AXA\nBXB",
        )

        self.assertEqual(
            specs,
            [
                {"source_text": "AAA", "replacement_text": "AXA"},
                {"source_text": "BBB", "replacement_text": "BXB"},
            ],
        )

    def test_output_applier_inserts_hwpx_fragments_in_order_after_delete(self) -> None:
        applier = OutputApplier()
        calls: list[tuple[str, str]] = []

        def fake_delete(_hwp):
            calls.append(("delete", ""))

        def fake_insert(_hwp, path):
            calls.append(("insert", Path(path).name))
            return True

        applier._delete_hwp_selection = fake_delete
        applier._insert_hwpx_file = fake_insert

        applier._replace_hwp_selection_with_fragments(
            hwp=object(),
            fragment_paths=[Path("first.hwpx"), Path("second.hwpx")],
        )

        self.assertEqual(calls, [("delete", ""), ("insert", "first.hwpx"), ("insert", "second.hwpx")])

    def test_output_applier_parses_hwp_selected_position(self) -> None:
        applier = OutputApplier()

        parsed = applier._parse_hwp_selected_pos((True, 0, 6, 0, 0, 6, 19))

        self.assertEqual(parsed["start_para"], 6)
        self.assertEqual(parsed["start_pos"], 0)
        self.assertEqual(parsed["end_para"], 6)
        self.assertEqual(parsed["end_pos"], 19)

    def test_output_applier_hwpml_path_avoids_saveas(self) -> None:
        applier = OutputApplier()
        calls: list[tuple[str, str]] = []
        hwpml = (
            "<HWPML><BODY><SECTION>"
            "<P><TEXT><CHAR>2. TITLE</CHAR></TEXT></P>"
            "<P><TEXT><CHAR>BODY</CHAR></TEXT></P>"
            "</SECTION></BODY></HWPML>"
        )

        class FakeHwp:
            def GetTextFile(self, fmt, option):
                if fmt == "TEXT":
                    return "1. 2. TITLE\nBODY"
                if fmt == "HWPML2X":
                    return hwpml
                return ""

            def SaveAs(self, *_args):
                raise AssertionError("SaveAs should not be called when HWPML2X succeeds")

        def fake_delete(_hwp):
            calls.append(("delete", ""))

        def fake_insert(_hwp, path):
            calls.append(("insert", Path(path).suffix))
            self.assertEqual(Path(path).suffix, ".hml")
            text = Path(path).read_text(encoding="utf-8")
            self.assertIn("<CHAR>2. TITLE</CHAR>", text)
            self.assertIn("<CHAR>BODY fixed</CHAR>", text)
            return True

        applier._delete_hwp_selection = fake_delete
        applier._insert_hwpml_file = fake_insert

        self.assertTrue(
            applier._apply_to_hwp_hwpx_selection(
                FakeHwp(),
                "1. 2. TITLE\nBODY fixed",
                {"_source_text": "1. 2. TITLE\nBODY"},
            )
        )

        self.assertEqual(calls, [("delete", ""), ("insert", ".hml")])

    def test_save_hwpx_prefers_full_document_when_selection_text_is_unique(self) -> None:
        applier = OutputApplier()

        class FakeHwp:
            def SaveAs(self, path, _format, option):
                if option == "":
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>BEFORE</hp:t></hp:run></hp:p>
  <hp:p><hp:run><hp:t>UNIQUE</hp:t></hp:run></hp:p>
  <hp:p><hp:run><hp:t>AFTER</hp:t></hp:run></hp:p>
</hs:sec>""")
                else:
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>UNIQUE</hp:t></hp:run></hp:p>
</hs:sec>""")
                return True

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "chosen.hwpx"

            self.assertTrue(
                applier._save_hwp_document_as_hwpx(
                    FakeHwp(),
                    path,
                    expected_text="UNIQUE",
                    prefer_full_document=True,
                )
            )

            self.assertEqual(extract_hwpx_text(path), "BEFORE\nUNIQUE\nAFTER")

    def test_save_hwpx_falls_back_to_selection_when_full_document_is_ambiguous(self) -> None:
        applier = OutputApplier()

        class FakeHwp:
            def SaveAs(self, path, _format, option):
                if option == "":
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
</hs:sec>""")
                else:
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
</hs:sec>""")
                return True

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "chosen.hwpx"

            self.assertTrue(
                applier._save_hwp_document_as_hwpx(
                    FakeHwp(),
                    path,
                    expected_text="SAME",
                    prefer_full_document=True,
                )
            )

            self.assertEqual(extract_hwpx_text(path), "SAME")

    def test_save_hwpx_prefers_full_document_when_duplicate_is_positioned(self) -> None:
        applier = OutputApplier()

        class FakeHwp:
            def SaveAs(self, path, _format, option):
                if option == "":
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
</hs:sec>""")
                else:
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
</hs:sec>""")
                return True

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "chosen.hwpx"
            applier._current_hwp_selection_pos = {"start_para": 1}
            try:
                self.assertTrue(
                    applier._save_hwp_document_as_hwpx(
                        FakeHwp(),
                        path,
                        expected_text="SAME",
                        prefer_full_document=True,
                    )
                )
            finally:
                applier._current_hwp_selection_pos = {}

            self.assertEqual(extract_hwpx_text(path), "SAME\nSAME")

    def test_save_hwpx_prefers_full_document_when_same_paragraph_duplicate_is_positioned(self) -> None:
        applier = OutputApplier()

        class FakeHwp:
            def SaveAs(self, path, _format, option):
                if option == "":
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SAME SAME</hp:t></hp:run></hp:p>
</hs:sec>""")
                else:
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SAME</hp:t></hp:run></hp:p>
</hs:sec>""")
                return True

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "chosen.hwpx"
            applier._current_hwp_selection_pos = {"start_para": 0, "start_pos": 5}
            try:
                self.assertTrue(
                    applier._save_hwp_document_as_hwpx(
                        FakeHwp(),
                        path,
                        expected_text="SAME",
                        prefer_full_document=True,
                    )
                )
            finally:
                applier._current_hwp_selection_pos = {}

            self.assertEqual(extract_hwpx_text(path), "SAME SAME")

    def test_save_hwpx_prefers_full_document_when_position_fallback_matches(self) -> None:
        applier = OutputApplier()

        class FakeHwp:
            def SaveAs(self, path, _format, option):
                if option == "":
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>BEFORE</hp:t></hp:run></hp:p>
  <hp:p><hp:run><hp:t>2. NUMBERED TARGET</hp:t></hp:run></hp:p>
  <hp:p><hp:run><hp:t>AFTER</hp:t></hp:run></hp:p>
</hs:sec>""")
                else:
                    create_hwpx(Path(path), """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run><hp:t>SELECTION EXPORT</hp:t></hp:run></hp:p>
</hs:sec>""")
                return True

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "chosen.hwpx"
            applier._current_hwp_selection_pos = {"start_para": 1, "start_pos": 0, "end_para": 1, "end_pos": 18}
            try:
                self.assertTrue(
                    applier._save_hwp_document_as_hwpx(
                        FakeHwp(),
                        path,
                        expected_text="1. 2. NUMBERED TARGET",
                        prefer_full_document=True,
                    )
                )
            finally:
                applier._current_hwp_selection_pos = {}

            self.assertEqual(extract_hwpx_text(path), "BEFORE\n2. NUMBERED TARGET\nAFTER")

    def test_output_applier_position_fallback_keeps_paragraph_structure(self) -> None:
        applier = OutputApplier()
        xml = """<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="x" xmlns:hp="y">
  <hp:p><hp:run charPrIDRef="1"><hp:t>TITLE</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="2"><hp:t>FIRST ITEM</hp:t></hp:run></hp:p>
  <hp:p><hp:run charPrIDRef="3"><hp:t>SECOND ITEM</hp:t></hp:run></hp:p>
</hs:sec>"""
        with tempfile.TemporaryDirectory() as temp_dir:
            work_dir = Path(temp_dir)
            path = work_dir / "document.hwpx"
            create_hwpx(path, xml)

            fragments = applier._create_hwp_hwpx_fragments(
                document_path=path,
                work_dir=work_dir,
                document_text=extract_hwpx_text(path),
                selection_pos={"start_para": 1, "start_pos": 0, "end_para": 2, "end_pos": 11},
                fragment_specs=[
                    {
                        "source_text": "1. FIRST ITEM\n2. SECOND ITEM",
                        "replacement_text": "1. FIRST ITEM fixed\n2. SECOND ITEM fixed",
                    }
                ],
            )

            self.assertEqual(len(fragments), 1)
            self.assertEqual(extract_hwpx_text(fragments[0]), "1. FIRST ITEM fixed\n2. SECOND ITEM fixed")
            updated = read_section(fragments[0])
            self.assertIn('charPrIDRef="2"', updated)
            self.assertIn('charPrIDRef="3"', updated)
            self.assertNotIn("TITLE", updated)


def first_existing(text: str, candidates: list[str]) -> str:
    for candidate in candidates:
        if candidate in text:
            return candidate
    return ""


def find_real_sample() -> Path | None:
    for root in (Path.home() / "OneDrive", Path.home() / "Documents"):
        if not root.exists():
            continue
        for path in root.rglob("s.hwpx"):
            if path.is_file():
                return path
    return None


if __name__ == "__main__":
    unittest.main()
