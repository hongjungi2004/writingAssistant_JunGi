from pathlib import Path
import tempfile
import unittest


from client.input.hwpml_document import create_hwpml_fragment_from_selection, extract_hwpml_text


class HwpmlDocumentTest(unittest.TestCase):
    def test_replaces_char_nodes_preserving_paragraph_shells(self) -> None:
        xml = (
            '<?xml version="1.0" encoding="UTF-16" standalone="no" ?>'
            '<HWPML><BODY><SECTION>'
            '<P ParaShape="1"><TEXT CharShape="1"><CHAR>2. TITLE</CHAR></TEXT></P>'
            '<P ParaShape="2"><TEXT CharShape="2"><CHAR>BODY</CHAR></TEXT></P>'
            '</SECTION></BODY></HWPML>'
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(source, "2. TITLE fixed\nBODY fixed", output)

            updated = output.read_text(encoding="utf-8")
            self.assertEqual(extract_hwpml_text(output), "2. TITLE fixed\nBODY fixed")
            self.assertIn('ParaShape="1"', updated)
            self.assertIn('CharShape="2"', updated)

    def test_strips_visual_auto_number_prefix(self) -> None:
        xml = (
            "<HWPML><BODY><SECTION>"
            '<P><TEXT><CHAR>2. TITLE</CHAR></TEXT></P>'
            '<P><TEXT><CHAR>BODY</CHAR></TEXT></P>'
            "</SECTION></BODY></HWPML>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(source, "1. 2. TITLE\nBODY fixed", output)

            self.assertEqual(extract_hwpml_text(output), "2. TITLE\nBODY fixed")

    def test_strips_visual_prefixes_on_each_line(self) -> None:
        xml = (
            "<HWPML><BODY><SECTION>"
            '<P><TEXT><CHAR>3. 목록 서식</CHAR></TEXT></P>'
            '<P><TEXT><CHAR>가. 글머리 기호 목록</CHAR></TEXT></P>'
            '<P><TEXT><CHAR>· 첫 번째 항목</CHAR></TEXT></P>'
            '<P><TEXT><CHAR>1. 문서 열기</CHAR></TEXT></P>'
            "</SECTION></BODY></HWPML>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(
                source,
                "2. 3. 목록 서식\n가. 가. 글머리 기호 목록\n· · 첫 번째 항목\n1. 1. 문서 열기 [done]",
                output,
            )

            self.assertEqual(
                extract_hwpml_text(output),
                "3. 목록 서식\n가. 글머리 기호 목록\n· 첫 번째 항목\n1. 문서 열기 [done]",
            )

    def test_strips_visual_prefixes_when_original_has_no_prefix(self) -> None:
        xml = (
            "<HWPML><BODY><SECTION>"
            '<P><TEXT><CHAR>글머리 기호 목록</CHAR></TEXT></P>'
            '<P><TEXT><CHAR>첫 번째 항목</CHAR></TEXT></P>'
            "</SECTION></BODY></HWPML>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(
                source,
                "가. 글머리 기호 목록\n· 첫 번째 항목 [done]",
                output,
            )

            self.assertEqual(extract_hwpml_text(output), "글머리 기호 목록\n첫 번째 항목 [done]")

    def test_preserves_inline_char_shapes_when_replacement_collapses_paragraph(self) -> None:
        xml = (
            "<HWPML><BODY><SECTION>"
            '<P><TEXT CharShape="1"><CHAR>일반 텍스트 </CHAR></TEXT>'
            '<TEXT CharShape="2"><CHAR>굵게</CHAR></TEXT>'
            '<TEXT CharShape="1"><CHAR>, </CHAR></TEXT>'
            '<TEXT CharShape="3"><CHAR>기울임</CHAR></TEXT>'
            '<TEXT CharShape="1"><CHAR> 문장</CHAR></TEXT></P>'
            "</SECTION></BODY></HWPML>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(source, "일반 텍스트 굵게, 기울임 문장 [done]", output)

            updated = output.read_text(encoding="utf-8")
            self.assertEqual(extract_hwpml_text(output), "일반 텍스트 굵게, 기울임 문장 [done]")
            self.assertIn('CharShape="2"><CHAR>굵게</CHAR>', updated)
            self.assertIn('CharShape="3"><CHAR>기울임</CHAR>', updated)
            self.assertIn('CharShape="1"><CHAR> 문장 [done]</CHAR>', updated)

    def test_keeps_replacement_near_changed_inline_run(self) -> None:
        xml = (
            "<HWPML><BODY><SECTION>"
            '<P><TEXT CharShape="1"><CHAR>일반 </CHAR></TEXT>'
            '<TEXT CharShape="2"><CHAR>굵께</CHAR></TEXT>'
            '<TEXT CharShape="1"><CHAR> 문장</CHAR></TEXT></P>'
            "</SECTION></BODY></HWPML>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(source, "일반 굵게 문장", output)

            updated = output.read_text(encoding="utf-8")
            self.assertEqual(extract_hwpml_text(output), "일반 굵게 문장")
            self.assertIn('CharShape="2"><CHAR>굵게</CHAR>', updated)

    def test_decodes_numeric_character_references_in_replacement(self) -> None:
        xml = (
            "<HWPML><BODY><SECTION>"
            '<P><TEXT><CHAR>체크박스</CHAR></TEXT></P>'
            "</SECTION></BODY></HWPML>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.hml"
            output = Path(temp_dir) / "output.hml"
            source.write_text(xml, encoding="utf-8")

            create_hwpml_fragment_from_selection(source, "체크박스 &#9744; &#x2610;", output)

            self.assertEqual(extract_hwpml_text(output), "체크박스 ☐ ☐")
            self.assertIn("<CHAR>체크박스 ☐ ☐</CHAR>", output.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
