"""Unit tests for the markdown writer and the write confirmation gate."""

import xml.etree.ElementTree as ET

import pytest

from onenote_lib import write_gate
from onenote_lib.markdown_writer import (
    NSMAP,
    apply_markdown,
    inline_to_onenote,
    set_page_title,
)
from onenote_lib.xml_parser import parse_page_to_markdown

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

EXISTING_PAGE = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-1" name="Existing" lastModifiedTime="2026-03-01T10:00:00.000Z">
  <one:QuickStyleDef index="0" name="PageTitle"/>
  <one:QuickStyleDef index="1" name="p"/>
  <one:Title><one:OE objectID="t-1"><one:T><![CDATA[Existing]]></one:T></one:OE></one:Title>
  <one:Outline objectID="o-1"><one:OEChildren>
    <one:OE objectID="oe-1"><one:T><![CDATA[original line]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""


class TestInlineConversion:
    def test_bold_and_italic(self):
        assert "font-weight:bold" in inline_to_onenote("**x**")
        assert "font-style:italic" in inline_to_onenote("*x*")

    def test_link(self):
        out = inline_to_onenote("see [docs](https://example.com/a)")
        assert '<a href="https://example.com/a">docs</a>' in out

    def test_plain_text_is_escaped(self):
        # Text that looks like markup must not become markup.
        assert inline_to_onenote("a < b & c") == "a &lt; b &amp; c"

    def test_round_trip_through_parser(self):
        xml = apply_markdown(EXISTING_PAGE, "**bold** and [x](https://e.com)", "replace")
        md, _ = parse_page_to_markdown(xml)
        assert "**bold**" in md
        assert "[x](https://e.com)" in md

    def test_list_item_keeps_its_emphasis(self):
        # A list marker must not suppress inline formatting the way a heading does.
        xml = apply_markdown(EXISTING_PAGE, "- item with **bold**", "replace")
        md, _ = parse_page_to_markdown(xml)
        assert "- item with **bold**" in md


class TestApplyMarkdown:
    def test_append_keeps_existing_content_and_ids(self):
        xml = apply_markdown(EXISTING_PAGE, "added line", "append")
        root = ET.fromstring(xml)
        oes = root.findall(".//one:Outline/one:OEChildren/one:OE", NSMAP)
        # The original OE survives with its objectID, so OneNote matches rather
        # than replaces it; only the new OE is untagged.
        assert oes[0].get("objectID") == "oe-1"
        assert oes[-1].get("objectID") is None
        md, _ = parse_page_to_markdown(xml)
        assert "original line" in md
        assert "added line" in md

    def test_replace_discards_existing_content(self):
        xml = apply_markdown(EXISTING_PAGE, "brand new", "replace")
        md, _ = parse_page_to_markdown(xml)
        assert "original line" not in md
        assert "brand new" in md

    def test_page_id_is_preserved(self):
        root = ET.fromstring(apply_markdown(EXISTING_PAGE, "x", "append"))
        assert root.get("ID") == "page-1"

    def test_rejects_unknown_mode(self):
        with pytest.raises(ValueError):
            apply_markdown(EXISTING_PAGE, "x", "clobber")

    def test_heading_round_trips(self):
        xml = apply_markdown(EXISTING_PAGE, "## Section Two", "replace")
        md, _ = parse_page_to_markdown(xml)
        assert "## Section Two" in md

    def test_nested_list_round_trips(self):
        xml = apply_markdown(EXISTING_PAGE, "- top\n  - child", "replace")
        md, _ = parse_page_to_markdown(xml)
        assert "- top" in md
        assert "    - child" in md

    def test_set_page_title(self):
        root = ET.fromstring(set_page_title(EXISTING_PAGE, "Renamed"))
        assert root.find("one:Title//one:T", NSMAP).text == "Renamed"


class TestWriteGate:
    def setup_method(self):
        write_gate._pending.clear()

    def test_valid_token_is_accepted(self):
        token = write_gate.issue_token("p1", "body", "append")
        write_gate.consume_token(token, "p1", "body", "append")

    def test_missing_token_is_rejected(self):
        with pytest.raises(PermissionError):
            write_gate.consume_token("made-up", "p1", "body", "append")

    def test_token_is_single_use(self):
        token = write_gate.issue_token("p1", "body", "append")
        write_gate.consume_token(token, "p1", "body", "append")
        with pytest.raises(PermissionError):
            write_gate.consume_token(token, "p1", "body", "append")

    def test_token_is_bound_to_content(self):
        # The whole point of the gate: approving one diff must not authorise
        # writing something else.
        token = write_gate.issue_token("p1", "body", "append")
        with pytest.raises(PermissionError):
            write_gate.consume_token(token, "p1", "different body", "append")

    def test_token_is_bound_to_page(self):
        token = write_gate.issue_token("p1", "body", "append")
        with pytest.raises(PermissionError):
            write_gate.consume_token(token, "p2", "body", "append")

    def test_token_is_bound_to_mode(self):
        token = write_gate.issue_token("p1", "body", "append")
        with pytest.raises(PermissionError):
            write_gate.consume_token(token, "p1", "body", "replace")

    def test_expired_token_is_rejected(self, monkeypatch):
        token = write_gate.issue_token("p1", "body", "append")
        monkeypatch.setattr(
            write_gate.time, "time", lambda: 1e12
        )
        with pytest.raises(PermissionError):
            write_gate.consume_token(token, "p1", "body", "append")


class TestDiffRendering:
    def test_diff_shows_added_line(self):
        diff = write_gate.render_diff("a\nb", "a\nb\nc", "My Page")
        assert "+c" in diff
        assert "My Page" in diff
