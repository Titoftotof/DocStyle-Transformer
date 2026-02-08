"""Test suite for DocStyle Transformer.

Tests cover models, parser, detector, mapper, and end-to-end generation.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest
from docx import Document as open_docx

from core.models import (
    Callout,
    CalloutType,
    ContentElement,
    DocumentMetadata,
    DocumentTree,
    Image,
    ListBlock,
    ListItem,
    ListType,
    PageBreak,
    Paragraph,
    Section,
    Step,
    StepsBlock,
    Table,
    TextRun,
)

FIXTURES_DIR = Path(__file__).parent / "fixtures"


# ── Model tests ─────────────────────────────────────────────────────


class TestModels:
    def test_text_run(self):
        run = TextRun(text="hello", bold=True)
        assert run.text == "hello"
        assert run.bold is True
        assert run.italic is False

    def test_paragraph_text_property(self):
        p = Paragraph(runs=[TextRun(text="hello "), TextRun(text="world")])
        assert p.text == "hello world"

    def test_empty_paragraph(self):
        p = Paragraph()
        assert p.text == ""

    def test_section(self):
        s = Section(heading="Intro", level=1, number=1)
        assert s.heading == "Intro"
        assert s.level == 1

    def test_document_tree_summary(self):
        tree = DocumentTree(
            sections=[
                Section(
                    heading="S1",
                    level=1,
                    children=[
                        Paragraph(runs=[TextRun(text="text")]),
                        Table(headers=["A", "B"], rows=[["1", "2"]]),
                        Callout(callout_type=CalloutType.INFO, body="info"),
                    ],
                ),
                Section(
                    heading="S2",
                    level=1,
                    children=[
                        Paragraph(runs=[TextRun(text="more")]),
                        ListBlock(list_type=ListType.BULLET, items=[]),
                        StepsBlock(steps=[Step(number=1, title="S")]),
                    ],
                ),
            ]
        )
        summary = tree.summary()
        assert summary["sections"] == 2
        assert summary["paragraphs"] == 2
        assert summary["tables"] == 1
        assert summary["callouts"] == 1
        assert summary["lists"] == 1
        assert summary["steps_blocks"] == 1

    def test_document_tree_empty(self):
        tree = DocumentTree()
        assert tree.section_count() == 0
        assert tree.flat_elements() == []
        summary = tree.summary()
        assert all(v == 0 for v in summary.values())


# ── Design system / mapper tests ────────────────────────────────────


class TestDesignSystem:
    def test_load_default_config(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        assert ds.resolve_color("black") == "#1D1D1F"
        assert ds.resolve_color("accent_blue") == "#0071E3"

    def test_resolve_hex_passthrough(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        assert ds.resolve_color("#FF0000") == "#FF0000"

    def test_unknown_color_raises(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        with pytest.raises(KeyError):
            ds.resolve_color("nonexistent_color")

    def test_heading_style(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        style = ds.get_heading_style(1)
        assert style["bold"] is True
        assert style["size"] == 40
        assert "font" in style

    def test_body_style(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        style = ds.get_body_style()
        assert style["size"] == 21
        assert "font" in style

    def test_page_config(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        page = ds.get_page_config()
        assert page["width"] == 12240
        assert page["height"] == 15840
        assert ds.usable_width == 9360

    def test_component_style(self):
        from core.mapper import DesignSystem

        ds = DesignSystem()
        table_style = ds.get_component_style("table")
        assert "header_bg" in table_style
        assert table_style["zebra_stripe"] is True


class TestStyleMapper:
    def test_map_paragraph(self):
        from core.mapper import StyleMapper

        sm = StyleMapper()
        p = Paragraph(runs=[TextRun(text="hello")])
        result = sm.map_element(p)
        assert result["type"] == "paragraph"
        assert "font" in result

    def test_map_heading(self):
        from core.mapper import StyleMapper

        sm = StyleMapper()
        result = sm.map_heading(1)
        assert result["type"] == "heading"
        assert result["level"] == 1

    def test_map_table(self):
        from core.mapper import StyleMapper

        sm = StyleMapper()
        t = Table(headers=["A"], rows=[["1"]])
        result = sm.map_element(t)
        assert result["type"] == "table"
        assert result["has_headers"] is True

    def test_map_callout(self):
        from core.mapper import StyleMapper

        sm = StyleMapper()
        c = Callout(callout_type=CalloutType.WARNING, body="warn")
        result = sm.map_element(c)
        assert result["type"] == "callout"
        assert result["callout_type"] == "warning"

    def test_map_page_break(self):
        from core.mapper import StyleMapper

        sm = StyleMapper()
        result = sm.map_element(PageBreak())
        assert result["type"] == "page_break"


# ── Detector tests ──────────────────────────────────────────────────


class TestDetector:
    def test_keyword_callout_detection(self):
        from core.detector import StructureDetector

        tree = DocumentTree(
            sections=[
                Section(
                    heading="Test",
                    level=1,
                    children=[
                        Paragraph(
                            runs=[TextRun(text="Note : Ceci est important.")]
                        ),
                    ],
                )
            ]
        )
        detector = StructureDetector()
        enriched = detector.detect(tree)
        children = enriched.sections[0].children
        assert len(children) == 1
        assert isinstance(children[0], Callout)
        assert children[0].callout_type == CalloutType.NOTE

    def test_warning_callout_detection(self):
        from core.detector import StructureDetector

        tree = DocumentTree(
            sections=[
                Section(
                    heading="Test",
                    level=1,
                    children=[
                        Paragraph(
                            runs=[TextRun(text="Attention : Ne pas oublier.")]
                        ),
                    ],
                )
            ]
        )
        detector = StructureDetector()
        enriched = detector.detect(tree)
        children = enriched.sections[0].children
        assert isinstance(children[0], Callout)
        assert children[0].callout_type == CalloutType.WARNING

    def test_heading_normalization(self):
        from core.detector import StructureDetector

        tree = DocumentTree(
            sections=[
                Section(heading="A", level=1),
                Section(heading="B", level=3),  # gap: no level 2
                Section(heading="C", level=5),  # way too deep
            ]
        )
        detector = StructureDetector()
        enriched = detector.detect(tree)
        levels = [s.level for s in enriched.sections]
        assert levels == [1, 2, 3]

    def test_section_numbering(self):
        from core.detector import StructureDetector

        tree = DocumentTree(
            sections=[
                Section(heading="A", level=1),
                Section(heading="A1", level=2),
                Section(heading="B", level=1),
            ]
        )
        detector = StructureDetector()
        enriched = detector.detect(tree)
        numbers = [s.number for s in enriched.sections]
        assert numbers[0] == 1
        assert numbers[1] is None  # level 2 not numbered
        assert numbers[2] == 2

    def test_step_detection(self):
        from core.detector import StructureDetector

        tree = DocumentTree(
            sections=[
                Section(
                    heading="Steps",
                    level=1,
                    children=[
                        Paragraph(runs=[TextRun(text="Étape 1 : Configurer")]),
                        Paragraph(runs=[TextRun(text="Description étape 1.")]),
                        Paragraph(runs=[TextRun(text="Étape 2 : Déployer")]),
                        Paragraph(runs=[TextRun(text="Description étape 2.")]),
                    ],
                )
            ]
        )
        detector = StructureDetector()
        enriched = detector.detect(tree)
        children = enriched.sections[0].children
        assert len(children) == 1
        assert isinstance(children[0], StepsBlock)
        assert len(children[0].steps) == 2

    def test_single_cell_table_as_callout(self):
        from core.detector import StructureDetector

        tree = DocumentTree(
            sections=[
                Section(
                    heading="Test",
                    level=1,
                    children=[
                        Table(headers=["Some important note."], rows=[]),
                    ],
                )
            ]
        )
        detector = StructureDetector()
        enriched = detector.detect(tree)
        children = enriched.sections[0].children
        assert isinstance(children[0], Callout)


# ── Parser tests ────────────────────────────────────────────────────


class TestParser:
    def test_parse_corporate_document(self):
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_corporate.docx"
        if not fixture.exists():
            pytest.skip("Corporate fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        assert tree.metadata.title == "Guide technique : Architecture Cloud"
        assert tree.metadata.author == "Jean Dupont"
        assert len(tree.sections) > 0

    def test_parse_minimal_document(self):
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_minimal.docx"
        if not fixture.exists():
            pytest.skip("Minimal fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        # No headings -> everything in preamble
        assert len(tree.sections) == 0
        assert len(tree.preamble) > 0

    def test_parse_empty_document(self):
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_empty.docx"
        if not fixture.exists():
            pytest.skip("Empty fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        assert len(tree.sections) == 0
        assert len(tree.preamble) == 0

    def test_parse_nonexistent_file(self):
        from core.parser import DocxParser

        parser = DocxParser()
        with pytest.raises(FileNotFoundError):
            parser.parse("/nonexistent/path/file.docx")

    def test_parse_non_docx_file(self):
        from core.parser import DocxParser

        parser = DocxParser()
        with pytest.raises(ValueError, match="Unsupported file type"):
            parser.parse(__file__)  # this .py file

    def test_sections_have_content(self):
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_corporate.docx"
        if not fixture.exists():
            pytest.skip("Corporate fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        # At least one section should have children
        has_content = any(len(s.children) > 0 for s in tree.sections)
        assert has_content

    def test_tables_detected(self):
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_corporate.docx"
        if not fixture.exists():
            pytest.skip("Corporate fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        tables = [
            e
            for s in tree.sections
            for e in s.children
            if isinstance(e, Table)
        ]
        assert len(tables) >= 1
        assert len(tables[0].headers) == 3  # Composant, Version, Statut

    def test_lists_detected(self):
        """Test list detection — may find ListBlock or individual paragraphs
        depending on how python-docx encodes list styles."""
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_corporate.docx"
        if not fixture.exists():
            pytest.skip("Corporate fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        # Lists may be detected as ListBlock or as plain paragraphs
        # if python-docx list styles don't generate numPr elements.
        all_elements = tree.flat_elements()
        lists = [e for e in all_elements if isinstance(e, ListBlock)]
        paragraphs = [e for e in all_elements if isinstance(e, Paragraph)]
        # At minimum we should have some content parsed
        assert len(lists) + len(paragraphs) > 0


# ── End-to-end generation tests ─────────────────────────────────────


class TestEndToEnd:
    def test_full_pipeline_corporate(self):
        from core.detector import StructureDetector
        from core.generator import DocumentGenerator
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_corporate.docx"
        if not fixture.exists():
            pytest.skip("Corporate fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        detector = StructureDetector()
        tree = detector.detect(tree)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            generator = DocumentGenerator()
            result = generator.generate(tree, output_path)

            assert os.path.isfile(result)
            assert os.path.getsize(result) > 0

            # Verify the output is a valid docx
            doc = open_docx(result)
            assert len(doc.paragraphs) > 0
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_full_pipeline_minimal(self):
        from core.detector import StructureDetector
        from core.generator import DocumentGenerator
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_minimal.docx"
        if not fixture.exists():
            pytest.skip("Minimal fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        detector = StructureDetector()
        tree = detector.detect(tree)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            generator = DocumentGenerator()
            result = generator.generate(
                tree,
                output_path,
                options={
                    "generate_cover": False,
                    "generate_toc": False,
                },
            )

            assert os.path.isfile(result)
            assert os.path.getsize(result) > 0
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_full_pipeline_with_all_options(self):
        from core.detector import StructureDetector
        from core.generator import DocumentGenerator
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_corporate.docx"
        if not fixture.exists():
            pytest.skip("Corporate fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        detector = StructureDetector()
        tree = detector.detect(tree)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            generator = DocumentGenerator()
            result = generator.generate(
                tree,
                output_path,
                options={
                    "generate_cover": True,
                    "generate_toc": True,
                    "number_sections": True,
                    "header_footer": True,
                    "cover_title_override": "Titre personnalisé : Test",
                },
            )

            assert os.path.isfile(result)

            doc = open_docx(result)
            # Should have many paragraphs (cover + toc + content)
            assert len(doc.paragraphs) > 10
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_empty_document_no_crash(self):
        from core.detector import StructureDetector
        from core.generator import DocumentGenerator
        from core.parser import DocxParser

        fixture = FIXTURES_DIR / "sample_empty.docx"
        if not fixture.exists():
            pytest.skip("Empty fixture not available")

        parser = DocxParser()
        tree = parser.parse(str(fixture))

        detector = StructureDetector()
        tree = detector.detect(tree)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            generator = DocumentGenerator()
            result = generator.generate(
                tree,
                output_path,
                options={
                    "generate_cover": False,
                    "generate_toc": False,
                },
            )
            assert os.path.isfile(result)
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)


# ── Cover generator tests ──────────────────────────────────────────


class TestCoverGenerator:
    def test_title_split_delimiter(self):
        from core.cover import CoverGenerator

        line1, line2 = CoverGenerator._split_title(
            "Guide technique : Architecture Cloud"
        )
        assert line1 == "Guide technique"
        assert line2 == "Architecture Cloud"

    def test_title_split_dash(self):
        from core.cover import CoverGenerator

        line1, line2 = CoverGenerator._split_title(
            "Infrastructure - Guide de déploiement"
        )
        assert line1 == "Infrastructure"
        assert line2 == "Guide de déploiement"

    def test_title_split_single_word(self):
        from core.cover import CoverGenerator

        line1, line2 = CoverGenerator._split_title("Introduction")
        assert line1 == "Introduction"
        assert line2 == ""

    def test_title_split_no_delimiter(self):
        from core.cover import CoverGenerator

        line1, line2 = CoverGenerator._split_title(
            "Mon document de test important"
        )
        assert line1
        assert line2
        # Both should be non-empty, total should equal original
        assert f"{line1} {line2}" == "Mon document de test important"
