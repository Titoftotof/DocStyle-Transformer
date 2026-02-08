"""DOCX file parser that produces a DocumentTree intermediate representation.

Reads a Microsoft Word .docx file using python-docx for high-level access and
lxml for low-level XML inspection (numbering definitions, hyperlinks, page
breaks, etc.).  The result is a fully populated :class:`DocumentTree` that
downstream generators can consume without any knowledge of the DOCX format.
"""

from __future__ import annotations

import logging
import os
import re
from typing import Optional

from docx import Document as open_docx
from docx.document import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from lxml import etree

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

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────

# EMU (English Metric Unit) to DXA (twentieth of a point) conversion factor.
# 1 inch = 914400 EMU = 1440 DXA  =>  1 EMU = 1440 / 914400
_EMU_TO_DXA = 1440 / 914400

# Heading style prefix used by python-docx (e.g. "Heading 1", "Heading 2").
_HEADING_RE = re.compile(r"^Heading\s+(\d)$", re.IGNORECASE)

# Minimum font-size (in half-points) to heuristically treat a bold paragraph
# as a heading when it does not carry a formal heading style.
_HEURISTIC_HEADING_MIN_SIZE = 28  # 14 pt

# Word XML namespaces used for low-level queries.
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# Callout keyword mapping (case-insensitive first word of certain styled
# paragraphs or text patterns).
_CALLOUT_KEYWORDS: dict[str, CalloutType] = {
    "info": CalloutType.INFO,
    "warning": CalloutType.WARNING,
    "note": CalloutType.NOTE,
    "tip": CalloutType.TIP,
}


# ── Helpers ────────────────────────────────────────────────────────────


def _emu_to_dxa(emu: int) -> int:
    """Convert English Metric Units to DXA (twentieths of a point)."""
    return round(emu * _EMU_TO_DXA)


def _color_str(color_elem) -> Optional[str]:
    """Extract a hex colour string from an ``<w:color>`` element."""
    if color_elem is None:
        return None
    val = color_elem.get(qn("w:val"))
    if val and val.lower() not in ("auto", "none"):
        return f"#{val}" if not val.startswith("#") else val
    return None


def _font_size_from_rpr(rpr) -> Optional[int]:
    """Return font size in half-points from a ``<w:rPr>`` element."""
    if rpr is None:
        return None
    sz = rpr.find(qn("w:sz"))
    if sz is not None:
        val = sz.get(qn("w:val"))
        if val and val.isdigit():
            return int(val)
    return None


def _has_break(run_elem, break_type: str = "page") -> bool:
    """Return *True* if a ``<w:r>`` element contains a break of *break_type*."""
    for br in run_elem.findall(qn("w:br")):
        if br.get(qn("w:type"), "textWrapping") == break_type:
            return True
    return False


def _paragraph_has_page_break(para_elem) -> bool:
    """Check whether a paragraph contains an explicit page or section break."""
    # Run-level <w:br w:type="page"/>
    for r in para_elem.findall(qn("w:r")):
        if _has_break(r, "page") or _has_break(r, "column"):
            return True
    # Paragraph-level section break in <w:pPr><w:sectPr>
    ppr = para_elem.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
        return True
    return False


def _last_rendered_page_break(para_elem) -> bool:
    """Detect ``<w:lastRenderedPageBreak/>`` inside runs."""
    for r in para_elem.findall(qn("w:r")):
        if r.find(qn("w:lastRenderedPageBreak")) is not None:
            return True
    return False


# ── Main parser ────────────────────────────────────────────────────────


class DocxParser:
    """Parse a ``.docx`` file into a :class:`DocumentTree`.

    Usage::

        parser = DocxParser()
        tree = parser.parse("document.docx")
    """

    def __init__(self) -> None:
        self._doc: Optional[Document] = None
        self._rels: dict[str, str] = {}  # rId -> target URL (hyperlinks)
        self._numbering_map: dict[str, dict] = {}  # numId -> abstractNum info
        self._style_name_map: dict[str, str] = {}  # styleId -> canonical name
        self._style_outline_lvl: dict[str, int] = {}  # styleId -> outline level (0-based)

    # ── Public API ─────────────────────────────────────────────────

    def parse(self, file_path: str) -> DocumentTree:
        """Parse *file_path* and return the corresponding :class:`DocumentTree`.

        Raises
        ------
        FileNotFoundError
            If *file_path* does not exist.
        ValueError
            If *file_path* is not a ``.docx`` file or is corrupted.
        """
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        if not file_path.lower().endswith(".docx"):
            raise ValueError(
                f"Unsupported file type (expected .docx): {file_path}"
            )

        try:
            self._doc = open_docx(file_path)
        except Exception as exc:
            raise ValueError(
                f"Failed to open document (file may be corrupted): {exc}"
            ) from exc

        logger.info("Parsing document: %s", file_path)

        self._build_hyperlink_map()
        self._build_numbering_map()
        self._build_style_map()

        metadata = self._extract_metadata()
        elements = self._walk_body()
        tree = self._build_tree(metadata, elements)

        # Heuristic title detection: if the document's core properties had no
        # title, try to infer one from the first heading or prominent paragraph.
        if not tree.metadata.title:
            tree.metadata.title = self._infer_title(tree)

        logger.info(
            "Parsed %d section(s), %d preamble element(s)",
            len(tree.sections),
            len(tree.preamble),
        )
        return tree

    # ── Metadata ───────────────────────────────────────────────────

    def _extract_metadata(self) -> DocumentMetadata:
        assert self._doc is not None
        props = self._doc.core_properties
        title = props.title or ""
        author = props.author or ""
        date = ""
        if props.created:
            date = props.created.strftime("%Y-%m-%d")
        elif props.modified:
            date = props.modified.strftime("%Y-%m-%d")
        version = props.version or ""
        # 'reference' is not a standard OPC property; leave blank.
        return DocumentMetadata(
            title=title,
            author=author,
            date=date,
            version=version,
        )

    @staticmethod
    def _infer_title(tree: DocumentTree) -> str:
        """Try to infer a document title from the tree structure.

        Checks, in order:
        1. The first section heading (the first heading in the document is
           typically the document title, regardless of its level).
        2. The first preamble paragraph with bold text (likely a title).
        """
        if tree.sections:
            return tree.sections[0].heading

        # Check preamble for a prominent paragraph
        for elem in tree.preamble:
            if isinstance(elem, Paragraph) and elem.text.strip():
                for run in elem.runs:
                    if run.text.strip():
                        if run.bold:
                            return elem.text.strip()
                        break
        return ""

    # ── Hyperlink / numbering maps ─────────────────────────────────

    def _build_hyperlink_map(self) -> None:
        """Populate ``self._rels`` from the main document-part relationships."""
        assert self._doc is not None
        self._rels = {}
        try:
            for rel in self._doc.part.rels.values():
                if rel.reltype == RT.HYPERLINK:
                    self._rels[rel.rId] = rel._target  # noqa: SLF001
        except Exception:
            logger.debug("Could not read hyperlink relationships", exc_info=True)

    def _build_numbering_map(self) -> None:
        """Build a lookup from ``numId`` to abstract numbering metadata.

        This allows us to know whether a given numId represents a bullet
        list or a numbered list, plus the indentation level mapping.
        """
        assert self._doc is not None
        self._numbering_map = {}

        numbering_part = None
        try:
            numbering_part = self._doc.part.numbering_part
        except Exception:
            logger.debug("No numbering part found in document", exc_info=True)
            return

        if numbering_part is None:
            return

        numbering_xml = numbering_part.element
        # Map abstractNumId -> format of level 0 (to distinguish bullet / numbered)
        abstract_map: dict[str, str] = {}
        for abstract_num in numbering_xml.findall(qn("w:abstractNum")):
            abs_id = abstract_num.get(qn("w:abstractNumId"))
            lvl0 = abstract_num.find(qn("w:lvl"))
            fmt = "bullet"
            if lvl0 is not None:
                num_fmt = lvl0.find(qn("w:numFmt"))
                if num_fmt is not None:
                    fmt_val = num_fmt.get(qn("w:val"), "bullet")
                    fmt = fmt_val
            if abs_id is not None:
                abstract_map[abs_id] = fmt

        for num_elem in numbering_xml.findall(qn("w:num")):
            num_id = num_elem.get(qn("w:numId"))
            abs_ref = num_elem.find(qn("w:abstractNumId"))
            if num_id is not None and abs_ref is not None:
                abs_id = abs_ref.get(qn("w:val"))
                fmt = abstract_map.get(abs_id or "", "bullet")
                self._numbering_map[num_id] = {"format": fmt}

    def _build_style_map(self) -> None:
        """Build mappings from style ID to canonical name and outline level.

        This allows heading detection to work regardless of whether the style
        ID is localised (e.g. "Titre1" in French) or differs from the
        canonical English name (e.g. "Heading1" vs "Heading 1").
        """
        assert self._doc is not None
        self._style_name_map = {}
        self._style_outline_lvl = {}

        try:
            styles_element = self._doc.styles.element
            for style_elem in styles_element.findall(qn("w:style")):
                style_id = style_elem.get(qn("w:styleId"), "")
                if not style_id:
                    continue

                # Canonical name from <w:name w:val="..."/>
                name_elem = style_elem.find(qn("w:name"))
                if name_elem is not None:
                    name = name_elem.get(qn("w:val"), "")
                    if name:
                        self._style_name_map[style_id] = name

                # Outline level from <w:pPr><w:outlineLvl w:val="N"/>
                ppr = style_elem.find(qn("w:pPr"))
                if ppr is not None:
                    outline_lvl = ppr.find(qn("w:outlineLvl"))
                    if outline_lvl is not None:
                        val = outline_lvl.get(qn("w:val"), "")
                        if val.isdigit():
                            self._style_outline_lvl[style_id] = int(val)

            logger.debug(
                "Style map built: %d styles, %d with outline levels",
                len(self._style_name_map),
                len(self._style_outline_lvl),
            )
        except Exception:
            logger.debug("Could not build style map", exc_info=True)

    # ── Body walk ──────────────────────────────────────────────────

    def _walk_body(self) -> list[dict]:
        """Walk all top-level body elements and return a flat list of dicts.

        Each dict has a ``"type"`` key (``"heading"``, ``"paragraph"``,
        ``"table"``, ``"image"``, ``"page_break"``, ``"list_item"``, etc.)
        plus type-specific payload keys.
        """
        assert self._doc is not None
        body = self._doc.element.body
        results: list[dict] = []

        for child in body:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

            if tag == "p":
                results.extend(self._process_paragraph(child))
            elif tag == "tbl":
                results.append(self._process_table(child))
            elif tag == "sdt":
                # Structured document tags wrap content; recurse into them.
                sdt_content = child.find(qn("w:sdtContent"))
                if sdt_content is not None:
                    for sdt_child in sdt_content:
                        sdt_tag = (
                            etree.QName(sdt_child.tag).localname
                            if isinstance(sdt_child.tag, str)
                            else ""
                        )
                        if sdt_tag == "p":
                            results.extend(self._process_paragraph(sdt_child))
                        elif sdt_tag == "tbl":
                            results.append(self._process_table(sdt_child))
            elif tag == "sectPr":
                # Final section properties -- may imply a section break but
                # typically appears at the very end; we emit a page break only
                # if there is content after it (handled in tree building).
                pass
            else:
                logger.debug("Skipping unknown body element: %s", tag)

        return results

    # ── Paragraph processing ───────────────────────────────────────

    def _process_paragraph(self, para_elem) -> list[dict]:
        """Process a single ``<w:p>`` element, returning one or more result dicts.

        A single XML paragraph can yield multiple logical elements when it
        contains a page break before the text content.
        """
        results: list[dict] = []

        # Check for page break / section break
        if _paragraph_has_page_break(para_elem) or _last_rendered_page_break(para_elem):
            results.append({"type": "page_break"})

        # Check for embedded images
        images = self._extract_images_from_paragraph(para_elem)
        for img in images:
            results.append({"type": "image", "image": img})

        # Determine heading level
        heading_level = self._detect_heading_level(para_elem)

        # Determine list properties
        list_info = self._detect_list_properties(para_elem)

        # Build text runs
        runs = self._parse_runs(para_elem)

        # Skip completely empty paragraphs that are not headings and have no
        # images (images already emitted above).
        plain_text = "".join(r.text for r in runs).strip()
        if not plain_text and heading_level is None and list_info is None:
            return results

        if heading_level is not None:
            results.append({
                "type": "heading",
                "level": heading_level,
                "text": plain_text,
                "runs": runs,
            })
        elif list_info is not None:
            results.append({
                "type": "list_item",
                "runs": runs,
                "level": list_info["level"],
                "list_type": list_info["list_type"],
            })
        else:
            if plain_text:
                results.append({
                    "type": "paragraph",
                    "runs": runs,
                })

        return results

    def _detect_heading_level(self, para_elem) -> Optional[int]:
        """Return heading level (1-6) if the paragraph is a heading, else *None*.

        Detection order:
        1. Paragraph-level outline level (``<w:outlineLvl>`` in ``<w:pPr>``).
        2. Style outline level from the document's style definitions.
        3. Style name / style ID pattern matching (English + French).
        4. Heuristic: bold text with large font size.
        """
        assert self._doc is not None

        ppr = para_elem.find(qn("w:pPr"))
        style_id = ""

        if ppr is not None:
            # 0) Direct outline level on the paragraph itself
            outline_lvl = ppr.find(qn("w:outlineLvl"))
            if outline_lvl is not None:
                val = outline_lvl.get(qn("w:val"), "")
                if val.isdigit():
                    lvl = int(val)
                    if lvl <= 5:
                        return lvl + 1  # outlineLvl is 0-based

            pstyle = ppr.find(qn("w:pStyle"))
            if pstyle is not None:
                style_id = pstyle.get(qn("w:val"), "")

        # 1) Check style outline level from the style definition
        if style_id and style_id in self._style_outline_lvl:
            outline = self._style_outline_lvl[style_id]
            if outline <= 5:
                return outline + 1

        # 2) Check style name (from style map) and style ID for heading patterns
        if style_id:
            style_name = self._style_name_map.get(style_id, "")
            for candidate in (style_name, style_id):
                if not candidate:
                    continue
                # "Heading 1" with space (canonical name)
                m = _HEADING_RE.match(candidate)
                if m:
                    return min(int(m.group(1)), 6)
                # "heading1" / "Heading1" without space (style ID)
                lower = candidate.lower().replace(" ", "")
                if lower.startswith("heading") and lower[7:].isdigit():
                    return min(int(lower[7:]), 6)
                # French: "Titre1", "titre 1", "Titre 1"
                lower_nospace = candidate.lower().replace(" ", "")
                if lower_nospace.startswith("titre"):
                    rest = lower_nospace[5:]
                    if rest and rest.isdigit():
                        return min(int(rest), 6)

            # Title / Subtitle (exact match on name or ID)
            for candidate in (style_name, style_id):
                if not candidate:
                    continue
                lc = candidate.lower()
                if lc in ("title", "titre"):
                    return 1
                if lc in ("subtitle", "sous-titre", "soustitre"):
                    return 2

        # 3) Heuristic: bold text with large font size
        runs = para_elem.findall(qn("w:r"))
        if not runs:
            return None

        all_bold = True
        max_size: Optional[int] = None

        for r in runs:
            rpr = r.find(qn("w:rPr"))
            t = r.find(qn("w:t"))
            text = t.text if t is not None else ""
            if not text or not text.strip():
                continue

            # Check bold
            bold = False
            if rpr is not None:
                b_elem = rpr.find(qn("w:b"))
                if b_elem is not None:
                    val = b_elem.get(qn("w:val"), "true")
                    bold = val.lower() not in ("false", "0")
            if not bold:
                all_bold = False

            # Check font size
            sz = _font_size_from_rpr(rpr)
            if sz is not None:
                if max_size is None or sz > max_size:
                    max_size = sz

        if all_bold and max_size is not None and max_size >= _HEURISTIC_HEADING_MIN_SIZE:
            # Map font sizes heuristically:
            # >= 48 half-pts (24pt) -> level 1
            # >= 36 half-pts (18pt) -> level 2
            # >= 28 half-pts (14pt) -> level 3
            if max_size >= 48:
                return 1
            if max_size >= 36:
                return 2
            return 3

        return None

    def _detect_list_properties(self, para_elem) -> Optional[dict]:
        """Return list info ``{"level": int, "list_type": ListType}`` or *None*."""
        ppr = para_elem.find(qn("w:pPr"))
        if ppr is None:
            return None

        num_pr = ppr.find(qn("w:numPr"))
        if num_pr is None:
            return None

        ilvl_elem = num_pr.find(qn("w:ilvl"))
        num_id_elem = num_pr.find(qn("w:numId"))

        level = 0
        if ilvl_elem is not None:
            val = ilvl_elem.get(qn("w:val"), "0")
            level = int(val) if val.isdigit() else 0

        num_id = "0"
        if num_id_elem is not None:
            num_id = num_id_elem.get(qn("w:val"), "0")

        # numId 0 means "no numbering" in Word
        if num_id == "0":
            return None

        # Determine bullet vs numbered from the numbering definitions
        list_type = ListType.BULLET
        num_info = self._numbering_map.get(num_id)
        if num_info is not None:
            fmt = num_info.get("format", "bullet")
            if fmt in ("decimal", "lowerLetter", "upperLetter",
                        "lowerRoman", "upperRoman", "ordinal",
                        "cardinalText", "ordinalText"):
                list_type = ListType.NUMBERED

        return {"level": level, "list_type": list_type}

    # ── Run parsing ────────────────────────────────────────────────

    def _parse_runs(self, para_elem) -> list[TextRun]:
        """Parse all runs within a paragraph element, including hyperlinks."""
        result: list[TextRun] = []

        for child in para_elem:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

            if tag == "hyperlink":
                href = self._resolve_hyperlink(child)
                for r in child.findall(qn("w:r")):
                    run = self._parse_single_run(r)
                    if run is not None:
                        run.hyperlink = href
                        result.append(run)

            elif tag == "r":
                run = self._parse_single_run(child)
                if run is not None:
                    result.append(run)

        return result

    def _parse_single_run(self, r_elem) -> Optional[TextRun]:
        """Parse a ``<w:r>`` element into a :class:`TextRun`."""
        t_elem = r_elem.find(qn("w:t"))
        text = t_elem.text if t_elem is not None else ""
        if text is None:
            text = ""

        # Also pick up tab and break characters as whitespace
        if not text:
            if r_elem.find(qn("w:tab")) is not None:
                text = "\t"
            elif r_elem.find(qn("w:br")) is not None:
                # Soft line break (not page break -- those are handled above)
                br = r_elem.find(qn("w:br"))
                br_type = br.get(qn("w:type"), "textWrapping")
                if br_type == "textWrapping":
                    text = "\n"
                else:
                    return None
            else:
                return None

        rpr = r_elem.find(qn("w:rPr"))
        bold = False
        italic = False
        underline = False
        strikethrough = False
        color: Optional[str] = None
        font_size: Optional[int] = None

        if rpr is not None:
            bold = self._flag_value(rpr, "w:b")
            italic = self._flag_value(rpr, "w:i")
            underline = self._underline_value(rpr)
            strikethrough = (
                self._flag_value(rpr, "w:strike")
                or self._flag_value(rpr, "w:dstrike")
            )
            color = _color_str(rpr.find(qn("w:color")))
            font_size = _font_size_from_rpr(rpr)

        return TextRun(
            text=text,
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            color=color,
            font_size=font_size,
        )

    @staticmethod
    def _flag_value(rpr, tag: str) -> bool:
        """Read a boolean toggle element like ``<w:b/>`` or ``<w:b w:val="0"/>``."""
        elem = rpr.find(qn(tag))
        if elem is None:
            return False
        val = elem.get(qn("w:val"), "true")
        return val.lower() not in ("false", "0", "none")

    @staticmethod
    def _underline_value(rpr) -> bool:
        """Return *True* if ``<w:u>`` specifies an underline style other than none."""
        u = rpr.find(qn("w:u"))
        if u is None:
            return False
        val = u.get(qn("w:val"), "none")
        return val.lower() not in ("none", "false", "0")

    def _resolve_hyperlink(self, hyperlink_elem) -> Optional[str]:
        """Resolve a ``<w:hyperlink>`` element to its target URL."""
        # External link via r:id
        rid = hyperlink_elem.get(qn("r:id"))
        if rid and rid in self._rels:
            return self._rels[rid]
        # Bookmark / anchor link
        anchor = hyperlink_elem.get(qn("w:anchor"))
        if anchor:
            return f"#{anchor}"
        return None

    # ── Image extraction ───────────────────────────────────────────

    def _extract_images_from_paragraph(self, para_elem) -> list[Image]:
        """Extract all images embedded in a paragraph via ``<w:drawing>``."""
        assert self._doc is not None
        images: list[Image] = []

        for drawing in para_elem.iter(qn("w:drawing")):
            images.extend(self._parse_drawing(drawing))

        # Legacy VML images (<w:pict>)
        for pict in para_elem.iter(qn("w:pict")):
            img = self._parse_pict(pict)
            if img is not None:
                images.append(img)

        return images

    def _parse_drawing(self, drawing_elem) -> list[Image]:
        """Parse a ``<w:drawing>`` element for inline/anchor images."""
        assert self._doc is not None
        images: list[Image] = []

        # Both inline and anchor images share the same nested structure.
        for blip in drawing_elem.iter(qn("a:blip")):
            embed_id = blip.get(qn("r:embed"))
            if embed_id is None:
                continue

            try:
                image_part = self._doc.part.related_parts.get(embed_id)
                if image_part is None:
                    continue
            except Exception:
                logger.debug("Could not load image part %s", embed_id, exc_info=True)
                continue

            data = image_part.blob
            filename = image_part.partname.split("/")[-1] if hasattr(image_part, "partname") else "image.png"

            width, height = self._drawing_extent(drawing_elem)

            alt_text = self._drawing_alt_text(drawing_elem)

            images.append(Image(
                data=data,
                width=width,
                height=height,
                filename=filename,
                alt_text=alt_text,
            ))

        return images

    @staticmethod
    def _drawing_extent(drawing_elem) -> tuple[Optional[int], Optional[int]]:
        """Read ``<wp:extent>`` from a drawing element (EMU -> DXA)."""
        for extent in drawing_elem.iter(qn("wp:extent")):
            cx = extent.get("cx")
            cy = extent.get("cy")
            w = _emu_to_dxa(int(cx)) if cx and cx.isdigit() else None
            h = _emu_to_dxa(int(cy)) if cy and cy.isdigit() else None
            return w, h
        return None, None

    @staticmethod
    def _drawing_alt_text(drawing_elem) -> str:
        """Extract alt text from ``<wp:docPr>`` in a drawing element."""
        for doc_pr in drawing_elem.iter(qn("wp:docPr")):
            descr = doc_pr.get("descr", "")
            name = doc_pr.get("name", "")
            return descr or name
        return ""

    def _parse_pict(self, pict_elem) -> Optional[Image]:
        """Best-effort parse of legacy ``<w:pict>`` / VML image."""
        assert self._doc is not None
        # VML images use <v:imagedata r:id="rIdX" />
        ns_v = "urn:schemas-microsoft-com:vml"
        for img_data in pict_elem.iter(f"{{{ns_v}}}imagedata"):
            rid = img_data.get(qn("r:id"))
            if rid is None:
                continue
            try:
                image_part = self._doc.part.related_parts.get(rid)
                if image_part is None:
                    continue
            except Exception:
                continue

            data = image_part.blob
            filename = image_part.partname.split("/")[-1] if hasattr(image_part, "partname") else "image.png"
            return Image(data=data, filename=filename)

        return None

    # ── Table processing ───────────────────────────────────────────

    def _process_table(self, tbl_elem) -> dict:
        """Process a ``<w:tbl>`` element into a table result dict."""
        assert self._doc is not None

        # Use the python-docx Table wrapper for convenience.
        docx_table = DocxTable(tbl_elem, self._doc)

        headers: list[str] = []
        header_runs: list[list[TextRun]] = []
        rows: list[list[str]] = []
        cell_runs: list[list[list[TextRun]]] = []

        for row_idx, row in enumerate(docx_table.rows):
            row_texts: list[str] = []
            row_runs: list[list[TextRun]] = []

            for cell in row.cells:
                cell_text = cell.text.strip()
                row_texts.append(cell_text)

                # Collect runs from all paragraphs in the cell
                runs: list[TextRun] = []
                for para in cell.paragraphs:
                    runs.extend(self._parse_runs(para._element))  # noqa: SLF001
                row_runs.append(runs)

            if row_idx == 0:
                headers = row_texts
                header_runs = row_runs
            else:
                rows.append(row_texts)
                cell_runs.append(row_runs)

        return {
            "type": "table",
            "table": Table(
                headers=headers,
                rows=rows,
                header_runs=header_runs,
                cell_runs=cell_runs,
            ),
        }

    # ── Tree building ──────────────────────────────────────────────

    def _build_tree(
        self,
        metadata: DocumentMetadata,
        elements: list[dict],
    ) -> DocumentTree:
        """Organise flat element dicts into a hierarchical :class:`DocumentTree`."""

        tree = DocumentTree(metadata=metadata)

        # Group consecutive list items before building sections.
        elements = self._group_list_items(elements)

        # Track the current section (if any).
        current_section: Optional[Section] = None
        section_counter = 0

        for elem in elements:
            etype = elem["type"]

            if etype == "heading":
                # Start a new section.
                section_counter += 1
                current_section = Section(
                    heading=elem["text"],
                    level=elem["level"],
                    number=section_counter,
                )
                tree.sections.append(current_section)

            elif etype == "paragraph":
                content = Paragraph(runs=elem["runs"])
                self._append_content(tree, current_section, content)

            elif etype == "table":
                self._append_content(tree, current_section, elem["table"])

            elif etype == "image":
                self._append_content(tree, current_section, elem["image"])

            elif etype == "page_break":
                self._append_content(tree, current_section, PageBreak())

            elif etype == "list_block":
                self._append_content(tree, current_section, elem["block"])

            else:
                logger.debug("Unhandled element type during tree build: %s", etype)

        return tree

    @staticmethod
    def _append_content(
        tree: DocumentTree,
        section: Optional[Section],
        content: ContentElement,
    ) -> None:
        """Add *content* to the current section or the preamble."""
        if section is not None:
            section.children.append(content)
        else:
            tree.preamble.append(content)

    # ── List grouping ──────────────────────────────────────────────

    def _group_list_items(self, elements: list[dict]) -> list[dict]:
        """Merge consecutive ``list_item`` dicts into ``list_block`` dicts."""
        grouped: list[dict] = []
        buffer: list[dict] = []
        buffer_type: Optional[ListType] = None

        def flush() -> None:
            nonlocal buffer, buffer_type
            if buffer:
                items = self._nest_list_items(buffer)
                grouped.append({
                    "type": "list_block",
                    "block": ListBlock(list_type=buffer_type or ListType.BULLET, items=items),
                })
                buffer = []
                buffer_type = None

        for elem in elements:
            if elem["type"] == "list_item":
                lt = elem.get("list_type", ListType.BULLET)
                # If the list type changes, flush the current buffer first.
                if buffer and lt != buffer_type:
                    flush()
                buffer.append(elem)
                buffer_type = lt
            else:
                flush()
                grouped.append(elem)

        flush()
        return grouped

    @staticmethod
    def _nest_list_items(items: list[dict]) -> list[ListItem]:
        """Convert a flat list of item dicts into a nested :class:`ListItem` tree.

        Items with ``level > 0`` are attached as children of the nearest
        preceding item with a lower level.
        """
        root_items: list[ListItem] = []
        # Stack of (level, ListItem) to track nesting parents.
        stack: list[tuple[int, ListItem]] = []

        for item_dict in items:
            li = ListItem(
                runs=item_dict["runs"],
                level=item_dict["level"],
            )

            if not stack or item_dict["level"] == 0:
                root_items.append(li)
                stack = [(item_dict["level"], li)]
            else:
                # Pop stack until we find a parent with a strictly lower level.
                while stack and stack[-1][0] >= item_dict["level"]:
                    stack.pop()

                if stack:
                    stack[-1][1].children.append(li)
                else:
                    root_items.append(li)

                stack.append((item_dict["level"], li))

        return root_items
