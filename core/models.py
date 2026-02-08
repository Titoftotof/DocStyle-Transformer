"""Data models for the DocStyle Transformer intermediate representation.

These models represent the parsed document structure independently of any
source formatting, serving as the bridge between parsing and generation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class CalloutType(Enum):
    INFO = "info"
    WARNING = "warning"
    NOTE = "note"
    TIP = "tip"


class ListType(Enum):
    BULLET = "bullet"
    NUMBERED = "numbered"


# ── Inline formatting ───────────────────────────────────────────────


@dataclass
class TextRun:
    """A contiguous run of text sharing the same inline formatting."""
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    color: Optional[str] = None
    font_size: Optional[int] = None  # half-points
    hyperlink: Optional[str] = None


# ── Block-level elements ────────────────────────────────────────────


@dataclass
class Paragraph:
    """A body paragraph composed of styled text runs."""
    runs: list[TextRun] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Table:
    """A table with optional headers and data rows."""
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    header_runs: list[list[TextRun]] = field(default_factory=list)
    cell_runs: list[list[list[TextRun]]] = field(default_factory=list)


@dataclass
class Image:
    """An embedded image with its binary data and dimensions."""
    data: bytes = b""
    width: Optional[int] = None   # DXA
    height: Optional[int] = None  # DXA
    filename: str = "image.png"
    alt_text: str = ""


@dataclass
class Callout:
    """An info/warning/note callout box."""
    callout_type: CalloutType = CalloutType.INFO
    title: str = ""
    body: str = ""
    body_runs: list[TextRun] = field(default_factory=list)


@dataclass
class ListItem:
    """A single item in a list, possibly with nested children."""
    runs: list[TextRun] = field(default_factory=list)
    level: int = 0
    children: list[ListItem] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class ListBlock:
    """A contiguous list (bullet or numbered)."""
    list_type: ListType = ListType.BULLET
    items: list[ListItem] = field(default_factory=list)


@dataclass
class Step:
    """A single step in a numbered procedure."""
    number: int = 1
    title: str = ""
    description: str = ""
    description_runs: list[TextRun] = field(default_factory=list)


@dataclass
class StepsBlock:
    """A sequence of procedural steps."""
    steps: list[Step] = field(default_factory=list)


@dataclass
class PageBreak:
    """An explicit page break."""
    pass


# ── Content type union ──────────────────────────────────────────────

ContentElement = (
    Paragraph | Table | Image | Callout | ListBlock | StepsBlock | PageBreak
)


# ── Document structure ──────────────────────────────────────────────


@dataclass
class Section:
    """A document section starting with a heading."""
    heading: str = ""
    level: int = 1
    number: Optional[int] = None  # Auto-assigned section number (01, 02...)
    children: list[ContentElement] = field(default_factory=list)


@dataclass
class DocumentMetadata:
    """Document-level metadata."""
    title: str = ""
    author: str = ""
    date: str = ""
    version: str = ""
    reference: str = ""


@dataclass
class DocumentTree:
    """The complete intermediate representation of a parsed document."""
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    sections: list[Section] = field(default_factory=list)
    preamble: list[ContentElement] = field(default_factory=list)

    def section_count(self) -> int:
        return len(self.sections)

    def flat_elements(self) -> list[ContentElement]:
        """Return all content elements in document order."""
        elements: list[ContentElement] = list(self.preamble)
        for section in self.sections:
            elements.extend(section.children)
        return elements

    def summary(self) -> dict:
        """Return a summary of the document structure."""
        stats: dict = {
            "sections": len(self.sections),
            "paragraphs": 0,
            "tables": 0,
            "images": 0,
            "callouts": 0,
            "lists": 0,
            "steps_blocks": 0,
        }
        for elem in self.flat_elements():
            match elem:
                case Paragraph():
                    stats["paragraphs"] += 1
                case Table():
                    stats["tables"] += 1
                case Image():
                    stats["images"] += 1
                case Callout():
                    stats["callouts"] += 1
                case ListBlock():
                    stats["lists"] += 1
                case StepsBlock():
                    stats["steps_blocks"] += 1
        return stats
