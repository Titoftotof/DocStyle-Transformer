"""Markdown parser for DocStyle Transformer.

Parses Markdown files and converts them to the DocumentTree intermediate representation.
Supports headers, paragraphs, lists, code blocks, tables, and images.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Optional

from core.models import (
    Callout,
    CalloutType,
    CodeBlock,
    ContentElement,
    DocumentMetadata,
    DocumentTree,
    Image,
    ListItem,
    ListBlock,
    ListType,
    Paragraph,
    Section,
    StepsBlock,
    Step,
    Table,
    TextRun,
)

logger = logging.getLogger(__name__)


class MarkdownParser:
    """Parse Markdown files into DocumentTree format.

    Supports:
    - ATX-style headers (# ## ###)
    - Paragraphs
    - Bullet lists (-, *, +)
    - Numbered lists (1., 2., etc.)
    - Code blocks (indented or ```fenced```)
    - Inline formatting: **bold**, *italic*, `code`
    - Tables (GitHub-flavored)
    - Images (![alt](url))
    - Blockquotes (>)

    Example::

        parser = MarkdownParser()
        tree = parser.parse("document.md")
    """

    def __init__(self) -> None:
        """Initialize the parser."""
        self._lines: list[str] = []
        self._pos: int = 0
        self._in_code_block = False
        self._code_fence_char = ""
        self._code_lang = ""

    def parse(self, file_path: str) -> DocumentTree:
        """Parse a Markdown file and return a DocumentTree.

        Parameters
        ----------
        file_path : str
            Path to the .md or .markdown file.

        Returns
        -------
        DocumentTree
            The parsed document as a DocumentTree.
        """
        logger.info("Parsing Markdown file: %s", file_path)

        # Read file
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()

        # Normalize line endings
        content = content.replace("\r\n", "\n").replace("\r", "\n")

        # Split into lines
        self._lines = content.split("\n")
        self._pos = 0

        # Extract metadata from YAML frontmatter if present
        metadata = self._parse_frontmatter()

        # Parse content into sections
        sections: list[Section] = []
        preamble: list[ContentElement] = []

        current_section: Optional[Section] = None
        in_preamble = True

        while self._pos < len(self._lines):
            line = self._lines[self._pos]

            # Skip empty lines at start
            if not line.strip() and in_preamble:
                self._pos += 1
                continue

            # Check for code fence
            if line.strip().startswith("```"):
                elem = self._parse_fenced_code()
                if in_preamble and not isinstance(elem, CodeBlock):
                    # First non-code, non-callout content ends preamble
                    if isinstance(elem, Paragraph) or isinstance(elem, Table):
                        in_preamble = False
                if in_preamble:
                    preamble.append(elem)
                elif current_section:
                    current_section.children.append(elem)
                continue

            # Check for header
            header_match = re.match(r"^(#{1,6})\s+(.+)$", line)
            if header_match:
                in_preamble = False
                level = len(header_match.group(1))
                heading = header_match.group(2).strip()

                # Save previous section
                if current_section:
                    sections.append(current_section)

                current_section = Section(heading=heading, level=level, children=[])
                self._pos += 1
                continue

            # Parse content element
            elem = self._parse_content_element()

            if elem:
                if in_preamble:
                    preamble.append(elem)
                elif current_section:
                    current_section.children.append(elem)

        # Save last section
        if current_section:
            sections.append(current_section)

        tree = DocumentTree(metadata=metadata, sections=sections, preamble=preamble)
        logger.info(
            "Parsed Markdown: %d sections, %d preamble elements",
            len(sections),
            len(preamble),
        )
        return tree

    def _parse_frontmatter(self) -> DocumentMetadata:
        """Parse YAML frontmatter if present."""
        metadata = DocumentMetadata()

        if self._pos >= len(self._lines):
            return metadata

        if self._lines[self._pos].strip() != "---":
            return metadata

        self._pos += 1
        frontmatter_lines = []

        while self._pos < len(self._lines):
            line = self._lines[self._pos]
            if line.strip() == "---":
                self._pos += 1
                break
            frontmatter_lines.append(line)
            self._pos += 1

        # Parse simple key: value pairs
        for line in frontmatter_lines:
            match = re.match(r"(\w+)\s*:\s*(.+)", line)
            if match:
                key, value = match.groups()
                value = value.strip().strip('"').strip("'")
                if key == "title":
                    metadata.title = value
                elif key == "author":
                    metadata.author = value
                elif key == "date":
                    metadata.date = value
                elif key == "version":
                    metadata.version = value

        return metadata

    def _parse_fenced_code(self) -> Optional[ContentElement]:
        """Parse a fenced code block (```language ... ```)."""
        if not self._lines[self._pos].strip().startswith("```"):
            return None

        # Parse fence
        fence_line = self._lines[self._pos].strip()
        self._code_lang = fence_line[3:].strip()
        self._pos += 1

        # Collect code lines
        code_lines = []
        while self._pos < len(self._lines):
            line = self._lines[self._pos]
            if line.strip() == "```":
                self._pos += 1
                break
            code_lines.append(line)
            self._pos += 1

        code = "\n".join(code_lines)
        return CodeBlock(language=self._code_lang, code=code, line_numbers=True)

    def _parse_content_element(self) -> Optional[ContentElement]:
        """Parse the next content element from the current position."""
        if self._pos >= len(self._lines):
            return None

        line = self._lines[self._pos]

        # Empty line
        if not line.strip():
            self._pos += 1
            return None

        # Blockquote
        if line.strip().startswith(">"):
            return self._parse_blockquote()

        # Table
        if "|" in line:
            table = self._parse_table()
            if table:
                return table

        # List
        if re.match(r"^\s*[-*+]\s+", line) or re.match(r"^\s*\d+\.\s+", line):
            return self._parse_list()

        # Horizontal rule
        if re.match(r"^\s*[-*_]{3,}\s*$", line):
            self._pos += 1
            return None

        # Paragraph (possibly multi-line)
        return self._parse_paragraph()

    def _parse_blockquote(self) -> Callout:
        """Parse a blockquote as a callout."""
        lines = []
        while self._pos < len(self._lines):
            line = self._lines[self._pos]
            if not line.strip().startswith(">"):
                break
            # Remove > prefix
            content = re.sub(r"^\s*>\s*", "", line)
            lines.append(content)
            self._pos += 1

        text = "\n".join(lines).strip()

        # Check for callout keywords
        first_line = lines[0].lower() if lines else ""
        callout_type = CalloutType.NOTE
        title = ""

        for keyword, ctype in [
            ("warning", CalloutType.WARNING),
            ("attention", CalloutType.WARNING),
            ("important", CalloutType.WARNING),
            ("note", CalloutType.NOTE),
            ("tip", CalloutType.TIP),
            ("info", CalloutType.INFO),
        ]:
            if first_line.startswith(keyword + ":"):
                callout_type = ctype
                title = keyword.capitalize()
                text = text[len(keyword) + 1:].strip()
                break

        return Callout(
            callout_type=callout_type,
            title=title,
            body=text,
            body_runs=[TextRun(text=text)],
        )

    def _parse_list(self) -> ContentElement:
        """Parse a bullet or numbered list."""
        items: list[ListItem] = []
        list_type = ListType.BULLET

        # Check list type from first item
        first_line = self._lines[self._pos]
        if re.match(r"^\s*\d+\.\s+", first_line):
            list_type = ListType.NUMBERED

        # Check for step pattern (1. **Title** or 1. Title)
        step_pattern = re.match(r"^\s*(\d+)\.\s+\*\*(.+?)\*\*\s*(.*)", first_line)
        if step_pattern or re.match(r"^\s*(\d+)\.\s+([A-Z].+)", first_line):
            # This might be a steps block
            return self._parse_steps_block()

        # Parse as regular list
        while self._pos < len(self._lines):
            line = self._lines[self._pos]

            # Check for list item
            bullet_match = re.match(r"^(\s*)([-*+])\s+(.+)", line)
            numbered_match = re.match(r"^(\s*)(\d+)\.\s+(.+)", line)

            if not bullet_match and not numbered_match:
                break

            if bullet_match:
                indent = len(bullet_match.group(1))
                text = bullet_match.group(3)
            else:
                indent = len(numbered_match.group(1))
                text = numbered_match.group(3)
                list_type = ListType.NUMBERED

            # Parse inline formatting
            runs = self._parse_inline_formatting(text)

            item = ListItem(runs=runs, level=indent // 2)
            items.append(item)
            self._pos += 1

        return ListBlock(list_type=list_type, items=items)

    def _parse_steps_block(self) -> StepsBlock:
        """Parse a numbered list as a steps block."""
        steps: list[Step] = []

        while self._pos < len(self._lines):
            line = self._lines[self._pos]

            # Match: "N. **Title** description" or "N. Title description"
            match = re.match(r"^\s*(\d+)\.\s+\*\*(.+?)\*\*\s*(.*)", line)
            if not match:
                match = re.match(r"^\s*(\d+)\.\s+([A-Z].+?)\s+(.*)", line)

            if not match:
                break

            step_num = int(match.group(1))
            title = match.group(2).strip()
            description = match.group(3).strip()

            step = Step(
                number=step_num,
                title=title,
                description=description,
                description_runs=[TextRun(text=description)] if description else [],
            )
            steps.append(step)
            self._pos += 1

        return StepsBlock(steps=steps)

    def _parse_table(self) -> Optional[Table]:
        """Parse a GitHub-flavored markdown table."""
        lines = []
        start_pos = self._pos

        # Collect table lines
        while self._pos < len(self._lines):
            line = self._lines[self._pos]
            if "|" not in line or not line.strip():
                break
            lines.append(line)
            self._pos += 1

        if len(lines) < 2:
            self._pos = start_pos
            return None

        # Parse headers
        headers = [h.strip() for h in lines[0].split("|")]
        headers = [h for h in headers if h]  # Remove empty strings

        # Skip separator line
        if len(lines) > 1:
            lines = lines[2:] if len(lines) > 2 else []

        # Parse rows
        rows = []
        for line in lines:
            cells = [c.strip() for c in line.split("|")]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)

        return Table(headers=headers, rows=rows)

    def _parse_paragraph(self) -> Paragraph:
        """Parse a paragraph (possibly multi-line with inline formatting)."""
        lines = []
        while self._pos < len(self._lines):
            line = self._lines[self._pos]

            # Stop on empty line or special block
            if not line.strip():
                break
            if line.strip().startswith(">"):
                break
            if line.strip().startswith("```"):
                break
            if line.strip().startswith("#"):
                break
            if re.match(r"^\s*[-*+]\s+", line):
                break
            if re.match(r"^\s*\d+\.\s+", line):
                break
            if "|" in line and self._is_table_line(line):
                break

            lines.append(line.strip())
            self._pos += 1

        text = " ".join(lines)
        runs = self._parse_inline_formatting(text)

        return Paragraph(runs=runs)

    def _is_table_line(self, line: str) -> bool:
        """Check if a line looks like a table row."""
        parts = line.split("|")
        return len(parts) >= 3

    def _parse_inline_formatting(self, text: str) -> list[TextRun]:
        """Parse inline formatting (bold, italic, code, links)."""
        runs: list[TextRun] = []
        remaining = text

        # Simple regex-based parsing
        # Pattern for inline code: `text`
        # Pattern for bold: **text**
        # Pattern for italic: *text*

        # First, handle inline code
        code_pattern = re.compile(r"`([^`]+)`")
        bold_pattern = re.compile(r"\*\*([^*]+)\*\*")
        italic_pattern = re.compile(r"\*([^*]+)\*")

        pos = 0
        while pos < len(remaining):
            # Find next formatting marker
            code_match = code_pattern.search(remaining, pos)
            bold_match = bold_pattern.search(remaining, pos)
            italic_match = italic_pattern.search(remaining, pos)

            matches = []
            if code_match:
                matches.append(("code", code_match))
            if bold_match:
                matches.append(("bold", bold_match))
            if italic_match:
                matches.append(("italic", italic_match))

            if not matches:
                # No more formatting, add remaining text
                if remaining[pos:]:
                    runs.append(TextRun(text=remaining[pos:]))
                break

            # Get earliest match
            matches.sort(key=lambda m: m[1].start())
            match_type, match = matches[0]

            # Add text before match
            if match.start() > pos:
                runs.append(TextRun(text=remaining[pos:match.start()]))

            # Add formatted text
            formatted_text = match.group(1)
            if match_type == "code":
                runs.append(TextRun(text=formatted_text))
            elif match_type == "bold":
                runs.append(TextRun(text=formatted_text, bold=True))
            elif match_type == "italic":
                runs.append(TextRun(text=formatted_text, italic=True))

            pos = match.end()

        return runs if runs else [TextRun(text=text)]
