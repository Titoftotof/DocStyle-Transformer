"""Structure detection for parsed document trees.

Analyzes document elements to identify callouts, step sequences, and
structural patterns.  Works on the :class:`DocumentTree` produced by the
parser and returns an enriched copy with detected patterns replacing the
raw block-level elements they were derived from.
"""

from __future__ import annotations

import logging
import re
from typing import Sequence

from core.models import (
    Callout,
    CalloutType,
    ContentElement,
    DocumentTree,
    Paragraph,
    Section,
    Step,
    StepsBlock,
    Table,
    TextRun,
)

logger = logging.getLogger(__name__)

# ── Keyword-to-callout-type mapping ────────────────────────────────

_CALLOUT_KEYWORD_MAP: dict[str, CalloutType] = {
    # French keywords
    "attention": CalloutType.WARNING,
    "important": CalloutType.WARNING,
    "conseil": CalloutType.TIP,
    "bon à savoir": CalloutType.TIP,
    "note": CalloutType.NOTE,
    # English keywords
    "warning": CalloutType.WARNING,
    "tip": CalloutType.TIP,
    "info": CalloutType.INFO,
}

# Pattern that matches any of the recognised keyword prefixes at the start
# of a paragraph.  The keywords are tried longest-first so that
# "Bon à savoir" is matched before "Bon".
_KEYWORD_PATTERN: re.Pattern[str] = re.compile(
    r"^(?P<keyword>"
    + "|".join(
        re.escape(kw)
        for kw in sorted(_CALLOUT_KEYWORD_MAP, key=len, reverse=True)
    )
    + r")\s*[:]\s*(?P<body>.*)",
    re.IGNORECASE | re.DOTALL,
)

# Patterns for step detection: "Étape N", "Step N", or "N."
_STEP_PREFIX_PATTERN: re.Pattern[str] = re.compile(
    r"^(?:(?:[EÉ]tape|Step)\s+(?P<named_num>\d+))"
    r"|^(?P<dotted_num>\d+)\.\s",
    re.IGNORECASE,
)


# ── Helper utilities ───────────────────────────────────────────────


def _callout_type_for_keyword(keyword: str) -> CalloutType:
    """Return the :class:`CalloutType` for a recognised *keyword*.

    Falls back to :attr:`CalloutType.INFO` when no explicit mapping exists.
    """
    return _CALLOUT_KEYWORD_MAP.get(keyword.lower(), CalloutType.INFO)


def _is_single_cell_table(table: Table) -> bool:
    """Return ``True`` if the table has exactly one row and one column."""
    if table.headers and len(table.headers) == 1 and not table.rows:
        return True
    if len(table.rows) == 1 and len(table.rows[0]) == 1 and not table.headers:
        return True
    if (
        len(table.rows) == 1
        and len(table.rows[0]) == 1
        and len(table.headers) == 1
    ):
        return True
    return False


def _single_cell_text(table: Table) -> str:
    """Extract the plain text from a single-cell table."""
    if table.rows and table.rows[0]:
        return table.rows[0][0]
    if table.headers:
        return table.headers[0]
    return ""


def _single_cell_runs(table: Table) -> list[TextRun]:
    """Extract the :class:`TextRun` list from a single-cell table."""
    if table.cell_runs and table.cell_runs[0]:
        return list(table.cell_runs[0][0])
    if table.header_runs:
        return list(table.header_runs[0])
    return []


def _has_left_border(paragraph: Paragraph) -> bool:
    """Heuristic: a paragraph whose first run is indented or styled in a way
    that suggests a left-border callout.

    Since the IR does not carry border metadata directly, we look for a
    leading ``"|"`` or ``"▎"`` character that some converters inject to
    represent left-bordered blocks.
    """
    text = paragraph.text.lstrip()
    return text.startswith("|") or text.startswith("\u258e")  # ▎


def _strip_border_marker(text: str) -> str:
    """Remove a leading border-marker character and surrounding whitespace."""
    stripped = text.lstrip()
    if stripped.startswith("|") or stripped.startswith("\u258e"):
        stripped = stripped[1:].lstrip()
    return stripped


def _extract_step_number(text: str) -> int | None:
    """Return the step number from *text* if it matches a step pattern."""
    m = _STEP_PREFIX_PATTERN.match(text.strip())
    if m is None:
        return None
    num_str = m.group("named_num") or m.group("dotted_num")
    if num_str is None:
        return None
    try:
        return int(num_str)
    except ValueError:
        return None


def _strip_step_prefix(text: str) -> str:
    """Remove the step-number prefix from *text*."""
    m = _STEP_PREFIX_PATTERN.match(text.strip())
    if m is None:
        return text
    return text.strip()[m.end():].strip()


def _paragraph_has_bold_start(paragraph: Paragraph) -> bool:
    """Return ``True`` if the paragraph begins with bold text."""
    for run in paragraph.runs:
        if not run.text.strip():
            continue
        return run.bold
    return False


def _extract_bold_title(paragraph: Paragraph) -> str:
    """Return the leading bold segment of a paragraph as a plain string."""
    parts: list[str] = []
    for run in paragraph.runs:
        if run.bold:
            parts.append(run.text)
        elif parts:
            break
    return "".join(parts).strip()


# ── Main detector class ───────────────────────────────────────────


class StructureDetector:
    """Analyses a :class:`DocumentTree` and enriches it with higher-level
    structural elements (callouts, step sequences, normalised headings).

    Usage::

        detector = StructureDetector()
        enriched_tree = detector.detect(tree)
    """

    # ── public API ─────────────────────────────────────────────────

    def detect(self, tree: DocumentTree) -> DocumentTree:
        """Run all detection passes on *tree* and return the enriched tree.

        The detection pipeline is:

        1. Heading hierarchy normalisation (fill gaps, cap at level 3).
        2. Section numbering for level-1 sections.
        3. Callout detection (single-cell tables, keyword prefixes, borders).
        4. Step-sequence detection.
        """
        logger.info("Starting structure detection on document tree")

        self._normalize_heading_hierarchy(tree)
        self._assign_section_numbers(tree)

        # Process preamble elements
        tree.preamble = self._detect_elements(tree.preamble)

        # Process each section's children
        for section in tree.sections:
            section.children = self._detect_elements(section.children)

        summary = tree.summary()
        logger.info(
            "Detection complete: %d callouts, %d steps blocks identified",
            summary.get("callouts", 0),
            summary.get("steps_blocks", 0),
        )
        return tree

    # ── heading normalisation ──────────────────────────────────────

    @staticmethod
    def _normalize_heading_hierarchy(tree: DocumentTree) -> None:
        """Fill heading-level gaps and cap at level 3.

        If the document jumps from H1 to H3 with no H2, this method
        remaps all levels so that no gap exists.  Levels above 3 are
        clamped to 3.
        """
        if not tree.sections:
            return

        # Collect used levels
        used_levels = sorted({s.level for s in tree.sections})
        if not used_levels:
            return

        # Build a mapping from original level to compressed level
        level_map: dict[int, int] = {}
        new_level = 1
        for lvl in used_levels:
            mapped = min(new_level, 3)
            level_map[lvl] = mapped
            new_level += 1

        changed = any(k != v for k, v in level_map.items())
        if changed:
            logger.debug("Heading level mapping: %s", level_map)

        for section in tree.sections:
            original = section.level
            section.level = level_map.get(original, min(original, 3))
            if section.level != original:
                logger.debug(
                    "Remapped heading '%s' from level %d to %d",
                    section.heading,
                    original,
                    section.level,
                )

    # ── section numbering ──────────────────────────────────────────

    @staticmethod
    def _assign_section_numbers(tree: DocumentTree) -> None:
        """Assign sequential numbers (1, 2, 3, ...) to level-1 sections."""
        counter = 0
        for section in tree.sections:
            if section.level == 1:
                counter += 1
                section.number = counter
                logger.debug(
                    "Assigned number %d to section '%s'",
                    counter,
                    section.heading,
                )

    # ── element-level detection pipeline ───────────────────────────

    def _detect_elements(
        self, elements: list[ContentElement]
    ) -> list[ContentElement]:
        """Run callout and step detection on a list of elements.

        Returns a new list with raw elements replaced by their detected
        higher-level counterparts where applicable.
        """
        result = self._detect_callouts(elements)
        result = self._detect_steps(result)
        return result

    # ── callout detection ──────────────────────────────────────────

    def _detect_callouts(
        self, elements: Sequence[ContentElement]
    ) -> list[ContentElement]:
        """Scan *elements* for callout patterns and return a new list with
        matching elements replaced by :class:`Callout` instances.
        """
        result: list[ContentElement] = []

        for elem in elements:
            callout = self._try_table_callout(elem)
            if callout is not None:
                result.append(callout)
                continue

            callout = self._try_keyword_callout(elem)
            if callout is not None:
                result.append(callout)
                continue

            callout = self._try_border_callout(elem)
            if callout is not None:
                result.append(callout)
                continue

            result.append(elem)

        return result

    @staticmethod
    def _try_table_callout(elem: ContentElement) -> Callout | None:
        """If *elem* is a single-cell table, convert it to a callout."""
        if not isinstance(elem, Table):
            return None
        if not _is_single_cell_table(elem):
            return None

        text = _single_cell_text(elem).strip()
        if not text:
            return None

        runs = _single_cell_runs(elem)

        # Try to extract a keyword prefix from the cell text
        m = _KEYWORD_PATTERN.match(text)
        if m:
            keyword = m.group("keyword")
            body = m.group("body").strip()
            callout_type = _callout_type_for_keyword(keyword)
            title = keyword.capitalize()
        else:
            callout_type = CalloutType.INFO
            title = ""
            body = text

        logger.debug(
            "Detected table callout (%s): '%s'",
            callout_type.value,
            title or body[:40],
        )

        return Callout(
            callout_type=callout_type,
            title=title,
            body=body,
            body_runs=runs if runs else [TextRun(text=body)],
        )

    @staticmethod
    def _try_keyword_callout(elem: ContentElement) -> Callout | None:
        """If *elem* is a paragraph starting with a callout keyword, convert
        it to a callout.
        """
        if not isinstance(elem, Paragraph):
            return None

        text = elem.text.strip()
        if not text:
            return None

        m = _KEYWORD_PATTERN.match(text)
        if m is None:
            return None

        keyword = m.group("keyword")
        body = m.group("body").strip()
        callout_type = _callout_type_for_keyword(keyword)
        title = keyword.capitalize()

        logger.debug(
            "Detected keyword callout (%s): '%s'",
            callout_type.value,
            title,
        )

        # Build body_runs by dropping the keyword prefix from the original
        # runs if possible; otherwise fall back to a single plain run.
        body_runs = _build_body_runs_after_keyword(elem.runs, keyword)

        return Callout(
            callout_type=callout_type,
            title=title,
            body=body,
            body_runs=body_runs,
        )

    @staticmethod
    def _try_border_callout(elem: ContentElement) -> Callout | None:
        """If *elem* is a paragraph with a left-border marker, convert it to
        a callout.
        """
        if not isinstance(elem, Paragraph):
            return None
        if not _has_left_border(elem):
            return None

        raw_text = _strip_border_marker(elem.text)
        if not raw_text:
            return None

        # Attempt keyword detection within the bordered block
        m = _KEYWORD_PATTERN.match(raw_text)
        if m:
            keyword = m.group("keyword")
            body = m.group("body").strip()
            callout_type = _callout_type_for_keyword(keyword)
            title = keyword.capitalize()
        else:
            callout_type = CalloutType.NOTE
            title = ""
            body = raw_text

        logger.debug(
            "Detected border callout (%s): '%s'",
            callout_type.value,
            title or body[:40],
        )

        body_runs = _build_body_runs_stripped(elem.runs)

        return Callout(
            callout_type=callout_type,
            title=title,
            body=body,
            body_runs=body_runs,
        )

    # ── step sequence detection ────────────────────────────────────

    def _detect_steps(
        self, elements: Sequence[ContentElement]
    ) -> list[ContentElement]:
        """Scan *elements* for consecutive step paragraphs and group them
        into :class:`StepsBlock` instances.

        A step is recognised when a paragraph matches a step-prefix pattern
        (e.g. "Étape 1", "Step 2", "3.") optionally followed by bold title
        text.  A subsequent non-step paragraph is absorbed as the step's
        description.  Consecutive steps are grouped into one block.
        """
        result: list[ContentElement] = []
        i = 0
        n = len(elements)

        while i < n:
            step, consumed = self._try_parse_step(elements, i)
            if step is None:
                result.append(elements[i])
                i += 1
                continue

            # We found at least one step; collect consecutive steps.
            steps: list[Step] = [step]
            i += consumed

            while i < n:
                next_step, next_consumed = self._try_parse_step(elements, i)
                if next_step is None:
                    break
                steps.append(next_step)
                i += next_consumed

            block = StepsBlock(steps=steps)
            logger.debug(
                "Detected steps block with %d step(s) starting at '%s'",
                len(steps),
                steps[0].title or f"Step {steps[0].number}",
            )
            result.append(block)

        return result

    @staticmethod
    def _try_parse_step(
        elements: Sequence[ContentElement], index: int
    ) -> tuple[Step | None, int]:
        """Try to parse a step starting at *index* in *elements*.

        Returns ``(step, consumed_count)`` where *consumed_count* is the
        number of elements consumed from *elements*, or ``(None, 0)`` if no
        step was recognised.
        """
        if index >= len(elements):
            return None, 0

        elem = elements[index]
        if not isinstance(elem, Paragraph):
            return None, 0

        text = elem.text.strip()
        step_num = _extract_step_number(text)
        if step_num is None:
            return None, 0

        remainder = _strip_step_prefix(text)

        # The remainder may contain a bold title.  Otherwise the whole
        # remainder is treated as the title.
        if _paragraph_has_bold_start(elem):
            title = _extract_bold_title(elem)
            # If the bold title is embedded inside the step prefix line,
            # strip it from the remainder to obtain the inline description.
            if title and remainder.startswith(title):
                description_inline = remainder[len(title):].strip()
            else:
                description_inline = remainder
                if not title:
                    title = remainder
        else:
            title = remainder
            description_inline = ""

        consumed = 1
        description = description_inline
        description_runs: list[TextRun] = []

        # Look ahead: if the next element is a plain paragraph (not itself a
        # step), absorb it as the step's description.
        if index + 1 < len(elements):
            next_elem = elements[index + 1]
            if isinstance(next_elem, Paragraph):
                next_text = next_elem.text.strip()
                if next_text and _extract_step_number(next_text) is None:
                    if description:
                        description = description + " " + next_text
                    else:
                        description = next_text
                    description_runs = list(next_elem.runs)
                    consumed = 2

        # If no separate description paragraph was absorbed but we have
        # inline description text, build runs from the original paragraph.
        if not description_runs and description:
            description_runs = [TextRun(text=description)]

        step = Step(
            number=step_num,
            title=title,
            description=description,
            description_runs=description_runs,
        )
        return step, consumed


# ── Module-level helpers (used by static methods above) ────────────


def _build_body_runs_after_keyword(
    runs: list[TextRun], keyword: str
) -> list[TextRun]:
    """Return a new run list with the keyword prefix and separator removed.

    Walks through *runs* and drops text corresponding to the keyword and the
    ``":"`` separator, preserving formatting on the remaining content.
    """
    to_skip = keyword.lower()
    body_runs: list[TextRun] = []
    skipped = False
    accumulated = ""

    for run in runs:
        if skipped:
            body_runs.append(run)
            continue

        accumulated += run.text
        lower_acc = accumulated.lower()

        # Check if we have passed the keyword + separator
        sep_idx = lower_acc.find(to_skip)
        if sep_idx == -1:
            continue

        after_keyword = accumulated[sep_idx + len(to_skip):]
        # Skip optional whitespace and colon
        m = re.match(r"\s*:\s*", after_keyword)
        if m:
            leftover = after_keyword[m.end():]
        else:
            leftover = after_keyword

        skipped = True
        if leftover.strip():
            body_runs.append(
                TextRun(
                    text=leftover,
                    bold=run.bold,
                    italic=run.italic,
                    underline=run.underline,
                    strikethrough=run.strikethrough,
                    color=run.color,
                    font_size=run.font_size,
                    hyperlink=run.hyperlink,
                )
            )

    if not body_runs:
        # Fallback: could not split; return a plain run with the body text
        full_text = "".join(r.text for r in runs)
        m_kw = _KEYWORD_PATTERN.match(full_text.strip())
        if m_kw:
            return [TextRun(text=m_kw.group("body").strip())]
        return [TextRun(text=full_text.strip())]

    return body_runs


def _build_body_runs_stripped(runs: list[TextRun]) -> list[TextRun]:
    """Return a new run list with the leading border marker removed."""
    if not runs:
        return []

    first = runs[0]
    stripped_text = first.text.lstrip()
    if stripped_text.startswith("|") or stripped_text.startswith("\u258e"):
        stripped_text = stripped_text[1:].lstrip()

    new_first = TextRun(
        text=stripped_text,
        bold=first.bold,
        italic=first.italic,
        underline=first.underline,
        strikethrough=first.strikethrough,
        color=first.color,
        font_size=first.font_size,
        hyperlink=first.hyperlink,
    )

    result = [new_first] if new_first.text else []
    result.extend(runs[1:])
    return result
