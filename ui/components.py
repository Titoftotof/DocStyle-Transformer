"""Reusable tkinter/ttkbootstrap widgets for the DocStyle Transformer desktop UI.

Provides five main widget classes:

- FileSelector: file chooser for .docx files
- OptionsPanel: transformation options (theme, flags, overrides)
- StructurePreview: document structure tree viewer
- ResultPanel: transformation result display
- ProgressBar: determinate progress bar with label
"""

from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Any, Optional

try:
    import ttkbootstrap as ttkb
    from ttkbootstrap.constants import *  # noqa: F401, F403
    HAS_TTKB = True
except ImportError:
    HAS_TTKB = False

from core.models import (
    Callout,
    ContentElement,
    DocumentTree,
    Image,
    ListBlock,
    PageBreak,
    Paragraph,
    Section,
    StepsBlock,
    Table,
)

# ── Resolve project paths ────────────────────────────────────────────

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_THEMES_DIR = _PROJECT_ROOT / "themes"


# ── Helper ────────────────────────────────────────────────────────────


def _available_themes() -> list[str]:
    """Return a sorted list of theme names found in the themes/ directory.

    Each theme name is derived from a ``.yaml`` file by stripping the
    extension.  Returns an empty list if the directory does not exist.
    """
    if not _THEMES_DIR.is_dir():
        return []
    return sorted(
        p.stem for p in _THEMES_DIR.glob("*.yaml") if p.is_file()
    )


# ── FileSelector ──────────────────────────────────────────────────────


class FileSelector(ttk.Frame):
    """A frame containing a label, a read-only file path display, and a
    Browse button that opens a file dialog restricted to ``.docx`` files.

    Parameters
    ----------
    master : tk widget
        Parent widget.
    label : str
        Label text displayed to the left of the path entry.
    on_select : callable or None
        Optional callback invoked with the selected file path string
        whenever a new file is chosen.
    **kwargs
        Additional keyword arguments forwarded to ``ttk.Frame``.
    """

    def __init__(
        self,
        master: Any,
        label: str = "Input file:",
        on_select: Optional[Any] = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(master, **kwargs)
        self._on_select = on_select
        self._file_path_var = tk.StringVar(value="")

        # Label
        lbl = ttk.Label(self, text=label)
        lbl.pack(side=tk.LEFT, padx=(0, 5))

        # Path entry (read-only)
        self._entry = ttk.Entry(
            self, textvariable=self._file_path_var, state="readonly", width=60
        )
        self._entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Browse button
        if HAS_TTKB:
            btn = ttkb.Button(
                self, text="Browse", command=self._browse, bootstyle="outline"
            )
        else:
            btn = ttk.Button(self, text="Browse", command=self._browse)
        btn.pack(side=tk.LEFT)

    # -- properties -----------------------------------------------------

    @property
    def file_path(self) -> str:
        """The currently selected file path, or an empty string."""
        return self._file_path_var.get()

    @file_path.setter
    def file_path(self, value: str) -> None:
        self._file_path_var.set(value)

    # -- internals ------------------------------------------------------

    def _browse(self) -> None:
        """Open a file dialog and update the path variable."""
        initial_dir = ""
        current = self._file_path_var.get()
        if current:
            parent = Path(current).parent
            if parent.is_dir():
                initial_dir = str(parent)

        path = filedialog.askopenfilename(
            title="Select a Word document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialdir=initial_dir or None,
        )
        if path:
            self._file_path_var.set(path)
            if self._on_select is not None:
                self._on_select(path)


# ── OptionsPanel ──────────────────────────────────────────────────────


class OptionsPanel(ttk.LabelFrame):
    """A labelled frame exposing transformation options.

    Options include:
    - Theme selection (combobox populated from ``themes/*.yaml``)
    - Boolean flags: generate_cover, generate_toc, number_sections, header_footer
    - Text overrides: cover_title_override, mention

    Parameters
    ----------
    master : tk widget
        Parent widget.
    **kwargs
        Forwarded to ``ttk.LabelFrame``.
    """

    def __init__(self, master: Any, **kwargs: Any) -> None:
        kwargs.setdefault("text", "Options")
        super().__init__(master, **kwargs)

        # Internal variables
        self._theme_var = tk.StringVar(value="")
        self._generate_cover_var = tk.BooleanVar(value=True)
        self._generate_toc_var = tk.BooleanVar(value=True)
        self._number_sections_var = tk.BooleanVar(value=True)
        self._header_footer_var = tk.BooleanVar(value=True)
        self._cover_title_var = tk.StringVar(value="")
        self._mention_var = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self) -> None:
        """Construct all child widgets."""
        row = 0

        # ---- Theme selector ----
        ttk.Label(self, text="Theme:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=(5, 2)
        )
        themes = _available_themes()
        self._theme_combo = ttk.Combobox(
            self,
            textvariable=self._theme_var,
            values=themes,
            state="readonly",
            width=30,
        )
        self._theme_combo.grid(
            row=row, column=1, sticky=tk.W, padx=5, pady=(5, 2)
        )
        if themes:
            self._theme_combo.current(0)
        row += 1

        # ---- Boolean flags ----
        flags = [
            ("Generate cover page", self._generate_cover_var),
            ("Generate table of contents", self._generate_toc_var),
            ("Number sections", self._number_sections_var),
            ("Include header/footer", self._header_footer_var),
        ]
        for label_text, var in flags:
            cb = ttk.Checkbutton(self, text=label_text, variable=var)
            cb.grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
            row += 1

        # ---- Cover title override ----
        ttk.Label(self, text="Cover title override:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=(8, 2)
        )
        ttk.Entry(self, textvariable=self._cover_title_var, width=40).grid(
            row=row, column=1, sticky=tk.W + tk.E, padx=5, pady=(8, 2)
        )
        row += 1

        # ---- Mention text ----
        ttk.Label(self, text="Mention:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=(2, 8)
        )
        ttk.Entry(self, textvariable=self._mention_var, width=40).grid(
            row=row, column=1, sticky=tk.W + tk.E, padx=5, pady=(2, 8)
        )
        row += 1

        # Allow the entry column to stretch
        self.columnconfigure(1, weight=1)

    # -- public API -----------------------------------------------------

    def get_options(self) -> dict[str, Any]:
        """Return a dictionary of all current option values.

        Keys
        ----
        theme : str
            Selected theme name (stem, no extension), or empty string.
        generate_cover : bool
        generate_toc : bool
        number_sections : bool
        header_footer : bool
        cover_title_override : str
        mention : str
        """
        theme_name = self._theme_var.get()
        # Resolve to a full path if a theme was selected
        theme_path = ""
        if theme_name:
            candidate = _THEMES_DIR / f"{theme_name}.yaml"
            if candidate.is_file():
                theme_path = str(candidate)

        return {
            "theme": theme_name,
            "theme_path": theme_path,
            "generate_cover": self._generate_cover_var.get(),
            "generate_toc": self._generate_toc_var.get(),
            "number_sections": self._number_sections_var.get(),
            "header_footer": self._header_footer_var.get(),
            "cover_title_override": self._cover_title_var.get().strip(),
            "mention": self._mention_var.get().strip(),
        }

    def refresh_themes(self) -> None:
        """Re-scan the themes directory and update the combobox."""
        themes = _available_themes()
        self._theme_combo["values"] = themes
        if themes and not self._theme_var.get():
            self._theme_combo.current(0)


# ── StructurePreview ──────────────────────────────────────────────────


class StructurePreview(ttk.LabelFrame):
    """A labelled frame showing the document structure as a tree.

    Uses a ``ttk.Treeview`` to display sections, paragraphs, tables,
    images, callouts, lists, steps, and page breaks parsed from a
    :class:`DocumentTree`.

    Parameters
    ----------
    master : tk widget
        Parent widget.
    **kwargs
        Forwarded to ``ttk.LabelFrame``.
    """

    def __init__(self, master: Any, **kwargs: Any) -> None:
        kwargs.setdefault("text", "Document Structure")
        super().__init__(master, **kwargs)

        self._tree = ttk.Treeview(self, show="tree", selectmode="none")
        self._tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Scrollbar
        scrollbar = ttk.Scrollbar(
            self._tree, orient=tk.VERTICAL, command=self._tree.yview
        )
        self._tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # -- public API -----------------------------------------------------

    def update(self, tree: DocumentTree) -> None:
        """Clear the treeview and repopulate it from *tree*.

        Parameters
        ----------
        tree : DocumentTree
            The parsed document intermediate representation.
        """
        # Clear existing items
        for item in self._tree.get_children():
            self._tree.delete(item)

        # Metadata node
        meta = tree.metadata
        meta_id = self._tree.insert(
            "", tk.END, text=f"Metadata: {meta.title or '(untitled)'}"
        )
        if meta.author:
            self._tree.insert(meta_id, tk.END, text=f"Author: {meta.author}")
        if meta.date:
            self._tree.insert(meta_id, tk.END, text=f"Date: {meta.date}")
        if meta.version:
            self._tree.insert(meta_id, tk.END, text=f"Version: {meta.version}")

        # Preamble elements
        if tree.preamble:
            preamble_id = self._tree.insert("", tk.END, text="Preamble")
            for elem in tree.preamble:
                self._insert_element(preamble_id, elem)

        # Sections
        for section in tree.sections:
            section_label = self._section_label(section)
            section_id = self._tree.insert("", tk.END, text=section_label)
            for child in section.children:
                self._insert_element(section_id, child)

    # -- internals ------------------------------------------------------

    @staticmethod
    def _section_label(section: Section) -> str:
        """Build a display label for a section node."""
        prefix = f"H{section.level}"
        if section.number is not None:
            prefix += f" ({section.number:02d})"
        heading = section.heading if section.heading else "(empty)"
        return f"[{prefix}] {heading}"

    def _insert_element(
        self, parent_id: str, elem: ContentElement
    ) -> None:
        """Insert a single content element as a child of *parent_id*."""
        if isinstance(elem, Paragraph):
            text = elem.text
            preview = text[:80] + "..." if len(text) > 80 else text
            self._tree.insert(
                parent_id, tk.END, text=f"Paragraph: {preview}"
            )

        elif isinstance(elem, Table):
            cols = len(elem.headers) if elem.headers else (
                len(elem.rows[0]) if elem.rows else 0
            )
            rows = len(elem.rows)
            self._tree.insert(
                parent_id, tk.END,
                text=f"Table: {rows} row(s) x {cols} col(s)",
            )

        elif isinstance(elem, Image):
            self._tree.insert(
                parent_id, tk.END,
                text=f"Image: {elem.filename} ({elem.alt_text or 'no alt'})",
            )

        elif isinstance(elem, Callout):
            ctype = elem.callout_type.value.upper()
            title = elem.title or elem.body[:40]
            self._tree.insert(
                parent_id, tk.END, text=f"Callout [{ctype}]: {title}"
            )

        elif isinstance(elem, ListBlock):
            ltype = elem.list_type.value
            count = len(elem.items)
            self._tree.insert(
                parent_id, tk.END,
                text=f"List ({ltype}): {count} item(s)",
            )

        elif isinstance(elem, StepsBlock):
            count = len(elem.steps)
            self._tree.insert(
                parent_id, tk.END, text=f"Steps: {count} step(s)"
            )

        elif isinstance(elem, PageBreak):
            self._tree.insert(parent_id, tk.END, text="--- Page Break ---")

        else:
            self._tree.insert(
                parent_id, tk.END,
                text=f"Unknown: {type(elem).__name__}",
            )


# ── ResultPanel ───────────────────────────────────────────────────────


class ResultPanel(ttk.LabelFrame):
    """A labelled frame displaying the outcome of a transformation.

    Shows a success/failure indicator, document statistics, and provides
    an *Open folder* button to reveal the output file in the system file
    manager.

    Parameters
    ----------
    master : tk widget
        Parent widget.
    **kwargs
        Forwarded to ``ttk.LabelFrame``.
    """

    def __init__(self, master: Any, **kwargs: Any) -> None:
        kwargs.setdefault("text", "Results")
        super().__init__(master, **kwargs)

        self._status_var = tk.StringVar(value="No transformation run yet.")
        self._stats_var = tk.StringVar(value="")
        self._output_path: Optional[str] = None

        # Status label
        self._status_label = ttk.Label(
            self, textvariable=self._status_var, wraplength=500
        )
        self._status_label.pack(anchor=tk.W, padx=5, pady=(5, 2))

        # Stats label
        self._stats_label = ttk.Label(
            self, textvariable=self._stats_var, wraplength=500
        )
        self._stats_label.pack(anchor=tk.W, padx=5, pady=(0, 2))

        # Open folder button
        if HAS_TTKB:
            self._open_btn = ttkb.Button(
                self,
                text="Open folder",
                command=self._open_folder,
                state=tk.DISABLED,
                bootstyle="info-outline",
            )
        else:
            self._open_btn = ttk.Button(
                self,
                text="Open folder",
                command=self._open_folder,
                state=tk.DISABLED,
            )
        self._open_btn.pack(anchor=tk.W, padx=5, pady=(2, 8))

    # -- public API -----------------------------------------------------

    def show_success(self, output_path: str, stats: dict[str, Any]) -> None:
        """Display a successful transformation result.

        Parameters
        ----------
        output_path : str
            Path to the generated output file.
        stats : dict
            Summary statistics from ``DocumentTree.summary()``.
        """
        self._output_path = output_path
        self._status_var.set(f"Transformation successful: {output_path}")
        self._status_label.configure(foreground="green")

        parts: list[str] = []
        for key, value in stats.items():
            parts.append(f"{key}: {value}")
        self._stats_var.set(" | ".join(parts))

        self._open_btn.configure(state=tk.NORMAL)

    def show_failure(self, error_message: str) -> None:
        """Display a failed transformation result.

        Parameters
        ----------
        error_message : str
            Description of the error that occurred.
        """
        self._output_path = None
        self._status_var.set(f"Transformation failed: {error_message}")
        self._status_label.configure(foreground="red")
        self._stats_var.set("")
        self._open_btn.configure(state=tk.DISABLED)

    def reset(self) -> None:
        """Reset the panel to its initial empty state."""
        self._output_path = None
        self._status_var.set("No transformation run yet.")
        self._status_label.configure(foreground="")
        self._stats_var.set("")
        self._open_btn.configure(state=tk.DISABLED)

    # -- internals ------------------------------------------------------

    def _open_folder(self) -> None:
        """Open the containing folder of the output file in the system
        file manager."""
        if self._output_path is None:
            return

        folder = str(Path(self._output_path).parent)
        if not Path(folder).is_dir():
            return

        # Cross-platform open
        if sys.platform == "win32":
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])


# ── ProgressBar ───────────────────────────────────────────────────────


class ProgressBar(ttk.Frame):
    """A simple determinate progress bar with an accompanying text label.

    Parameters
    ----------
    master : tk widget
        Parent widget.
    **kwargs
        Forwarded to ``ttk.Frame``.
    """

    def __init__(self, master: Any, **kwargs: Any) -> None:
        super().__init__(master, **kwargs)

        self._label_var = tk.StringVar(value="Ready")

        self._label = ttk.Label(self, textvariable=self._label_var)
        self._label.pack(anchor=tk.W, padx=5, pady=(2, 0))

        if HAS_TTKB:
            self._bar = ttkb.Progressbar(
                self, mode="determinate", length=400, bootstyle="info-striped"
            )
        else:
            self._bar = ttk.Progressbar(
                self, mode="determinate", length=400
            )
        self._bar.pack(fill=tk.X, padx=5, pady=(0, 5))

    # -- public API -----------------------------------------------------

    def set_progress(self, value: float, label: Optional[str] = None) -> None:
        """Update the progress bar value and optionally the label text.

        Parameters
        ----------
        value : float
            Progress percentage (0-100).
        label : str or None
            If provided, replaces the current label text.
        """
        self._bar["value"] = max(0.0, min(100.0, value))
        if label is not None:
            self._label_var.set(label)

    def reset(self) -> None:
        """Reset progress to zero and label to *Ready*."""
        self._bar["value"] = 0
        self._label_var.set("Ready")

    @property
    def value(self) -> float:
        """The current progress value (0-100)."""
        return float(self._bar["value"])
