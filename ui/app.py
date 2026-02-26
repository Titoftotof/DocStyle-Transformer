"""Main application window for the DocStyle Transformer desktop UI.

Assembles the reusable widgets from :mod:`ui.components` into a complete
application.  The transformation is executed in a background thread so that
the UI remains responsive, with progress updates delivered through a
:class:`queue.Queue`.
"""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Optional

try:
    import ttkbootstrap as ttkb
    HAS_TTKB = True
except ImportError:
    HAS_TTKB = False

from core.detector import StructureDetector
from core.models import DocumentTree
from core.parser import DocxParser
from core.markdown_parser import MarkdownParser
from core.settings import SettingsManager

from ui.components import (
    FileSelector,
    OptionsPanel,
    ProgressBar,
    ResultPanel,
    StructurePreview,
)

logger = logging.getLogger(__name__)

# ── Constants ────────────────────────────────────────────────────────

_APP_TITLE = "DocStyle Transformer"
_WINDOW_MIN_WIDTH = 800
_WINDOW_MIN_HEIGHT = 620
_POLL_INTERVAL_MS = 100  # how often the UI checks the progress queue


# ── Progress message types ────────────────────────────────────────────


class _ProgressMsg:
    """A message sent from the worker thread to the UI via the queue."""

    def __init__(
        self,
        kind: str,
        value: float = 0.0,
        label: str = "",
        data: Any = None,
    ) -> None:
        self.kind = kind      # "progress", "done", "error"
        self.value = value    # 0-100
        self.label = label    # human-readable status text
        self.data = data      # arbitrary payload (output path, stats, error msg)


# ── DocStyleApp ───────────────────────────────────────────────────────


class DocStyleApp:
    """Main application class that wires together all UI components.

    Usage::

        app = DocStyleApp()
        app.run()
    """

    def __init__(self) -> None:
        # ---- load settings first to determine theme ----
        self._settings = SettingsManager()
        use_dark_mode = self._settings.get_dark_mode()
        theme_name = "darkly" if use_dark_mode else "cosmo"
        self._current_theme = theme_name

        # ---- main window ----
        if HAS_TTKB:
            self.root = ttkb.Window(
                title=_APP_TITLE,
                themename=theme_name,
                minsize=(_WINDOW_MIN_WIDTH, _WINDOW_MIN_HEIGHT),
            )
        else:
            self.root = tk.Tk()
            self.root.title(_APP_TITLE)
            self.root.minsize(_WINDOW_MIN_WIDTH, _WINDOW_MIN_HEIGHT)

        # Centre on screen
        self.root.geometry(
            f"{_WINDOW_MIN_WIDTH}x{_WINDOW_MIN_HEIGHT}"
            f"+{(self.root.winfo_screenwidth() - _WINDOW_MIN_WIDTH) // 2}"
            f"+{(self.root.winfo_screenheight() - _WINDOW_MIN_HEIGHT) // 2}"
        )

        # ---- state ----
        self._parser = DocxParser()
        self._detector = StructureDetector()
        self._current_tree: Optional[DocumentTree] = None
        self._progress_queue: queue.Queue[_ProgressMsg] = queue.Queue()
        self._worker_thread: Optional[threading.Thread] = None

        # ---- build UI ----
        self._build_layout()
        self._build_menu()

        # ---- keyboard shortcuts ----
        self._bind_shortcuts()

        # ---- load saved options ----
        self._load_saved_options()

    # ── Layout ────────────────────────────────────────────────────────

    def _build_menu(self) -> None:
        """Create the application menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open... (Ctrl+O)", command=self._file_selector._browse)
        file_menu.add_command(label="Transform (Ctrl+T)", command=self._shortcut_transform)
        file_menu.add_separator()
        file_menu.add_command(label="Exit (Ctrl+Q)", command=self.root.quit)

        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)

        # Dark mode toggle
        self._dark_mode_var = tk.BooleanVar(value=self._settings.get_dark_mode())
        view_menu.add_checkbutton(
            label="Dark mode",
            variable=self._dark_mode_var,
            command=self._toggle_dark_mode,
        )

        # Separator
        view_menu.add_separator()

        # Recent files submenu
        self._recent_menu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Recent files", menu=self._recent_menu)
        self._update_recent_menu()

    def _update_recent_menu(self) -> None:
        """Update the recent files menu with current list."""
        # Clear existing items
        self._recent_menu.delete(0, tk.END)

        recent_files = self._settings.get_recent_files()
        if not recent_files:
            self._recent_menu.add_command(label="No recent files", state=tk.DISABLED)
        else:
            for i, file_path in enumerate(recent_files[:10], 1):
                # Show just filename with path as tooltip-ish effect
                name = Path(file_path).name
                self._recent_menu.add_command(
                    label=f"{i}. {name}",
                    command=lambda p=file_path: self._load_recent_file(p),
                )
            # Separator and clear option
            self._recent_menu.add_separator()
            self._recent_menu.add_command(
                label="Clear history",
                command=self._clear_recent_files,
            )

    def _load_recent_file(self, file_path: str) -> None:
        """Load a file from the recent files menu."""
        if Path(file_path).exists():
            self._file_selector.file_path = file_path
            self._on_file_selected(file_path)
        else:
            messagebox.showwarning(
                "File Not Found",
                f"The file '{file_path}' no longer exists.",
                parent=self.root,
            )
            # Remove from history
            recent = [p for p in self._settings.get_recent_files() if p != file_path]
            # Update settings
            for p in self._settings.get_recent_files():
                if p not in recent:
                    # Clear and rebuild
                    self._settings.clear_recent_files()
                    for rp in recent:
                        self._settings.add_recent_file(rp)
                    break
            self._update_recent_menu()

    def _clear_recent_files(self) -> None:
        """Clear the recent files history."""
        self._settings.clear_recent_files()
        self._update_recent_menu()

    def _toggle_dark_mode(self) -> None:
        """Toggle dark mode on/off."""
        current = self._dark_mode_var.get()
        self._apply_theme(dark=current)
        self._settings.set_dark_mode(current)

    def _apply_theme(self, dark: bool) -> None:
        """Apply the light or dark theme.

        Note: ttkbootstrap requires restarting the app to change themes properly.
        The settings are saved and will apply on next launch.
        """
        target_theme = "darkly" if dark else "cosmo"
        self._current_theme = target_theme

        if dark:
            messagebox.showinfo(
                "Dark Mode Activated",
                "Dark mode has been enabled.\n\n"
                "Please restart the application for the theme to be fully applied.",
                parent=self.root,
            )
        else:
            messagebox.showinfo(
                "Light Mode Activated",
                "Light mode has been enabled.\n\n"
                "Please restart the application for the theme to be fully applied.",
                parent=self.root,
            )

    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts for common actions."""
        self.root.bind("<Control-o>", lambda e: self._file_selector._browse())
        self.root.bind("<Control-O>", lambda e: self._file_selector._browse())
        self.root.bind("<Control-t>", self._shortcut_transform)
        self.root.bind("<Control-T>", self._shortcut_transform)
        self.root.bind("<Control-q>", lambda e: self.root.quit())
        self.root.bind("<Control-Q>", lambda e: self.root.quit())

    def _shortcut_transform(self, event) -> None:
        """Handle Ctrl+T shortcut - only enable if button is active."""
        if self._transform_btn["state"] == tk.NORMAL:
            self._on_transform()

    def _load_saved_options(self) -> None:
        """Load saved options from settings into the UI."""
        options = self._settings.load_options()

        # Apply to options panel
        self._options_panel._theme_var.set(options.get("theme", ""))
        self._options_panel._generate_cover_var.set(options.get("generate_cover", True))
        self._options_panel._generate_toc_var.set(options.get("generate_toc", True))
        self._options_panel._number_sections_var.set(options.get("number_sections", True))
        self._options_panel._header_footer_var.set(options.get("header_footer", True))
        self._options_panel._cover_title_var.set(options.get("cover_title_override", ""))
        self._options_panel._mention_var.set(options.get("mention", ""))

    def _save_options(self) -> None:
        """Save current options from the UI to settings."""
        options = self._options_panel.get_options()
        self._settings.save_options(options)

    def _build_layout(self) -> None:
        """Create and arrange all widgets inside the main window."""
        pad = {"padx": 10, "pady": 5}

        # -- top: file selector --
        self._file_selector = FileSelector(
            self.root,
            label="Input file (.docx, .md):",
            on_select=self._on_file_selected,
        )
        self._file_selector.pack(fill=tk.X, **pad)

        # -- middle: options + structure preview side by side --
        middle_frame = ttk.Frame(self.root)
        middle_frame.pack(fill=tk.BOTH, expand=True, **pad)

        self._options_panel = OptionsPanel(middle_frame)
        self._options_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))

        self._structure_preview = StructurePreview(middle_frame)
        self._structure_preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # -- transform button --
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, **pad)

        if HAS_TTKB:
            self._transform_btn = ttkb.Button(
                btn_frame,
                text="Transform",
                command=self._on_transform,
                bootstyle="success",
                state=tk.DISABLED,
            )
        else:
            self._transform_btn = ttk.Button(
                btn_frame,
                text="Transform",
                command=self._on_transform,
                state=tk.DISABLED,
            )
        self._transform_btn.pack(side=tk.LEFT)

        # -- progress bar --
        self._progress_bar = ProgressBar(btn_frame)
        self._progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # -- bottom: result panel --
        self._result_panel = ResultPanel(self.root)
        self._result_panel.pack(fill=tk.X, **pad)

    # ── File selection callback ───────────────────────────────────────

    def _on_file_selected(self, path: str) -> None:
        """Parse the selected file and update the structure preview.

        Called by the :class:`FileSelector` widget when the user picks a
        new file (supports .docx, .md, .markdown).
        """
        self._result_panel.reset()
        self._progress_bar.reset()
        self._current_tree = None
        self._transform_btn.configure(state=tk.DISABLED)

        if not path or not Path(path).is_file():
            return

        self._progress_bar.set_progress(10, "Parsing document...")

        try:
            # Choose parser based on file extension
            file_ext = Path(path).suffix.lower()
            if file_ext in (".md", ".markdown"):
                parser = MarkdownParser()
            else:
                parser = self._parser

            tree = parser.parse(path)
            tree = self._detector.detect(tree)
        except Exception as exc:
            logger.exception("Failed to parse %s", path)
            messagebox.showerror(
                "Parse Error",
                f"Could not parse the selected file:\n\n{exc}",
                parent=self.root,
            )
            self._progress_bar.reset()
            return

        self._current_tree = tree
        self._structure_preview.update(tree)
        self._progress_bar.set_progress(100, "Document parsed successfully")
        self._transform_btn.configure(state=tk.NORMAL)

        logger.info(
            "Parsed '%s': %d sections, %d preamble elements",
            path,
            tree.section_count(),
            len(tree.preamble),
        )

    # ── Transform callback ────────────────────────────────────────────

    def _on_transform(self) -> None:
        """Kick off the transformation in a background thread.

        Disables the transform button, resets results, and starts polling
        the progress queue for updates from the worker.
        """
        input_path = self._file_selector.file_path
        if not input_path or self._current_tree is None:
            return

        # Prevent double-clicks
        if self._worker_thread is not None and self._worker_thread.is_alive():
            return

        self._transform_btn.configure(state=tk.DISABLED)
        self._result_panel.reset()
        self._progress_bar.set_progress(0, "Starting transformation...")

        options = self._options_panel.get_options()

        self._worker_thread = threading.Thread(
            target=self._transform_worker,
            args=(input_path, self._current_tree, options),
            daemon=True,
        )
        self._worker_thread.start()

        # Begin polling the queue
        self.root.after(_POLL_INTERVAL_MS, self._poll_progress)

    # ── Background worker ─────────────────────────────────────────────

    def _transform_worker(
        self,
        input_path: str,
        tree: DocumentTree,
        options: dict[str, Any],
    ) -> None:
        """Run the transformation pipeline in a background thread.

        Communicates progress back to the UI via ``self._progress_queue``.
        """
        q = self._progress_queue

        try:
            q.put(_ProgressMsg("progress", 5, "Step 1/4: Loading design system..."))

            from core.generator import DocumentGenerator

            theme_path = options.get("theme_path") or None
            generator = DocumentGenerator(theme_path=theme_path)

            q.put(_ProgressMsg("progress", 20, "Step 2/4: Preparing document structure..."))

            # Build output path
            input_p = Path(input_path)
            output_dir = input_p.parent
            output_name = input_p.stem + "_transformed.docx"
            output_path = str(output_dir / output_name)

            # Build generator options
            gen_options: dict = {
                "generate_cover": options.get("generate_cover", True),
                "generate_toc": options.get("generate_toc", True),
                "number_sections": options.get("number_sections", True),
                "header_footer": options.get("header_footer", True),
            }
            if options.get("cover_title_override"):
                gen_options["cover_title_override"] = options["cover_title_override"]

            q.put(_ProgressMsg("progress", 40, "Step 3/4: Applying design system and generating content..."))

            generator.generate(tree, output_path, options=gen_options)

            q.put(_ProgressMsg("progress", 90, "Step 4/4: Finalizing document..."))

            stats = tree.summary()
            q.put(_ProgressMsg(
                "done", 100, "✓ Transformation complete!",
                data={"output_path": output_path, "stats": stats},
            ))

        except Exception as exc:
            logger.exception("Transformation failed")
            q.put(_ProgressMsg(
                "error", 0, "✗ Transformation failed",
                data=str(exc),
            ))

    # ── Progress polling ──────────────────────────────────────────────

    def _poll_progress(self) -> None:
        """Drain the progress queue and update the UI accordingly.

        Reschedules itself until a terminal message (*done* or *error*)
        is received.
        """
        keep_polling = True

        while True:
            try:
                msg: _ProgressMsg = self._progress_queue.get_nowait()
            except queue.Empty:
                break

            if msg.kind == "progress":
                self._progress_bar.set_progress(msg.value, msg.label)

            elif msg.kind == "done":
                self._progress_bar.set_progress(100, msg.label)
                data = msg.data or {}
                output_path = data.get("output_path", "")
                self._result_panel.show_success(
                    output_path=output_path,
                    stats=data.get("stats", {}),
                )
                self._transform_btn.configure(state=tk.NORMAL)
                # Save options and add to recent files
                self._save_options()
                input_path = self._file_selector.file_path
                if input_path:
                    self._settings.add_recent_file(input_path)
                    self._update_recent_menu()
                keep_polling = False

            elif msg.kind == "error":
                self._progress_bar.set_progress(0, msg.label)
                self._result_panel.show_failure(str(msg.data))
                self._transform_btn.configure(state=tk.NORMAL)
                keep_polling = False

        if keep_polling:
            self.root.after(_POLL_INTERVAL_MS, self._poll_progress)

    # ── Public entry point ────────────────────────────────────────────

    def run(self) -> None:
        """Start the Tk main loop."""
        logger.info("Starting %s", _APP_TITLE)
        self.root.mainloop()


# ── Convenience launcher ──────────────────────────────────────────────


def main() -> None:
    """Entry point for running the application directly."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    app = DocStyleApp()
    app.run()


if __name__ == "__main__":
    main()
