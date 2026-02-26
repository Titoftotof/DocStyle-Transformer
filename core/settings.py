"""Settings management for DocStyle Transformer.

Provides persistent storage for user preferences and options using INI files.
"""

from __future__ import annotations

import configparser
import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# Default settings file path
_SETTINGS_DIR = Path.home() / ".docstyle-transformer"
_SETTINGS_FILE = _SETTINGS_DIR / "settings.ini"


class SettingsManager:
    """Manages persistent user settings for the application.

    Handles loading and saving options such as:
    - Theme selection
    - Transformation flags (cover, TOC, numbering, header/footer)
    - Custom overrides (cover title, mention)
    - UI preferences (dark mode, window geometry)
    - Recent files history

    Example::

        settings = SettingsManager()
        options = settings.load_options()
        settings.save_options({"theme": "apple-minimal", "generate_cover": True})
    """

    def __init__(self, settings_path: Path | str = _SETTINGS_FILE) -> None:
        """Initialize the settings manager.

        Parameters
        ----------
        settings_path : Path or str
            Path to the settings INI file. Defaults to
            ~/.docstyle-transformer/settings.ini. Creates the file and
            directory if they don't exist.
        """
        self._settings_path = Path(settings_path)
        self._settings_path.parent.mkdir(parents=True, exist_ok=True)

        self._config = configparser.ConfigParser()

        # Load existing settings or create defaults
        if self._settings_path.exists():
            self._config.read(self._settings_path, encoding="utf-8")
        else:
            self._create_defaults()
            self._save()

    def _create_defaults(self) -> None:
        """Create default sections and keys in the config."""
        self._config.add_section("options")
        self._config.set("options", "theme", "")
        self._config.set("options", "generate_cover", "true")
        self._config.set("options", "generate_toc", "true")
        self._config.set("options", "number_sections", "true")
        self._config.set("options", "header_footer", "true")
        self._config.set("options", "cover_title_override", "")
        self._config.set("options", "mention", "")

        self._config.add_section("ui")
        self._config.set("ui", "dark_mode", "false")
        self._config.set("ui", "window_width", "800")
        self._config.set("ui", "window_height", "620")

        self._config.add_section("history")
        self._config.set("history", "recent_files", "")
        self._config.set("history", "max_recent", "10")

    def _save(self) -> None:
        """Write the current configuration to disk."""
        with open(self._settings_path, "w", encoding="utf-8") as f:
            self._config.write(f)

    # ── Options ───────────────────────────────────────────────────────────

    def load_options(self) -> dict[str, Any]:
        """Load all transformation options from settings.

        Returns
        -------
        dict
            Dictionary with keys: theme, theme_path, generate_cover,
            generate_toc, number_sections, header_footer,
            cover_title_override, mention.
        """
        return {
            "theme": self._config.get("options", "theme"),
            "theme_path": "",  # Resolved separately
            "generate_cover": self._config.getboolean("options", "generate_cover"),
            "generate_toc": self._config.getboolean("options", "generate_toc"),
            "number_sections": self._config.getboolean("options", "number_sections"),
            "header_footer": self._config.getboolean("options", "header_footer"),
            "cover_title_override": self._config.get("options", "cover_title_override"),
            "mention": self._config.get("options", "mention"),
        }

    def save_options(self, options: dict[str, Any]) -> None:
        """Save transformation options to settings.

        Parameters
        ----------
        options : dict
            Dictionary with keys matching load_options() output.
        """
        self._config.set("options", "theme", options.get("theme", ""))
        self._config.set("options", "generate_cover", str(options.get("generate_cover", True)))
        self._config.set("options", "generate_toc", str(options.get("generate_toc", True)))
        self._config.set("options", "number_sections", str(options.get("number_sections", True)))
        self._config.set("options", "header_footer", str(options.get("header_footer", True)))
        self._config.set("options", "cover_title_override", options.get("cover_title_override", ""))
        self._config.set("options", "mention", options.get("mention", ""))

        self._save()
        logger.debug("Options saved to %s", self._settings_path)

    # ── UI Preferences ────────────────────────────────────────────────────

    def get_dark_mode(self) -> bool:
        """Return the dark mode preference."""
        return self._config.getboolean("ui", "dark_mode", fallback=False)

    def set_dark_mode(self, enabled: bool) -> None:
        """Set the dark mode preference."""
        self._config.set("ui", "dark_mode", str(enabled))
        self._save()

    def get_window_size(self) -> tuple[int, int]:
        """Return the saved window size as (width, height)."""
        width = self._config.getint("ui", "window_width", fallback=800)
        height = self._config.getint("ui", "window_height", fallback=620)
        return width, height

    def set_window_size(self, width: int, height: int) -> None:
        """Save the window size."""
        self._config.set("ui", "window_width", str(width))
        self._config.set("ui", "window_height", str(height))
        self._save()

    # ── Recent Files ───────────────────────────────────────────────────────

    def get_recent_files(self) -> list[str]:
        """Return the list of recent file paths."""
        recent_str = self._config.get("history", "recent_files", fallback="")
        if not recent_str:
            return []
        return [p for p in recent_str.split("|") if p]

    def add_recent_file(self, file_path: str) -> None:
        """Add a file to the recent files list.

        The file is added to the beginning of the list. Duplicates are
        removed. The list is truncated to max_recent entries.
        """
        recent = self.get_recent_files()

        # Remove duplicate if exists
        recent = [p for p in recent if p != file_path]

        # Add to beginning
        recent.insert(0, file_path)

        # Truncate to max
        max_recent = self._config.getint("history", "max_recent", fallback=10)
        recent = recent[:max_recent]

        # Save
        self._config.set("history", "recent_files", "|".join(recent))
        self._save()
        logger.debug("Added to recent files: %s", file_path)

    def clear_recent_files(self) -> None:
        """Clear the recent files list."""
        self._config.set("history", "recent_files", "")
        self._save()
