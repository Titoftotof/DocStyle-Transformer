"""Style mapper that bridges the design-system YAML and document generation.

Loads the design-system configuration and provides resolved style lookups for
every content-element type.  Color names are transparently resolved to hex
codes, and font selections fall back through the configured chain until an
available font is found.

Classes
-------
DesignSystem
    Loads and queries the ``design-system.yaml`` configuration.
StyleMapper
    Maps :class:`ContentElement` instances to concrete style dictionaries
    consumable by downstream document generators.
"""

from __future__ import annotations

import logging
import re
from copy import deepcopy
from pathlib import Path
from typing import Any, Optional

import yaml

from core.models import (
    Callout,
    CalloutType,
    ContentElement,
    Image,
    ListBlock,
    ListType,
    PageBreak,
    Paragraph,
    StepsBlock,
    Table,
)

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────

_HEX_COLOR_RE = re.compile(r"^#(?:[0-9a-fA-F]{3}){1,2}$")

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_CONFIG_PATH = _PROJECT_ROOT / "config" / "design-system.yaml"

# Keys inside the YAML component configs whose values are color references.
# Used by the recursive color-resolution logic to know which fields to resolve.
_COLOR_SUFFIXES = ("color", "bg", "background", "border_color",
                   "border_left_color", "border_top_color",
                   "border_bottom_color", "border_right_color",
                   "text_color", "header_text_color", "header_bg",
                   "zebra_color", "title_color", "body_color",
                   "number_color", "bullet_color", "accent_color",
                   "title_accent_color", "accent_bar_color",
                   "metadata_color", "subtitle_color")


# ── Font availability check ───────────────────────────────────────────


def _is_font_available(font_name: str) -> bool:
    """Best-effort check for whether *font_name* is installed on the system.

    Uses ``matplotlib.font_manager`` when available.  Falls back to checking
    common system font directories.  Returns ``True`` if the check cannot be
    performed (optimistic fallback).
    """
    # Strategy 1: matplotlib (if installed)
    try:
        from matplotlib import font_manager  # type: ignore[import-untyped]
        matches = font_manager.findSystemFonts()
        lower_name = font_name.lower()
        for fpath in matches:
            if lower_name in Path(fpath).stem.lower():
                return True
        # matplotlib is available but font was not found
        return False
    except ImportError:
        pass

    # Strategy 2: scan common OS font directories
    font_dirs = [
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path.home() / ".fonts",
        Path.home() / ".local/share/fonts",
        # macOS
        Path("/Library/Fonts"),
        Path.home() / "Library/Fonts",
        Path("/System/Library/Fonts"),
        # Windows
        Path("C:/Windows/Fonts"),
    ]
    lower_name = font_name.lower().replace(" ", "")
    for font_dir in font_dirs:
        if not font_dir.is_dir():
            continue
        try:
            for fpath in font_dir.rglob("*"):
                if fpath.is_file() and lower_name in fpath.stem.lower().replace(" ", ""):
                    return True
        except PermissionError:
            continue

    # Cannot determine -- assume available so the caller uses the preferred font.
    logger.debug(
        "Could not determine availability of font '%s'; assuming available",
        font_name,
    )
    return True


# ── DesignSystem ──────────────────────────────────────────────────────


class DesignSystem:
    """Loads and queries the design-system YAML configuration.

    Parameters
    ----------
    config_path : str or Path, optional
        Path to the YAML configuration file.  Defaults to
        ``config/design-system.yaml`` relative to the project root.
    theme_path : str or Path or None, optional
        Path to an optional theme overlay YAML.  Values in the theme file
        are deep-merged on top of the base configuration.

    Raises
    ------
    FileNotFoundError
        If the requested configuration file does not exist.
    yaml.YAMLError
        If the YAML is malformed.
    """

    def __init__(
        self,
        config_path: str | Path | None = None,
        theme_path: str | Path | None = None,
    ) -> None:
        self._config_path = Path(config_path) if config_path else _DEFAULT_CONFIG_PATH
        self._raw: dict[str, Any] = {}
        self._colors: dict[str, str] = {}
        self._typography: dict[str, Any] = {}
        self._page: dict[str, Any] = {}
        self._components: dict[str, Any] = {}
        self._header_footer: dict[str, Any] = {}
        self._cover: dict[str, Any] = {}

        self._load(self._config_path)

        if theme_path is not None:
            self._apply_theme(Path(theme_path))

        self._resolved_font_cache: dict[str, str] = {}

        logger.info("DesignSystem loaded from %s", self._config_path)

    # ── Loading / merging ──────────────────────────────────────────

    def _load(self, path: Path) -> None:
        """Load the base YAML configuration from *path*."""
        if not path.is_file():
            raise FileNotFoundError(
                f"Design-system configuration not found: {path}"
            )

        with open(path, "r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh)

        if not isinstance(data, dict):
            raise ValueError(
                f"Expected a YAML mapping at the top level in {path}"
            )

        self._raw = data
        self._colors = data.get("colors", {})
        self._typography = data.get("typography", {})
        self._page = data.get("page", {})
        self._components = data.get("components", {})
        self._header_footer = data.get("header_footer", {})
        self._cover = data.get("cover", {})

        logger.debug(
            "Loaded %d colors, %d components from %s",
            len(self._colors),
            len(self._components),
            path,
        )

    def _apply_theme(self, theme_path: Path) -> None:
        """Deep-merge a theme overlay on top of the current configuration."""
        if not theme_path.is_file():
            raise FileNotFoundError(
                f"Theme configuration not found: {theme_path}"
            )

        with open(theme_path, "r", encoding="utf-8") as fh:
            theme_data = yaml.safe_load(fh)

        if not isinstance(theme_data, dict):
            logger.warning("Theme file %s does not contain a YAML mapping; skipping", theme_path)
            return

        self._raw = self._deep_merge(self._raw, theme_data)
        self._colors = self._raw.get("colors", self._colors)
        self._typography = self._raw.get("typography", self._typography)
        self._page = self._raw.get("page", self._page)
        self._components = self._raw.get("components", self._components)
        self._header_footer = self._raw.get("header_footer", self._header_footer)
        self._cover = self._raw.get("cover", self._cover)

        logger.info("Applied theme overlay from %s", theme_path)

    @staticmethod
    def _deep_merge(base: dict, overlay: dict) -> dict:
        """Recursively merge *overlay* into a copy of *base*.

        Overlay values take precedence.  Nested dicts are merged rather than
        replaced outright.
        """
        result = deepcopy(base)
        for key, value in overlay.items():
            if (
                key in result
                and isinstance(result[key], dict)
                and isinstance(value, dict)
            ):
                result[key] = DesignSystem._deep_merge(result[key], value)
            else:
                result[key] = deepcopy(value)
        return result

    # ── Color resolution ───────────────────────────────────────────

    def resolve_color(self, name: str) -> str:
        """Resolve a color name to its hex code.

        Parameters
        ----------
        name : str
            A color name defined in the ``colors`` section of the YAML
            (e.g. ``"accent_blue"``) or an already-resolved hex string
            (e.g. ``"#0071E3"``).

        Returns
        -------
        str
            The hex color code, including the ``#`` prefix.

        Raises
        ------
        KeyError
            If *name* is not a known color name and is not a hex color.
        """
        if not name:
            return "#000000"

        # Already a hex code -- return as-is.
        if _HEX_COLOR_RE.match(name):
            return name

        if name in self._colors:
            value = self._colors[name]
            # Resolve chains (a color name pointing to another name).
            if _HEX_COLOR_RE.match(value):
                return value
            return self.resolve_color(value)

        raise KeyError(f"Unknown color name: '{name}'")

    def _resolve_colors_in_dict(self, d: dict[str, Any]) -> dict[str, Any]:
        """Return a copy of *d* with all colour-reference values resolved.

        Any key whose name ends with one of the recognised colour suffixes
        and whose value is a string present in the colour palette will be
        replaced with the corresponding hex code.
        """
        resolved = dict(d)
        for key, value in resolved.items():
            if isinstance(value, str) and not _HEX_COLOR_RE.match(value):
                # Check if this key looks like a colour field or if the value
                # is a known colour name.
                is_color_key = any(key.endswith(suffix) or key == suffix for suffix in _COLOR_SUFFIXES)
                is_known_color = value in self._colors
                if is_color_key or is_known_color:
                    try:
                        resolved[key] = self.resolve_color(value)
                    except KeyError:
                        logger.warning(
                            "Could not resolve color '%s' for key '%s'",
                            value,
                            key,
                        )
            elif isinstance(value, dict):
                resolved[key] = self._resolve_colors_in_dict(value)
        return resolved

    # ── Font resolution ────────────────────────────────────────────

    def get_font(self, role: str = "body") -> str:
        """Return the best available font name for the given *role*.

        Parameters
        ----------
        role : str
            ``"display"`` for headings / titles, or ``"body"`` for body text.
            Defaults to ``"body"``.

        Returns
        -------
        str
            The font-family name to use.
        """
        if role in self._resolved_font_cache:
            return self._resolved_font_cache[role]

        if role == "display":
            primary = self._typography.get("display_font", "")
        else:
            primary = self._typography.get("body_font", "")

        fallbacks: list[str] = self._typography.get("fallback_fonts", [])

        # Try primary first, then each fallback in order.
        candidates = [primary] + fallbacks if primary else list(fallbacks)

        for font_name in candidates:
            if _is_font_available(font_name):
                logger.debug("Font for role '%s' resolved to '%s'", role, font_name)
                self._resolved_font_cache[role] = font_name
                return font_name

        # Ultimate fallback -- use the primary regardless of availability so
        # that the document at least declares the intended font.
        result = primary or (fallbacks[0] if fallbacks else "Arial")
        logger.warning(
            "No available font found for role '%s'; falling back to '%s'",
            role,
            result,
        )
        self._resolved_font_cache[role] = result
        return result

    def get_font_family(self, role: str = "body") -> list[str]:
        """Return the full font-family chain (primary + fallbacks) for *role*.

        Useful for CSS or contexts where a fallback list is supported.
        """
        if role == "display":
            primary = self._typography.get("display_font", "")
        else:
            primary = self._typography.get("body_font", "")

        fallbacks: list[str] = list(self._typography.get("fallback_fonts", []))
        chain = [primary] + fallbacks if primary else fallbacks
        return [f for f in chain if f]

    # ── Typography styles ──────────────────────────────────────────

    def get_heading_style(self, level: int) -> dict[str, Any]:
        """Return the heading style configuration for the given *level*.

        Colors are resolved to hex codes.  The ``font`` key is added with the
        display font resolved through the fallback chain.

        Parameters
        ----------
        level : int
            Heading level (1, 2, or 3).

        Returns
        -------
        dict
            Style dictionary with keys like ``size``, ``bold``, ``color``,
            ``spacing_before``, ``spacing_after``, ``keep_with_next``, and
            ``font``.
        """
        key = f"heading{level}"
        raw_style = self._typography.get(key, {})
        if not raw_style:
            logger.warning("No typography style defined for heading level %d", level)
            raw_style = self._typography.get("heading3", {})

        style = self._resolve_colors_in_dict(dict(raw_style))
        style["font"] = self.get_font("display")
        style.setdefault("level", level)
        return style

    def get_body_style(self) -> dict[str, Any]:
        """Return the body text style with resolved colors.

        Returns
        -------
        dict
            Style dictionary with keys like ``size``, ``bold``, ``color``,
            ``line_spacing``, ``spacing_after``, and ``font``.
        """
        raw_style = self._typography.get("body", {})
        style = self._resolve_colors_in_dict(dict(raw_style))
        style["font"] = self.get_font("body")
        return style

    def get_caption_style(self) -> dict[str, Any]:
        """Return the caption text style with resolved colors."""
        raw_style = self._typography.get("caption", {})
        style = self._resolve_colors_in_dict(dict(raw_style))
        style["font"] = self.get_font("body")
        return style

    def get_small_style(self) -> dict[str, Any]:
        """Return the small text style with resolved colors."""
        raw_style = self._typography.get("small", {})
        style = self._resolve_colors_in_dict(dict(raw_style))
        style["font"] = self.get_font("body")
        return style

    # ── Component styles ───────────────────────────────────────────

    def get_component_style(self, component: str) -> dict[str, Any]:
        """Return the component style configuration with resolved colors.

        Parameters
        ----------
        component : str
            Component name as defined in the ``components`` section of the
            YAML (e.g. ``"table"``, ``"info_box"``, ``"steps"``).

        Returns
        -------
        dict
            The fully resolved style dictionary for the component.

        Raises
        ------
        KeyError
            If the component name is not found in the configuration.
        """
        raw_style = self._components.get(component)
        if raw_style is None:
            raise KeyError(f"Unknown component: '{component}'")
        return self._resolve_colors_in_dict(dict(raw_style))

    # ── Page configuration ──────────────────────────────────────────

    def get_page_config(self) -> dict[str, Any]:
        """Return page dimensions and margins.

        Returns
        -------
        dict
            Keys include ``format``, ``width``, ``height``, ``margins``
            (a nested dict), and ``usable_width``.
        """
        return deepcopy(self._page)

    @property
    def usable_width(self) -> int:
        """The usable content width in DXA (page width minus margins)."""
        explicit = self._page.get("usable_width")
        if explicit is not None:
            return int(explicit)

        width = self._page.get("width", 12240)
        margins = self._page.get("margins", {})
        left = margins.get("left", 1440)
        right = margins.get("right", 1440)
        return int(width - left - right)

    # ── Cover configuration ────────────────────────────────────────

    def get_cover_config(self) -> dict[str, Any]:
        """Return cover-page configuration with resolved colors.

        Returns
        -------
        dict
            All cover settings with color names replaced by hex codes.
        """
        config = self._resolve_colors_in_dict(dict(self._cover))
        # Ensure the title font is resolved through the fallback chain.
        title_font = config.get("title_font")
        if title_font and not _is_font_available(title_font):
            config["title_font"] = self.get_font("display")
        elif not title_font:
            config["title_font"] = self.get_font("display")
        return config

    # ── Header / footer ────────────────────────────────────────────

    def get_header_config(self) -> dict[str, Any]:
        """Return header configuration with resolved colors."""
        raw = self._header_footer.get("header", {})
        config = self._resolve_colors_in_dict(dict(raw))
        config.setdefault("font", self.get_font("body"))
        return config

    def get_footer_config(self) -> dict[str, Any]:
        """Return footer configuration with resolved colors."""
        raw = self._header_footer.get("footer", {})
        config = self._resolve_colors_in_dict(dict(raw))
        config.setdefault("font", self.get_font("body"))
        return config


# ── StyleMapper ───────────────────────────────────────────────────────


class StyleMapper:
    """Maps :class:`ContentElement` instances to style instructions.

    Uses a :class:`DesignSystem` to look up the appropriate style for any
    content element, producing a self-contained dictionary that a document
    generator can consume directly.

    Parameters
    ----------
    design_system : DesignSystem or None, optional
        The design system to use for style lookups.  If ``None``, a default
        ``DesignSystem`` is created using the standard config path.
    """

    def __init__(self, design_system: DesignSystem | None = None) -> None:
        self._ds = design_system or DesignSystem()

    @property
    def design_system(self) -> DesignSystem:
        """The underlying :class:`DesignSystem` instance."""
        return self._ds

    def map_element(self, element: ContentElement) -> dict[str, Any]:
        """Return the style dictionary for the given *element*.

        The returned dict always contains a ``"type"`` key indicating the
        element kind, plus type-specific style information.

        Parameters
        ----------
        element : ContentElement
            Any content element from the document IR.

        Returns
        -------
        dict
            A style dictionary ready for consumption by a generator.
        """
        match element:
            case Paragraph():
                return self._map_paragraph(element)
            case Table():
                return self._map_table(element)
            case Image():
                return self._map_image(element)
            case Callout():
                return self._map_callout(element)
            case ListBlock():
                return self._map_list(element)
            case StepsBlock():
                return self._map_steps(element)
            case PageBreak():
                return self._map_page_break()
            case _:
                logger.warning("No style mapping for element type: %s", type(element).__name__)
                return {"type": "unknown"}

    # ── Heading mapping (called from Section level) ────────────────

    def map_heading(self, level: int) -> dict[str, Any]:
        """Return the style dictionary for a heading of the given *level*.

        This is intended to be called when processing a :class:`Section`,
        which carries the heading text and level but is not itself a
        ``ContentElement``.
        """
        style = self._ds.get_heading_style(level)
        return {
            "type": "heading",
            "level": level,
            **style,
        }

    # ── Private mapping methods ────────────────────────────────────

    def _map_paragraph(self, paragraph: Paragraph) -> dict[str, Any]:
        """Map a paragraph to body-text style."""
        style = self._ds.get_body_style()
        return {
            "type": "paragraph",
            **style,
        }

    def _map_table(self, table: Table) -> dict[str, Any]:
        """Map a table element to table component style."""
        style = self._ds.get_component_style("table")
        style["font"] = self._ds.get_font("body")
        return {
            "type": "table",
            "has_headers": bool(table.headers),
            "row_count": len(table.rows),
            "col_count": len(table.headers) if table.headers else (
                len(table.rows[0]) if table.rows else 0
            ),
            **style,
        }

    def _map_image(self, image: Image) -> dict[str, Any]:
        """Map an image element to image placement style."""
        page_config = self._ds.get_page_config()
        usable_width = self._ds.usable_width

        return {
            "type": "image",
            "max_width": usable_width,
            "page_width": page_config.get("width", 12240),
            "original_width": image.width,
            "original_height": image.height,
            "caption_style": self._ds.get_caption_style(),
        }

    def _map_callout(self, callout: Callout) -> dict[str, Any]:
        """Map a callout element to the matching callout-box style."""
        component_name = self._callout_component_name(callout.callout_type)
        try:
            style = self._ds.get_component_style(component_name)
        except KeyError:
            # Fall back to info_box for unrecognised callout types.
            logger.debug(
                "No component style for '%s'; falling back to 'info_box'",
                component_name,
            )
            style = self._ds.get_component_style("info_box")

        style["font"] = self._ds.get_font("body")
        return {
            "type": "callout",
            "callout_type": callout.callout_type.value,
            **style,
        }

    def _map_list(self, list_block: ListBlock) -> dict[str, Any]:
        """Map a list block to the matching list component style."""
        if list_block.list_type == ListType.NUMBERED:
            component_name = "numbered_list"
        else:
            component_name = "bullet_list"

        style = self._ds.get_component_style(component_name)
        style["font"] = self._ds.get_font("body")
        return {
            "type": "list",
            "list_type": list_block.list_type.value,
            "item_count": len(list_block.items),
            **style,
        }

    def _map_steps(self, steps_block: StepsBlock) -> dict[str, Any]:
        """Map a steps block to the steps component style."""
        style = self._ds.get_component_style("steps")
        style["display_font"] = self._ds.get_font("display")
        style["body_font"] = self._ds.get_font("body")
        return {
            "type": "steps",
            "step_count": len(steps_block.steps),
            **style,
        }

    @staticmethod
    def _map_page_break() -> dict[str, Any]:
        """Map a page-break element."""
        return {"type": "page_break"}

    # ── Helpers ────────────────────────────────────────────────────

    @staticmethod
    def _callout_component_name(callout_type: CalloutType) -> str:
        """Map a :class:`CalloutType` to the YAML component key."""
        mapping = {
            CalloutType.INFO: "info_box",
            CalloutType.NOTE: "info_box",
            CalloutType.TIP: "info_box",
            CalloutType.WARNING: "warning_box",
        }
        return mapping.get(callout_type, "info_box")
