"""CLI entry point for DocStyle Transformer.

Usage::

    python main.py input.docx [-o output.docx] [--theme path/to/theme.yaml] \\
        [--no-cover] [--no-toc] [--no-numbering] [--no-header-footer] \\
        [--cover-title "Custom Title"] [--mention "Confidentiel"] [-v]
"""

import argparse
import logging
import sys
import os
from pathlib import Path

from core.parser import DocxParser
from core.detector import StructureDetector
from core.generator import DocumentGenerator

logger = logging.getLogger("docstyle-transformer")


def _build_argument_parser() -> argparse.ArgumentParser:
    """Create and return the CLI argument parser."""
    parser = argparse.ArgumentParser(
        prog="docstyle-transformer",
        description="Transform .docx documents with professional design system styling.",
    )

    parser.add_argument(
        "input",
        help="Path to the input .docx file.",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help=(
            "Path to the output .docx file. "
            "Defaults to {input_stem}_transformed.docx in the same directory."
        ),
    )
    parser.add_argument(
        "--theme",
        default=None,
        help="Path to a custom theme YAML file.",
    )

    # Feature toggles
    parser.add_argument(
        "--no-cover",
        action="store_true",
        default=False,
        help="Disable cover page generation.",
    )
    parser.add_argument(
        "--no-toc",
        action="store_true",
        default=False,
        help="Disable table of contents generation.",
    )
    parser.add_argument(
        "--no-numbering",
        action="store_true",
        default=False,
        help="Disable automatic section numbering.",
    )
    parser.add_argument(
        "--no-header-footer",
        action="store_true",
        default=False,
        help="Disable header and footer generation.",
    )

    # Overrides
    parser.add_argument(
        "--cover-title",
        default=None,
        help="Custom title for the cover page (overrides detected title).",
    )
    parser.add_argument(
        "--mention",
        default=None,
        help='Mention text for the header/footer (e.g. "Confidentiel").',
    )

    # Verbosity
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        default=False,
        help="Enable verbose (DEBUG) logging output.",
    )

    return parser


def _setup_logging(verbose: bool) -> None:
    """Configure the root logger for the application."""
    level = logging.DEBUG if verbose else logging.INFO

    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    handlers: list[logging.Handler] = [
        logging.StreamHandler(sys.stderr),
        logging.FileHandler(log_dir / "docstyle.log", encoding="utf-8"),
    ]

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=handlers,
    )


def _resolve_output_path(input_path: str, output_arg: str | None) -> str:
    """Determine the output file path.

    If *output_arg* is provided it is returned as-is.  Otherwise the output
    is placed alongside the input file with a ``_transformed`` suffix.
    """
    if output_arg:
        return output_arg

    src = Path(input_path)
    return str(src.with_name(f"{src.stem}_transformed{src.suffix}"))


def _print_summary(summary: dict) -> None:
    """Print a human-readable document summary to stdout."""
    print("\n--- Transformation Summary ---")
    print(f"  Sections   : {summary.get('sections', 0)}")
    print(f"  Paragraphs : {summary.get('paragraphs', 0)}")
    print(f"  Tables     : {summary.get('tables', 0)}")
    print(f"  Images     : {summary.get('images', 0)}")
    print(f"  Callouts   : {summary.get('callouts', 0)}")
    print(f"  Lists      : {summary.get('lists', 0)}")
    print(f"  Steps      : {summary.get('steps_blocks', 0)}")
    print("------------------------------\n")


def main() -> None:
    """Run the DocStyle Transformer pipeline."""
    parser = _build_argument_parser()
    args = parser.parse_args()

    # -- Logging -----------------------------------------------------------
    _setup_logging(args.verbose)

    # -- Validate input ----------------------------------------------------
    input_path = args.input
    if not os.path.isfile(input_path):
        logger.error("Input file not found: %s", input_path)
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if not input_path.lower().endswith(".docx"):
        logger.error("Input file is not a .docx file: %s", input_path)
        print(f"Error: Input file must be a .docx file: {input_path}", file=sys.stderr)
        sys.exit(1)

    output_path = _resolve_output_path(input_path, args.output)
    logger.info("Input : %s", input_path)
    logger.info("Output: %s", output_path)

    try:
        # -- Step 1: Parse -------------------------------------------------
        logger.info("Parsing document...")
        doc_parser = DocxParser()
        tree = doc_parser.parse(input_path)
        logger.debug("Parsed %d section(s) from input.", tree.section_count())

        # -- Step 2: Detect ------------------------------------------------
        logger.info("Detecting document structure...")
        detector = StructureDetector()
        tree = detector.detect(tree)
        logger.debug("Detection complete.")

        # -- Step 3: Generate ----------------------------------------------
        logger.info("Generating styled document...")

        generator_kwargs: dict = {}
        if args.theme:
            generator_kwargs["theme_path"] = args.theme

        generator = DocumentGenerator(**generator_kwargs)

        options: dict = {
            "generate_cover": not args.no_cover,
            "generate_toc": not args.no_toc,
            "number_sections": not args.no_numbering,
            "header_footer": not args.no_header_footer,
        }
        if args.cover_title:
            options["cover_title_override"] = args.cover_title

        generator.generate(tree, output_path, options=options)

        # -- Summary -------------------------------------------------------
        summary = tree.summary()
        _print_summary(summary)
        print(f"Document saved to: {output_path}")
        logger.info("Transformation complete: %s", output_path)

    except FileNotFoundError as exc:
        logger.error("File not found: %s", exc)
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
    except ValueError as exc:
        logger.error("Invalid input: %s", exc)
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        logger.exception("Unexpected error during transformation.")
        print(
            f"Error: An unexpected error occurred: {exc}\n"
            "Run with -v for detailed debug output.",
            file=sys.stderr,
        )
        sys.exit(2)


if __name__ == "__main__":
    main()
