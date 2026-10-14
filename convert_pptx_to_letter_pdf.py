#!/usr/bin/env python3
"""Convert PowerPoint presentations to top-left-aligned PDF pages.

Requires LibreOffice Impress and the Python package ``pypdf``
(``python -m pip install -r requirements.txt``). The Windows release ZIP
includes LibreOffice. Impress renders slides to an intermediate PDF; Python
then applies the page size and placement. Defaults to 17-by-10-inch pages, with
slides limited to 75% of page height.
"""

from __future__ import annotations

import argparse
import os
import posixpath
import shutil
import subprocess
import sys
import tempfile
import textwrap
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from pypdf import PageObject, PdfReader, PdfWriter, Transformation
from pypdf.generic import (
    ContentStream,
    DictionaryObject,
    IndirectObject,
    NameObject,
)

LIBREOFFICE_TIMEOUT_SECONDS = 300
PDF_POINTS_PER_PIXEL = 72 / 96
DEFAULT_RULE_SPACING_PIXELS = 49.0
NOTE_FONT_SIZE_POINTS = 12
NOTE_CHARACTER_WIDTH_FACTOR = 0.5
RULE_LINE_WIDTH_POINTS = 0.5
RULE_HORIZONTAL_MARGIN_POINTS = 24
RULE_SLIDE_GUTTER_POINTS = 6
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def find_libreoffice(explicit_path: Path | None = None) -> Path:
    """Find the LibreOffice command-line executable.

    Args:
        explicit_path: Optional path to ``soffice.com`` or ``soffice.exe``.

    Raises:
        FileNotFoundError: If LibreOffice cannot be found.

    Notes:
        The application searches beside the script or executable first, then
        the system PATH and standard Windows installation folders.
    """
    if explicit_path is not None:
        executable = explicit_path.expanduser().resolve()
        if executable.is_file():
            return executable
        raise FileNotFoundError(f"LibreOffice executable not found: {executable}")

    if getattr(sys, "frozen", False):
        application_dir = Path(sys.executable).resolve().parent
        bundle_dir = Path(getattr(sys, "_MEIPASS", application_dir))
        candidates = [
            bundle_dir / "LibreOffice" / "program" / "soffice.com",
            bundle_dir / "LibreOffice" / "program" / "soffice.exe",
            application_dir / "LibreOffice" / "program" / "soffice.com",
            application_dir / "LibreOffice" / "program" / "soffice.exe",
        ]
    else:
        application_dir = Path(__file__).resolve().parent
        candidates = [
            application_dir / "LibreOffice" / "program" / "soffice.com",
            application_dir / "LibreOffice" / "program" / "soffice.exe",
        ]
    for name in ("soffice.com", "soffice.exe"):
        path_from_path = shutil.which(name)
        if path_from_path:
            candidates.append(Path(path_from_path))

    for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        program_files = os.environ.get(variable)
        if program_files:
            candidates.extend(
                Path(program_files) / "LibreOffice" / "program" / name
                for name in ("soffice.com", "soffice.exe")
            )

    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(
        "LibreOffice was not found. Build or extract the complete release "
        "package, install LibreOffice, or pass its executable with "
        "--libreoffice."
    )


def export_with_libreoffice(
    pptx_path: Path,
    pdf_path: Path,
    libreoffice_path: Path,
) -> None:
    """Render a PowerPoint presentation to PDF using LibreOffice Impress.

    Args:
        pptx_path: Path to the source PowerPoint presentation.
        pdf_path: Destination path for the intermediate PDF.
        libreoffice_path: Path to ``soffice.com`` or ``soffice.exe``.

    Raises:
        RuntimeError: If LibreOffice fails, times out, or creates no PDF.
        OSError: If temporary directories cannot be created or the PDF cannot
            be copied to its destination.

    Notes:
        A temporary LibreOffice profile isolates each conversion from running
        LibreOffice instances. Export is limited to 300 seconds.
    """
    with tempfile.TemporaryDirectory(prefix="pptx_pdf_") as temp_dir:
        temp_path = Path(temp_dir)
        profile_path = temp_path / "profile"
        exported_pdf = temp_path / f"{pptx_path.stem}.pdf"
        command = [
            str(libreoffice_path),
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--norestore",
            f"-env:UserInstallation={profile_path.as_uri()}",
            "--convert-to",
            "pdf:impress_pdf_Export",
            "--outdir",
            str(temp_path),
            str(pptx_path),
        ]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                timeout=LIBREOFFICE_TIMEOUT_SECONDS,
                cwd=libreoffice_path.parent,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                f"LibreOffice PDF export timed out after "
                f"{LIBREOFFICE_TIMEOUT_SECONDS} seconds for {pptx_path}."
            ) from exc

        if result.returncode != 0:
            detail = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(f"LibreOffice PDF export failed: {detail}")
        if not exported_pdf.is_file():
            detail = result.stderr.strip() or result.stdout.strip()
            message = f"LibreOffice did not create a PDF for {pptx_path}."
            if detail:
                message = f"{message} {detail}"
            raise RuntimeError(message)
        shutil.copy2(exported_pdf, pdf_path)


def extract_long_description_notes(pptx_path: Path) -> dict[int, str]:
    """Return long-description notes keyed by presentation slide number.

    Args:
        pptx_path: PowerPoint package to inspect.

    Returns:
        Long-description notes keyed by one-based slide order.

    Notes:
        Only text in the notes body placeholder is considered. The associated
        slide is resolved through the package relationships, rather than by
        assuming the notes and slide filenames have matching numbers.
    """
    presentation_part = "ppt/presentation.xml"
    presentation_relationships_part = "ppt/_rels/presentation.xml.rels"
    notes_directory = "ppt/notesSlides"
    relationship_tag = f"{{{PACKAGE_REL_NS}}}Relationship"
    notes_body_tag = f"{{{PRESENTATION_NS}}}sp"
    slide_list_tag = f"{{{PRESENTATION_NS}}}sldIdLst"
    slide_tag = f"{{{PRESENTATION_NS}}}sldId"
    drawing_paragraph_tag = f"{{{DRAWING_NS}}}p"
    drawing_text_tag = f"{{{DRAWING_NS}}}t"

    with zipfile.ZipFile(pptx_path) as package:
        presentation_root = ET.fromstring(package.read(presentation_part))
        presentation_relationships_root = ET.fromstring(
            package.read(presentation_relationships_part)
        )
        presentation_relationships = {
            relationship.get("Id"): relationship.get("Target")
            for relationship in presentation_relationships_root.findall(relationship_tag)
        }
        slide_numbers: dict[str, int] = {}
        slide_list = presentation_root.find(slide_list_tag)
        if slide_list is None:
            return {}
        for slide_number, slide_element in enumerate(slide_list.findall(slide_tag), start=1):
            relationship_id = slide_element.get(f"{{{OFFICE_REL_NS}}}id")
            target = presentation_relationships.get(relationship_id)
            if target is not None:
                slide_part = _resolve_package_target(presentation_part, target)
                slide_numbers[slide_part] = slide_number

        notes: dict[int, str] = {}
        for notes_part in package.namelist():
            notes_filename = posixpath.basename(notes_part)
            if (
                posixpath.dirname(notes_part) != notes_directory
                or not notes_filename.startswith("notesSlide")
                or not notes_filename.endswith(".xml")
            ):
                continue

            notes_relationships_part = posixpath.join(
                notes_directory,
                "_rels",
                f"{notes_filename}.rels",
            )
            if notes_relationships_part not in package.namelist():
                continue
            notes_relationships_root = ET.fromstring(
                package.read(notes_relationships_part)
            )
            for relationship in notes_relationships_root.findall(relationship_tag):
                relationship_type = relationship.get("Type", "")
                target = relationship.get("Target")
                if not relationship_type.endswith("/slide") or target is None:
                    continue
                slide_number = slide_numbers.get(
                    _resolve_package_target(notes_part, target)
                )
                if slide_number is None:
                    continue

                notes_root = ET.fromstring(package.read(notes_part))
                paragraphs: list[str] = []
                for shape in notes_root.findall(f".//{notes_body_tag}"):
                    placeholder = shape.find(
                        f"./{{{PRESENTATION_NS}}}nvSpPr/"
                        f"{{{PRESENTATION_NS}}}nvPr/"
                        f"{{{PRESENTATION_NS}}}ph"
                    )
                    if placeholder is None or placeholder.get("type") != "body":
                        continue
                    paragraphs.extend(
                        "".join(text.text or "" for text in paragraph.findall(f".//{drawing_text_tag}"))
                        for paragraph in shape.findall(f".//{drawing_paragraph_tag}")
                    )
                note_text = "\n".join(paragraphs).strip()
                if note_text.casefold().startswith("long description:"):
                    notes[slide_number] = note_text
                break

    return notes


def _resolve_package_target(source_part: str, target: str) -> str:
    """Resolve a package relationship target to a normalized package path."""
    target_path = target.lstrip("/")
    return posixpath.normpath(
        target_path
        if target.startswith("/")
        else posixpath.join(posixpath.dirname(source_part), target_path)
    )


def fit_to_pages(
    source_path: Path,
    output_path: Path,
    page_width: float,
    page_height: float,
    max_vertical_ratio: float,
    line_spacing_px: float = DEFAULT_RULE_SPACING_PIXELS,
    slide_notes: dict[int, str] | None = None,
) -> None:
    """Fit source pages, add ruled writing space, and append a ruled page.

    Args:
        source_path: PDF to read. Page dimensions are interpreted in points.
        output_path: Destination PDF path. An existing file is overwritten. The
            output always ends with one additional ruled page.
        page_width: Output page width in inches.
        page_height: Output page height in inches.
        max_vertical_ratio: Maximum fraction of output page height occupied by
            slide content; must be greater than 0 and at most 1.
        line_spacing_px: Distance between writing lines in CSS pixels at 96 DPI.
        slide_notes: Optional long-description notes keyed by one-based slide
            number. A lined page is inserted after each slide with a note.

    Raises:
        ValueError: If page dimensions or the vertical ratio are invalid, or a
            source page has invalid dimensions or line spacing is less than
            one pixel. A note that exceeds one lined page or contains
            unsupported characters also raises ``ValueError``.
        OSError: If the source cannot be read or destination cannot be written.
    """
    if page_width <= 0 or page_height <= 0:
        raise ValueError("Page width and height must be greater than zero.")
    if not 0 < max_vertical_ratio <= 1:
        raise ValueError("Maximum vertical ratio must be greater than 0 and at most 1.")
    if line_spacing_px < 1:
        raise ValueError("Line spacing must be at least 1 pixel.")

    line_spacing_points = line_spacing_px * PDF_POINTS_PER_PIXEL
    page_width_points = page_width * 72
    page_height_points = page_height * 72
    max_slide_height = page_height_points * max_vertical_ratio
    reader = PdfReader(str(source_path))
    writer = PdfWriter()
    for slide_number, slide in enumerate(reader.pages, start=1):
        # Use the visible page bounds. pypdf reports coordinates in points.
        box = slide.cropbox
        width = float(box.width)
        height = float(box.height)
        if width <= 0 or height <= 0:
            raise ValueError("A slide has invalid PDF dimensions.")

        scale = min(1.0, page_width_points / width, max_slide_height / height)
        scaled_height = height * scale

        # Account for a non-zero crop-box origin when moving the content.
        tx = -float(box.left) * scale
        ty = page_height_points - scaled_height - float(box.bottom) * scale
        page = writer.add_blank_page(width=page_width_points, height=page_height_points)
        page.merge_transformed_page(
            slide,
            Transformation().scale(scale).translate(tx, ty),
            over=True,
            expand=False,
        )
        _add_rule_lines(
            page,
            writer,
            page_width_points,
            page_height_points,
            (
                -float(box.left) * scale,
                page_height_points - scaled_height,
                -float(box.left) * scale + width * scale,
                page_height_points,
            ),
            line_spacing_points,
        )
        note_text = (slide_notes or {}).get(slide_number)
        if note_text is not None:
            _add_note_page(
                writer,
                note_text,
                page_width_points,
                page_height_points,
                line_spacing_points,
            )

    blank_page = writer.add_blank_page(
        width=page_width_points,
        height=page_height_points,
    )
    _add_rule_lines(
        blank_page,
        writer,
        page_width_points,
        page_height_points,
        None,
        line_spacing_points,
    )

    with output_path.open("wb") as f:
        writer.write(f)


def _add_note_page(
    writer: PdfWriter,
    note_text: str,
    page_width_points: float,
    page_height_points: float,
    line_spacing_points: float,
) -> None:
    """Append a ruled page containing a note in black 12-point sans-serif."""
    page = writer.add_blank_page(
        width=page_width_points,
        height=page_height_points,
    )
    _add_rule_lines(
        page,
        writer,
        page_width_points,
        page_height_points,
        None,
        line_spacing_points,
    )

    available_width = (
        page_width_points - 2 * RULE_HORIZONTAL_MARGIN_POINTS
    )
    wrap_width = max(
        1,
        int(available_width / (NOTE_FONT_SIZE_POINTS * NOTE_CHARACTER_WIDTH_FACTOR)),
    )
    lines = [
        line
        for paragraph in note_text.splitlines()
        for line in (textwrap.wrap(paragraph, width=wrap_width) or [""])
    ]
    rule_positions: list[float] = []
    y = page_height_points - line_spacing_points
    while y >= line_spacing_points:
        rule_positions.append(y)
        y -= line_spacing_points
    if len(lines) > len(rule_positions):
        raise ValueError(
            f"Long-description note needs {len(lines)} writing lines, but the "
            f"page has room for {len(rule_positions)}."
        )

    resources = page.get("/Resources")
    if isinstance(resources, IndirectObject):
        resources = resources.get_object()
    if resources is None:
        resources = DictionaryObject()
        page[NameObject("/Resources")] = resources
    font_resources = resources.get("/Font")
    if isinstance(font_resources, IndirectObject):
        font_resources = font_resources.get_object()
    if font_resources is None:
        font_resources = DictionaryObject()
        resources[NameObject("/Font")] = font_resources
    font_resources[NameObject("/Fnote")] = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
            NameObject("/Encoding"): NameObject("/WinAnsiEncoding"),
        }
    )

    commands = []
    for line, rule_y in zip(lines, rule_positions):
        try:
            encoded_line = line.encode("cp1252")
        except UnicodeEncodeError as exc:
            raise ValueError(
                "Long-description note contains characters not supported by "
                "the standard PDF font."
            ) from exc
        escaped_line = (
            encoded_line.replace(b"\\", b"\\\\")
            .replace(b"(", b"\\(")
            .replace(b")", b"\\)")
        )
        baseline_y = rule_y + 2
        commands.append(
            b"BT /Fnote "
            + str(NOTE_FONT_SIZE_POINTS).encode("ascii")
            + b" Tf 1 0 0 1 "
            + str(RULE_HORIZONTAL_MARGIN_POINTS).encode("ascii")
            + b" "
            + str(baseline_y).encode("ascii")
            + b" Tm ("
            + escaped_line
            + b") Tj ET\n"
        )

    existing = page.get_contents()
    content = existing.get_data() if existing is not None else b""
    stream = ContentStream(None, writer)
    stream.set_data(content + b"".join(commands))
    page.replace_contents(stream)


def _add_rule_lines(
    page: PageObject,
    writer: PdfWriter,
    page_width_points: float,
    page_height_points: float,
    slide_bounds: tuple[float, float, float, float] | None,
    line_spacing_points: float,
) -> None:
    """Draw rules outside the slide bounds, or across a blank page."""
    if line_spacing_points <= 0:
        raise ValueError("Line spacing must be greater than zero.")

    half_line_width = RULE_LINE_WIDTH_POINTS / 2
    segments: list[tuple[float, float, float]] = []
    y = page_height_points - line_spacing_points
    while y >= line_spacing_points:
        if slide_bounds is None:
            segments.append(
                (RULE_HORIZONTAL_MARGIN_POINTS, page_width_points - RULE_HORIZONTAL_MARGIN_POINTS, y)
            )
        else:
            _, slide_bottom, slide_right, slide_top = slide_bounds
            overlaps_slide_height = (
                y + half_line_width > slide_bottom
                and y - half_line_width < slide_top
            )
            if overlaps_slide_height:
                line_start = slide_right + RULE_SLIDE_GUTTER_POINTS
                line_end = page_width_points - RULE_HORIZONTAL_MARGIN_POINTS
                if line_start < line_end:
                    segments.append((line_start, line_end, y))
            elif y + half_line_width < slide_bottom:
                segments.append(
                    (
                        RULE_HORIZONTAL_MARGIN_POINTS,
                        page_width_points - RULE_HORIZONTAL_MARGIN_POINTS,
                        y,
                    )
                )
        y -= line_spacing_points

    if not segments:
        return

    existing = page.get_contents()
    content = existing.get_data() if existing is not None else b""
    commands = [
        "\nq\n0.72 0.78 0.84 RG\n",
        f"{RULE_LINE_WIDTH_POINTS} w\n",
    ]
    for line_start, line_end, line_y in segments:
        commands.append(f"{line_start} {line_y} m\n{line_end} {line_y} l\n")
    commands.append("S\nQ\n")

    stream = ContentStream(None, writer)
    stream.set_data(content + "".join(commands).encode("ascii"))
    page.replace_contents(stream)


def main() -> None:
    """Parse command-line arguments and convert one presentation or a folder.

    Raises:
        SystemExit: If arguments are invalid or one or more conversions fail.
    """
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Required .pptx file to convert, or folder containing .pptx files",
    )
    parser.add_argument(
        "--page-width",
        type=float,
        default=17.0,
        metavar="INCHES",
        help="PDF page width in inches",
    )
    parser.add_argument(
        "--page-height",
        type=float,
        default=10.0,
        metavar="INCHES",
        help="PDF page height in inches",
    )
    parser.add_argument(
        "--max-vertical-ratio",
        type=float,
        default=0.75,
        metavar="FRACTION",
        help="Maximum fraction of page height that a slide may occupy (0 to 1)",
    )
    parser.add_argument(
        "--line-spacing-px",
        type=float,
        default=DEFAULT_RULE_SPACING_PIXELS,
        metavar="PIXELS",
        help=(
            "Spacing between horizontal writing lines in CSS pixels "
            "at 96 DPI (minimum 1)"
        ),
    )
    parser.add_argument(
        "--recurse",
        action="store_true",
        help="When input is a folder, search its subfolders for .pptx files",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing PDF output files",
    )
    parser.add_argument(
        "--libreoffice",
        type=Path,
        metavar="PATH",
        help="Path to LibreOffice soffice.com or soffice.exe",
    )
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    if args.page_width <= 0 or args.page_height <= 0:
        parser.error("--page-width and --page-height must be greater than zero.")
    if not 0 < args.max_vertical_ratio <= 1:
        parser.error("--max-vertical-ratio must be greater than 0 and at most 1.")
    if args.line_spacing_px < 1:
        parser.error("--line-spacing-px must be at least 1.")
    if not input_path.exists():
        parser.error(f"Input path does not exist: {input_path}")
    if input_path.is_file():
        if input_path.suffix.lower() != ".pptx":
            parser.error(f"Input file must have a .pptx extension: {input_path}")
        presentations = [input_path]
    elif input_path.is_dir():
        iterator = input_path.rglob("*") if args.recurse else input_path.glob("*")
        presentations = sorted(
            path for path in iterator if path.is_file() and path.suffix.lower() == ".pptx"
        )
        if not presentations:
            parser.error(f"No .pptx files found in {input_path}")
    else:
        parser.error(f"Input path is neither a file nor a folder: {input_path}")

    try:
        libreoffice_path = find_libreoffice(args.libreoffice)
    except FileNotFoundError as exc:
        parser.error(str(exc))

    failures = 0
    for pptx_path in presentations:
        output_path = pptx_path.with_suffix(".pdf")
        try:
            if output_path.exists() and not args.overwrite:
                raise FileExistsError(
                    f"Output already exists: {output_path} "
                    "(pass --overwrite to replace it)"
                )
            slide_notes = extract_long_description_notes(pptx_path)
            with tempfile.TemporaryDirectory(prefix="pptx_pdf_") as temp_dir:
                intermediate = Path(temp_dir) / "libreoffice_export.pdf"
                export_with_libreoffice(pptx_path, intermediate, libreoffice_path)
                fit_to_pages(
                    intermediate,
                    output_path,
                    args.page_width,
                    args.page_height,
                    args.max_vertical_ratio,
                    args.line_spacing_px,
                    slide_notes,
                )
            print(f"Created {output_path}")
        except Exception as exc:
            failures += 1
            print(f"Failed to convert {pptx_path}: {exc}")

    if failures:
        raise SystemExit(f"{failures} of {len(presentations)} file(s) failed to convert.")


if __name__ == "__main__":
    main()
