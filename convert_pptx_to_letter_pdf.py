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
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from pypdf import PdfReader, PdfWriter, Transformation

LIBREOFFICE_TIMEOUT_SECONDS = 300


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


def fit_to_pages(
    source_path: Path,
    output_path: Path,
    page_width: float,
    page_height: float,
    max_vertical_ratio: float,
) -> None:
    """Fit each source PDF page on a custom-sized page, anchored at upper left.

    Args:
        source_path: PDF to read. Page dimensions are interpreted in points.
        output_path: Destination PDF path. An existing file is overwritten.
        page_width: Output page width in inches.
        page_height: Output page height in inches.
        max_vertical_ratio: Maximum fraction of output page height occupied by
            slide content; must be greater than 0 and at most 1.

    Raises:
        ValueError: If page dimensions or the vertical ratio are invalid, or a
            source page has invalid dimensions.
        OSError: If the source cannot be read or destination cannot be written.
    """
    if page_width <= 0 or page_height <= 0:
        raise ValueError("Page width and height must be greater than zero.")
    if not 0 < max_vertical_ratio <= 1:
        raise ValueError("Maximum vertical ratio must be greater than 0 and at most 1.")

    page_width_points = page_width * 72
    page_height_points = page_height * 72
    max_slide_height = page_height_points * max_vertical_ratio
    reader = PdfReader(str(source_path))
    writer = PdfWriter()
    for slide in reader.pages:
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

    with output_path.open("wb") as f:
        writer.write(f)


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
            with tempfile.TemporaryDirectory(prefix="pptx_pdf_") as temp_dir:
                intermediate = Path(temp_dir) / "libreoffice_export.pdf"
                export_with_libreoffice(pptx_path, intermediate, libreoffice_path)
                fit_to_pages(
                    intermediate,
                    output_path,
                    args.page_width,
                    args.page_height,
                    args.max_vertical_ratio,
                )
            print(f"Created {output_path}")
        except Exception as exc:
            failures += 1
            print(f"Failed to convert {pptx_path}: {exc}")

    if failures:
        raise SystemExit(f"{failures} of {len(presentations)} file(s) failed to convert.")


if __name__ == "__main__":
    main()
