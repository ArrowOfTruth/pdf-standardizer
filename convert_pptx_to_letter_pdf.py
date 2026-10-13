#!/usr/bin/env python3
"""Convert PowerPoint presentations to top-left-aligned PDF pages.

Requires Microsoft PowerPoint for Windows and the Python package ``pypdf``
(``python -m pip install -r requirements.txt``). PowerPoint renders the slides
to an intermediate PDF; Python then applies the page size and placement. Defaults
to 17-by-10-inch pages, with slides limited to 75% of page height.
"""

from __future__ import annotations

import argparse
import subprocess
import tempfile
from pathlib import Path

from pypdf import PdfReader, PdfWriter, Transformation

POWERPOINT_TIMEOUT_SECONDS = 300


def export_with_powerpoint(pptx_path: Path, pdf_path: Path) -> None:
    """Render a PowerPoint presentation to PDF using PowerPoint for Windows.

    Args:
        pptx_path: Path to the source PowerPoint presentation.
        pdf_path: Destination path for PowerPoint's intermediate PDF.

    Raises:
        RuntimeError: If PowerPoint fails, times out, or does not create the PDF.
        OSError: If the temporary PowerShell script cannot be created or removed.

    Notes:
        Microsoft PowerPoint and Windows PowerShell must be installed. Export
        is limited to 300 seconds to avoid hanging indefinitely.
    """
    powershell_script = r'''
param([string]$InputPath, [string]$OutputPath)
$ErrorActionPreference = 'Stop'
$powerpoint = $null
$presentation = $null
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerpoint.Presentations.Open($InputPath, $true, $false, $false)
    # ppSaveAsPDF = 32. Preserve PowerPoint's rendering of slide content.
    $presentation.SaveAs($OutputPath, 32)
}
finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($powerpoint -ne $null) { $powerpoint.Quit() }
    if ($presentation -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($powerpoint -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) }
}
'''
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", encoding="utf-8", delete=False) as f:
        f.write(powershell_script)
        script_path = Path(f.name)
    try:
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-STA",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(pptx_path),
                str(pdf_path),
            ],
            capture_output=True,
            text=True,
            check=False,
            timeout=POWERPOINT_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"PowerPoint PDF export timed out after "
            f"{POWERPOINT_TIMEOUT_SECONDS} seconds for {pptx_path}."
        ) from exc
    finally:
        script_path.unlink(missing_ok=True)
    if result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(f"PowerPoint PDF export failed: {detail}")
    if not pdf_path.exists():
        raise RuntimeError("PowerPoint completed without creating the intermediate PDF.")


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
                intermediate = Path(temp_dir) / "powerpoint_export.pdf"
                export_with_powerpoint(pptx_path, intermediate)
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
