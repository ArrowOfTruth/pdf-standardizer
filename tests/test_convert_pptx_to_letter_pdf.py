"""Tests for PDF fitting and command-line output handling."""

from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream

import convert_pptx_to_letter_pdf as converter


class LibreOfficeExportTests(unittest.TestCase):
    """Tests for LibreOffice process handling."""

    def test_exports_pdf_with_isolated_profile(self) -> None:
        """Write the generated intermediate PDF to the requested path."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "presentation.pptx"
            input_path.touch()
            output_path = Path(temp_dir) / "presentation.pdf"

            def export_pdf(
                command: list[str], **kwargs: object
            ) -> subprocess.CompletedProcess:
                self.assertEqual(kwargs["timeout"], converter.LIBREOFFICE_TIMEOUT_SECONDS)
                export_dir = Path(command[command.index("--outdir") + 1])
                writer = PdfWriter()
                writer.add_blank_page(width=72, height=72)
                with (export_dir / "presentation.pdf").open("wb") as pdf_file:
                    writer.write(pdf_file)
                return subprocess.CompletedProcess(command, 0, "", "")

            soffice_path = Path(temp_dir) / "soffice.com"
            with patch.object(
                converter.subprocess, "run", side_effect=export_pdf
            ) as run:
                converter.export_with_libreoffice(input_path, output_path, soffice_path)

            self.assertEqual(len(PdfReader(str(output_path)).pages), 1)
            command = run.call_args.args[0]
            self.assertEqual(command[0], str(soffice_path))
            self.assertIn("--headless", command)
            self.assertTrue(
                any(arg.startswith("-env:UserInstallation=file:") for arg in command)
            )
            self.assertEqual(
                run.call_args.kwargs["timeout"],
                converter.LIBREOFFICE_TIMEOUT_SECONDS,
            )

    def test_reports_timeout(self) -> None:
        """Raise a useful error when LibreOffice exceeds its timeout."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "presentation.pptx"
            input_path.touch()
            output_path = Path(temp_dir) / "presentation.pdf"
            timeout = subprocess.TimeoutExpired(
                cmd="soffice.com",
                timeout=converter.LIBREOFFICE_TIMEOUT_SECONDS,
            )
            with patch.object(converter.subprocess, "run", side_effect=timeout):
                with self.assertRaisesRegex(RuntimeError, "timed out after 300 seconds"):
                    converter.export_with_libreoffice(
                        input_path, output_path, Path("soffice.com")
                    )

    def test_finds_explicit_libreoffice_executable(self) -> None:
        """Accept an explicitly selected LibreOffice executable."""
        with tempfile.TemporaryDirectory() as temp_dir:
            executable = Path(temp_dir) / "soffice.com"
            executable.touch()

            self.assertEqual(
                converter.find_libreoffice(executable),
                executable.resolve(),
            )

    def test_finds_libreoffice_bundled_inside_executable(self) -> None:
        """Find LibreOffice in PyInstaller's one-file extraction directory."""
        with tempfile.TemporaryDirectory() as temp_dir:
            application_dir = Path(temp_dir) / "application"
            bundle_dir = Path(temp_dir) / "bundle"
            executable = bundle_dir / "LibreOffice" / "program" / "soffice.com"
            executable.parent.mkdir(parents=True)
            executable.touch()
            with (
                patch.object(converter.sys, "frozen", True, create=True),
                patch.object(
                    converter.sys,
                    "executable",
                    str(application_dir / "pptx-to-letter-pdf.exe"),
                ),
                patch.object(converter.sys, "_MEIPASS", str(bundle_dir), create=True),
            ):
                self.assertEqual(converter.find_libreoffice(), executable)


class FitToPagesTests(unittest.TestCase):
    """Tests for PDF page sizing and slide transformations."""

    def test_fits_pages_and_preserves_page_count(self) -> None:
        """Fit content while preserving source page count and output dimensions."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.pdf"
            output_path = Path(temp_dir) / "output.pdf"
            writer = PdfWriter()
            for width, height in ((288, 144), (144, 72)):
                page = writer.add_blank_page(width=width, height=height)
                content = ContentStream(None, writer)
                content.set_data(b"0 0 10 10 re f")
                page.replace_contents(content)
            with source_path.open("wb") as source_file:
                writer.write(source_file)

            converter.fit_to_pages(source_path, output_path, 4, 3, 0.5)

            reader = PdfReader(str(output_path))
            self.assertEqual(len(reader.pages), 2)
            for page in reader.pages:
                self.assertEqual(float(page.mediabox.width), 288)
                self.assertEqual(float(page.mediabox.height), 216)

            matrices = []
            for page in reader.pages:
                content = ContentStream(page.get_contents(), reader)
                matrices.extend(
                    tuple(float(value) for value in operands)
                    for operands, operator in content.operations
                    if operator == b"cm"
                )
            self.assertEqual(
                matrices,
                [
                    (0.75, 0, 0, 0.75, 0, 108),
                    (1, 0, 0, 1, 0, 144),
                ],
            )

    def test_rejects_invalid_page_configuration(self) -> None:
        """Reject invalid dimensions and content-height limits."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.pdf"
            output_path = Path(temp_dir) / "output.pdf"
            for width, height, ratio in (
                (0, 3, 0.5),
                (4, -1, 0.5),
                (4, 3, 0),
                (4, 3, 1.1),
            ):
                with self.subTest(width=width, height=height, ratio=ratio):
                    with self.assertRaises(ValueError):
                        converter.fit_to_pages(
                            source_path, output_path, width, height, ratio
                        )


class OutputHandlingTests(unittest.TestCase):
    """Tests for protecting existing conversion results."""

    def test_does_not_overwrite_existing_pdf_without_flag(self) -> None:
        """Preserve an existing PDF and report that overwrite was not enabled."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "presentation.pptx"
            output_path = input_path.with_suffix(".pdf")
            input_path.touch()
            output_path.write_bytes(b"existing PDF")
            with (
                patch.object(sys, "argv", ["converter", str(input_path)]),
                patch.object(
                    converter,
                    "find_libreoffice",
                    return_value=Path("soffice.com"),
                ),
                patch.object(converter, "export_with_libreoffice") as export,
                contextlib.redirect_stdout(io.StringIO()),
                self.assertRaises(SystemExit),
            ):
                converter.main()

            export.assert_not_called()
            self.assertEqual(output_path.read_bytes(), b"existing PDF")

    def test_overwrite_flag_replaces_existing_pdf(self) -> None:
        """Allow an existing PDF to be replaced when explicitly requested."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "presentation.pptx"
            output_path = input_path.with_suffix(".pdf")
            input_path.touch()
            output_path.write_bytes(b"existing PDF")

            def write_output(*args: object) -> None:
                Path(args[1]).write_bytes(b"new PDF")

            with (
                patch.object(
                    sys,
                    "argv",
                    [
                        "converter",
                        str(input_path),
                        "--overwrite",
                    ],
                ),
                patch.object(converter, "export_with_libreoffice"),
                patch.object(converter, "fit_to_pages", side_effect=write_output),
                patch.object(
                    converter,
                    "find_libreoffice",
                    return_value=Path("soffice.com"),
                ),
                contextlib.redirect_stdout(io.StringIO()),
            ):
                converter.main()

            self.assertEqual(output_path.read_bytes(), b"new PDF")


if __name__ == "__main__":
    unittest.main()
