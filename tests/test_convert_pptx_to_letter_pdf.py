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


class PowerPointExportTests(unittest.TestCase):
    """Tests for PowerPoint process failure handling."""

    def test_reports_timeout_and_removes_temporary_script(self) -> None:
        """Raise a useful error and clean up when PowerPoint exceeds its timeout."""
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "presentation.pptx"
            input_path.touch()
            output_path = Path(temp_dir) / "presentation.pdf"
            timeout = subprocess.TimeoutExpired(
                cmd="powershell.exe",
                timeout=converter.POWERPOINT_TIMEOUT_SECONDS,
            )
            with patch.object(
                converter.subprocess, "run", side_effect=timeout
            ) as run:
                with self.assertRaisesRegex(RuntimeError, "timed out after 300 seconds"):
                    converter.export_with_powerpoint(input_path, output_path)

            self.assertEqual(
                run.call_args.kwargs["timeout"],
                converter.POWERPOINT_TIMEOUT_SECONDS,
            )
            script_path = Path(run.call_args.args[0][6])
            self.assertFalse(script_path.exists())


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
                patch.object(converter, "export_with_powerpoint") as export,
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
                    sys, "argv", ["converter", str(input_path), "--overwrite"]
                ),
                patch.object(converter, "export_with_powerpoint"),
                patch.object(converter, "fit_to_pages", side_effect=write_output),
                contextlib.redirect_stdout(io.StringIO()),
            ):
                converter.main()

            self.assertEqual(output_path.read_bytes(), b"new PDF")


if __name__ == "__main__":
    unittest.main()
