"""Tests for PDF fitting and command-line output handling."""

from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from pypdf import PageObject, PdfReader, PdfWriter
from pypdf.generic import ContentStream

import convert_pptx_to_letter_pdf as converter


class PowerPointNotesTests(unittest.TestCase):
    """Tests for extracting PowerPoint long-description notes."""

    def test_extracts_long_description_by_presentation_slide_order(self) -> None:
        """Resolve notes to presentation order instead of slide filenames."""
        with tempfile.TemporaryDirectory() as temp_dir:
            pptx_path = Path(temp_dir) / "presentation.pptx"
            with zipfile.ZipFile(pptx_path, "w") as package:
                package.writestr(
                    "ppt/presentation.xml",
                    """<p:presentation
                        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <p:sldIdLst>
                        <p:sldId id="1" r:id="rId1"/>
                        <p:sldId id="2" r:id="rId2"/>
                      </p:sldIdLst>
                    </p:presentation>""",
                )
                package.writestr(
                    "ppt/_rels/presentation.xml.rels",
                    """<Relationships
                        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide6.xml"/>
                    </Relationships>""",
                )
                package.writestr(
                    "ppt/notesSlides/notesSlide1.xml",
                    """<p:notes
                        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <p:cSld><p:spTree>
                        <p:sp>
                          <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                          <p:txBody><a:p><a:r><a:t>Long Description:</a:t></a:r></a:p>
                            <a:p><a:r><a:t>A side view of the brain.</a:t></a:r></a:p>
                          </p:txBody>
                        </p:sp>
                      </p:spTree></p:cSld>
                    </p:notes>""",
                )
                package.writestr(
                    "ppt/notesSlides/_rels/notesSlide1.xml.rels",
                    """<Relationships
                        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide6.xml"/>
                    </Relationships>""",
                )
                package.writestr(
                    "ppt/notesSlides/notesSlide2.xml",
                    """<p:notes
                        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <p:cSld><p:spTree>
                        <p:sp>
                          <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                          <p:txBody><a:p><a:r><a:t>Ordinary note</a:t></a:r></a:p></p:txBody>
                        </p:sp>
                      </p:spTree></p:cSld>
                    </p:notes>""",
                )
                package.writestr(
                    "ppt/notesSlides/_rels/notesSlide2.xml.rels",
                    """<Relationships
                        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/>
                    </Relationships>""",
                )

            self.assertEqual(
                converter.extract_long_description_notes(pptx_path),
                {2: "Long Description:\nA side view of the brain."},
            )


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

    def get_horizontal_rule_positions(
        self,
        page: PageObject,
        reader: PdfReader,
    ) -> list[tuple[float, float, float]]:
        """Return each horizontal rule's x start, x end, and y position."""
        operations = ContentStream(page.get_contents(), reader).operations
        rules = []
        for index, (operands, operator) in enumerate(operations):
            if operator != b"m":
                continue
            self.assertEqual(operations[index + 1][1], b"l")
            line_operands = operations[index + 1][0]
            self.assertEqual(float(operands[1]), float(line_operands[1]))
            rules.append(
                (
                    float(operands[0]),
                    float(line_operands[0]),
                    float(operands[1]),
                )
            )
        return rules

    def test_default_line_spacing_is_49_pixels(self) -> None:
        """Use the requested default spacing between writing lines."""
        self.assertEqual(converter.DEFAULT_RULE_SPACING_PIXELS, 49.0)

    def test_adds_empty_page_when_source_has_no_pages(self) -> None:
        """Produce a ruled page even for an empty source PDF."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.pdf"
            output_path = Path(temp_dir) / "output.pdf"
            with source_path.open("wb") as source_file:
                PdfWriter().write(source_file)

            converter.fit_to_pages(source_path, output_path, 4, 3, 0.5)

            reader = PdfReader(str(output_path))
            self.assertEqual(len(reader.pages), 1)
            self.assertEqual(float(reader.pages[0].mediabox.width), 288)
            self.assertEqual(float(reader.pages[0].mediabox.height), 216)
            self.assertEqual(
                self.get_horizontal_rule_positions(reader.pages[0], reader),
                [
                    (24, 264, 179.25),
                    (24, 264, 142.5),
                    (24, 264, 105.75),
                    (24, 264, 69.0),
                ],
            )

    def test_fits_pages_and_appends_empty_page(self) -> None:
        """Fit source pages and append one empty page with matching dimensions."""
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
            self.assertEqual(len(reader.pages), 3)
            for page in reader.pages:
                self.assertEqual(float(page.mediabox.width), 288)
                self.assertEqual(float(page.mediabox.height), 216)
            self.assertEqual(
                self.get_horizontal_rule_positions(
                    reader.pages[0], reader
                ),
                [
                    (222, 264, 179.25),
                    (222, 264, 142.5),
                    (24, 264, 105.75),
                    (24, 264, 69.0),
                ],
            )
            self.assertEqual(
                self.get_horizontal_rule_positions(
                    reader.pages[1], reader
                ),
                [
                    (150, 264, 179.25),
                    (24, 264, 142.5),
                    (24, 264, 105.75),
                    (24, 264, 69.0),
                ],
            )
            blank_page_rules = self.get_horizontal_rule_positions(
                reader.pages[-1], reader
            )
            self.assertEqual(
                [y for _, _, y in blank_page_rules],
                [179.25, 142.5, 105.75, 69.0],
            )
            self.assertTrue(
                all(
                    x_start >= 222 or y + 0.25 < 108
                    for x_start, _, y in self.get_horizontal_rule_positions(
                        reader.pages[0], reader
                    )
                )
            )
            self.assertTrue(
                all(
                    x_start >= 150 or y + 0.25 < 144
                    for x_start, _, y in self.get_horizontal_rule_positions(
                        reader.pages[1], reader
                    )
                )
            )

            matrices = []
            for page in reader.pages[:-1]:
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

    def test_uses_custom_line_spacing_in_pixels(self) -> None:
        """Convert CSS pixel spacing to PDF points."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.pdf"
            output_path = Path(temp_dir) / "output.pdf"
            with source_path.open("wb") as source_file:
                PdfWriter().write(source_file)

            converter.fit_to_pages(
                source_path,
                output_path,
                4,
                3,
                0.5,
                line_spacing_px=48,
            )

            reader = PdfReader(str(output_path))
            self.assertEqual(
                [y for _, _, y in self.get_horizontal_rule_positions(
                    reader.pages[0], reader
                )],
                [180, 144, 108, 72, 36],
            )

    def test_inserts_ruled_note_page_after_matching_slide(self) -> None:
        """Write long-description text on a ruled page after its slide."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "source.pdf"
            output_path = Path(temp_dir) / "output.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=288, height=144)
            writer.add_blank_page(width=288, height=144)
            with source_path.open("wb") as source_file:
                writer.write(source_file)

            converter.fit_to_pages(
                source_path,
                output_path,
                4,
                3,
                0.5,
                slide_notes={2: "Long Description: A side view of the brain."},
            )

            reader = PdfReader(str(output_path))
            self.assertEqual(len(reader.pages), 4)
            self.assertEqual(
                " ".join(reader.pages[2].extract_text().split()),
                "Long Description: A side view of the brain.",
            )
            self.assertEqual(
                len(self.get_horizontal_rule_positions(reader.pages[2], reader)),
                4,
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
            with self.assertRaisesRegex(ValueError, "at least 1 pixel"):
                converter.fit_to_pages(
                    source_path,
                    output_path,
                    4,
                    3,
                    0.5,
                    line_spacing_px=0.5,
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
                patch.object(converter, "extract_long_description_notes", return_value={}),
                contextlib.redirect_stdout(io.StringIO()),
            ):
                converter.main()

            self.assertEqual(output_path.read_bytes(), b"new PDF")


if __name__ == "__main__":
    unittest.main()
