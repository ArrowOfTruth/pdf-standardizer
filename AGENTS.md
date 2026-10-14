# Repository guidance

## Project behavior

- This is a Windows-focused Python CLI that converts `.pptx` files to PDFs.
- LibreOffice Impress renders presentations; `pypdf` then fits the pages to the
  requested output size. Do not reintroduce a Microsoft PowerPoint dependency.
- The release executable bundles Python, `pypdf`, and LibreOffice. Preserve
  LibreOffice's third-party license notices when changing its packaging.
- Conversion must not overwrite existing PDFs unless `--overwrite` is supplied.

## Python changes

- Follow existing formatting and add Google-style docstrings when documenting
  functions or classes.
- Keep dependencies pinned in the existing requirements files; update the
  release workflow and tests when changing packaged dependencies.
- Add or update focused tests for behavior changes. Run them with:
  `python -m unittest discover -s tests -v`
- Keep tests portable to the Windows GitHub Actions runner. In particular,
  compare normalized paths rather than relying on Windows path casing or
  8.3-name spelling.
- Run the sample end-to-end check with the bundled executable:
  `.\dist\pptx-to-letter-pdf.exe .\test-data\test.pptx --overwrite`
  The generated `test-data\test.pdf` is ignored; verify it is valid, every page
  has the configured dimensions, writing lines appear to the right of and below
  slides without overlapping them, and the final page is ruled. `--line-spacing-px`
  controls rule spacing in CSS pixels.

## Builds and releases

- The GitHub Actions workflow builds and tests on pushes to `main`. Preserve the
  bundled LibreOffice conversion check using `test-data/test.pptx`.
- For a local release-like build on Windows, install dependencies with
  `python -m pip install -r requirements-build.txt`, then follow the workflow's
  LibreOffice download, SHA-256 verification, and extraction steps to stage the
  runtime at `dist\LibreOffice`. Build with:
  `pyinstaller --noconfirm --clean --onefile --name pptx-to-letter-pdf --add-data "dist\LibreOffice;LibreOffice" convert_pptx_to_letter_pdf.py`
  The executable is written to `dist\pptx-to-letter-pdf.exe`. Run the fixture
  check above after building. The README also documents a simpler build using
  an installed LibreOffice runtime.
- Verify the LibreOffice download checksum before extracting or packaging it.
- Use Conventional Commit prefixes so Release Please can select the SemVer
  bump: `fix:` for patch, `feat:` for minor, and `!` or a `BREAKING CHANGE:`
  footer for major changes.
- Keep the README and `THIRD-PARTY-NOTICES.txt` aligned with changes to
  packaging, runtime requirements, and release behavior.
