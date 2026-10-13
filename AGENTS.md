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

## Builds and releases

- The GitHub Actions workflow builds and tests on pushes to `main`. Preserve the
  bundled LibreOffice conversion check using `test-data/test.pptx`.
- Verify the LibreOffice download checksum before extracting or packaging it.
- Use Conventional Commit prefixes so Release Please can select the SemVer
  bump: `fix:` for patch, `feat:` for minor, and `!` or a `BREAKING CHANGE:`
  footer for major changes.
- Keep the README and `THIRD-PARTY-NOTICES.txt` aligned with changes to
  packaging, runtime requirements, and release behavior.
