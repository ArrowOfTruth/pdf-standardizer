# Repository guidance

## Project behavior

- This is a Windows-focused Python CLI that converts `.pptx` presentations and
  `.pdf` files to ruled note-taking PDFs.
- When same-stem `.pdf` and `.pptx` files are available, use the PDF for page
  artwork and the PPTX for long-description notes. Use PDF files directly when
  unpaired; render unpaired PPTX files with LibreOffice Impress. `pypdf` fits
  pages to the requested output size. Do not reintroduce a Microsoft PowerPoint
  dependency.
- The release executable bundles Python, `pypdf`, and LibreOffice. Preserve
  LibreOffice's third-party license notices when changing its packaging.
- Conversion must not overwrite existing PDFs unless `--overwrite` is supplied.

## Python changes

- Follow existing formatting and add Google-style docstrings when documenting
  functions or classes.
- Keep dependencies pinned in the existing requirements files; update the
  release workflow and tests when changing packaged dependencies.
- Add or update focused tests for behavior changes, but do not execute tests
  unless the user explicitly asks for them. When requested, run:
  `python -m unittest discover -s tests -v`
- Keep tests portable to the Windows GitHub Actions runner. In particular,
  compare normalized paths rather than relying on Windows path casing or
  8.3-name spelling.
- When the user explicitly requests the sample end-to-end check, run it with
  the bundled executable:
  `.\dist\pptx-to-letter-pdf.exe .\test-data\test.pptx --overwrite`
  The generated `test-data\test.note.taking.format.pdf` is ignored; verify it is
  valid, every page has the configured dimensions, writing lines appear to the
  right of and below slides without overlapping them, and the final page is
  ruled. `--line-spacing-px` controls rule spacing in CSS pixels.

## Agent command execution

- Do not run builds, tests, sample conversions, or other verification commands
  unless the user explicitly requests them.
- When a user requests a command, prefer a synchronous run and wait for it to
  finish. Do not repeatedly poll long-running command output; wait for its
  completion notification when available.
- Surface command output to the user whenever possible. Show concise output
  verbatim; for lengthy output, include the relevant result and any warnings or
  errors rather than silently omitting them.

## Builds and releases

- The GitHub Actions workflow builds and tests on pushes to `main`. Preserve the
  bundled LibreOffice conversion check using `test-data/test.pptx`.
- When the user explicitly requests a local release-like build, require Python
  3 on `PATH` and run `build-locally\build.ps1` from the repository root. It
  creates/reuses a local `.venv`, installs build dependencies, runs tests,
  verifies and extracts LibreOffice, builds the executable, and checks the
  bundled executable. The executable is written to
  `dist\pptx-to-letter-pdf.exe`.
- Verify the LibreOffice download checksum before extracting or packaging it.
- Use Conventional Commit prefixes so Release Please can select the SemVer
  bump: `fix:` for patch, `feat:` for minor, and `!` or a `BREAKING CHANGE:`
  footer for major changes.
- Keep the README and `THIRD-PARTY-NOTICES.txt` aligned with changes to
  packaging, runtime requirements, and release behavior.
