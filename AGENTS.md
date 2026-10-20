# Repository guidance

## Project overview

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

## Project layout

- `convert_pptx_to_letter_pdf.py` contains the CLI and conversion pipeline.
- `tests/` contains the portable Python unit tests.
- `test-data/` contains conversion fixtures.
- `build-locally/` contains the Windows release build and LibreOffice setup
  scripts.
- `.github/workflows/` contains CI and release automation.

## Development tips

- Follow existing formatting and add Google-style docstrings when documenting
  functions or classes.
- Keep dependencies pinned in the existing requirements files; update the
  release workflow and tests when changing packaged dependencies.

## Developer commands

| Purpose | Command | Guidance |
| --- | --- | --- |
| Focused converter tests | `python -m unittest discover -s tests -p test_convert_pptx_to_letter_pdf.py -v` | Prefer for a targeted change; run only when explicitly requested. |
| Full unit suite | `python -m unittest discover -s tests -v` | Run only when explicitly requested. |
| Local release build | `.\build-locally\build.ps1` | Downloads and verifies LibreOffice, builds the executable, and runs checks; run only when explicitly requested. |
| Bundled sample conversion | `.\dist\pptx-to-letter-pdf.exe .\test-data\test.pptx --overwrite` | Uses the bundled executable; run only when explicitly requested. |

## Testing instructions

- Add a focused test for each behavior change when practical. Follow nearby test
  style, check whether existing tests already cover the behavior, and avoid
  duplicate or tautological cases.
- Prefer running the focused test command above over the full suite. Do not run
  tests or other verification commands unless the user explicitly requests
  them.
- Keep tests portable to the Windows GitHub Actions runner. In particular,
  compare normalized paths rather than relying on Windows path casing or
  8.3-name spelling.
- When the user explicitly requests the sample end-to-end check, verify the
  generated `test-data\test.note.taking.format.pdf` is valid, every page has
  the configured dimensions, writing lines appear to the right of and below
  slides without overlapping them, and the final page is ruled.
  `--line-spacing-px` controls rule spacing in CSS pixels.

## Agent workflow and PR guidance

- For improvement requests, first present a concise implementation plan and
  wait for the user's review and explicit approval before editing files or
  running commands.
- After making changes, suggest a concise Conventional Commit message that
  summarizes the complete set of changes. Do not create a commit unless the
  user explicitly asks.

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
- The local release build requires Python 3 on `PATH`; it creates/reuses `.venv`,
  installs build dependencies, verifies and extracts LibreOffice, builds the
  executable, and runs the bundled checks. The executable is written to
  `dist\pptx-to-letter-pdf.exe`. Run this build only on explicit user request.
- Verify the LibreOffice download checksum before extracting or packaging it.
- Use Conventional Commit prefixes so Release Please can select the SemVer
  bump: `fix:` for patch, `feat:` for minor, and `!` or a `BREAKING CHANGE:`
  footer for major changes.
- Keep the README and `THIRD-PARTY-NOTICES.txt` aligned with changes to
  packaging, runtime requirements, and release behavior.
