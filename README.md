# pdf-standardizer

## Download the Windows executable

The release executable bundles Python, `pypdf`, and the LibreOffice runtime. It
does not require Microsoft PowerPoint or a separate LibreOffice installation.

Download `pptx-to-letter-pdf.exe` from the
[latest GitHub Release](https://github.com/ArrowOfTruth/pdf-standardizer/releases/latest).
Also download `THIRD-PARTY-NOTICES.txt` from the same release to retain the
bundled LibreOffice license notices. Run the executable from PowerShell or
Command Prompt:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx
```

The executable extracts its bundled LibreOffice runtime to a temporary
directory when launched, so startup takes longer and requires over 1 GB of
temporary disk space. The executable is large (about 643 MB in the current
build). LibreOffice renders the slides; output appearance can differ from
Microsoft PowerPoint, especially when presentations rely on fonts or features
that LibreOffice does not support. The bundled LibreOffice distribution is
unmodified.

Light gray horizontal writing lines are spaced 49 CSS pixels apart by
default. On slide pages, lines appear in the unused area to the right of and
below the slide; they do not cross the slide's page bounds. The PDF ends with
one additional fully ruled page using the same page size.

When a slide's speaker notes start with `Long Description:`, a lined page is
inserted immediately after that slide with the note text in black 12-point
sans-serif type. If the note needs more lines than fit on one page, additional
lined pages are inserted consecutively until the note is complete. Note pages
use the same rule spacing as the rest of the PDF.

Use `--line-spacing-px` to set the spacing in CSS pixels (96 DPI; 1 pixel =
0.75 PDF points). For example:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx --overwrite --line-spacing-px 48
```

The default is 49 pixels (36.75 PDF points); values below 1 pixel are
rejected.

Existing PDF outputs are preserved by default. Pass `--overwrite` to replace
them. For example:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx --overwrite
```

Every push to `main` runs unit tests and builds the Windows x64 executable.
Normal commits upload `pptx-to-letter-pdf.exe` as a GitHub Actions artifact,
available for 14 days. GitHub downloads workflow artifacts as ZIP archives;
this artifact contains the executable directly, without a project-created
package ZIP. When a release PR is merged, the same build run attaches the raw
executable and `THIRD-PARTY-NOTICES.txt` to the new GitHub Release.

Releases use Semantic Versioning and Conventional Commits. Release Please
opens or updates a release PR with the version bump and changelog. Merge that
PR to create the version tag and GitHub Release; that run builds and uploads the
executable. The initial version is `0.1.0`.

For Release Please to open its PR, enable **Settings → Actions → General →
Workflow permissions → Allow GitHub Actions to create and approve pull
requests** in the repository.

Use commit prefixes to select the version bump:

- `fix: ...` for a patch release, such as `0.1.0` to `0.1.1`
- `feat: ...` for a minor release, such as `0.1.0` to `0.2.0`
- Add `!` after the type (for example, `feat!:`) or a `BREAKING CHANGE:`
  footer for a major release, such as `0.1.0` to `1.0.0`

GitHub Actions installs the pinned Python dependencies, downloads and verifies
the LibreOffice MSI checksum, extracts its runtime, and embeds it in the
executable. If `test-data/test.pptx` is present in the repository, the workflow
converts it with the packaged executable and checks the resulting PDF before
publishing or uploading the build.

## Build locally

Install Docker Desktop on Windows and switch it to **Windows containers**.
The Windows container engine must be running, and the host must support Windows
containers. From the repository root, run:

```powershell
.\build-locally\build.ps1
```

The script builds a Windows container with Python 3.12, then runs the same
dependency installation, unit tests, LibreOffice MSI checksum verification and
extraction, PyInstaller build, executable help check, and sample conversion
check used by GitHub Actions. Docker writes the build outputs back to the
working tree; the executable is `dist\pptx-to-letter-pdf.exe`. The first run
downloads the Python container image and LibreOffice MSI.

### Test with the sample presentation

The tracked sample presentation is `test-data\test.pptx`. Convert it with the
local executable:

```powershell
.\dist\pptx-to-letter-pdf.exe .\test-data\test.pptx --overwrite --line-spacing-px 48
```

This writes `test-data\test.pdf`. The generated PDF is ignored by Git. Check
that the PDF opens, its pages have the configured output dimensions, and it
ends with one ruled page matching the other pages' size. The GitHub Actions
workflow currently checks that this conversion produces a non-empty PDF.

To use a separately installed LibreOffice without embedding it, pass the path
to `soffice.com`:

```powershell
.\dist\pptx-to-letter-pdf.exe presentation.pptx --libreoffice "C:\Program Files\LibreOffice\program\soffice.com"
```