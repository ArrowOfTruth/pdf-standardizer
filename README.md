# pdf-standardizer

## Download the Windows executable

The executable bundles Python and the dependencies listed in `requirements.txt`.
Microsoft PowerPoint must still be installed on the Windows computer because
the converter uses PowerPoint to render presentations to PDF.

Download `pptx-to-letter-pdf.exe` from the
[latest GitHub Release](https://github.com/ArrowOfTruth/pdf-standardizer/releases/latest).
Run it from PowerShell or Command Prompt:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx
```

Existing PDF outputs are preserved by default. Pass `--overwrite` to replace
them. For example:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx --overwrite
```

To create a new release, push a version tag such as `v1.0.0`. GitHub Actions
installs the pinned direct dependencies, runs unit tests, builds the Windows
x64 executable, checks that it starts, and attaches it to a GitHub Release.
The unit tests cover PDF fitting and output protection without requiring
PowerPoint. A full conversion test still requires Windows with PowerPoint
installed.

## Build locally

On Windows with Python installed:

```powershell
python -m pip install -r requirements-build.txt
pyinstaller --noconfirm --clean --onefile --name pptx-to-letter-pdf convert_pptx_to_letter_pdf.py
```

The executable is written to `dist\pptx-to-letter-pdf.exe`.