# pdf-standardizer

## Download the Windows package

The release ZIP contains the converter executable with Python, `pypdf`, and the
LibreOffice runtime embedded. It does not require Microsoft PowerPoint or a
separate LibreOffice installation.

Download `pdf-standardizer-windows-x64.zip` from the
[latest GitHub Release](https://github.com/ArrowOfTruth/pdf-standardizer/releases/latest).
After extracting, run from PowerShell or Command Prompt:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx
```

The executable extracts its bundled LibreOffice runtime to a temporary
directory when launched, so startup takes longer and requires over 1 GB of
temporary disk space. The executable is large (about 643 MB in the current
build). LibreOffice renders the slides; output appearance can differ from
Microsoft PowerPoint, especially when presentations rely on fonts or features
that LibreOffice does not support. The bundled LibreOffice distribution is
unmodified and its license notices are included in the ZIP.

Existing PDF outputs are preserved by default. Pass `--overwrite` to replace
them. For example:

```powershell
.\pptx-to-letter-pdf.exe presentation.pptx --overwrite
```

To create a new release, push a version tag such as `v1.0.0`. GitHub Actions
installs the pinned Python dependencies, runs unit tests, downloads and
verifies the LibreOffice MSI checksum, extracts its runtime, and embeds it in
the Windows x64 executable. If `test-data/test.pptx` is present in the
repository, the workflow converts it with the packaged executable and checks
the resulting PDF before attaching the executable ZIP to a GitHub Release.

## Build locally

On Windows with Python and LibreOffice installed at the default path:

```powershell
python -m pip install -r requirements-build.txt
pyinstaller --noconfirm --clean --onefile --name pptx-to-letter-pdf --add-data "$env:ProgramFiles\LibreOffice;LibreOffice" convert_pptx_to_letter_pdf.py
```

The executable is written to `dist\pptx-to-letter-pdf.exe` with LibreOffice
embedded. To embed a non-standard LibreOffice installation, replace the
`$env:ProgramFiles\LibreOffice` source path with its install directory.

To use a separately installed LibreOffice without embedding it, pass the path
to `soffice.com`:

```powershell
.\dist\pptx-to-letter-pdf.exe presentation.pptx --libreoffice "C:\Program Files\LibreOffice\program\soffice.com"
```