$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

Set-Location 'C:\workspace'

python -m pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to install build dependencies.'
}

python -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) {
    throw 'Unit tests failed.'
}

& (Join-Path $PSScriptRoot 'prepare-libreoffice.ps1')
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to download, verify, and extract LibreOffice.'
}

pyinstaller --noconfirm --clean --onefile `
    --name pptx-to-letter-pdf `
    --add-data 'dist\LibreOffice;LibreOffice' `
    convert_pptx_to_letter_pdf.py
if ($LASTEXITCODE -ne 0) {
    throw 'PyInstaller failed to bundle LibreOffice.'
}

& '.\dist\pptx-to-letter-pdf.exe' --help
if ($LASTEXITCODE -ne 0) {
    throw 'Bundled executable failed its --help check.'
}

$fixture = 'test-data\test.pptx'
if (Test-Path -LiteralPath $fixture) {
    $fixturePath = (Resolve-Path $fixture).Path
    & '.\dist\pptx-to-letter-pdf.exe' $fixturePath --overwrite
    if ($LASTEXITCODE -ne 0) {
        throw 'The bundled executable failed to convert the PowerPoint fixture.'
    }
    python -c "from pypdf import PdfReader; reader = PdfReader(r'test-data\test.pdf'); assert len(reader.pages) > 0, 'Conversion produced an empty PDF'; print(f'Validated {len(reader.pages)} PDF pages')"
    if ($LASTEXITCODE -ne 0) {
        throw 'The generated fixture PDF is invalid.'
    }
}

Write-Host 'Build complete: dist\pptx-to-letter-pdf.exe'
