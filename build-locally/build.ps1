$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $repoRoot

$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if ($null -eq $pythonCommand) {
    throw 'Python 3 is required. Install Python and add it to PATH, then rerun this script.'
}

Write-Host 'Checking the installed Python version...'
& python --version
if ($LASTEXITCODE -ne 0) {
    throw 'The python command could not run. Install Python 3 and add it to PATH, then rerun this script.'
}

$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host 'Creating the local Python virtual environment...'
    & python -m venv '.venv'
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to create the local Python virtual environment.'
    }
} else {
    Write-Host 'Reusing the local Python virtual environment.'
}

Write-Host 'Installing build dependencies (this may take a few minutes)...'
& $venvPython -m pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to install build dependencies.'
}

Write-Host 'Running the unit test suite...'
& $venvPython -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) {
    throw 'Unit tests failed.'
}

Write-Host 'Preparing LibreOffice (download, checksum verification, and extraction may take several minutes)...'
& (Join-Path $PSScriptRoot 'prepare-libreoffice.ps1')
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to download, verify, and extract LibreOffice.'
}

$pyInstallerArguments = @(
    '--noconfirm',
    '--clean',
    '--onefile',
    '--name',
    'pptx-to-letter-pdf',
    '--add-data',
    'dist\LibreOffice;LibreOffice',
    '--add-data',
    '.release-please-manifest.json;.',
    'convert_pptx_to_letter_pdf.py'
)
$officialBuildMarker = Join-Path $env:TEMP "pdf-standardizer-official-build-$PID.marker"
if ($env:GITHUB_ACTIONS -eq 'true') {
    Set-Content -LiteralPath $officialBuildMarker -Value 'GitHub Actions' -NoNewline
    $pyInstallerArguments += @('--add-data', "$officialBuildMarker;.")
}

try {
    Write-Host 'Packaging the executable with PyInstaller (this may take several minutes)...'
    & $venvPython -m PyInstaller @pyInstallerArguments
    if ($LASTEXITCODE -ne 0) {
        throw 'PyInstaller failed to bundle LibreOffice.'
    }
}
finally {
    if (Test-Path -LiteralPath $officialBuildMarker) {
        Remove-Item -LiteralPath $officialBuildMarker
    }
}

Write-Host 'Checking the bundled executable help output...'
& '.\dist\pptx-to-letter-pdf.exe' --help
if ($LASTEXITCODE -ne 0) {
    throw 'Bundled executable failed its --help check.'
}

Write-Host 'Checking the bundled executable version output...'
& '.\dist\pptx-to-letter-pdf.exe' --version
if ($LASTEXITCODE -ne 0) {
    throw 'Bundled executable failed its --version check.'
}

$fixture = 'test-data\test.pptx'
if (Test-Path -LiteralPath $fixture) {
    Write-Host 'Converting the bundled sample presentation (this may take a few minutes)...'
    $fixturePath = (Resolve-Path $fixture).Path
    & '.\dist\pptx-to-letter-pdf.exe' $fixturePath --overwrite
    if ($LASTEXITCODE -ne 0) {
        throw 'The bundled executable failed to convert the PowerPoint fixture.'
    }

    Write-Host 'Validating the generated sample PDF...'
    & $venvPython -c "from pypdf import PdfReader; reader = PdfReader(r'test-data\test.note.taking.format.pdf'); assert len(reader.pages) > 0, 'Conversion produced an empty PDF'; print(f'Validated {len(reader.pages)} PDF pages')"
    if ($LASTEXITCODE -ne 0) {
        throw 'The generated fixture PDF is invalid.'
    }
} else {
    Write-Host 'Skipping the sample conversion check because test-data\test.pptx was not found.'
}

Write-Host 'Build complete: dist\pptx-to-letter-pdf.exe'
