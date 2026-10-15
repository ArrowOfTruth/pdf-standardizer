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

$version = '26.8.1'
$filename = "LibreOffice_${version}_Win_x86-64.msi"
$baseUrl = "https://download.documentfoundation.org/libreoffice/stable/$version/win/x86_64"
$msiPath = Join-Path $env:TEMP $filename
Invoke-WebRequest "$baseUrl/$filename" -OutFile $msiPath -UseBasicParsing
$expectedHash = 'c6298b10cfa1748bcbedad94c62ce771e240e5021bfd6d7a967cd571b04b5312'
$actualHash = (Get-FileHash $msiPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualHash -ne $expectedHash) {
    throw 'LibreOffice MSI SHA-256 verification failed.'
}

$stagePath = Join-Path $env:TEMP 'libreoffice-admin-image'
if (Test-Path -LiteralPath $stagePath) {
    Remove-Item -LiteralPath $stagePath -Recurse -Force
}
New-Item -ItemType Directory -Path $stagePath | Out-Null
$process = Start-Process msiexec.exe `
    -ArgumentList "/a `"$msiPath`" /qn TARGETDIR=`"$stagePath`"" `
    -Wait -PassThru
if ($process.ExitCode -ne 0) {
    throw "LibreOffice administrative extraction failed with exit code $($process.ExitCode)."
}
$soffice = Get-ChildItem $stagePath -Filter soffice.com -Recurse -File |
    Select-Object -First 1
if ($null -eq $soffice) {
    throw 'LibreOffice extraction did not contain soffice.com.'
}
$installRoot = $soffice.Directory.Parent
$destination = Join-Path (Get-Location) 'dist\LibreOffice'
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
if (Test-Path -LiteralPath $destination) {
    Remove-Item -LiteralPath $destination -Recurse -Force
}
Move-Item -LiteralPath $installRoot.FullName -Destination $destination
if (-not (Test-Path -LiteralPath (Join-Path $destination 'program\soffice.com') -PathType Leaf)) {
    throw 'LibreOffice runtime was not placed at the expected destination.'
}
& (Join-Path $destination 'program\soffice.com') --headless --version
if ($LASTEXITCODE -ne 0) {
    throw 'The bundled LibreOffice runtime failed its version check.'
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
