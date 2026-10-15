$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$stableUrl = 'https://download.documentfoundation.org/libreoffice/stable/'
$versionMatches = [regex]::Matches(
    (Invoke-WebRequest $stableUrl -UseBasicParsing).Content,
    'href="(?<version>\d+\.\d+\.\d+)/"'
)
$versions = @(
    $versionMatches |
        ForEach-Object { [version]$_.Groups['version'].Value } |
        Sort-Object -Descending
)
if ($versions.Count -eq 0) {
    throw 'No stable LibreOffice releases were found in the official download index.'
}

$version = $versions[0].ToString()
$filename = "LibreOffice_${version}_Win_x86-64.msi"
$baseUrl = "${stableUrl}${version}/win/x86_64"
$msiPath = Join-Path $env:TEMP $filename
$checksumContent = (Invoke-WebRequest "$baseUrl/$filename.sha256" -UseBasicParsing).Content
$checksumMatch = [regex]::Match($checksumContent, '(?im)^\s*(?<hash>[0-9a-f]{64})\s+')
if (-not $checksumMatch.Success) {
    throw "Could not read the published SHA-256 checksum for LibreOffice $version."
}
$expectedHash = $checksumMatch.Groups['hash'].Value.ToLowerInvariant()

Invoke-WebRequest "$baseUrl/$filename" -OutFile $msiPath -UseBasicParsing
$actualHash = (Get-FileHash $msiPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualHash -ne $expectedHash) {
    throw "LibreOffice $version MSI SHA-256 verification failed."
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
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$destination = Join-Path $repoRoot 'dist\LibreOffice'
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
if (Test-Path -LiteralPath $destination) {
    Remove-Item -LiteralPath $destination -Recurse -Force
}
Move-Item -LiteralPath $installRoot.FullName -Destination $destination
if (-not (Test-Path -LiteralPath (Join-Path $destination 'program\soffice.com') -PathType Leaf)) {
    throw 'LibreOffice runtime was not placed at the expected destination.'
}

Write-Host "Extracted LibreOffice $version."
& (Join-Path $destination 'program\soffice.com') --headless --version
if ($LASTEXITCODE -ne 0) {
    throw 'The bundled LibreOffice runtime failed its version check.'
}
