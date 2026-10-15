$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$imageTag = 'pptx-to-letter-pdf-local:latest'
$dockerfilePath = Join-Path $PSScriptRoot 'Dockerfile'

$engineOs = & docker info --format '{{.OSType}}'
if ($LASTEXITCODE -ne 0) {
    throw 'Docker is unavailable. Start Docker Desktop with Windows containers enabled.'
}
if ($engineOs.Trim() -ne 'windows') {
    throw 'Docker is using a Linux container engine. Switch Docker Desktop to Windows containers and retry.'
}

& docker build --tag $imageTag --file $dockerfilePath $PSScriptRoot
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to build the local Windows build image.'
}

& docker run --rm `
    --mount "type=bind,source=$repoRoot,target=C:\workspace" `
    --workdir C:\workspace `
    $imageTag
if ($LASTEXITCODE -ne 0) {
    throw 'The build inside the Docker container failed.'
}
