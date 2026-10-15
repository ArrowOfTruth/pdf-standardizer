$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$imageTag = 'pptx-to-letter-pdf-local:latest'
$dockerfilePath = Join-Path $PSScriptRoot 'Dockerfile'

# These checks and Docker commands run on the Windows host.
$engineOs = & docker info --format '{{.OSType}}'
if ($LASTEXITCODE -ne 0) {
    throw 'Docker is unavailable. Start Docker Desktop with Windows containers enabled.'
}
if ($engineOs.Trim() -ne 'windows') {
    throw 'Docker is using a Linux container engine. Switch Docker Desktop to Windows containers and retry.'
}

# Build the local image from the Dockerfile; the Python build steps run in its container.
& docker build --tag $imageTag --file $dockerfilePath $PSScriptRoot
if ($LASTEXITCODE -ne 0) {
    throw 'Failed to build the local Windows build image.'
}

# The repository is mounted into the container, so dist outputs persist on the host.
# LibreOffice is extracted and checked inside that container, not on the host.
& docker run --rm `
    --mount "type=bind,source=$repoRoot,target=C:\workspace" `
    --workdir C:\workspace `
    $imageTag
if ($LASTEXITCODE -ne 0) {
    throw 'The build inside the Docker container failed.'
}
