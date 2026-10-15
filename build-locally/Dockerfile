FROM python:windowsservercore-ltsc2025

WORKDIR C:\\workspace

ENTRYPOINT ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "C:\\workspace\\build-locally\\build-in-container.ps1"]
