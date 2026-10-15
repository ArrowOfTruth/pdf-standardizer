FROM python:3.12-windowsservercore-ltsc2022

WORKDIR C:\workspace

ENTRYPOINT ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "C:\\workspace\\build-locally\\build-in-container.ps1"]
