@echo off
setlocal enabledelayedexpansion
REM Build the Google Maps GUI client executable
REM This script uses Docker to ensure cross-platform compatibility

if not exist bin mkdir bin

REM Try Docker first (recommended)
where docker >nul 2>nul
if !ERRORLEVEL! EQU 0 (
    echo [*] Building with Docker...
    docker build -f Dockerfile.client -t gmaps-client-builder:latest . >nul 2>&1
    if !ERRORLEVEL! EQU 0 (
        docker run --rm gmaps-client-builder:latest > bin\gmaps-client.exe 2>nul
        if !ERRORLEVEL! EQU 0 (
            echo [+] Success: bin\gmaps-client.exe
            exit /b 0
        )
    )
    echo [-] Docker build failed, trying native build...
)

REM Fallback: Try native build with MSYS64 if Docker unavailable
echo [*] Attempting native build with MSYS64...
set GOOS=windows
set GOARCH=amd64
set CGO_ENABLED=1
if exist "C:\msys64\mingw64\bin" (
    set PATH=C:\msys64\mingw64\bin;!PATH!
)
go build -v -o bin\gmaps-client.exe .\cmd\gmapsclient\ > build.log 2>&1

if !ERRORLEVEL! EQU 0 (
    echo [+] Success: bin\gmaps-client.exe
    exit /b 0
) else (
    echo [-] Build failed. See build.log for details.
    echo.
    echo To fix this, install one of:
    echo   1. Docker Desktop: https://www.docker.com/products/docker-desktop
    echo   2. MSYS2 64-bit: https://www.msys2.org/
    exit /b 1
)
