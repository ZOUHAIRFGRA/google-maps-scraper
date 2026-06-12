@echo off
REM Build standalone Windows .exe from Python
REM One-time setup: pip install pyinstaller openpyxl pandas requests

echo Building Google Maps Client .exe...

REM Install dependencies if needed
pip install -q pyinstaller openpyxl pandas requests

REM Build exe (one-file, windowed)
pyinstaller ^
  --onefile ^
  --windowed ^
  --name gmaps-client ^
  --distpath=bin ^
  --workpath=build ^
  --specpath=build ^
  gmaps_client.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUCCESS] Executable: bin\gmaps-client.exe
    echo.
    echo You can now:
    echo   - Run it: .\bin\gmaps-client.exe
    echo   - Distribute it to others (standalone, no Python required)
) else (
    echo [FAILED] Check the output above for errors
)
