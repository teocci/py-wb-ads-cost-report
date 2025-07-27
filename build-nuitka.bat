@echo off
setlocal enabledelayedexpansion

REM === Settings ===
set "SCRIPT=build_wb_ads_report.py"
set "NAME=wb-ads-report"
set "OUTDIR=dist"

REM === Sanity checks ===
if not exist "%SCRIPT%" (
  echo [ERROR] %SCRIPT% not found in %CD%.
  exit /b 1
)

where python >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python not found in PATH.
  echo Install Python 3.10+ and ensure 'python' is available.
  exit /b 1
)

echo [STEP] Upgrading pip/wheel...
python -m pip install -U pip wheel

echo [STEP] Installing project requirements...
if exist requirements.txt (
  python -m pip install -r requirements.txt
) else (
  echo [WARN] requirements.txt not found, skipping.
)

echo [STEP] Installing Nuitka and helpers...
python -m pip install -U nuitka orderedset zstandard

REM === Choose compiler: MSVC (cl) preferred; fallback to MinGW (gcc) if available ===
set "COMPILER_FLAGS="
where cl >nul 2>&1
if errorlevel 1 (
  where gcc >nul 2>&1
  if errorlevel 1 (
    echo [WARN] No C/C++ compiler detected. Build may fail.
    echo        Install "Visual Studio Build Tools 2022" (C++ workload) OR MinGW (and add gcc to PATH).
  ) else (
    echo [INFO] Using MinGW64 toolchain (gcc).
    set "COMPILER_FLAGS=--mingw64"
  )
) else (
  echo [INFO] Using MSVC toolchain (cl.exe).
)

echo [STEP] Building with Nuitka...
python -m nuitka "%SCRIPT%" ^
  --onefile ^
  --standalone ^
  --follow-imports ^
  --plugin-enable=numpy ^
  --include-package=pandas ^
  --include-package=openpyxl ^
  --include-package=XlsxWriter ^
  --windows-console ^
  --output-dir="%OUTDIR%" ^
  --output-filename="%NAME%.exe" ^
  --remove-output ^
  %COMPILER_FLAGS%

set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
  echo [ERROR] Build failed with exit code %EXITCODE%.
  exit /b %EXITCODE%
)

echo.
echo [SUCCESS] Built "%OUTDIR%\%NAME%.exe"
echo   Run example:
echo     %OUTDIR%\%NAME%.exe --supplier-id 3925272 --date 2025-06-21
echo.
endlocal
