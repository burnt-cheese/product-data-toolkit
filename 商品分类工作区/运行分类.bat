@echo off
chcp 65001 >nul 2>&1
setlocal
rem keep default console codepage; no chcp switch needed
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
REM ============================================
REM  Product Auto-Classifier (China site)
REM  Double-click : reads the newest .xlsx in input/
REM  Drag a file  : classifies that specific .xlsx
REM  Result is always written to the output\ folder.
REM ============================================
set "HERE=%~dp0"

REM Python: prefer this project .venv, then WorkBuddy, then system python
set "PY=%HERE%..\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo ============================================
echo   Product classifier (China site)
echo ============================================
echo.

REM Only check existence when PY is a path; PY=python relies on PATH
if not "%PY%"=="python" (
    if not exist "%PY%" (
        echo [ERROR] Python not found: %PY%
        echo Please ensure this project's .venv exists, or WorkBuddy is installed.
        echo.
        pause
        exit /b 1
    )
)

if "%~1"=="" (
    echo Mode: scan input\ folder for newest file
    echo Running classifier...
    "%PY%" "%HERE%classify.py"
) else (
    echo Mode: classify dropped file
    echo   %~1
    echo Running classifier...
    "%PY%" "%HERE%classify.py" "%~1"
)

if errorlevel 1 (
    echo.
    echo [NOTE] The script returned an error. Check the messages above.
)

echo.
echo ============================================
echo   DONE. Result saved to output\ folder
echo ============================================
echo.
if defined CHILD_NO_PAUSE goto :skip_pause
pause
:skip_pause
