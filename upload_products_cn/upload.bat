@echo off
chcp 65001 >nul 2>&1
setlocal
rem proxy strategy handled by upload.py auto-detect; do not clear vars here
rem --- 登录凭据：一律从环境变量读取，切勿写死到本文件 ---
if not defined SITE_USER ( echo ERROR: SITE_USER not set. & echo   run: set SITE_USER=your_account & pause ^& exit /b 5 )
if not defined SITE_PASS ( echo ERROR: SITE_PASS not set. & echo   run: set SITE_PASS=your_password & pause ^& exit /b 5 )
rem use %USERPROFILE% to adapt to current user
set "PY=%~dp0..\.venv\Scripts\python.exe"
rem absolute project path fallback (bat next to upload.py uses HERE)
set "SCRIPT_ABS=%~dp0upload.py"

rem upload.py next to this bat? else fall back to SCRIPT_ABS
set "HERE=%~dp0upload.py"
if exist "%HERE%" ( set "SCRIPT=%HERE%" ) else ( set "SCRIPT=%SCRIPT_ABS%" )

echo [upload.bat] python : %PY%
echo [upload.bat] script : %SCRIPT%
if not exist "%PY%" (
  echo ERROR: python not found at %PY%
  echo Please ensure the .workbuddy env is installed, or point PY to your python
  pause & exit /b 3
)
if not exist "%SCRIPT%" (
  echo ERROR: upload.py not found at %SCRIPT%
  pause & exit /b 3
)

rem precheck deps: openpyxl / playwright missing -> install hint
"%PY%" -c "import openpyxl, playwright" >NUL 2>&1
if errorlevel 1 (
  echo Missing deps: openpyxl / playwright.
  echo Run first: "%PY%" -m pip install playwright openpyxl
  pause & exit /b 4
)

set "TARGET=%~1"
rem 2nd/3rd args: target site (cn/es/gr) and dry run (order-free)
set "STATION=cn"
set "DRY="
if /I "%~2"=="dry" ( set "DRY=1" ) else if not "%~2"=="" ( set "STATION=%~2" )
if /I "%~3"=="dry" ( set "DRY=1" )

if "%TARGET%"=="" (
    for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Filter='Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*'; $d.Title='Select product template'; if($d.ShowDialog()-eq 'OK'){ $d.FileName }"`) do set "TARGET=%%F"
)
if "%TARGET%"=="" ( echo No file selected. & pause & exit /b 2 )

if defined DRY (
    "%PY%" "%SCRIPT%" --file "%TARGET%" --station %STATION% --dry-run
) else (
    "%PY%" "%SCRIPT%" --file "%TARGET%" --station %STATION%
)
set "RC=%ERRORLEVEL%"
echo [upload.bat] python exit code: %RC%
if not %RC%==0 ( echo Upload may have failed - check upload.log / last_run.json & pause & exit /b %RC% )
if defined CHILD_NO_PAUSE goto :skip_pause
pause
:skip_pause
