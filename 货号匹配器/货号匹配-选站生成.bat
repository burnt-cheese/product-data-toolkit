@echo off
chcp 65001 >nul 2>&1
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

REM ============================================
REM  Goods-code matcher + auto classify + station template generator
REM  Drag a goods-code xlsx onto this bat.
REM  Step 1: match against SAP (A006)
REM  Step 2: pick station (cn/es/gr) in popup
REM  Step 3: auto classify + review / fillback (shared loop)
REM  Step 4: generate that station's upload template
REM  (No auto upload. Lightweight tool; classify loop enabled.)
REM ============================================

set "HERE=%~dp0"
set "ROOT=%~dp0.."
set "PY=%ROOT%\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

set "MATCH=%HERE%scripts\run_pipeline.py"
set "CLASSIFY_BAT=%ROOT%商品分类工作区\运行分类.bat"
set "FB=%ROOT%商品分类工作区\fillback.py"
set "OPENWAIT=%ROOT%open_wait.py"
set "REVIEW=%ROOT%review_confirm.ps1"
set "PICK=%ROOT%station_pick.ps1"
set "ADAPT=%ROOT%\upload_products_cn\site_adapt.py"
set "OUT=%HERE%输出"
set "F1=%OUT%\货号匹配结果.xlsx"
set "F2=%ROOT%商品分类工作区\output\货号匹配结果_已分类.xlsx"
set "PICKFILE=%OUT%\station_pick.txt"
set "SNAP=%ROOT%商品分类工作区\output\货号匹配结果_待审核快照.json"
set "CAND=%ROOT%商品分类工作区\output\货号匹配结果_填回候选.json"

if "%~1"=="" (
    echo [ERROR] Drag a goods-code xlsx onto this bat to start.
    pause
    exit /b 1
)
if not exist "%~1" (
    echo [ERROR] File not found: %~1
    pause
    exit /b 1
)

echo ============================================
echo  [Step 1/4] Matching goods codes against SAP
echo ============================================
"%PY%" "%MATCH%" match "%~1"
if not exist "%F1%" (
    echo [FAIL] Matching failed, no output generated. Check messages above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  [Step 2/4] Pick target station
echo ============================================
if exist "%PICKFILE%" del /f "%PICKFILE%" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%PICK%" "%PICKFILE%"
set "STATION=cn"
if exist "%PICKFILE%" ( set /p STATION=<"%PICKFILE%" )
echo  Selected station: %STATION%

echo.
echo ============================================
echo  [Step 3/4] Classify + review / fillback
echo ============================================
call "%CLASSIFY_BAT%" "%F1%"
if not exist "%F2%" (
    echo [FAIL] %F2% not generated. Aborted.
    pause
    exit /b 1
)
echo.
echo  [3.1] Snapshot before review ...
"%PY%" "%FB%" scan --classified "%F2%" --out "%SNAP%"
echo.
echo  [3.2] Excel opened. Review / edit / save, then close to continue.
echo.
"%PY%" "%OPENWAIT%" "%F2%"
echo.
echo  [3.3] Preparing fillback candidates ...
"%PY%" "%FB%" prepare --classified "%F2%" --snapshot "%SNAP%" --out "%CAND%"
echo.
echo  [3.4] Confirm window: review then click Confirm to fillback + continue.
echo        Cancel aborts (no fillback, no template).
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%REVIEW%" "%CAND%"
if errorlevel 1 (
    echo.
    echo [CANCELLED] Fillback aborted. Result kept at: %F2%
    pause
    exit /b 1
)
echo.
echo  [3.5] Fillback to master DB + cleanup ...
"%PY%" "%FB%" apply --classified "%F2%" --snapshot "%SNAP%"

echo.
echo ============================================
echo  [Step 4/4] Generating %STATION% upload template
echo ============================================
"%PY%" "%ADAPT%" --station %STATION% --input "%F2%" --out "%OUT%"
if errorlevel 1 (
    echo [FAIL] Template generation failed. Check messages above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  DONE. %STATION% upload template generated in:
echo    %OUT%
echo  (No auto upload. Upload later via upload_products_cn\upload.bat)
echo ============================================
pause
