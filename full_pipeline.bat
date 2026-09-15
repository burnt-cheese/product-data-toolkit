@echo off
chcp 65001 >nul 2>&1
set "CHILD_NO_PAUSE=1"
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

REM ============================================
REM  后台 pipeline: match -> classify -> adapt (upload only for cn)
REM  Drag an xlsx onto this bat to start.
REM  Pick station at step 0:
REM    cn = classify + review/fillback + upload
REM    es/gr = classify + review/fillback + generate station table (NO upload)
REM  All stations now share the same classify/review/fillback loop.
REM ============================================

set "ROOT=%~dp0"
set "MATCH_BAT=%ROOT%.venv\Scripts\python.exe"
set "MATCH_SCRIPT=%ROOT%sku_matcher\scripts\run_pipeline.py"
set "CLASSIFY_BAT=%ROOT%product_classifier\run_classify.bat"
set "UPLOAD_BAT=%ROOT%upload_products_cn\upload.bat"
set "PY=%ROOT%.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
set "FB=%ROOT%product_classifier\fillback.py"
set "OPENWAIT=%ROOT%open_wait.py"
set "SNAP=%ROOT%product_classifier\output\sku_match_result_review_snapshot.json"
set "CAND=%ROOT%product_classifier\output\sku_match_result_fillback_candidates.json"
set "REVIEW=%ROOT%review_confirm.ps1"
set "STATION_PICK=%ROOT%station_pick.ps1"
set "ADAPT=%ROOT%upload_products_cn\site_adapt.py"
set "OUTDIR=%ROOT%sku_matcher\output"

set "SRC=%~1"
if "%SRC%"=="" (
    echo [ERROR] Drag a goods-code xlsx onto this bat to start.
    echo         Or run: full_pipeline.bat "货号清单.xlsx"
    pause
    exit /b 1
)
if not exist "%SRC%" (
    echo [ERROR] File not found: %SRC%
    pause
    exit /b 1
)

set "F1=%ROOT%sku_matcher\output\sku_match_result.xlsx"
set "F2=%ROOT%product_classifier\output\sku_match_result_classified.xlsx"

REM ============================================
REM  [0] pick station right after drop
REM ============================================
echo.
echo ============================================
echo  [0] Select target station
echo ============================================
set "STATION_FILE=%ROOT%product_classifier\output\station_pick.txt"
if exist "%STATION_FILE%" del /f "%STATION_FILE%" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%STATION_PICK%" "%STATION_FILE%"
set "STATION=cn"
if exist "%STATION_FILE%" ( set /p STATION=<"%STATION_FILE%" )
echo  Selected station: %STATION%
echo.

REM All stations share the same flow below (classify + review/fillback),
REM upload step only runs for cn.

echo ============================================
echo  [1/4] Matching goods codes
echo ============================================
call "%MATCH_BAT%" "%MATCH_SCRIPT%" match "%SRC%"
if not exist "%F1%" (
    echo [FAIL] %F1% not generated. Aborted.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  [2/4] Classifying products
echo ============================================
call "%CLASSIFY_BAT%" "%F1%"
if not exist "%F2%" (
    echo [FAIL] %F2% not generated. Aborted.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  [3/4] Review classify + fillback
echo ============================================
echo  Classification result:
echo    %F2%
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
echo        Cancel aborts (no fillback, no further steps).
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%REVIEW%" "%CAND%"
if errorlevel 1 (
    echo.
    echo [CANCELLED] Fillback aborted. Result kept at: %F2%
    echo   To continue later, drag the file onto upload_products_cn\upload.bat
    pause
    exit /b 1
)
echo.
echo  [3.5] Fillback to master DB + cleanup ...
"%PY%" "%FB%" apply --classified "%F2%" --snapshot "%SNAP%"
echo.
echo ============================================
echo  [4/4] Generating %STATION% upload table
echo ============================================
set "GEN_FILE=%ROOT%product_classifier\output\last_generated.txt"
if exist "%GEN_FILE%" del /f "%GEN_FILE%" >nul 2>&1
"%PY%" "%ADAPT%" --station %STATION% --input "%F2%" --out "%OUTDIR%" --pathfile "%GEN_FILE%"
set "F3="
if exist "%GEN_FILE%" ( set /p F3=<"%GEN_FILE%" )
if not defined F3 (
    echo [FAIL] Station template not generated. Aborted.
    pause
    exit /b 1
)
echo  Generated: %F3%
echo.

if /I "%STATION%"=="cn" (
    echo ============================================
    echo  [4.5] Uploading (cn) ...
    echo ============================================
    call "%UPLOAD_BAT%" "%F3%" cn
    if errorlevel 1 (
        echo [NOTE] Upload returned non-zero. See upload.log / last_run.json
        pause
        exit /b 1
    )
) else (
    echo ============================================
    echo  DONE. %STATION% upload table generated & validated.
    echo  Auto upload is DISABLED for %STATION%.
    echo  To upload later, drag the table onto upload_products_cn\upload.bat
    echo    Template: %F3%
    echo ============================================
    pause
    exit /b 0
)

echo.
echo ============================================
echo  Pipeline complete.
echo ============================================
pause
