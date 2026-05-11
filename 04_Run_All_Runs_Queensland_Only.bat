@echo off
setlocal

echo ==========================================
echo LUTO Economics Dashboard State Exporter v2
echo Run mode: All runs / Queensland only
echo ==========================================
echo.

set TOOL_DIR=%~dp0
set SCRIPT=%TOOL_DIR%luto_economics_dashboard_state_exporter_v2.py
set OUTPUT_DIR=%TOOL_DIR%outputs

if not exist "%SCRIPT%" (
    echo ERROR: Could not find the exporter script:
    echo %SCRIPT%
    echo.
    pause
    exit /b 1
)

if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
)

echo Paste the folder that contains all Run_* folders.
echo Example:
echo S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data
echo.
set /p REPORTS_BASE_DIR=Report_Data folder path: 

set REPORTS_BASE_DIR=%REPORTS_BASE_DIR:"=%

if not exist "%REPORTS_BASE_DIR%" (
    echo ERROR: This folder does not exist:
    echo %REPORTS_BASE_DIR%
    pause
    exit /b 1
)

python "%SCRIPT%" ^
  --reports-base-dir "%REPORTS_BASE_DIR%" ^
  --output-dir "%OUTPUT_DIR%" ^
  --states Queensland ^
  --start-year 2020 ^
  --end-year 2050

if errorlevel 1 (
    echo.
    echo WARNING: Some runs may have failed. Check the messages above.
    echo.
    pause
    exit /b 1
)

echo.
echo Extraction finished.
echo Queensland outputs are here:
echo %OUTPUT_DIR%\Queensland
echo.
pause
