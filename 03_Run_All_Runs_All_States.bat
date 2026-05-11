@echo off
setlocal

echo ==========================================
echo LUTO Economics Dashboard State Exporter v2
echo Run mode: All runs / all states
echo ==========================================
echo.

set TOOL_DIR=%~dp0
set SCRIPT=%TOOL_DIR%luto_economics_dashboard_state_exporter.py
set OUTPUT_DIR=%TOOL_DIR%outputs

if not exist "%SCRIPT%" (
    echo ERROR: Could not find the exporter script:
    echo %SCRIPT%
    echo.
    echo Make sure luto_economics_dashboard_state_exporter.py is in the same folder as this BAT file.
    echo.
    pause
    exit /b 1
)

if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
)

echo Paste the folder that contains all Run_* folders.
echo.
echo Correct example:
echo S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data
echo.
echo This folder should contain:
echo Run_G0001\DATA_REPORT\data
echo Run_G0002\DATA_REPORT\data
echo Run_G0013\DATA_REPORT\data
echo.
set /p REPORTS_BASE_DIR=Report_Data folder path: 

set REPORTS_BASE_DIR=%REPORTS_BASE_DIR:"=%

if not exist "%REPORTS_BASE_DIR%" (
    echo.
    echo ERROR: This folder does not exist:
    echo %REPORTS_BASE_DIR%
    echo.
    pause
    exit /b 1
)

echo.
echo Running batch extraction for all Run_* folders...
echo Input base folder:
echo %REPORTS_BASE_DIR%
echo.
echo Output folder:
echo %OUTPUT_DIR%
echo.

python "%SCRIPT%" ^
  --reports-base-dir "%REPORTS_BASE_DIR%" ^
  --output-dir "%OUTPUT_DIR%" ^
  --start-year 2020 ^
  --end-year 2050

if errorlevel 1 (
    echo.
    echo WARNING: Some runs may have failed. Check the messages above.
    echo Successful runs should still have output files.
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo Extraction finished.
echo Output files are here:
echo %OUTPUT_DIR%
echo ==========================================
echo.
pause

