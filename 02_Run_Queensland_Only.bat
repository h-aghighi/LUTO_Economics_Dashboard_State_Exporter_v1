@echo off
setlocal

echo ==========================================
echo LUTO Economics Dashboard State Exporter v1
echo Run mode: Queensland only
echo ==========================================
echo.

set TOOL_DIR=%~dp0
set SCRIPT=%TOOL_DIR%luto_economics_dashboard_state_exporter_v1.py
set OUTPUT_DIR=%TOOL_DIR%outputs

if not exist "%SCRIPT%" (
    echo ERROR: Could not find the extractor script:
    echo %SCRIPT%
    echo.
    echo Make sure luto_economics_dashboard_state_exporter_v1.py is in the same folder as this BAT file.
    echo.
    pause
    exit /b 1
)

if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
)

echo Paste the path to the LUTO DATA folder.
echo.
echo Correct example:
echo C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data
echo.
echo Do NOT paste the map_layers folder.
echo Wrong example:
echo C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data\map_layers
echo.
set /p DATA_DIR=DATA folder path: 

set DATA_DIR=%DATA_DIR:"=%

if not exist "%DATA_DIR%" (
    echo.
    echo ERROR: This folder does not exist:
    echo %DATA_DIR%
    echo.
    pause
    exit /b 1
)

echo.
set /p PREFIX=Output prefix, e.g. Run_G0001: 

if "%PREFIX%"=="" (
    set PREFIX=LUTO
)

echo.
echo Running Queensland-only extraction...
echo Input folder:
echo %DATA_DIR%
echo.
echo Output folder:
echo %OUTPUT_DIR%
echo.

python "%SCRIPT%" ^
  --data-dir "%DATA_DIR%" ^
  --output-dir "%OUTPUT_DIR%" ^
  --output-prefix "%PREFIX%" ^
  --states Queensland ^
  --start-year 2020 ^
  --end-year 2050

if errorlevel 1 (
    echo.
    echo ERROR: Extraction failed.
    echo Check the messages above.
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo Extraction finished.
echo Queensland output is here:
echo %OUTPUT_DIR%\Queensland
echo ==========================================
echo.
pause
