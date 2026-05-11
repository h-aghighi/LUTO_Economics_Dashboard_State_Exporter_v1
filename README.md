# LUTO Economics Dashboard State Exporter

## Purpose

This tool extracts the **LUTO dashboard economics overview data** and exports final dashboard-style Excel workbooks.

The production script is:

```text
luto_economics_dashboard_state_exporter.py
```

The exporter reads:

```text
DATA_REPORT\data\Economics_overview_sum.js
```

It saves only the final Excel outputs. It does not save diagnostic or intermediate extraction files.

---

## Final repository contents

```text
LUTO_Economics_Dashboard_State_Exporter_v1
│
├── install_requirements_helper.py
├── 00_Install_Requirements.bat
├── luto_economics_dashboard_state_exporter.py
├── 01_Run_All_States.bat
├── 02_Run_Queensland_Only.bat
├── 03_Run_All_Runs_All_States.bat
├── 04_Run_All_Runs_Queensland_Only.bat
├── README.md
├── README_User_Guide.txt
└── outputs
    └── .gitkeep
```

There should be no public `v1.py`, `v2.py`, or `v3.py` scripts in the final repository.

---

## Python requirements

This tool requires Python and these packages:

```text
pandas
openpyxl
```

If they are already installed, no installation is needed.

To check or install them, double-click:

```text
00_Install_Requirements.bat
```

or run:

```powershell
python .\install_requirements_helper.py
```

Manual alternatives:

```powershell
python -m pip install pandas openpyxl
python -m pip install --user pandas openpyxl
py -m pip install pandas openpyxl
py -m pip install --user pandas openpyxl
conda install pandas openpyxl
conda install -c conda-forge pandas openpyxl
```

Verify installation:

```powershell
python -c "import pandas, openpyxl; print('Packages installed successfully')"
```

---

## Correct input folder

For a single run, use the folder ending with:

```text
DATA_REPORT\data
```

Example:

```text
C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data
```

Do **not** use:

```text
C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data\map_layers
```

For multiple runs, use the folder that contains all `Run_*` folders.

Example:

```text
S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data
```

This folder should contain:

```text
Run_G0001\DATA_REPORT\data
Run_G0002\DATA_REPORT\data
Run_G0013\DATA_REPORT\data
```

---

## Output structure

### Single-run mode

```text
outputs
│
├── Queensland
│   └── Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx
│
├── Victoria
│   └── Run_G0001_Victoria_Economics_Dashboard_Final_Table.xlsx
│
└── ...
```

### Batch-run mode

```text
outputs
│
├── Queensland
│   ├── Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx
│   ├── Run_G0002_Queensland_Economics_Dashboard_Final_Table.xlsx
│   └── Run_G0013_Queensland_Economics_Dashboard_Final_Table.xlsx
│
├── Victoria
│   ├── Run_G0001_Victoria_Economics_Dashboard_Final_Table.xlsx
│   ├── Run_G0002_Victoria_Economics_Dashboard_Final_Table.xlsx
│   └── Run_G0013_Victoria_Economics_Dashboard_Final_Table.xlsx
│
└── ...
```

---

## Workbook structure

Each state workbook contains:

1. One sheet per region.
2. A final state-total sheet with the same name as the state.

Example:

```text
Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx

Sheets:
1. Burdekin
2. Burnett Mary
3. Cape York
4. Fitzroy
...
last sheet: Queensland
```

The final state sheet is the row-wise sum of all regional sheets in that workbook.

```text
Queensland row for 2020 =
    Burdekin row for 2020
  + Burnett Mary row for 2020
  + Cape York row for 2020
  + Fitzroy row for 2020
  + ...
```

The same logic is applied for every year and every economics column.

---

## Running from BAT files

### Check/install packages

Double-click:

```text
00_Install_Requirements.bat
```

### Run one data folder for all states

Double-click:

```text
01_Run_All_States.bat
```

### Run one data folder for Queensland only

Double-click:

```text
02_Run_Queensland_Only.bat
```

### Run all Run_* folders for all states

Double-click:

```text
03_Run_All_Runs_All_States.bat
```

### Run all Run_* folders for Queensland only

Double-click:

```text
04_Run_All_Runs_Queensland_Only.bat
```

---

## Running directly from PowerShell

Open PowerShell and go to the tool folder:

```powershell
cd C:\LUTO_Extractor
```

### Single run: all states

```powershell
python ".\luto_economics_dashboard_state_exporter.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --start-year 2020 --end-year 2050
```

### Single run: Queensland only

```powershell
python ".\luto_economics_dashboard_state_exporter.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --states Queensland --start-year 2020 --end-year 2050
```

### Single run: selected regions

```powershell
python ".\luto_economics_dashboard_state_exporter.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --regions Fitzroy "Burnett Mary" "South East Queensland" --start-year 2020 --end-year 2050
```

### Batch mode: all runs, all states

```powershell
python ".\luto_economics_dashboard_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data" --output-dir "C:\LUTO_Extractor\outputs" --start-year 2020 --end-year 2050
```

### Batch mode: all runs, Queensland only

```powershell
python ".\luto_economics_dashboard_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data" --output-dir "C:\LUTO_Extractor\outputs" --states Queensland --start-year 2020 --end-year 2050
```

### Batch mode: selected runs only

```powershell
python ".\luto_economics_dashboard_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data" --output-dir "C:\LUTO_Extractor\outputs" --run-names Run_G0001 Run_G0013 --start-year 2020 --end-year 2050
```

---

## Excel table structure

Each regional and state-total sheet follows this structure:

```text
Row 1: Chart title
Row 2: Category and economics series names
Rows below: Yearly values from 2020 to 2050
```

The exported values correspond to dashboard economics series such as:

- Agricultural Land-use (revenue)
- Agricultural Management (revenue)
- Non-Agricultural Land-use (revenue)
- Agricultural Land-use (cost)
- Agricultural Management (cost)
- Non-Agricultural Land-use (cost)
- Transition cost (Ag2Ag)
- Transition cost (Ag2Non-Ag)
- Profit

---

## Common errors

### Script not found

Make sure this file exists in the tool folder:

```text
luto_economics_dashboard_state_exporter.py
```

### Wrong data folder

Use:

```text
DATA_REPORT\data
```

Do **not** use:

```text
DATA_REPORT\data\map_layers
```

### No output created

Check that the selected data folder contains:

```text
Economics_overview_sum.js
```

Expected location:

```text
DATA_REPORT\data\Economics_overview_sum.js
```

### Missing Python package

Run:

```powershell
python .\install_requirements_helper.py
```

or:

```powershell
python -m pip install pandas openpyxl
```

If permission is blocked:

```powershell
python -m pip install --user pandas openpyxl
```

---

## GitHub notes

Do not commit generated Excel outputs.

The repository should include only the tool files and an empty `outputs` folder placeholder:

```text
outputs\.gitkeep
```

The `.gitignore` file should exclude generated outputs:

```text
*.xlsx
*.xls
*.csv
outputs/*
!outputs/.gitkeep
```
