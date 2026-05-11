LUTO Economics Dashboard State Exporter
======================================

Purpose
-------
This tool extracts the LUTO dashboard economics overview data and exports final dashboard-style Excel workbooks.

The production script is:

luto_economics_dashboard_state_exporter.py

There should be no public v1.py, v2.py, or v3.py scripts in the final shared folder.

Output structure
----------------
The tool saves outputs by state.

For one run:

outputs
│
├── Queensland
│   └── Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx
│
├── Victoria
│   └── Run_G0001_Victoria_Economics_Dashboard_Final_Table.xlsx
│
└── ...

For multiple runs:

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

Workbook structure
------------------
Each state workbook contains:

1. One sheet per region.
2. A final state-total sheet with the same name as the state.

Example for Queensland:

Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx

Sheets:
- Burdekin
- Burnett Mary
- Cape York
- Fitzroy
- ...
- Queensland

The final Queensland sheet is the sum of all regional sheets in that workbook.

Required files
--------------
Place these files in the same folder:

1. install_requirements_helper.py
2. 00_Install_Requirements.bat
3. luto_economics_dashboard_state_exporter.py
4. 01_Run_All_States.bat
5. 02_Run_Queensland_Only.bat
6. 03_Run_All_Runs_All_States.bat
7. 04_Run_All_Runs_Queensland_Only.bat
8. README.md
9. README_User_Guide.txt
10. outputs folder

Python requirements
-------------------
This tool requires Python and the following packages:

- pandas
- openpyxl

Install/check packages
----------------------
Double-click:

00_Install_Requirements.bat

Or run from PowerShell:

python .\install_requirements_helper.py

Manual package install commands:

python -m pip install pandas openpyxl
python -m pip install --user pandas openpyxl
conda install pandas openpyxl
conda install -c conda-forge pandas openpyxl

Correct input folder
--------------------
For a single run, use the folder ending with:

DATA_REPORT\data

Do NOT use:

DATA_REPORT\data\map_layers

For multiple runs, use the folder that contains all Run_* folders.

Example:

S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data

Run using BAT files
-------------------
1. 00_Install_Requirements.bat
2. 01_Run_All_States.bat
3. 02_Run_Queensland_Only.bat
4. 03_Run_All_Runs_All_States.bat
5. 04_Run_All_Runs_Queensland_Only.bat

Run directly from PowerShell
----------------------------
Open PowerShell and go to the tool folder:

cd C:\LUTO_Extractor

Single run, all states:

python ".\luto_economics_dashboard_state_exporter.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --start-year 2020 --end-year 2050

Single run, Queensland only:

python ".\luto_economics_dashboard_state_exporter.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --states Queensland --start-year 2020 --end-year 2050

All runs, all states:

python ".\luto_economics_dashboard_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data" --output-dir "C:\LUTO_Extractor\outputs" --start-year 2020 --end-year 2050

All runs, Queensland only:

python ".\luto_economics_dashboard_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data" --output-dir "C:\LUTO_Extractor\outputs" --states Queensland --start-year 2020 --end-year 2050

Selected runs only:

python ".\luto_economics_dashboard_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES3 30 April\Report_Data" --output-dir "C:\LUTO_Extractor\outputs" --run-names Run_G0001 Run_G0013 --start-year 2020 --end-year 2050

What the Excel file contains
----------------------------
Each regional and state-total sheet follows this structure:

Row 1: Chart title
Row 2: Category and economics series names
Rows below: Yearly values from 2020 to 2050

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

Common errors
-------------
1. Script not found

Make sure this file is in the same folder as the BAT files:

luto_economics_dashboard_state_exporter.py

2. Wrong data folder

Use:

DATA_REPORT\data

Do not use:

DATA_REPORT\data\map_layers

3. No output created

Check that the selected data folder contains:

Economics_overview_sum.js

4. Python package error

Run:

python .\install_requirements_helper.py

or:

python -m pip install pandas openpyxl

If permission is blocked:

python -m pip install --user pandas openpyxl
