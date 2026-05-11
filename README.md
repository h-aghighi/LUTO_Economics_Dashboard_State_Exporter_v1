LUTO Economics Dashboard State Exporter v1
================================================

Purpose
-------
This tool extracts the LUTO dashboard economics overview data and exports final dashboard-style Excel workbooks.

Output structure:
- Each state is saved as a separate folder.
- Each state folder contains one Excel workbook.
- Each region is saved as a separate sheet inside that state workbook.
- Only final Excel output files are saved. Diagnostic/intermediate files are not saved.

Required files
--------------
Place these files in the same folder:

1. luto_economics_dashboard_state_exporter_v1.py
2. 01_Run_All_States.bat
3. 02_Run_Queensland_Only.bat
4. README_User_Guide.txt

Recommended folder structure:

LUTO_Economics_Dashboard_State_Exporter_v1
│
├── luto_economics_dashboard_state_exporter_v1.py
├── README_User_Guide.txt
├── 01_Run_All_States.bat
├── 02_Run_Queensland_Only.bat
└── outputs

Python requirements
-------------------
This tool requires Python and the following packages:

- pandas
- openpyxl

If these packages are already installed, no installation is needed.

If they are missing, open PowerShell in the tool folder and run:

python -m pip install pandas openpyxl

Correct data folder
-------------------
When the BAT file asks for the data folder, paste the path ending with:

DATA_REPORT\data

Example:

C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data

Do NOT use:

C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data\map_layers

The script reads the parent data folder and finds the required economics overview file from there.

How to run all states
---------------------
1. Open the tool folder.
2. Double-click:

01_Run_All_States.bat

3. Paste the DATA_REPORT\data folder path when asked.
4. Enter an output prefix, for example:

Run_G0001

5. Open the outputs folder.

Expected output example:

outputs
│
├── Queensland
│   └── Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx
│
├── Victoria
│   └── Run_G0001_Victoria_Economics_Dashboard_Final_Table.xlsx
│
├── New_South_Wales
│   └── Run_G0001_New_South_Wales_Economics_Dashboard_Final_Table.xlsx
│
└── ...

How to run Queensland only
--------------------------
1. Open the tool folder.
2. Double-click:

02_Run_Queensland_Only.bat

3. Paste the DATA_REPORT\data folder path when asked.
4. Enter an output prefix, for example:

Run_G0001

5. Open:

outputs\Queensland

Expected output example:

outputs\Queensland\Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx

What the Excel file contains
----------------------------
Each workbook contains dashboard-style tables.

Each sheet represents one region.

For example, in the Queensland workbook, sheets may include:

- Fitzroy
- Burdekin
- Burnett Mary
- South East Queensland
- Wet Tropics
- Cape York
- Mackay Whitsunday

Each regional sheet follows the dashboard export structure:

Row 1: Chart title
Row 2: Category and economics series names
Rows below: Yearly values from 2020 to 2050

Common errors
-------------
1. Script not found

Make sure this file is in the same folder as the BAT files:

luto_economics_dashboard_state_exporter_v1.py

2. Data folder does not exist

Make sure the pasted path is correct and ends with:

DATA_REPORT\data

3. No output created

Check that the selected data folder contains:

Economics_overview_sum.js

This is usually located directly inside:

DATA_REPORT\data

4. Python package error

Run:

python -m pip install pandas openpyxl

Recommended sharing package
---------------------------
Share the folder as:

LUTO_Economics_Dashboard_State_Exporter_v1.zip

The folder should contain:

- luto_economics_dashboard_state_exporter_v1.py
- README_User_Guide.txt
- 01_Run_All_States.bat
- 02_Run_Queensland_Only.bat
- outputs
