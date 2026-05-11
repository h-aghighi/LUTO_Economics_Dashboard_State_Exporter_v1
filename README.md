# LUTO Economics Dashboard State Exporter v1

## Purpose

This tool extracts the **LUTO dashboard economics overview data** and exports final dashboard-style Excel workbooks.

The output is organised as:

- Each state is saved as a separate folder.
- Each state folder contains one Excel workbook.
- Each region is saved as a separate sheet inside the state workbook.
- Only final Excel output files are saved.
- Diagnostic and intermediate files are not saved.

---

## Repository contents

```text
LUTO_Economics_Dashboard_State_Exporter_v1
│
├── install_requirements_helper.py
├── 00_Install_Requirements.bat
├── luto_economics_dashboard_state_exporter_v1.py
├── 01_Run_All_States.bat
├── 02_Run_Queensland_Only.bat
├── README.md
├── README_User_Guide.txt
└── outputs
    └── .gitkeep
```

---

## Python requirements

This tool requires Python and these packages:

```text
pandas
openpyxl
```

If Python and these packages are already installed, no installation is needed.

---

## Installing or checking required packages

### Option 1: Double-click installer

Double-click:

```text
00_Install_Requirements.bat
```

This checks whether the required packages are installed. If they are missing, it tries several installation methods automatically.

### Option 2: Run helper from PowerShell

Open PowerShell in the tool folder and run:

```powershell
python .\install_requirements_helper.py
```

### Option 3: Manual installation commands

If the helper does not work, try these commands manually in PowerShell.

First try:

```powershell
python -m pip install pandas openpyxl
```

If that fails because of permission restrictions, try:

```powershell
python -m pip install --user pandas openpyxl
```

If the `py` launcher is available, try:

```powershell
py -m pip install pandas openpyxl
```

or:

```powershell
py -m pip install --user pandas openpyxl
```

If using Anaconda or Miniconda, try:

```powershell
conda install pandas openpyxl
```

or:

```powershell
conda install -c conda-forge pandas openpyxl
```

To verify installation:

```powershell
python -c "import pandas, openpyxl; print('Packages installed successfully')"
```

---

## Correct input folder

When the BAT file asks for the data folder, paste the path ending with:

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

The script reads the parent `data` folder and finds the required economics overview file from there.

The required source file is usually:

```text
DATA_REPORT\data\Economics_overview_sum.js
```

---

## Running from BAT files

### Run all states

1. Open the tool folder.
2. Double-click:

```text
01_Run_All_States.bat
```

3. Paste the `DATA_REPORT\data` folder path when asked.
4. Enter an output prefix, for example:

```text
Run_G0001
```

5. Open the `outputs` folder.

Expected output structure:

```text
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
```

### Run Queensland only

1. Open the tool folder.
2. Double-click:

```text
02_Run_Queensland_Only.bat
```

3. Paste the `DATA_REPORT\data` folder path when asked.
4. Enter an output prefix, for example:

```text
Run_G0001
```

5. Open:

```text
outputs\Queensland
```

Expected output:

```text
outputs\Queensland\Run_G0001_Queensland_Economics_Dashboard_Final_Table.xlsx
```

---

## Running directly from PowerShell

Open PowerShell and go to the tool folder:

```powershell
cd C:\LUTO_Extractor
```

Check that the script exists:

```powershell
dir
```

You should see:

```text
luto_economics_dashboard_state_exporter_v1.py
```

### Run all states

```powershell
python ".\luto_economics_dashboard_state_exporter_v1.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --output-prefix Run_G0001 --start-year 2020 --end-year 2050
```

### Run Queensland only

```powershell
python ".\luto_economics_dashboard_state_exporter_v1.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --output-prefix Run_G0001 --states Queensland --start-year 2020 --end-year 2050
```

### Run selected regions only

Example for Fitzroy:

```powershell
python ".\luto_economics_dashboard_state_exporter_v1.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --output-prefix Run_G0001 --regions Fitzroy --start-year 2020 --end-year 2050
```

Example for multiple regions:

```powershell
python ".\luto_economics_dashboard_state_exporter_v1.py" --data-dir "C:\Users\hosse\Downloads\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --output-prefix Run_G0001 --regions Fitzroy "Burnett Mary" "South East Queensland" --start-year 2020 --end-year 2050
```

Use quotation marks around region names that contain spaces.

### Run another scenario

Only change the input folder and output prefix.

Example for `Run_G0013`:

```powershell
python ".\luto_economics_dashboard_state_exporter_v1.py" --data-dir "C:\Users\hosse\Downloads\Run_G0013\DATA_REPORT\data" --output-dir "C:\LUTO_Extractor\outputs" --output-prefix Run_G0013 --start-year 2020 --end-year 2050
```

---

## What the Excel files contain

Each workbook contains dashboard-style economics tables.

Each sheet represents one region.

For example, the Queensland workbook may include sheets such as:

- Fitzroy
- Burdekin
- Burnett Mary
- South East Queensland
- Wet Tropics
- Cape York
- Mackay Whitsunday

Each regional sheet follows this structure:

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

If PowerShell says it cannot find:

```text
luto_economics_dashboard_state_exporter_v1.py
```

make sure you are in the correct folder:

```powershell
cd C:\LUTO_Extractor
```

Then run:

```powershell
dir
```

The script must be listed there.

### Data folder does not exist

Check that the pasted path exists and ends with:

```text
DATA_REPORT\data
```

Do **not** use the `map_layers` folder.

### No output created

Check that this file exists inside the selected data folder:

```text
Economics_overview_sum.js
```

Expected location:

```text
DATA_REPORT\data\Economics_overview_sum.js
```

If this file is missing, the selected folder is not the correct LUTO dashboard data folder.

### Missing Python package

If you see an error such as:

```text
ModuleNotFoundError: No module named 'pandas'
```

run:

```powershell
python .\install_requirements_helper.py
```

or:

```powershell
python -m pip install pandas openpyxl
```

If permission is blocked, try:

```powershell
python -m pip install --user pandas openpyxl
```

---

## Notes for GitHub users

Do not commit generated Excel outputs.

The repository should include only the tool files and an empty `outputs` folder placeholder:

```text
outputs\.gitkeep
```

The `.gitignore` file should exclude generated outputs such as:

```text
*.xlsx
*.xls
*.csv
outputs/*
!outputs/.gitkeep
```

---

## Recommended sharing package

Share the repository or zip folder as:

```text
LUTO_Economics_Dashboard_State_Exporter_v1.zip
```

The folder should contain:

```text
install_requirements_helper.py
00_Install_Requirements.bat
luto_economics_dashboard_state_exporter_v1.py
README.md
README_User_Guide.txt
01_Run_All_States.bat
02_Run_Queensland_Only.bat
outputs
```
