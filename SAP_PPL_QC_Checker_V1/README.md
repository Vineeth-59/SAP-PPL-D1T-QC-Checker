# SAP PPL QC Checker V1

Upload a PPL Excel and a D1T Excel export. Select T-Code, Company Code and optional Value Stream, then run QC.

For V_T001 the built-in mappings include:
- Currency -> Currency
- Country -> Country/Region Key
- Description -> Company Name
- City -> City
- Address -> Address
- Language Key -> Language Key

PASS/FAIL is deterministic Python logic. No client data is included in this project.

## Run on Windows
1. Extract the ZIP.
2. Open PowerShell in the extracted folder.
3. Run:
   `python -m pip install -r requirements.txt`
4. Run:
   `python -m streamlit run app.py`
5. Open the localhost address shown in the terminal.

You can also double-click `run_project.bat`.
