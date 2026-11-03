# Project check and fixes

## Checks completed
- Python source files passed `compileall` syntax compilation.
- Excel workbook was inspected: 600 inventory rows and all 18 expected columns are present.
- Reviewed Flask routes, inventory rules, Excel read/write flow, forms, and frontend JavaScript.
- Fixed delete-item matching so IDs entered with different letter casing still resolve to the actual stored ID.
- Added validation preventing quantity from exceeding maximum stock and maximum stock from being below reorder level.

## Important test limitation
The review environment did not have Flask/Pandas/OpenPyXL preinstalled and could not reach the package index, so live browser/route tests and Excel write-cycle tests could not be executed here. The included `run.bat` installs dependencies on the user's computer and starts the app. Please keep the included original workbook backup before testing destructive actions.

## Run in Antigravity
1. Extract this ZIP.
2. Open the extracted `Inventory_Stock_Tracker` folder in Antigravity.
3. Double-click `run.bat` (or run it in the terminal).
4. Open `http://127.0.0.1:5000` in your browser.

First launch requires internet to install Python packages. Python 3.9+ must be installed.
