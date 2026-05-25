# AddNewHistory

VBA macro module for FrontPage.xlsm to append revision history rows to product workbooks.

## File

- AddNewHitory.bas

## What It Does

- Reads product names from Sheets(2) column A in FrontPage.xlsm
- Skips products marked `New` in column N
- Opens corresponding files under products folder
- Appends one row to Sheets(2) history section
- Copies columns B:C from previous last row
- Writes fixed values:
  - D: 06/01/2026
  - E: Adopted a new format of the Front Page
- Applies formatting to appended row:
  - D cell horizontal center
  - Appended row height: 18
  - Next row height: 6
  - Grid borders on B:E

## Usage

1. Import AddNewHitory.bas into FrontPage.xlsm standard module.
2. Run macro `AddNewHitory`.
3. Review the result message and Immediate Window logs.

## Notes

- Target path is hardcoded in the module:
  - D:\Ｄ_Project\BSIフロントページ\products\
- Ensure file names match product names in column A.
