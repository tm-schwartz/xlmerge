# Merge Excel Workbooks/Sheets

Usage:

<img alt="Gif" src="https://raw.githubusercontent.com/tm-schwartz/xlmerge/main/example.gif" width="80%" />

---

## Caveats 

- This code was written to merge multiple workbooks with multiple worksheets
  into a single file. The original workbooks contained a large amount of macros
  as well. **The file type must be a non-XLSM extension**. It is fine for macros
  to be present, the extension just needs to be i.e. XLSX. 

- The result file is intended to function solely as a single source to copy
  consolidated data from. It will not contain any formatting as of now. After
  running the code, simply copy the data from the result into a final, formatted
  workbook, run any macros etc.

