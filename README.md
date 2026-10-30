Excel Merge & Duplicate Checker
A Python command-line utility for combining Excel worksheets and workbooks into a consolidated Excel output, with configurable key-column validation and duplicate reporting.
Features
- Merge worksheets from multiple Excel workbooks.
- Merge multiple worksheets from one workbook.
- Select input files or folders at runtime using file/folder dialogs.
- Align columns by header name.
- Select and validate one or more key columns.
- Identify duplicates within a source and across sources.
- Generate a detailed Duplicate Report with source file, folder, worksheet, and row information.
- Highlight duplicate records in the output workbook.
- Choose a single merged-data worksheet or organize output by source folder.
- Split large outputs across worksheets when Excel's row limit is reached.
- Generate timestamped output workbooks.
Requirements
- Python 3.9+
- pandas
- numpy
- openpyxl
- xlsxwriter
- tkinter (usually included with standard Windows Python installations)
Optional: python-calamine may improve Excel reading speed.
Installation
pip install pandas numpy openpyxl xlsxwriter
Optional faster reader:
pip install python-calamine
How to Run
python "excel_merge file.py"
1. Choose whether to merge multiple Excel files or worksheets from one workbook.
2. Select the input Excel files or folder when prompted.
3. Choose the worksheets to include.
4. Select the key column(s) used for duplicate checking.
5. Review the column alignment and merge settings.
6. Choose the output layout.
7. The script creates a timestamped Excel workbook containing the merged data and Duplicate Report.
