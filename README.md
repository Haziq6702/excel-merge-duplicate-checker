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
Input and Output Paths
Input files and folders are selected by the user when the program runs. No personal computer path is required in the public template.

In the Configuration section, OUTPUT_DIR is optional:
OUTPUT_DIR = ""  # Enter your preferred output folder path here.
- Leave it blank to save the result in the current working directory.
- Enter a folder path to save results in that folder.
Supported Input Formats
- .xlsx
- .xlsm
Notes
- The script does not upload data to an external service.
- The source workbooks are read; the merged result is written to a separate output workbook.
- Use test or non-confidential data when demonstrating the project.
- Review your employer's policies before publishing code developed as part of your employment.
Disclaimer
This is a general-purpose Excel automation template. Users are responsible for selecting appropriate input files, configuring the output location, and ensuring they have permission to process and share their data.
