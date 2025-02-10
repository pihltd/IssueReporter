# IssueReporter
A script that takes the isuue tracking spreadsheet and creates a simplified Word report

## Usage
python Excel2Doc.py -c configfile.yml

## Config file variables

**excelfile**: Full path to the Excel file with the original info
**workbooklist**: A list of sheet names in the workbook.  Should be 1 or more.
**wordfile**: Full path and name for the Word report file
**report_title**: The title of the report
**descwidth**: The width in inches of the Description column in the report
**recwidth**: The width in inches of the Recommendations column in the report
**headercolor**: The cell shading color in hex for the column headers
**sectioncolor**: The cell shading color in hex for the individual report sections
**linecolor**: The cell shading color in hex for the alternating lines
