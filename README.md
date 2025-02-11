# IssueReporter
A script that takes the isuue tracking spreadsheet and creates a simplified Word report

## Usage
python Excel2Doc.py -c configfile.yml

## Config file variables

**excelfile**: Full path to the Excel file with the original info

**workbooklist**: A list of sheet names in the workbook.  Should be 1 or more.

**wordfile**: Full path and name for the Word report file

**report_title**: The title of the report

**table_headers**: A list of the text that should be the header for each column

**content**: This is a bit complex - {cellindex:{[Excel column header: Table labels]}.  If the table label is None, the content is added to the cell as-is}

**colwidths**: A list of the column index and the widths in inches.  {column_index:width_inches}

**headercolor**: The cell shading color in hex for the column headers

**sectioncolor**: The cell shading color in hex for the individual report sections

**linecolor**: The cell shading color in hex for the alternating lines
