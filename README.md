# eCTD_File_Validator
eCTD file existence validator that checks existence of outputted files, in respect to .xmls that define the outputted structure.

The main advantage of this Python script is the vast amount of data which it can check. Not for one/dozens of dossier outputs, but thousands/millions in theory.

# Result
Output of this major scope validator are 3 Excel files:

eCTD_XML_Report.xlsx - contains the relevant file output paths from the index and regional .xmls

FilePaths.xlsx - contains the relevant file output paths from the inputted folder

eCTD_file_compare - contains the comparison of the real eCTD file output that is present in the defined folders vs relevant file output paths extracted from the index and regional .xmls

# Prerequisite libraries

xlsxwriter, lxml

