# Audit-Report-Program v.0.1
 Python program to collate multiple Excel sheets into one Excel report

HOW TO USE:
Note: In order to set up the data you will need to use the CTS Audit Tool and make no changes to it's format as of 2/2/2022.

1. Create a folder titled Audits to Process.
2. Move all completed audit files that need to be processed into this folder.
3. Run Audit Report program.
4. Navigate to the 'Audits to Process' folder.
5. Select "Choose".
6. The program will generate a new Excel file called Audit_Averages_Report_'currentdate'.xslx
7. Ensure that the Files Processed value in the report equals the amount of audit files

Troubleshooting:

If no file is created take the following steps:
1. Ensure that files in the folder are in the correct format for the CTS Audit Tool.
2. Ensure that you have selected the correct file folder.
3. Run again.

If the Files Processed does not match the number of files in the folder:
1. Determine if any files are listed for  Incomplete scores
2. Correct these files.
3. Rune again.

If all else fails email @telimektar3

Still to implement:
1. Identify the Policy number and discipline of "NO" items
2. Identify the staff person responsible for the missed items
 
