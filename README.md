# Audit-Report-Program
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

<!-- Will need to be able to identify: -->
<!-- 1. Specific cells on a specific sheet within a group of Excel documents
    a. Batch select the Excel spreadsheets with a Folder select pop-up
2. Store that information as a list (or do it as a dictionary since it will have a key and value? It should be easier to do as a list)
3. Manipulate the information to get average percent score per discipline
4. Output that information into another Excel sheet/report
    a. Have the file name be auto-generated based on a set name and an auto-generated date
    b. Should the file name include -2 if it's generated more than once on the same date? -->
 
