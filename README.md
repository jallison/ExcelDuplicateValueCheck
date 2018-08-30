# ExcelDuplicateValueCheck

VBA code that will monitor changes for specified columns on a worksheet. If there is a change in those columns it will search multiple worksheets to see if that value already exists.

This project is for bringing in NJUNS tickets to excel. There will be two main goals.

Goal 1: Import data into excel from NJUNS.
The process will consist of the following:

1. Export data from NJUNS into an excel file.
2. Copy the exported data and past data into the "Import" sheet starting on row 5.
3. Click the "Convert" button to rearrange the data into a column format that we use. 
	This will also check for duplicates in the existing worksheet.
	If a duplicate is found in the open ticket worksheet  - ticket will be removed
	If a duplicate is found in the completed ticket worksheet - it will be identified as a kick back in the notes field.
	The page then will sort based on the notes field.
4. Any ticket not identified as a kick back can be cut and pasted into the open ticket worksheet.
5. Any ticket identified as a kick back can be researched.

Goal 2: Any manual data entered into the open ticket worksheet will be checked for duplicates.

This process will do a duplicate check across all needed worksheets with a manual entry of data

I am using the FindAll function created by Chip Pearson.
www.cpearson.com/Excel/FindAll.aspx