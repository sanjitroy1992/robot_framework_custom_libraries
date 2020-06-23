ExcelLibraryExtended
Library version:	0.0.2
Library scope:	global
Named arguments:	supported
Introduction
This test library provides keywords to allow opening, reading, writing, and saving Excel files from Unified Test Framework.

Shortcuts
Add New Sheet · Add To Date · Check Cell Type · Create Excel Workbook · Edit Data Xlsx File · Get Column Count · Get Column Values · Get Number Of Sheets · Get Row Count · Get Row Values · Get Sheet Names · Get Sheet Values · Get Workbook Values · Modify Cell With · Open Excel · Open Excel Current Directory · Put Date To Cell · Put Number To Cell · Put String To Cell · Read Cell Data By Coordinates · Read Cell Data By Name · Save Excel · Save Excel Current Directory · Subtract From Date
Keywords
Keyword	Arguments	Documentation
Add New Sheet	newsheetname	
Creates and appends new Excel worksheet using the new sheet name to the current workbook.

Arguments:

New Sheet name (string)	The name of the new sheet added to the workbook.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Add New Sheet	NewSheet
Add To Date	sheetname, column, row, numdays	
Using the sheet name the number of days are added to the date in the indicated cell.

Arguments:

Sheet Name (string)	The selected sheet that the cell will be modified from.
Column (int)	The column integer value that will be used to modify the cell.
Row (int)	The row integer value that will be used to modify the cell.
Number of Days (int)	The integer value containing the number of days that will be added to the specified sheetname at the specified column and row.
Example:

Keywords	Parameters			
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls			
Add To Date	TestSheet1	0	0	4
Check Cell Type	sheetname, column, row	
Checks the type of value that is within the cell of the sheet name selected.

Arguments:

Sheet Name (string)	The selected sheet that the cell type will be checked from.
Column (int)	The column integer value that will be used to check the cell type.
Row (int)	The row integer value that will be used to check the cell type.
Example:

Keywords	Parameters		
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls		
Check Cell Type	TestSheet1	0	0
Create Excel Workbook	newsheetname	
Creates a new Excel workbook

Arguments:

New Sheet Name (string)	The name of the new sheet added to the new workbook.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Create Excel	NewExcelSheet
Edit Data Xlsx File	file_path, sheetname, coloumnheader, changedata, rownumber=None	
Usage
It updates the data in a xlsx file.

Arguments
'file_path' = xlsx file location.

'sheetname' = SheetName of the xlsx file.

'coloumnheader' = ColoumnHeader name.

'changedata' = data to be updated in the xlsx file.

'rownumber' [Optional]= If provided as an integer value, it will update the data into the corrosponding rownumber. By default it selects the first row of the corrosponding columnheader.

Example:

|***TestCases*** |

1. To Update Data Into The First Row of The Corrosponding Columnheader : |Edit Data Xlsx File | file_path=C:/example.xlsx | sheetname=Sheet1 | coloumnheader=Sample | changedata=hello world |

2. To Update Data Into The Fourth Row of The Corrosponding Columnheader : |Edit Data Xlsx File | file_path=C:/example.xlsx | sheetname=Sheet1 | coloumnheader=Sample | changedata=hello world | rownumber= 4

Get Column Count	sheetname	
Returns the specific number of columns of the sheet name specified.

Arguments:

Sheet Name (string)	The selected sheet that the column count will be returned from.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Get Column Count	TestSheet1
Get Column Values	sheetname, column, includeEmptyCells=True	
Returns the specific column values of the sheet name specified.

Arguments:

Sheet Name (string)	The selected sheet that the column values will be returned from.
Column (int)	The column integer value that will be used to select the column from which the values will be returned.
Include Empty Cells (default=True)	The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable.
Example:

Keywords	Parameters	
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls	
Get Column Values	TestSheet1	0
Get Number Of Sheets		
Returns the number of worksheets in the current workbook.

Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Get Number of Sheets	
Get Row Count	sheetname	
Returns the specific number of rows of the sheet name specified.

Arguments:

Sheet Name (string)	The selected sheet that the row count will be returned from.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Get Row Count	TestSheet1
Get Row Values	sheetname, row, includeEmptyCells=True	
Returns the specific row values of the sheet name specified.

Arguments:

Sheet Name (string)	The selected sheet that the row values will be returned from.
Row (int)	The row integer value that will be used to select the row from which the values will be returned.
Include Empty Cells (default=True)	The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable.
Example:

Keywords	Parameters	
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls	
Get Row Values	TestSheet1	0
Get Sheet Names		
Returns the names of all the worksheets in the current workbook.

Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Get Sheets Names	
Get Sheet Values	sheetname, includeEmptyCells=True	
Returns the values from the sheet name specified.

Arguments:

Sheet Name (string)	The selected sheet that the cell values will be returned from.
Include Empty Cells (default=True)	The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Get Sheet Values	TestSheet1
Get Workbook Values	includeEmptyCells=True	
Returns the values from each sheet of the current workbook.

Arguments:

Include Empty Cells (default=True)	The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Get Workbook Values	
Modify Cell With	sheetname, column, row, op, val	
Using the sheet name a cell is modified with the given operation and value.

Arguments:

Sheet Name (string)	The selected sheet that the cell will be modified from.
Column (int)	The column integer value that will be used to modify the cell.
Row (int)	The row integer value that will be used to modify the cell.
Operation (operator)	The operation that will be performed on the value within the cell located by the column and row values.
Value (int)	The integer value that will be used in conjuction with the operation parameter.
Example:

Keywords	Parameters				
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls				
Modify Cell With	TestSheet1	0	0	*	56
Open Excel	filename, useTempDir=False	
Opens the Excel file from the path provided in the file name parameter. If the boolean useTempDir is set to true, depending on the operating system of the computer running the test the file will be opened in the Temp directory if the operating system is Windows or tmp directory if it is not.

Arguments:

File Name (string)	The file name string value that will be used to open the excel file to perform tests upon.
Use Temporary Directory (default=False)	The file will not open in a temporary directory by default. To activate and open the file in a temporary directory, pass 'True' in the variable.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Open Excel Current Directory	filename	
Opens the Excel file from the current directory using the directory the test has been run from.

Arguments:

File Name (string)	The file name string value that will be used to open the excel file to perform tests upon.
Example:

Keywords	Parameters
Open Excel	ExcelRobotTest.xls
Put Date To Cell	sheetname, column, row, value	
Using the sheet name the value of the indicated cell is set to be the date given in the parameter.

Arguments:

Sheet Name (string)	The selected sheet that the cell will be modified from.
Column (int)	The column integer value that will be used to modify the cell.
Row (int)	The row integer value that will be used to modify the cell.
Value (int)	The integer value containing a date that will be added to the specified sheetname at the specified column and row.
Example:

Keywords	Parameters			
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls			
Put Date To Cell	TestSheet1	0	0	12.3.1999
Put Number To Cell	sheetname, column, row, value	
Using the sheet name the value of the indicated cell is set to be the number given in the parameter.

Arguments:

Sheet Name (string)	The selected sheet that the cell will be modified from.
Column (int)	The column integer value that will be used to modify the cell.
Row (int)	The row integer value that will be used to modify the cell.
Value (int)	The integer value that will be added to the specified sheetname at the specified column and row.
Example:

Keywords	Parameters			
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls			
Put Number To Cell	TestSheet1	0	0	34
Put String To Cell	sheetname, column, row, value	
Using the sheet name the value of the indicated cell is set to be the string given in the parameter.

Arguments:

Sheet Name (string)	The selected sheet that the cell will be modified from.
Column (int)	The column integer value that will be used to modify the cell.
Row (int)	The row integer value that will be used to modify the cell.
Value (string)	The string value that will be added to the specified sheetname at the specified column and row.
Example:

Keywords	Parameters			
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls			
Put String To Cell	TestSheet1	0	0	Hello
Read Cell Data By Coordinates	sheetname, column, row	
Uses the column and row to return the data from that cell.

Arguments:

Sheet Name (string)	The selected sheet that the cell value will be returned from.
Column (int)	The column integer value that the cell value will be returned from.
Row (int)	The row integer value that the cell value will be returned from.
Example:

Keywords	Parameters		
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls		
Read Cell	TestSheet1	0	0
Read Cell Data By Name	sheetname, cell_name	
Uses the cell name to return the data from that cell.

Arguments:

Sheet Name (string)	The selected sheet that the cell value will be returned from.
Cell Name (string)	The selected cell name that the value will be returned from.
Example:

Keywords	Parameters	
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls	
Get Cell Data	TestSheet1	A2
Save Excel	filename, useTempDir=False	
Saves the Excel file indicated by file name, the useTempDir can be set to true if the user needs the file saved in the temporary directory. If the boolean useTempDir is set to true, depending on the operating system of the computer running the test the file will be saved in the Temp directory if the operating system is Windows or tmp directory if it is not.

Arguments:

File Name (string)	The name of the of the file to be saved.
Use Temporary Directory (default=False)	The file will not be saved in a temporary directory by default. To activate and save the file in a temporary directory, pass 'True' in the variable.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Save Excel	NewExcelRobotTest.xls
Save Excel Current Directory	filename	
Saves the Excel file from the current directory using the directory the test has been run from.

Arguments:

File Name (string)	The name of the of the file to be saved.
Example:

Keywords	Parameters
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls
Save Excel Current Directory	NewTestCases.xls
Subtract From Date	sheetname, column, row, numdays	
Using the sheet name the number of days are subtracted from the date in the indicated cell.

Arguments:

Sheet Name (string)	The selected sheet that the cell will be modified from.
Column (int)	The column integer value that will be used to modify the cell.
Row (int)	The row integer value that will be used to modify the cell.
Number of Days (int)	The integer value containing the number of days that will be subtracted from the specified sheetname at the specified column and row.
Example:

Keywords	Parameters			
Open Excel	C:\Python27\ExcelRobotTest\ExcelRobotTest.xls			
Subtract From Date	TestSheet1	0	0	7
Altogether 24 keywords.
Generated by Libdoc on 2020-02-11 11:30:31.

