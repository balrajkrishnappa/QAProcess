' Get the file path from the command line argument
strFilePath = WScript.Arguments.Item(0)

' Create an Excel object
Set objExcel = CreateObject("Excel.Application")

' Open the Excel workbook
Set objWorkbook = objExcel.Workbooks.Open(strFilePath)

' Convert the active sheet of the Excel workbook to CSV format
objWorkbook.ActiveSheet.SaveAs Replace(strFilePath, ".xlsx", ".csv"), 6

' Close the workbook and quit Excel
objWorkbook.Close False
objExcel.Quit

' Release the objects from memory
Set objWorkbook = Nothing
Set objExcel = Nothing

' Inform the user that the conversion is complete
WScript.Echo "The file has been converted from XLSX to CSV format."
