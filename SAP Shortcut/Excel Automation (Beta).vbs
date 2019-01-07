Option Explicit
Dim objExcel, strExcelPath, objSheet

strExcelPath = "C:\Users\vntruongv\Desktop\Testing.xlsx"

' Open specified spreadsheet and select the first worksheet.
Set objExcel = CreateObject("Excel.Application")
objExcel.WorkBooks.Open strExcelPath
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

' Modify a cell.
objExcel.Application.Visible = True
objSheet.Cells(1, 1).Value = "Truong VanLong"
objSheet.Cells(1, 2).Value = "Chau Pham Khanh Ha"


' Add a row before row 5.
' objSheet.Range("B5").EntireRow.Insert

' Label the new rows. Row 4 is unchanged,
' but what was row 5 is now row 6.
' objSheet.Cells(4, 2).Value = "Row 4"
' objSheet.Cells(5, 2).Value = "Row 5"
' objSheet.Cells(6, 2).Value = "Row 6"

' Save and quit.
objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit