
'方法一 
Dim appExcel As Object
Set appExcel = CreateObject("excel.application")
appExcel.Visible = False
appExcel.Workbooks.Open (dailyFilename(9))
With Excel.Application.Workbooks
gaojing = Worksheets("数据").Cells(5, 15)
End With
appExcel.Quit


' 方法二
Dim appExcelYxzb As New Excel.Application
Dim wbExcelYxzb As Excel.Workbook
Set wbExceYxzb = appExcelYxzb.Workbooks.Open(dailyFilename(9))
gaojing = wbExceYxzb.Worksheets("数据").Cells(5, 15)
appExcelYxzb.Quit