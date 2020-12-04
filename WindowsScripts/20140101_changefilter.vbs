Function IsExist(mws, name1, name2,team As String) As Boolean
   Dim i As Integer
   IsExist = False

   If name1 = "" Or name2 = "" Then
      Exit Function
   End If
   
   For i = 3 To mws.UsedRange.Rows.Count Step 1
      If mws.Cells(i, 3).Value <> "" And (StrComp(mws.Cells(i, 3).Value, name1, vbTextCompare) = 0 Or StrComp(mws.Cells(i, 3).Value, name2, vbTextCompare) = 0) Then
		 team = mws.Cells(i,2).value
		 IsExist = True
         Exit Function
      End If
      
   Next i
   
End Function

Function Record(ws,t, team,s)
   Dim row As Integer
   Dim var As String
   Dim start As Integer
   
   start = s
   Select Case t
	  Case "应急":
		 row = start
	  Case "标准":
		 row = start + 1
	  Case "常规":
		 row = start + 2
	  Case "重大":
		 row = start + 3
	  Case Else:
		 row = start + 4
   End Select

   ws.Cells(row,1).Value = t
   ws.Cells(row,2).Value =  ws.Cells(row,2).Value + 1

   start = row + 1
   
   Select Case team
	  Case "应用支持":
		 row = start
		 var = "应用"
	  Case "风险管理":
		 row = start + 1
		 var = "风险"
	  Case "运行管理":
		 row = start + 2
		 var = "运行"
	  Case "网络支持":
		 row = start + 3
		 var = "网络"
	  Case "环境管理":
		 row = start + 4
		 var = "环境"
	  Case "基础设施":
		 row = start + 5
		 var = "基础设施"
	  Case Else:
		 row = start + 6
		 var = "未知团队"
   End Select

   ws.Cells(row,1).Value = team
   ws.Cells(row,2).Value = ws.Cells(row,2).Value  + 1
		 
End Function

Sub Filter()
   Dim ws As Worksheet
   Dim mws As Worksheet
   Dim name As String
   Dim i As Integer
   Dim start As Integer
   Dim team As String
   name = "C:\namelist.xlsx"
   Set ws = ThisWorkbook.Worksheets(1) '导出变更列表的第一个表格
   Set mws = Workbooks.Open(Filename:=name).Worksheets(1)
   i = 2
   start = ws.UsedRange.Rows.Count + 1
   
   While ws.Cells(i,1).Value  <> ""
      If Not IsExist(mws, ws.Cells(i, 7).Value, ws.Cells(i, 8).Value,team) Then
         ws.Rows(i).EntireRow.Delete Shift:=xlShiftUp
	  Else
		 'Record ws,ws.Cells(i,4).Value,team,start
		 'ws.Cells(i+50,1).Value = "start:" & start
		 i = i + 1
	  End If
   Wend
   
   mws.Parent.Close SaveChanges:=False
   
End Sub