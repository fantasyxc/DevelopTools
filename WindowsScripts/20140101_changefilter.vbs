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
	  Case "t1":
		 row = start
	  Case "t2":
		 row = start + 1
	  Case "t3":
		 row = start + 2
	  Case "t4":
		 row = start + 3
	  Case Else:
		 row = start + 4
   End Select

   ws.Cells(row,1).Value = t
   ws.Cells(row,2).Value =  ws.Cells(row,2).Value + 1

   start = row + 1
   
   Select Case team
	  Case "team1":
		 row = start
		 var = "team1"
	  Case "team2":
		 row = start + 1
		 var = "team2"
	  Case Else:
		 row = start + 6
		 var = "unknown team"
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
   Set ws = ThisWorkbook.Worksheets(1)
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