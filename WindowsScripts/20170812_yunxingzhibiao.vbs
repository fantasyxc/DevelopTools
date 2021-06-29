'
'功能：生成日报
'

Private Sub creat()

Dim i ' 临时变量

'检查文件是否存在
Dim myFilename(3)
myFilename(0) = ThisWorkbook.Path & "\alarm.xls"
myFilename(1) = ThisWorkbook.Path & "\change.xls"
myFilename(2) = ThisWorkbook.Path & "\team1.xlsm"
For i = 0 To 2
If Dir(myFilename(i)) = Empty Then
   MsgBox myFilename(i)
   Exit Sub
 End If
Next


'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
Dim alarmnum
alarmnum = 0
Set workbooksalarm = Workbooks.Open(Filename:=myFilename(0), ReadOnly:=True).Worksheets(1)

'清空alarm表中原告警信息
i = 2
While ThisWorkbook.Sheets("alarm").Cells(i, 2).Value <> ""
    ThisWorkbook.Sheets("alarm").Rows(i).EntireRow.Delete Shift:=xlShiftUp
Wend

'读取文件内容
i = 2
While workbooksalarm.Cells(i, 2).Value <> ""
	workbooksalarm.Cells(i, 1).EntireRow.Copy ThisWorkbook.Sheets("alarm").Cells(alarmnum + 2, 1)
	alarmnum = alarmnum + 1
	ThisWorkbook.Sheets("alarm").Cells(alarmnum + 1, 1).Value = alarmnum
    i = i + 1
Wend
ActiveWorkbook.Close True 


Dim changezhongda, changeyingji, changechanggui, changebiaozhun
changezhongda = 0
changeyingji = 0
changechanggui = 0
changebiaozhun = 0

'
i = 2
While ThisWorkbook.Sheets("change").Cells(i, 1).Value <> ""
    ThisWorkbook.Sheets("change").Rows(i).EntireRow.Delete Shift:=xlShiftUp
Wend

Set workbookschange = Workbooks.Open(Filename:=myFilename(1), ReadOnly:=True).Worksheets(1)
i = 2
While workbookschange.Cells(i, 1).Value <> ""
    workbookschange.Cells(i, 1).EntireRow.Copy ThisWorkbook.Sheets("change").Cells(changezhongda + changeyingji + changechanggui + changebiaozhun + 2, 1)
    If InStr(workbookschange.Cells(i, 4).Value, "重大") Then
      changezhongda = changezhongda + 1
    ElseIf InStr(workbookschange.Cells(i, 4).Value, "应急") Then 'Or InStr(workbookschange.Cells(i, 3).Value, "紧急") Then
      changeyingji = changeyingji + 1
    ElseIf InStr(workbookschange.Cells(i, 4).Value, "常规") Then
      changechanggui = changechanggui + 1
    ElseIf InStr(workbookschange.Cells(i, 4).Value, "标准") Then
      changebiaozhun = changebiaozhun + 1
    End If
    i = i + 1
Wend

ActiveWorkbook.Close True 


'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
pathdaji = ThisWorkbook.Path & "\team1.xlsm" '把路径赋值给一个字符串
Set workbooksdaji = Workbooks.Open(Filename:=pathdaji, ReadOnly:=True).Worksheets("time")
i = 2
While workbooksdaji.Cells(2, i).Value <> ""
    i = i + 1
Wend
ntime = workbooksdaji.Cells(3, i - 1).Value
stime = workbooksdaji.Cells(4, i - 1).Value
ActiveWorkbook.Close savechanges:=False

'处理批量格式
htime = Split(ntime, "小时")
mtime = Split(htime(1), "分")
ntime = htime(0) & ":" & mtime(0)
htime = Split(stime, "小时")
mtime = Split(htime(1), "分")
stime = htime(0) & ":" & mtime(0)


For j = 4 To 34
   For k = 2 To 14
     ThisWorkbook.Sheets(1).Cells(j, k) = ThisWorkbook.Sheets(1).Cells(j, k + 1)
   Next
Next

Dim nowtime As Date
nowtime = ThisWorkbook.Sheets("数据").Cells(4, 15) + 1
ThisWorkbook.Sheets("图").Cells(1, 1) = "daily （" & Year(nowtime) & "年" & Month(nowtime) & "月" & Day(nowtime) & "日)"

ThisWorkbook.Sheets("数据").Cells(4, 15) = ThisWorkbook.Sheets(1).Cells(4, 15) + 1
ThisWorkbook.Sheets("数据").Cells(5, 15) = alarmnum
ThisWorkbook.Sheets("数据").Cells(6, 15) = ""
ThisWorkbook.Sheets("数据").Cells(7, 15) = ""

ThisWorkbook.Sheets("数据").Cells(8, 15) = ThisWorkbook.Sheets(1).Cells(8, 15) + 1
ThisWorkbook.Sheets("数据").Cells(9, 15) = changechanggui
ThisWorkbook.Sheets("数据").Cells(10, 15) = changebiaozhun
ThisWorkbook.Sheets("数据").Cells(11, 15) = changeyingji
ThisWorkbook.Sheets("数据").Cells(12, 15) = changezhongda

ThisWorkbook.Sheets("数据").Cells(13, 15) = ThisWorkbook.Sheets(1).Cells(13, 15) + 1
ThisWorkbook.Sheets("数据").Cells(14, 15) = ntime
ThisWorkbook.Sheets("数据").Cells(15, 15) = stime

ThisWorkbook.Sheets("数据").Cells(17, 15) = ThisWorkbook.Sheets(1).Cells(17, 15) + 1
ThisWorkbook.Sheets("数据").Cells(18, 15) = ""

ThisWorkbook.Sheets("图").Cells(5, 2).Value = ""
ThisWorkbook.Sheets("图").Cells(6, 2).Value = ""


MsgBox "生成成功"

End Sub