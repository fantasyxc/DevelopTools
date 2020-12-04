'
'功能：生成日报
'2016-09-09 变更统计页面格式调整
'2016-09-10 变更时间和告警时间判断
'2017-08-12 更新
'

Private Sub creat()

Dim i ' 临时变量

'检查文件是否存在
Dim myFilename(3)
myFilename(0) = ThisWorkbook.Path & "\告警统计.xls"
myFilename(1) = ThisWorkbook.Path & "\变更统计.xls"
myFilename(2) = ThisWorkbook.Path & "\大机一线.xlsm"
For i = 0 To 2
If Dir(myFilename(i)) = Empty Then
   MsgBox myFilename(i)
   Exit Sub
 End If
Next


'''''''''''''''''''''''''''''''
''''''’''''处理告警''''''''''''
'''''''''''''''''''''''''''''''
Dim alarmnum
alarmnum = 0
Set workbooksalarm = Workbooks.Open(Filename:=myFilename(0), ReadOnly:=True).Worksheets(1)

'清空告警统计表中原告警信息
i = 2
While ThisWorkbook.Sheets("告警统计").Cells(i, 2).Value <> ""
    ThisWorkbook.Sheets("告警统计").Rows(i).EntireRow.Delete Shift:=xlShiftUp
Wend

'读取文件内容
i = 2
While workbooksalarm.Cells(i, 2).Value <> ""
	workbooksalarm.Cells(i, 1).EntireRow.Copy ThisWorkbook.Sheets("告警统计").Cells(alarmnum + 2, 1)
	alarmnum = alarmnum + 1
	ThisWorkbook.Sheets("告警统计").Cells(alarmnum + 1, 1).Value = alarmnum
    i = i + 1
Wend
ActiveWorkbook.Close True 


'''''''''''''''''''''''''''''''
''''''’''''处理变更''''''''''''
'''''''''''''''''''''''''''''''
Dim changezhongda, changeyingji, changechanggui, changebiaozhun
changezhongda = 0
changeyingji = 0
changechanggui = 0
changebiaozhun = 0

'清空变更统计表中原变更信息
i = 2
While ThisWorkbook.Sheets("变更统计").Cells(i, 1).Value <> ""
    ThisWorkbook.Sheets("变更统计").Rows(i).EntireRow.Delete Shift:=xlShiftUp
Wend

Set workbookschange = Workbooks.Open(Filename:=myFilename(1), ReadOnly:=True).Worksheets(1)
i = 2
While workbookschange.Cells(i, 1).Value <> ""
    workbookschange.Cells(i, 1).EntireRow.Copy ThisWorkbook.Sheets("变更统计").Cells(changezhongda + changeyingji + changechanggui + changebiaozhun + 2, 1)
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
''''''读取大机一线数据'''''''''
'''''''''''''''''''''''''''''''
pathdaji = ThisWorkbook.Path & "\大机一线.xlsm" '把路径赋值给一个字符串
Set workbooksdaji = Workbooks.Open(Filename:=pathdaji, ReadOnly:=True).Worksheets("批量时间统计")
'批量
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


'''''''''''''''''''''''''''''''
''''''更新运行指标文件'''''''''
'''''''''''''''''''''''''''''''
'将后一列粘贴到前一列
For j = 4 To 34
   For k = 2 To 14
     ThisWorkbook.Sheets(1).Cells(j, k) = ThisWorkbook.Sheets(1).Cells(j, k + 1)
   Next
Next

'标题
Dim nowtime As Date
nowtime = ThisWorkbook.Sheets("数据").Cells(4, 15) + 1
ThisWorkbook.Sheets("图").Cells(1, 1) = "武汉数据中心运行日报 （" & Year(nowtime) & "年" & Month(nowtime) & "月" & Day(nowtime) & "日)"

'告警
ThisWorkbook.Sheets("数据").Cells(4, 15) = ThisWorkbook.Sheets(1).Cells(4, 15) + 1
ThisWorkbook.Sheets("数据").Cells(5, 15) = alarmnum
ThisWorkbook.Sheets("数据").Cells(6, 15) = ""
ThisWorkbook.Sheets("数据").Cells(7, 15) = ""

'变更
ThisWorkbook.Sheets("数据").Cells(8, 15) = ThisWorkbook.Sheets(1).Cells(8, 15) + 1
ThisWorkbook.Sheets("数据").Cells(9, 15) = changechanggui
ThisWorkbook.Sheets("数据").Cells(10, 15) = changebiaozhun
ThisWorkbook.Sheets("数据").Cells(11, 15) = changeyingji
ThisWorkbook.Sheets("数据").Cells(12, 15) = changezhongda

'批量时间
ThisWorkbook.Sheets("数据").Cells(13, 15) = ThisWorkbook.Sheets(1).Cells(13, 15) + 1
ThisWorkbook.Sheets("数据").Cells(14, 15) = ntime
ThisWorkbook.Sheets("数据").Cells(15, 15) = stime

'TD CPU
ThisWorkbook.Sheets("数据").Cells(17, 15) = ThisWorkbook.Sheets(1).Cells(17, 15) + 1
ThisWorkbook.Sheets("数据").Cells(18, 15) = ""

'清空原来图中的文字
ThisWorkbook.Sheets("图").Cells(5, 2).Value = ""
ThisWorkbook.Sheets("图").Cells(6, 2).Value = ""


MsgBox "日报生成成功"

End Sub