Sub Daily()
'
' Daily
'

Dim today, yesterday As Date
Dim dailyFilename(10)
Dim yesterdayMonth, yesterdayDay
Dim gaojing, yingji, shijian, biangeng
Dim changeChanggui, changeBiaozhun, changeYingji, changeZhongda

dailyFilename(0) = ActiveDocument.Path & "\index.png"
dailyFilename(1) = ActiveDocument.Path & "\team1.docx"
dailyFilename(2) = ActiveDocument.Path & "\team2.docx"
dailyFilename(3) = ActiveDocument.Path & "\team3.docx"
dailyFilename(4) = ActiveDocument.Path & "\team4.docx"
dailyFilename(5) = ActiveDocument.Path & "\team5.docx"
dailyFilename(6) = ActiveDocument.Path & "\team6.docx"
dailyFilename(7) = ActiveDocument.Path & "\team7.docx"
dailyFilename(8) = ActiveDocument.Path & "\team8.docx"

'check
For i = 0 To 7
    If Dir(dailyFilename(i)) = Empty Then
        MsgBox dailyFilename(i) & "不存在"
        Exit Sub
    End If
Next


'update
ActiveDocument.Paragraphs(2).Range.Delete '删除第2段
ActiveDocument.Paragraphs(2).Range.Select
today = Date
yesterday = Date - 1
Selection.TypeText Text:="--" & Year(yesterday) & "年" & Month(yesterday) & "月" & Day(yesterday) & "日" & Chr(13)

'获文件名
dailyFilename(9) = ActiveDocument.Path & "\"
If 10 > Month(yesterday) Then
    dailyFilename(9) = dailyFilename(9) & "0"
End If
dailyFilename(9) = dailyFilename(9) & Month(yesterday)
If 10 > Day(yesterday) Then
    dailyFilename(9) = dailyFilename(9) & "0"
End If
dailyFilename(9) = dailyFilename(9) & Day(yesterday) & "index.xlsm"

If Dir(dailyFilename(9)) = Empty Then
    MsgBox dailyFilename(9) & "不存在"
    Exit Sub
End If

Dim appExcel As New Excel.Application
Dim wbExcelYxzb, wbExcelGaojing As Excel.Workbook
Set wbExcelYxzb = appExcel.Workbooks.Open(dailyFilename(9))


wbExcelYxzb.Worksheets("alarm").Range("A1", "O" & (gaojing + 1)).Copy
Set wbExcelGaojing = appExcel.Workbooks.Add
wbExcelGaojing.ActiveSheet.Paste 
'appExcel.CutCopyMode = xlCopy
appExcel.ActiveSheet.Name = "alarm details"

wbExcelGaojing.Sheets.Add '创建透视图
appExcel.ActiveSheet.Name = "透视图"
wbExcelGaojing.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="alarm details!R2C1:R10C15").CreatePivotTable _
    TableDestination:="透视图!R3C1", TableName:="数据透视表1"
'wbExcelGaojing.Worksheets("Sheet2").Name = "透视图"
With wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").PivotFields("alarm type")
    .Orientation = xlRowField
    .Position = 1
End With
With wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").PivotFields("团队")
    .Orientation = xlColumnField
    .Position = 1
End With
wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").AddDataField wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").PivotFields("alarm概述"), "计数项:alarm概述", xlCount
With ActiveSheet.PivotTables("数据透视表1").PivotFields("alarm type")
    .PivotItems("(blank)").Visible = False
End With

wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").NullString = "0" '透视图中内容为空时填写0

ActiveWorkbook.SaveAs FileName:=ActiveDocument.Path & "\alarm统计表.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False '另存为文件

appExcel.Quit

Exit Sub

'生成整体情况
Dim strWholeOperationMsg
strWholeOperationMsg = Month(yesterday) & "月" & Day(yesterday) & "日8:00至" & Month(today) & "月" & Day(today) & "日8:00，"
If 0 = yingji Then
    strWholeOperationMsg = strWholeOperationMsg & "无"
Else
    strWholeOperationMsg = strWholeOperationMsg & yingji & "个"
End If
strWholeOperationMsg = strWholeOperationMsg & "应急事件，"

If 0 = shijian Then
    strWholeOperationMsg = strWholeOperationMsg & "无"
Else
    strWholeOperationMsg = strWholeOperationMsg & shijian & "张"
End If
strWholeOperationMsg = strWholeOperationMsg & "相关事件单，共"

If 0 = biangeng Then
    strWholeOperationMsg = strWholeOperationMsg & "无"
Else
    strWholeOperationMsg = strWholeOperationMsg & biangeng & "个"
End If
strWholeOperationMsg = strWholeOperationMsg & "change"

If 0 <> changeBiaozhun Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeBiaozhun & "个标准change"
End If

If 0 <> changeChanggui Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeChanggui & "个常规change"
End If

If 0 <> changeYingji Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeYingji & "个应急change"
End If

If 0 <> changeZhongda Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeZhongda & "个重大change"
End If

strWholeOperationMsg = strWholeOperationMsg & "。" & Chr(13)


ActiveDocument.Paragraphs(5).Range.Select
ActiveDocument.Paragraphs(5).Range.Delete '删除第5段
Selection.TypeText Text:=strWholeOperationMsg
ActiveDocument.Paragraphs(5).Range.Select
Selection.Style = ActiveDocument.Styles("正文")
Selection.Font.Name = "宋体"
Selection.Font.Size = 11

ActiveDocument.Tables.Item(1).Select '选择change表格
Selection.MoveUp Unit:=wdLine, Count:=1 '退回至上一行
Selection.HomeKey Unit:=wdLine '选择整行
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
Selection.Delete
Selection.TypeText Text:="alarm统计（共" & gaojing & "条）"

ActiveDocument.Tables.Item(2).Select '选择change表格
Selection.MoveUp Unit:=wdLine, Count:=1 '退回至上一行
Selection.HomeKey Unit:=wdLine '选择整行
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
Selection.Delete
Selection.TypeText Text:="change统计（共" & biangeng & "条）"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=17
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddPicture FileName:=dailyFilename(0), LinkToFile:=False, SaveWithDocument:=True

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=15
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(1), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team7.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=13
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(2), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team0.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=11
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(3), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team1.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=9
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(4), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team2.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=7
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(5), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team3.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=5
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(6), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team4.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=3
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(7), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team5.docx"

Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=1
Selection.Delete Unit:=wdCharacter, Count:=1
If Dir(dailyFilename(8)) = Empty Then
    Selection.TypeText Text:="无"
Else
    Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(8), LinkToFile:=False, DisplayAsIcon:=True, _
        IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="team6.docx"
End If


MsgBox "生成成功"


End Sub