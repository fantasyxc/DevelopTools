Sub Daily()
'
' Daily
'
' 汇总生成日报
'

Dim today, yesterday As Date
Dim dailyFilename(10)
Dim yesterdayMonth, yesterdayDay
Dim gaojing, yingji, shijian, biangeng
Dim changeChanggui, changeBiaozhun, changeYingji, changeZhongda

dailyFilename(0) = ActiveDocument.Path & "\运行指标.png"
dailyFilename(1) = ActiveDocument.Path & "\服务台日报.docx"
dailyFilename(2) = ActiveDocument.Path & "\运行一线日报.docx"
dailyFilename(3) = ActiveDocument.Path & "\大机一线日报.docx"
dailyFilename(4) = ActiveDocument.Path & "\平台日报.docx"
dailyFilename(5) = ActiveDocument.Path & "\环境日报.docx"
dailyFilename(6) = ActiveDocument.Path & "\网络日报.docx"
dailyFilename(7) = ActiveDocument.Path & "\基础设施日报.docx"
dailyFilename(8) = ActiveDocument.Path & "\风险日报.docx"

'检查日报文件是否存在
For i = 0 To 7
    If Dir(dailyFilename(i)) = Empty Then
        MsgBox dailyFilename(i) & "不存在"
        Exit Sub
    End If
Next


'更新日期
ActiveDocument.Paragraphs(2).Range.Delete '删除第2段
ActiveDocument.Paragraphs(2).Range.Select
today = Date
yesterday = Date - 1
Selection.TypeText Text:="--" & Year(yesterday) & "年" & Month(yesterday) & "月" & Day(yesterday) & "日" & Chr(13)

'获取运行指标文件名
dailyFilename(9) = ActiveDocument.Path & "\"
If 10 > Month(yesterday) Then
    dailyFilename(9) = dailyFilename(9) & "0"
End If
dailyFilename(9) = dailyFilename(9) & Month(yesterday)
If 10 > Day(yesterday) Then
    dailyFilename(9) = dailyFilename(9) & "0"
End If
dailyFilename(9) = dailyFilename(9) & Day(yesterday) & "运行指标.xlsm"

If Dir(dailyFilename(9)) = Empty Then
    MsgBox dailyFilename(9) & "不存在"
    Exit Sub
End If

'读取运行指标文档信息
Dim appExcel As New Excel.Application
Dim wbExcelYxzb, wbExcelGaojing As Excel.Workbook
Set wbExcelYxzb = appExcel.Workbooks.Open(dailyFilename(9))
'gaojing = wbExcelYxzb.Worksheets("数据").Cells(5, 15)
'yingji = wbExcelYxzb.Worksheets("数据").Cells(6, 15)
'shijian = wbExcelYxzb.Worksheets("数据").Cells(7, 15)
'changeChanggui = wbExcelYxzb.Worksheets("数据").Cells(9, 15)
'changeBiaozhun = wbExcelYxzb.Worksheets("数据").Cells(10, 15)
'changeYingji = wbExcelYxzb.Worksheets("数据").Cells(11, 15)
'changeZhongda = wbExcelYxzb.Worksheets("数据").Cells(12, 15)
'biangeng = changeBiaozhun + changeChangui + changeYingji + changeZhongda


'更新变更
'wbExcelYxzb.Worksheets("变更统计").Range("A1", "J" & (biangeng + 1)).Copy
'ActiveDocument.Tables.Item(2).Select
'ActiveDocument.Tables.Item(2).Delete '删除变更表格
'Selection.PasteExcelTable False, False, False
'ActiveDocument.Tables.Item(2).AutoFitBehavior (wdAutoFitWindow) '自动匹配窗格
'ActiveDocument.Tables.Item(2).Select
'Selection.Font.Shrink '字体缩小


'更新告警信息
wbExcelYxzb.Worksheets("告警统计").Range("A1", "O" & (gaojing + 1)).Copy
Set wbExcelGaojing = appExcel.Workbooks.Add
wbExcelGaojing.ActiveSheet.Paste '粘贴告警
'appExcel.CutCopyMode = xlCopy
appExcel.ActiveSheet.Name = "告警明细"

wbExcelGaojing.Sheets.Add '创建透视图
appExcel.ActiveSheet.Name = "透视图"
wbExcelGaojing.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="告警明细!R2C1:R10C15").CreatePivotTable _
    TableDestination:="透视图!R3C1", TableName:="数据透视表1"
'wbExcelGaojing.Worksheets("Sheet2").Name = "透视图"
With wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").PivotFields("告警分类")
    .Orientation = xlRowField
    .Position = 1
End With
With wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").PivotFields("团队")
    .Orientation = xlColumnField
    .Position = 1
End With
wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").AddDataField wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").PivotFields("告警概述"), "计数项:告警概述", xlCount
With ActiveSheet.PivotTables("数据透视表1").PivotFields("告警分类")
    .PivotItems("(blank)").Visible = False
End With

wbExcelGaojing.Worksheets("透视图").PivotTables("数据透视表2").NullString = "0" '透视图中内容为空时填写0

ActiveWorkbook.SaveAs FileName:=ActiveDocument.Path & "\告警统计表.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False '另存为文件

appExcel.Quit

Exit Sub

'生成整体运行情况
Dim strWholeOperationMsg
strWholeOperationMsg = Month(yesterday) & "月" & Day(yesterday) & "日8:00至" & Month(today) & "月" & Day(today) & "日8:00，"
If 0 = yingji Then
    strWholeOperationMsg = strWholeOperationMsg & "无"
Else
    strWholeOperationMsg = strWholeOperationMsg & yingji & "个"
End If
strWholeOperationMsg = strWholeOperationMsg & "南湖运行应急事件，"

If 0 = shijian Then
    strWholeOperationMsg = strWholeOperationMsg & "无"
Else
    strWholeOperationMsg = strWholeOperationMsg & shijian & "张"
End If
strWholeOperationMsg = strWholeOperationMsg & "南湖相关事件单，共"

If 0 = biangeng Then
    strWholeOperationMsg = strWholeOperationMsg & "无"
Else
    strWholeOperationMsg = strWholeOperationMsg & biangeng & "个"
End If
strWholeOperationMsg = strWholeOperationMsg & "变更"

If 0 <> changeBiaozhun Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeBiaozhun & "个标准变更"
End If

If 0 <> changeChanggui Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeChanggui & "个常规变更"
End If

If 0 <> changeYingji Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeYingji & "个应急变更"
End If

If 0 <> changeZhongda Then
    strWholeOperationMsg = strWholeOperationMsg & "，" & changeZhongda & "个重大变更"
End If

strWholeOperationMsg = strWholeOperationMsg & "。" & Chr(13)


'更新整体运行情况
ActiveDocument.Paragraphs(5).Range.Select
ActiveDocument.Paragraphs(5).Range.Delete '删除第5段
Selection.TypeText Text:=strWholeOperationMsg
ActiveDocument.Paragraphs(5).Range.Select
Selection.Style = ActiveDocument.Styles("正文")
Selection.Font.Name = "宋体"
Selection.Font.Size = 11

'更新告警统计标题
ActiveDocument.Tables.Item(1).Select '选择变更表格
Selection.MoveUp Unit:=wdLine, Count:=1 '退回至上一行
Selection.HomeKey Unit:=wdLine '选择整行
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
Selection.Delete
Selection.TypeText Text:="告警统计（共" & gaojing & "条）"

'更新变更统计标题
ActiveDocument.Tables.Item(2).Select '选择变更表格
Selection.MoveUp Unit:=wdLine, Count:=1 '退回至上一行
Selection.HomeKey Unit:=wdLine '选择整行
Selection.EndKey Unit:=wdLine, Extend:=wdExtend
Selection.Delete
Selection.TypeText Text:="变更统计（共" & biangeng & "条）"

'插入运行指标
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=17
Selection.Delete Unit:=wdCharacter, Count:=1
'Selection.TypeBackspace
Selection.InlineShapes.AddPicture FileName:=dailyFilename(0), LinkToFile:=False, SaveWithDocument:=True

'更新服务台日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=15
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(1), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="服务台日报.docx"

'更新运行一线日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=13
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(2), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="运行一线日报.docx"

'更新大机一线日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=11
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(3), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="大机一线日报.docx"

'更新平台日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=9
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(4), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="平台日报.docx"

'更新环境日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=7
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(5), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="环境日报.docx"

'更新网络日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=5
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(6), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="网络日报.docx"

'更新基础设施日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=3
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(7), LinkToFile:=False, DisplayAsIcon:=True, _
    IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="基础设施日报.docx"

'更新风险日报
Selection.EndKey Unit:=wdStory, Extend:=wdExtendEnd
Selection.MoveUp Unit:=wdLine, Count:=1
Selection.Delete Unit:=wdCharacter, Count:=1
If Dir(dailyFilename(8)) = Empty Then '不存在风险日报则显示无
    Selection.TypeText Text:="无"
Else '存在风险日报则导入
    Selection.InlineShapes.AddOLEObject ClassType:="Word.Document.12", FileName:=dailyFilename(8), LinkToFile:=False, DisplayAsIcon:=True, _
        IconFileName:="C:\Windows\Installer\{90150000-0011-0000-1000-0000000FF1CE}\wordicon.exe", IconIndex:=13, IconLabel:="风险日报.docx"
End If


'生成成功提示
MsgBox "日报生成成功"


End Sub