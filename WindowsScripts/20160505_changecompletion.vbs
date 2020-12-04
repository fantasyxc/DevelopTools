'完成情况统计
Sub Statistics()
 Dim ws As Worksheet
 Dim i As Integer
 Dim jinxingzhong As Integer
 Dim yiwancheng As Integer
 Dim chixuduori As Integer
 Dim daishishi As Integer
 Dim quxiao As Integer
 Dim weizhi As Integer
 Dim yingji As Integer
 Dim biaozhun As Integer
 Dim changgui As Integer
 Dim zhongda As Integer
 Dim biangengzongshu As Integer
 Dim huizong As String

 Set ws = ThisWorkbook.Worksheets(1)
 i = 2
 jinxingzhong = 0
 yiwancheng = 0
 chixuduori = 0
 daishishi = 0
 quxiao = 0
 
 yingji = 0
 biaozhun = 0
 changgui = 0
 biangengzongshu = 0
 zhongda = 0
 weizhi = 0

 While i <> (ws.UsedRange.Rows.Count + 1)
 If ws.Cells(i, 2).Value Like "C*" Then
 biangengzongshu = biangengzongshu + 1

 If ws.Cells(i, 5).Value Like "*应急*" Then
 yingji = yingji + 1
 ElseIf ws.Cells(i, 5).Value Like "*常规*" Then
 changgui = changgui + 1
 ElseIf ws.Cells(i, 5).Value Like "*标准*" Then
 biaozhun = biaozhun + 1
 ElseIf ws.Cells(i, 5).Value Like "*重大*" Then
 zhongda = zhongda + 1
 End If

 If ws.Cells(i, 1).Value Like "*进行*" Then
 jinxingzhong = jinxingzhong + 1
 ElseIf ws.Cells(i, 1).Value Like "*完成*" Then
 yiwancheng = yiwancheng + 1
 ElseIf ws.Cells(i, 1).Value Like "*持续多日*" Then
 chixuduori = chixuduori + 1
 ElseIf ws.Cells(i, 1).Value Like "*取消*" Then
 quxiao = quxiao + 1
 ElseIf ws.Cells(i, 1).Value Like "" Then
 daishishi = daishishi + 1
 Else
 weizhi = weizhi + 1
 End If


 End If
 i = i + 1
 Wend

 '变更情况
 huizong = "今日共" & biangengzongshu & "个变更，其中"
 If 0 <> yingji Then
 huizong = huizong & "应急变更" & yingji & "个（）"
 End If
 If 0 <> changgui Then
 huizong = huizong & "，常规变更" & changgui & "个"
 End If
 If 0 <> biaozhun Then
 huizong = huizong & "，标准变更" & biaozhun & "个，"
 End If
 If 0 <> zhongda Then
 huizong = huizong & "，重大变更" & zhongda & "个"
 End If

 '完成情况
 huizong = huizong & yiwancheng & "个已完成，" & chixuduori & "个持续多日，"& jinxingzhong & "个正在实施，" & daishishi & "个尚未开始，"  & quxiao & "个取消。"

 ws.Cells(biangengzongshu + 3, 1).Value = huizong
End Sub