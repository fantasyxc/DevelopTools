' 刷新工作簿中的所有数据透视表
Sub 刷新数据()
Workbooks("(20131231)Test.xlsm").RefreshAll
End Sub


' 刷新特意对应的数据透视表，注意数据透视表的名称一定要对应
Sub 刷新数据()
Workbooks("(20131231)Test.xlsm").Worksheets("透视").PivotTables("数据透视表1").PivotCache.Refresh
End Sub

