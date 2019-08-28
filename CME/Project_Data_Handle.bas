Sub Project_Data_Handle()
    Application.ScreenUpdating = False
    Dim i%, Src_Wkb As Workbook, Stat_Period$, SumNumb%

    Stat_Period = ""  'TODO:设置获取当前日期

    ' TODO:数值可能不统一，检查数据的统一性与真实性


    '------------------- 切换工作表 -------------------
    For i = 1 To Workbooks.Count
        If Not Workbooks(i).Name like "*personal*" Then Workbooks(i).Activate       
    Next
    If ActiveWorkbook.Name like "*personal*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If


'------------------- 【学习人数汇总】 -------------------
    Sheets("学习人数汇总").Activate
    For i = 3 To [A1048576].End(xlUp).Row - 1
        If Cells(i,1) <> "" Then Cells(i,2) = Sheets("学习基本情况").Cells(i,6)
    Next 
    Cells([A1048576].End(xlUp).Row,2) = Application.WorksheetFunction.Sum(Range([B3],Cells([A1048576].End(xlUp).Row-1,2)))
    SumNumb = Cells([A1048576].End(xlUp).Row,2)

'------------------- 【学习基本情况】 -------------------
'说明：总计 = 学习中 + 未申请 + 已申请，已获取学分数 ≤ 已申请
    Sheets("学习基本情况").Activate
    For i = 3 To [A1048576].End(xlUp).Row 
        If Cells(i,1) Like "*年*" Then
            If Cells(i,6)<> Cells(i,2) + Cells(i,3) + Cells(i,5) Then Msgbox "第 " & i & " 行 #总计数据# 有差异！" 
            If Cells(i,4) > Cells(i,5) Then Msgbox "第 " & i & " 行 #已获取学分数# 超出已申请数！" 
        End if
        If Cells(i,1) = "总计" Then Cells(i,6) = Application.WorksheetFunction.Sum(Range([F3],Cells(i-1,6)))
    Next


'------------------- 【专业分析】 -------------------
'TODO:按照专业数量进行排序
'TODO:检查是否有空值
    Dim TempOther%
    Sheets("专业分析").Activate
    ' Call BeautifySht
    With Range([D1],Cells([E1048576].End(xlUp).Row,4))
        .Replace " ",""
        .Replace "NULL","其他"
        .Replace "-请选择-","其他"
        .Replace "无职称","其他"
        .Replace "","其他"
    End With
    TempOther = Application.WorksheetFunction.Sumif([D:D],"其他",[E:E])
    Range([D1],Cells([E1048576].End(xlUp).Row,5)).Copy
    [A3].Insert Shift:=xlDown
    [D:E].ClearContents
    Cells([A1048576].End(xlUp).Row - 1,1) = "其他"
    Cells([A1048576].End(xlUp).Row - 1,2) = TempOther
    Call Delete_BlankRows
    For i = [A1048576].End(xlUp).Row - 2 To 3 Step -1
        If Cells(i,1) = "其他" Then Rows(i).EntireRow.Delete
    Next
    Cells([A1048576].End(xlUp).Row,2) = Application.WorksheetFunction.Sum(Range([B3],Cells([A1048576].End(xlUp).Row-1,2)))
    If Cells([A1048576].End(xlUp).Row,2) <> SumNumb Then Msgbox "表格 " & ActiveSheet.Name & " 总计数据存在差异！"

'------------------- 【职称分析】 -------------------
'TODO:按照职称级别进行排序
'TODO:检查是否有空值
    Sheets("职称分析").Activate
    ' Call BeautifySht
    With Range([D1],Cells([E1048576].End(xlUp).Row,4))
        .Replace " ",""
        .Replace "NULL","TEMP"
        .Replace "-请选择-","TEMP"
        .Replace "无职称","TEMP"
        .Replace "","TEMP"
        .Replace "其他","TEMP"
    End With
    Cells([E1048576].End(xlUp).Row + 1,4)="其他"
    Cells([E1048576].End(xlUp).Row + 1,5)=Application.WorksheetFunction.Sumif([D:D],"TEMP",[E:E])
    Range([D1],Cells([E1048576].End(xlUp).Row,5)).Copy
    [A3].Insert Shift:=xlDown
    [D:E].ClearContents
    Call Delete_BlankRows
    For i = [A1048576].End(xlUp).Row To 3 Step -1
        If Cells(i,1) = "TEMP" Then Rows(i).EntireRow.Delete
    Next
    Cells([A1048576].End(xlUp).Row,2) = Application.WorksheetFunction.Sum(Range([B3],Cells([A1048576].End(xlUp).Row-1,2)))
    If Cells([A1048576].End(xlUp).Row,2) <> SumNumb Then Msgbox "表格 " & ActiveSheet.Name & " 总计数据存在差异！"

'------------------- 【省市分布分析】 -------------------
'TODO:按照数量多少进行排序
'TODO:检查是否有空值
    Sheets("省市分布分析").Activate
    ' Call BeautifySht
    Columns(1).ColumnWidth = 20
    Columns(2).ColumnWidth = 20
    Columns(3).ColumnWidth = 28
    Range([E1],Cells([G1048576].End(xlUp).Row,7)).Copy
    [A3].Insert Shift:=xlDown
    [E:G].ClearContents
    Call Delete_BlankRows
    Cells([A1048576].End(xlUp).Row,3) = Application.WorksheetFunction.Sum(Range([C3],Cells([A1048576].End(xlUp).Row-1,3)))
    If Cells([A1048576].End(xlUp).Row,3) <> SumNumb Then Msgbox "表格 " & ActiveSheet.Name & " 总计数据存在差异！"
'TODO:进行职称的数据统计

'------------------- 【医院等级分析】 -------------------
    Sheets("医院等级分析").Activate
    ' Call BeautifySht
    With Range([D1],Cells([E1048576].End(xlUp).Row,4))
        .Replace "NULL","其他"
        .Replace "-请选择-","其他"
        .Replace "","其他"
        .Replace "一甲","一级甲等"
        .Replace "一乙","一级乙等"
        .Replace "二甲","二级甲等"
        .Replace "二乙","二级乙等"
        .Replace "三甲","二级甲等"
        .Replace "三乙","三级乙等"
    End With
    For i = 3 To 9
        Cells(i,2) = Application.WorksheetFunction.Sumif([D:D],Cells(i,1),[E:E])
    Next
    [B10] = Application.WorksheetFunction.Sum([B3:B9])
    [D:E].ClearContents
    If [B10] <> SumNumb Then Msgbox "表格 " & ActiveSheet.Name & " 总计数据存在差异！"



    ActiveWorkbook.Save
    Msgbox "项目数据已经统计完成！"

    'TODO:后续需要手工完成的有以下步骤
End Sub

Sub BeautifySht()
    With ActiveSheet.Selection
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone

        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Sub Delete_BlankRows()
    Dim i%
    For i = [A1048576].End(xlUp).Row To 3 Step -1
        If WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).EntireRow.Delete
    Next
End Sub


'TODO:有的表格没有美化成功
' 没有其他的时候不要插入
如果只有学习基本情况的表格和其他的不同，那么不报错，直接更改学习基本情况的表格数据，同时需要保证已获取学分的数量是小于已经申请的