'【功能】分析各城市的辉瑞问卷数据
'【准备工作】医院、职称、医院级别数据清洗；多选的选项整合与替换
Sub Questionnaire_Data_Analyse()
    Application.ScreenUpdating = False
    Dim cityname$, LastRow&, i&, j&, k&, Src_Wkb As Workbook, Dst_Wkb As Workbook, Src_sht As Worksheet, citynumb&
    Dim likeStr$
    
    '---------------- 切换数据表 --------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name Like "*问卷回答情况分析*" Then Workbooks(i).Activate
    Next
    If Not ActiveWorkbook.Name Like "*问卷回答情况分析*" Then
        MsgBox "Cannot find the workbook!"
        Exit Sub
    End If
    cityname = Left(ActiveWorkbook.Name, 2)
    Set Src_Wkb = Workbooks("辉瑞问卷-DataTool.xlsm")
    Set Dst_Wkb = Workbooks(ActiveWorkbook.Name)
    Set Src_sht = Src_Wkb.Sheets(cityname)
    citynumb = Src_sht.[A1048576].End(xlUp).Row - 1

    '---------------- 第四题 --------------
    Dst_Wkb.Sheets("4").Activate
    For i = 5 To 7
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[F:F], i - 4)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[F:F], i - 4)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(8,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(7, k)))
    Next k

    [B2].select
    If [I8] <> citynumb Or [P8] <> citynumb Then
        Msgbox "数据 4 出现问题！"
    Else
        Msgbox "4 成功！"
    End If


    '---------------- 第五题 --------------
    Dst_Wkb.Sheets("5").Activate
    LastRow = [B1048576].End(xlUp).Row
    For i = 5 To 12
        likeStr = "*" & Mid(Cells(i,2),3,1) & "*"
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[O:O], likeStr)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[O:O], likeStr)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(13,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(12, k)))
    Next k

    For i = 15 To LastRow - 1 
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[O:O], cells(i,2))
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[O:O], cells(i,2))
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i


    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(LastRow,k) = Application.WorksheetFunction.Sum(Range(Cells(15, k), Cells(LastRow - 1, k)))
    Next k

    Application.DisplayAlerts = False
    For i = LastRow To 15 Step -1
        If Cells(i,9) = 0 Then Rows(i).Delete
    Next i
    Application.DisplayAlerts = True

    [B2].select

    If [I13] <>  [P13] Then
        Msgbox "数据 5 出现问题！"
    Else
        Msgbox "5 成功！"
    End If


    '---------------- 第六题 --------------
    Dst_Wkb.Sheets("6").Activate
    For i = 5 To 7
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[P:P], i - 4)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[P:P], i - 4)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(8,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(7, k)))
    Next k

    [B2].select

    If [I8] <> citynumb Or [P8] <> citynumb Then
        Msgbox "数据 6 出现问题！"
    Else
        Msgbox "6 成功！"
    End If


    '---------------- 第七题 --------------
    Dst_Wkb.Sheets("7").Activate
    LastRow = [B1048576].End(xlUp).Row
    For i = 5 To 12
        likeStr = "*" & Mid(Cells(i,2),3,1) & "*"
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[Y:Y], likeStr)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[Y:Y], likeStr)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(13,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(12, k)))
    Next k

    For i = 15 To LastRow - 1 
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[Y:Y], cells(i,2))
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[Y:Y], cells(i,2))
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i


    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(LastRow,k) = Application.WorksheetFunction.Sum(Range(Cells(15, k), Cells(LastRow - 1, k)))
    Next k

    Application.DisplayAlerts = False
    For i = LastRow To 15 Step -1
        If Cells(i,9) = 0 Then Rows(i).Delete
    Next i
    Application.DisplayAlerts = True

    [B2].select

    If [I13] <>  [P13] Then
        Msgbox "数据 7 出现问题！"
    Else
        Msgbox "7 成功！"
    End If

'---------------- 第八题 --------------
    Dst_Wkb.Sheets("8").Activate
    LastRow = [B1048576].End(xlUp).Row
    For i = 5 To 11
        likeStr = "*" & Mid(Cells(i,2),3,1) & "*"
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[AG:AG], likeStr)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[AG:AG], likeStr)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(12,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(11, k)))
    Next k

    For i = 14 To LastRow - 1 
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[AG:AG], cells(i,2))
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[AG:AG], cells(i,2))
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i


    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(LastRow,k) = Application.WorksheetFunction.Sum(Range(Cells(14, k), Cells(LastRow - 1, k)))
    Next k

    Application.DisplayAlerts = False
    For i = LastRow To 14 Step -1
        If Cells(i,9) = 0 Then Rows(i).Delete
    Next i
    Application.DisplayAlerts = True

    [B2].select

    If [I12] <>  [P12] Then
        Msgbox "数据 8 出现问题！"
    Else
        Msgbox "8 成功！"
    End If


    '---------------- 第九题 --------------
    Dst_Wkb.Sheets("9").Activate
    For i = 5 To 8
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[AH:AH], i - 4)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[AH:AH], i - 4)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(9,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(8, k)))
    Next k

    [B2].select

    If [I9] <> citynumb Or [P9] <> citynumb Then
        Msgbox "数据 9 出现问题！"
    Else
        Msgbox "9 成功！"
    End If

'---------------- 第十题 --------------
    Dst_Wkb.Sheets("10").Activate
    LastRow = [B1048576].End(xlUp).Row
    For i = 5 To 8
        likeStr = "*" & Mid(Cells(i,2),3,1) & "*"
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[AM:AM], likeStr)
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[AM:AM], likeStr)
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i

    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(9,k) = Application.WorksheetFunction.Sum(Range(Cells(5, k), Cells(8, k)))
    Next k

    For i = 11 To LastRow - 1 
        For j = 3 To 8
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[E:E], Cells(4, j), Src_sht.[AM:AM], cells(i,2))
        Next j

        Cells(i, 9) = Application.WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 8)))

        For j = 12 To 15
            Cells(i, j) = Application.WorksheetFunction.CountIfs(Src_sht.[C:C], Cells(4, j), Src_sht.[AM:AM], cells(i,2))
        Next j
        Cells(i, 16) = Application.WorksheetFunction.Sum(Range(Cells(i, 12), Cells(i, 15)))
    Next i


    For k = 3 To 16 
        If k < 10 Or k > 11 Then Cells(LastRow,k) = Application.WorksheetFunction.Sum(Range(Cells(11, k), Cells(LastRow - 1, k)))
    Next k

    Application.DisplayAlerts = False
    For i = LastRow To 11 Step -1
        If Cells(i,9) = 0 Then Rows(i).Delete
    Next i
    Application.DisplayAlerts = True

    [B2].select

    If [I9] <>  [P9] Then
        Msgbox "数据 10 出现问题！"
    Else
        Msgbox "10 成功！"
    End If


    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Msgbox "数据已经处理完成！"
End Sub






' 需要提前处理的工作：
' [1] 仅保留目标城市数据，并将此表命名为 "Data"
' [2] 将以文本存储的数字转化为数值
Sub Data_Clean()

    With Sheets("Data").UsedRange
        .Replace " ",""
    
    End With


    LastRow = Sheets("Data").[A1048576].End(xlUp).Row
    '------------------- 删除没有用的数据 -------------------
    Columns(10).Delete '删除来源列
    Columns(9).Delete '删除ip列
    Columns(6).Delete '删除序号列

    '------------------- 处理来源端口数据 -------------------
    For i = 2 To LastRow
        If Left(Cells(i,8),2) = "pc" Then 
            Cells(i,8) = "PC"
        ElseIf Left(Cells(i,8),2) = "mo" Then
            Cells(i,8) = "Mobile"
        Else
            Cells(i,8) = "Others"
            DataError = True
        End If
    Next
    If DataError = True Then Msgbox "存在其他来源的问卷数据！"

    '------------------- 数据清洗 -------------------
    

    '------------------- 获取标准职称 -------------------

End Sub