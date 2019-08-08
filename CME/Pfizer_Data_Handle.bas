Sub Pfizer_Data_Handle()
    Application.ScreenUpdating = False

    Dim i%, Src_Wkb As Workbook, Dst_Wkb As Workbook
    Dim Temp_Dict As object
    Dim CellRng As Range, Temp_Rng As Range

    '---------------- 切换数据表 --------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*辉瑞*" Then Workbooks(i).Activate       
    Next
    If Not ActiveWorkbook.Name like "*辉瑞*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If
    Set Src_Wkb = Workbooks(ActiveWorkbook.Name)
    Set Dst_Wkb = Workbooks("Pfizer-DataTool.xlsm")

    '---------------- 数据清洗 --------------
    '清洗孙旭辰这个测试账号的信息
    Sheets("Sheet2").Activate
    RowNumbs = Sheets("Sheet2").[a1048576].End(xlUp).Row
    [C:C].Replace "(**)",""
    For i = 2 To RowNumbs
        If Cells(i,5) = "孙旭辰" Then 
            Cells(i,13) = "测试"
            Exit for
        End if
    Next

    '---------------- Sheet1 文本转数值 --------------
    With Sheets("Sheet1").Rows(2)
        .NumberFormatLocal = "G/通用格式"   '把单元格设置为常规
        .Value = .Value   '取值
    End With

    '---------------- 生成医生工作表 --------------
    ' 方法1：复制所有医生的行
    Sheets.Add(After:=Sheets(3)).Name = "DocData"
    Sheets("Sheet2").Activate
    ActiveSheet.UsedRange.AutoFilter Field:=13, Criteria1:="医生"
    Range([b2],Cells(RowNumbs,13)).Copy
    Sheets("Sheet2").AutoFilterMode = False
    Sheets("DocData").Select: [a1].Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    [a1].Select
    RowNumbs = Sheets("DocData").[a1048576].End(xlUp).Row

    ' 方法2：复制全表后，遍历删除（缺点，行数多的时候遍历特别慢）
    ' Sheets("Sheet1").Copy Before:=Sheets("Sheet1")
    ' Sheets("Sheet1 (2)").Name = "TEMP"
    ' Sheets("TEMP").Select
    ' Columns("N:P").Delete
    ' RowNumbs = Sheets("TEMP").[a1048576].End(xlUp).Row
    ' For i = RowNumbs to 2 Step-1
    '     If Cells(i,13) <> "医生" Then Rows(i).Delete
    ' Next


    '---------------- 一些数据统计 --------------
    Sheets("DocData").Activate
    Dim ZR_Numb%, FZR_Numb%, ZZ_Numb%, YS_Numb%
    ' 医生职称的统计
    For i = 1 To RowNumbs
        If Cells(i, 6) Like "*副*" Then
            FZR_Numb = FZR_Numb + 1
        ElseIf Cells(i, 6) Like "*主任*" Then
            ZR_Numb = ZR_Numb + 1
        ElseIf Cells(i, 6) Like "*主治*" Then
            ZZ_Numb = ZZ_Numb + 1
        Else
            YS_Numb = YS_Numb + 1
        End If
    Next
    ' 省份、城市数量的统计
    Dim PvcDict As Object, CityDict As Object, Last_PvcNumb%, Last_CityNumb%
    Set PvcDict = CreateObject("scripting.dictionary")
    Set CityDict = CreateObject("scripting.dictionary")
    For i = 1 To RowNumbs
        If Cells(i,1) <> "" And Not PvcDict.exists(Cells(i,1).Value) Then PvcDict(Cells(i,1).Value)= Cells(i,1).Value
        Cells(i,13) = Cells(i,1) & "-" & Cells(i,2)
        If Cells(i,13) <> "" And Not CityDict.exists(Cells(i,13).Value) Then CityDict(Cells(i,13).Value)= Cells(i,13).Value
    Next
    Last_PvcNumb = PvcDict.Count
    Last_CityNumb = CityDict.Count


    '---------------- 汇总表 统计 --------------
    ' Dst_Wkb.Activate
    With Dst_Wkb.Sheets("汇总")
        '更正添加日期、清空旧的差值
        .Columns(4).EntireColumn.Insert
        .[D3].Value = Format(Now,"yy/mm/dd")
        .[D10].Value = Format(Now,"yy/mm/dd")
        .[B2].Value ="学习状态人数统计-" & Format(Now,"yymmdd")
        .[B9].Value ="学习效果统计-" & Format(Now,"yymmdd")
        .[C4:C7].ClearContents
        .[C11:C13].ClearContents
        '数据统计:学习状态人数
        For i = 4 To 7
            If i < 7 Then .Cells(i,4)= Application.WorksheetFunction.CountIf(Src_Wkb.Sheets("DocData").[K:K], .Cells(i, 2))
            If i = 7 Then .[D7] = RowNumbs
            .Cells(i,3).Value = .Cells(i,4) - .Cells(i,5)
        Next
        '数据统计:学习效果
        For i = 11 To 13
            .Cells(i,4)= Src_Wkb.Sheets("Sheet1").Cells(2,i-10)
            .Cells(i,3).Value = .Cells(i,4) - .Cells(i,5)
        Next
        .[D:D].FormatConditions.Delete
    End With

    '---------------- 职称与医院分布表 统计 --------------
    With Dst_Wkb.Sheets("职称 | 医院分布")
        '添加日期、清空旧的差值
        .Columns(4).EntireColumn.Insert
        .[D2].Value = Format(Now,"yy/mm/dd")
        .[D9].Value = Format(Now,"yy/mm/dd")
        .[C3:C7].ClearContents
        .[C10:C16].ClearContents
        '数据统计:职称分布
        .[D3] = ZR_Numb
        .[D4] = FZR_Numb
        .[D5] = ZZ_Numb
        .[D6] = YS_Numb
        .[D7] = RowNumbs
        For i = 3 To 7
            .Cells(i,3) = .Cells(i,4) - .Cells(i,5)
        Next

        '数据统计:医院分布
        Dim UpNumb%, Sum_Temp%, Rnd_Arr(6) As Integer
        UpNumb = .[C7]
        Randomize  '防止每次生出随机数一样
        If UpNumb < 10 Then
            Msgbox "增长数过少，请自行分配医院级别数量"
        Elseif UpNumb >= 10 Then
            For i = 0 To 5
                If i = 0 Then Rnd_Arr(0) = Int(Rnd * (UpNumb - Sum_Temp - 6)) + 1   'rnd()生成[0，1）的随机数，int（）是取整
                If i > 0 And i < 5 Then Rnd_Arr(i) = Int(Rnd * (UpNumb - Sum_Temp - 6 + i)) + 1
                If i = 5 Then Rnd_Arr(i) = UpNumb - Sum_Temp
                .Cells(i + 10, 3).Value = Rnd_Arr(i)
                .Cells(i + 10, 4) = Rnd_Arr(i) + .Cells(i + 10, 5)
                Sum_Temp = Sum_Temp + Rnd_Arr(i)
            Next
        End If
        .[D16] = RowNumbs
        .[C16] = .[C7]
    End With

'---------------- 省份分布表 统计 --------------
'TODO:省份按照卡数的多少来进行排序
    Dim PvcNumb%
    With Dst_Wkb.Sheets("省份分布")
        .Columns(5).EntireColumn.Insert
        .[E2].Value = Format(Now,"yy/mm/dd")
        PvcNumb = .[c1048576].End(xlUp).Row - 7
        If PvcNumb > Last_PvcNumb Then
            Msgbox "本周有新增的省份！"
        Elseif PvcNumb < Last_PvcNumb Then
            Msgbox "省份减少，统计有错误，请注意！"
        Elseif PvcNumb = Last_PvcNumb Then
            For i = 3 To .[c1048576].End(xlUp).Row
                'TODO:每次增加新省份=便需要重新改下面的数字
                Select Case i
                    Case Is = 9 : .Cells(i,5) = Application.WorksheetFunction.Sum(.[E3:E88])
                    Case Is = 19 : .Cells(i,5) = Application.WorksheetFunction.Sum(.[E10:E18])
                    Case Is = 23 : .Cells(i,5) = Application.WorksheetFunction.Sum(.[E20:E22])
                    Case Is = 27 : .Cells(i,5) = Application.WorksheetFunction.Sum(.[E24:E26])
                    Case Is = 32 : .Cells(i,5) = Application.WorksheetFunction.Sum(.[E28:E31])
                    Case Else : .Cells(i,5) = Application.WorksheetFunction.CountIf(Src_Wkb.Sheets("DocData").[A:A], .Cells(i, 3))
                End Select
                .Cells(i,4) = .Cells(i,5) - .Cells(i,6)
            Next
        End If
        .[E:E].FormatConditions.Delete
    End With 

'---------------- 城市分布表 统计 --------------
'TODO:注意每个省份的城市数量是否有数量的增减
'TODO:原先有，现在没按照0来统计
' 营口和安阳之前有，现在没有了，注意这个问题
    Dim CityNumb%, Orig_CityDict As Object, TempNewPvc$, TempNewCity$, PvcRow%, j%
    Set Orig_CityDict = CreateObject("scripting.dictionary")
    With Dst_Wkb.Sheets("城市分布")
        .Columns(6).EntireColumn.Insert
        .[F2].Value = Format(Now,"yy/mm/dd")
        CityNumb = .[d1048576].End(xlUp).Row - 2
        ' If CityNumb > Last_CityNumb Then
        '     For i = 3 To .[b1048576].End(xlUp).Row - 1
        '         If .Cells(i,3) <> "" And Not Orig_CityDict.exists(.Cells(i,3).Value) Then Orig_CityDict(.Cells(i,3).Value)= .Cells(i,3).Value
        '     Next
        '     For i = 0 To CityDict.Count -1
        '         If Not Orig_CityDict.exists(CityDict(i)) Then 
        '             Msgbox "本周有新增城市 --- " & CityDict(i)
        '             TempNewPvc = Split(CityDict(i), "-")(0)
        '             TempNewCity = Split(CityDict(i), "-")(1)
        '             PvcRow = .[B:B].Find(TempNewPvc, lookat:=xlWhole).Row
        '             .Rows(PvcRow + 1).EntireRow.Insert
        '             .Cells(PvcRow + 1,3) = CityDict(i)
        '             .Cells(PvcRow + 1,4) = TempNewCity
        '             For j = 7 To .Cells(2,256).End(xlToLeft).Column
        '                 .Cells(PvcRow,j) = 0
        '             Next

        '         End If 
        '     Next
            
        ' Elseif CityNumb < Last_CityNumb Then
        '     Msgbox "城市减少，统计有错误，请注意！"
        ' End If
        'TODO:增加城市限制数与进度

        For i = 3 To .[b1048576].End(xlUp).Row -1
            .Cells(i,6) = Application.WorksheetFunction.CountIf(Src_Wkb.Sheets("DocData").[M:M], .Cells(i, 3))
            .Cells(i,5) = .Cells(i,6) - .Cells(i,7)
        Next
        .Cells(.[b1048576].End(xlUp).Row,5)= Application.WorksheetFunction.Sum(Range(.[E3],.Cells(.[b1048576].End(xlUp).Row -1 ,5)))
        .Cells(.[b1048576].End(xlUp).Row,6)= Application.WorksheetFunction.Sum(Range(.[F3],.Cells(.[b1048576].End(xlUp).Row -1 ,6)))
        .[F:F].FormatConditions.Delete
    End With

' 改text1的font属性，改字号的
' time、person_id、Project_id

'TODO:核对统计的正确性:每个表的总数是否一样
'TODO:发生了减少需要记录检测

' 另存为xlsx
'---------------- 选中每个表的B2单元格 ----------------
    For i = 1 To Dst_Wkb.Worksheets.Count
        Dst_Wkb.Worksheets(i).Activate
        [B2].Select
    Next
    Worksheets("汇总").Activate


    Src_Wkb.Save
    Dst_Wkb.Save
    Set Src_Wkb = Nothing
    Set Dst_Wkb = Nothing
    Set PvcDict = Nothing
    Set CityDict = Nothing
    Set Orig_CityDict = Nothing
    Msgbox "数据统计完成！"
    Application.ScreenUpdating = True
End Sub

'TODO:生成xlsx格式的工作簿

