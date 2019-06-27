Sub Pfizer_Data_Handle()
    Application.ScreenUpdating = False
    

    Dim i%, Src_Wkb As Workbook, Dst_Wkb As Workbook
    Dim Temp_Dict As object
    Dim CellRng As Range, Temp_Rng As Range

    '---------------- 切换数据表 --------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*辉瑞统计*" Then Workbooks(i).Activate       
    Next
    If Not ActiveWorkbook.Name like "*辉瑞统计*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If
    Set Src_Wkb = Workbooks(ActiveWorkbook.Name)
    Set Dst_Wkb = Workbooks("辉瑞-DataTool.xlsm")

    '---------------- 数据清洗 --------------
    '清洗孙旭辰这个测试账号的信息

    '---------------- 生成医生工作表 --------------
    ' 方法1：复制所有医生的行
    Sheets.Add(After:=Sheets(3)).Name = "DocData"
    Sheets("Sheet2").Activate
    RowNumbs = Sheets("Sheet2").[a1048576].End(xlUp).Row
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
    For i = 1 To [a999999].End(xlUp).Row
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
            Cells(i,3) = Cells(i,4)-Cells(i,5)
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
                .Cells(i + 10, 4) = Rnd_Arr(i) + Cells(i + 10, 5)
                Sum_Temp = Sum_Temp + Rnd_Arr(i)
            Next
        End If
        .[D16] = RowNumbs
    End With

'---------------- 省份分布表 统计 --------------
'TODO:省份按照卡数的多少来进行排序
    Dim PvcNumb%, 
    With Dst_Wkb.Sheets("省份分布")
        PvcNumb = .[c1048576].End(xlUp).Row - 7
    

    End With 

'---------------- 城市分布表 统计 --------------
'TODO:注意每个省份的城市数量是否有数量的增减
    Dim CityNumb%, 
    With Dst_Wkb.Sheets("城市分布")


    End With

改text1的font属性，改字号的
time、person_id、Project_id

'TODO:核对统计的正确性

鼠标位置
省份如果有增加

清除由于增加列导致的新列多余的条件格式
另存为xlsx

    Src_Wkb.Save
    Dst_Wkb.Save
    Set Src_Wkb = Nothing
    Set Dst_Wkb = Nothing
    Msgbox "数据统计完成！"
End Sub



