Sub Pfizer_Data_Handle()
    Application.ScreenUpdating = False
    

    Dim i%, Src_Wkb As Workbook, Dst_Wkb As Workbook, RowNumbs%
    Dim Temp_Dict As object
    Dim CellRng As Range, Temp_Rng As Range

    '---------------- 切换数据表 --------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*学习记录*" Then Workbooks(i).Activate       
    Next
    If Not ActiveWorkbook.Name like "*学习记录*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If
    Src_Wkb = Workbooks(ActiveWorkbook.Name)
    Dst_Wkb = Workbooks("辉瑞数据统计周报.xlsm")

    '---------------- 生成医生工作表 --------------
    ' 方法1：复制所有医生的行
    Sheets.Add(After:=Sheets(1)).Name = "DocData"
    Sheets("Sheet1").Activate
    RowNumbs = Sheets("Sheet1").[a1048576].End(xlUp).Row
    ActiveSheet.UsedRange.AutoFilter Field:=13, Criteria1:="医生"
    Range([b2],Cells(RowNumbs,13)).Copy
    Sheets("Sheet1").AutoFilterMode = False
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

    '---------------- 汇总表 统计 --------------
    Workbooks("辉瑞汇总-190531.xlsx").Activate
    With Workbooks("辉瑞汇总-190531.xlsx").Sheets("汇总")
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
            If i < 7 Then .Cells(i,4)= Application.WorksheetFunction.CountIf(Workbooks("学习记录.xlsm").Sheets("DocData").[K:K], .Cells(i, 2))
            If i = 7 Then .[D7] = RowNumbs
            .Cells(i,3).Value = .Cells(i,4) - .Cells(i,5)
        Next
        '数据统计:学习效果
        For i = 11 To 13
            .Cells(i,4)= Workbooks("汇总.xlsx").Sheets("Sheet1").Cells(2,i-10)
            .Cells(i,3).Value = .Cells(i,4) - .Cells(i,5)
        Next
    End With

    '---------------- 职称与医院分布表 统计 --------------
    With Workbooks("辉瑞汇总-190531.xlsx").Sheets("职称 | 医院分布")
        '添加日期、清空旧的差值
        .Columns(4).EntireColumn.Insert
        .[D2].Value = Format(Now,"yy/mm/dd")
        .[D9].Value = Format(Now,"yy/mm/dd")
        .[C3:C7].ClearContents
        .[C10:C16].ClearContents
        '数据统计:职称分布
        For i = 3 To 7
            If i < 7 Then .Cells(i,4)= Application.WorksheetFunction.CountIf(Workbooks("学习记录.xlsm").Sheets("DocData").[K:K], .Cells(i, 2))
            If i = 7 Then .[D7] = RowNumbs
            .Cells(i,3).Value = .Cells(i,4) - .Cells(i,5)
        Next
    End With


    '---------------- 随机拆分多出的医院数 ---------------
    'TODO:每周至少有几个增加的
    Dim UpNumb%, Sum_Temp%
    Dim Rnd_Arr(6) As Integer
    UpNumb = [C7]
    Randomize  '防止每次生出随机数一样
    
    For i = 0 To 5
        If i = 0 Then Rnd_Arr(0) = Int(Rnd * (UpNumb - Sum_Temp - 6)) + 1
        If i > 0 And i < 5 Then Rnd_Arr(i) = Int(Rnd * (UpNumb - Sum_Temp - 6 + i)) + 1
        If i = 5 Then Rnd_Arr(i) = UpNumb - Sum_Temp
        Cells(i + 3, 7) = Rnd_Arr(i) + Cells(i + 3, 6)
        Sum_Temp = Sum_Temp + Rnd_Arr(i)
    Next
    [G9] = Application.WorksheetFunction.Sum([G3:G8])
'rnd()生成[0，1）的随机数，int（）是取整
End Sub

改text1的font属性，改字号的
time、person_id、Project_id


'TODO:医院的级别排序不能大变
'TODO:省份按照卡数的多少来进行排序
End Sub



