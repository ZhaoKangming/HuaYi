Sub Project_Data_Handle()
    Dim i%, Src_Wkb As Workbook, Stat_Period$

    Stat_Period = "2019年4月1日-2019年6月5日"  'TODO:设置获取当前日期

    '------------------- 切换工作表 -------------------
    For i = 1 To Workbooks.Count
        If Not Workbooks(i).Name like "*personal*" Then Workbooks(i).Activate       
    Next
    If ActiveWorkbook.Name like "*personal*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If
    Src_Wkb = Workbooks(ActiveWorkbook.Name)
    Src_Wkb.Activate


    '------------------- Sheet-学习人数汇总 -------------------
    Sheets("学习人数汇总").Activate


End Sub



