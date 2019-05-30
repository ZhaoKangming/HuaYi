Sub Get_CardsNumb()
    Application.ScreenUpdating = False


    'TODO:检查市是否有增加？
    市
    导出的数据有些事文本格式需要转换成数字格式
    表格样式设置，字体，列宽，字号，对齐等，冻结首行

百分比，条件格式，进度预警，颜色变化
省市的增长进度的突然增加预警
广东，山西，海南，四川，百分比排序
数据的严谨与准确性， 程序的私有性
按销售人员统计，图表，默认是开筛选的

    Dim i%, Src_Wkb As workbook, Dst_Wkb As workbook, RowNumbs%
    Dim Temp_Dict As object
    Dim CellRng As Range, Temp_Rng As Range

    '------------------- 切换工作表 -------------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*学习卡*" Then Workbooks(i).Activate       
    Next
    If Not ActiveWorkbook.Name like "*学习卡*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If
    Src_Wkb = Workbooks(ActiveWorkbook.Name)

    '------------------- 复制工作表并简单处理 -------------------
    Sheets("Sheet1").Copy Before:=Sheets("Sheet1")
    Sheets("Sheet1 (2)").Name = "TEMP"
    Sheets("TEMP").Activate
    Columns(1).Delete


    '------------------- 将文本储存的数字转为数字 -------------------
    With Columns(3)
        .NumberFormatLocal = "G/通用格式"   '把单元格设置为常规
        .Value = .Value   '取值
    End With

    '------------------- 工作表按省份名称排序 -------------------

    

    '------------------- 获取各省份限制数 -------------------


    '------------------- 获取将之前周的统计数据 -------------------

    '------------------- 填写策略备注 -------------------
    For i = 2 To LastRow
        Select Case Cells(i,1)
            Case Is = "海南省" : Cells(i,2) = "已停止投放"
            Case Is = "四川省" : Cells(i,2) = "凉山州已停止投放"
    Next
        


    '------------------- 合并单元格 -------------------
    Application.DisplayAlerts = False
    LastRow = Sheets("Sheet1").[a99999].End(xlUp).Row
    
    For i = LastRow - 1 To 2 Step -1
        If Cells(i,1) = Cells(i-1,1) Then 
                Range(Cells(i-1,1),Cells(i,1)).Merge
                Range(Cells(i-1,2),Cells(i,2)).Merge
                Range(Cells(i-1,3),Cells(i,3)).Merge            
        End If
    Next
    Application.DisplayAlerts = True


    '------------------- 添加总计 -------------------


    '------------------- 核对计算正确性 -------------------
    Dim Amount_Arr
    Amount_Arr = Array(18666,18498,16882)


    Src_Wkb.Save
    Dst_Wkb.Save
    Application.ScreenUpdating = True
    Msgbox "Finished!"
End Sub


