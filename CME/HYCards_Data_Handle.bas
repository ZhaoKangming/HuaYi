Sub Get_CardsNumb()
    Application.ScreenUpdating = False

    'TODO:检查市是否有增加？
    ' 城市数据的统计
    ' 导出的数据有些事文本格式需要转换成数字格式
    ' 表格样式设置，字体，列宽，字号，对齐等，冻结首行
    ' 百分比，条件格式，进度预警，颜色变化
    ' 省市的增长进度的突然增加预警
    ' 按销售人员统计，图表，默认是开筛选的
    ' 增长刷超过150的进行提示，Max值

    Dim i%, Card_Wkb As workbook, Click_Wkb As workbook, Tool_Wkb As workbook, RowNumbs%
    
    '------------------- 切换工作表 -------------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*药企*" Then Set Card_Wkb = Workbooks(Workbooks(i).Name) 
        If Workbooks(i).Name like "*HYCards*" Then Set Tool_Wkb = Workbooks(Workbooks(i).Name)        
    Next

    '------------------- 复制工作表并简单处理 -------------------
    Sheets("Sheet1").Copy Before:=Sheets("Sheet1")
    Sheets("Sheet1 (2)").Name = "TEMP"
    Sheets("TEMP").Activate
    Columns(1).Delete
    For i = 1 To 4
        Columns(2).insert
    Next
    Columns(7).insert
    RowNumbs = Sheets("TEMP").[a99999].End(xlUp).Row


    '------------------- 将文本储存的数字转为数字 -------------------
    With Columns(8)
        .NumberFormatLocal = "G/通用格式"   '把单元格设置为常规
        .Value = .Value   '取值
    End With



    '------------------- 获取各省份限制数 -------------------
    

    '------------------- 获取剩余数、进度 -------------------


    进度如果>1,或者限制数为0的如何处理，新爆仓的怎么办
    保留两位小数



    '------------------- 获取将之前周的统计数据 -------------------
    Usedrange.replace "#N/A","0"

    '------------------- 工作表按省份名称排序 -------------------


    '------------------- 处理新表样式 -------------------

    '------------------- 合并单元格 -------------------

    '------------------- 添加总计 -------------------


    '------------------- 核对计算正确性 -------------------
    Dim Amount_Arr
    Amount_Arr = Array(18666,18498,16882)



    '------------------- 样式美化 -------------------
    ' 水平竖直居中
    ' 字体大小与字体
    ' 设置所有框线


    Card_Wkb.Save
    Click_Wkb.Save
    Learn_Wkb.Save
    Tool_Wkb.Save
    Set Card_Wkb = Nothing
    Set Click_Wkb = Nothing
    Set Learn_Wkb = Nothing
    Set Tool_Wkb = Nothing 
    Application.ScreenUpdating = True
    Msgbox "Finished!"
End Sub


