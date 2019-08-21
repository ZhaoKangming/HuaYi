Sub Get_CardsNumb()
    Application.ScreenUpdating = False

    'TODO:检查市是否有增加？
    'TODO:省份内/卡类由多到少排序
    'TODO:加分割双线
    ' 城市数据的统计
    ' 导出的数据有些事文本格式需要转换成数字格式
    ' 表格样式设置，字体，列宽，字号，对齐等，冻结首行
    ' 百分比，条件格式，进度预警，颜色变化
    ' 省市的增长进度的突然增加预警
    ' 按销售人员统计，图表，默认是开筛选的
    ' 增长刷超过150的进行提示，Max值
    ' 80% 预警

    Dim i%, j%, RowNumb%, ColNumb%
    Dim Prov_Dict As Object, Current_Card_Dict As Object
    
    Workbooks("HYCards-DataTools.xlsm").Activate

    '------------------- 初步处理原始 Data 与表格-------------------
    Sheets("Data").Activate
    Sheets("Data").UsedRange.Replace " ",""
    Rows(1).Delete
    Columns(1).ClearContents
    Columns(1).ColumnWidth = 45
    Columns(3).ColumnWidth = 35
    With Columns(4)
        .NumberFormatLocal = "G/通用格式"   '把单元格设置为常规
        .Value = .Value   '取值
    End With
    RowNumb = Sheets("Data").[B99999].End(xlUp).Row
    For i = 1 To RowNumb
        Cells(i,1) = Cells(i,2) & Cells(i,3)  ' 合并省份及学术卡类型
        If Application.WorksheetFunction.countif(Sheets().,Cells(i,5)   'TODO:分析新增地区卡类
    Next i


    '------------------- 省份统计表添加新增的地区卡类型 -------------------
    Dim ProvCell As Range
    ' Const NoCardsProv_Numb As Integer = 7  '此数量是
    For i = 1 To RowNumb
        ' 获取是否为本周新增
        If Cells(i,5)="Y" Then 
            Set ProvCell = Sheets("省份统计").Columns(1).Find(Sheets("Data").Cells(i,2),,,,,xlPrevious)
            If Not ProvCell Is Nothing Then
                With Sheets("省份统计")
                    ColNumb = .Cells(1, Columns.Count).End(xlToLeft).Column
                    .Rows(ProvCell.Row + 1).Insert
                    .Cells(ProvCell.Row + 1,1) = Sheets("Data").Cells(i,2)
                    .Cells(ProvCell.Row + 1,8) = Sheets("Data").Cells(i,3)
                    .Cells(ProvCell.Row + 1,7) = .Cells(ProvCell.Row + 1,1) & .Cells(ProvCell.Row + 1,8)
                    For j = 10 To ColNumb
                        .Cells(ProvCell.Row + 1,j) = 0
                    Next j
                End With               
            Else
                Sheets("Data").Cells(i,6) = "新增省份"
            End If
        End If
    Next i

    '------------------- 卡类统计表添加新增的卡类型地区 -------------------
    Dim CardTypeCell As Range
    ' Const NoCardsProv_Numb As Integer = 7  '此数量是
    For i = 1 To RowNumb
        ' 获取是否为本周新增
        If Cells(i,5)="Y" Then 
            Set CardTypeCell = Sheets("卡类统计").Columns(1).Find(Sheets("Data").Cells(i,3),,,,,xlPrevious)
            If Not CardTypeCell Is Nothing Then
                With Sheets("卡类统计")
                    ColNumb = .Cells(1, Columns.Count).End(xlToLeft).Column
                    .Rows(CardTypeCell.Row + 1).Insert
                    .Cells(CardTypeCell.Row + 1,1) = Sheets("Data").Cells(i,3)
                    .Cells(CardTypeCell.Row + 1,7) = Sheets("Data").Cells(i,2)
                    For j = 9 To ColNumb
                        .Cells(CardTypeCell.Row + 1,j) = 0
                    Next j
                End With               
            Else
                Sheets("Data").Cells(i,7) = "新增卡类型"
            End If
        End If
    Next i


    '------------------- 向表格中填充新增的地区卡类-------------------

    '------------------- 获取剩余数、进度 -------------------


    进度如果>1,或者限制数为0的如何处理，新爆仓的怎么办
    保留两位小数


前三名填充颜色

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


    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Msgbox "Finished!"
End Sub


' 完成后的处理工作
