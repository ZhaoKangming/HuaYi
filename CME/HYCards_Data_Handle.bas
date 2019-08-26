Sub Get_CardsNumb()
    Application.ScreenUpdating = False

    'TODO:城市数据的统计、检查市是否有增加？
    'TODO:省份内/卡类由多到少排序
    
    ' 表格样式设置，字体，列宽，字号，对齐等，冻结首行
    ' 百分比，条件格式，进度预警，颜色变化
    ' 省市的增长进度的突然增加预警
    ' 增长刷超过150的进行提示，Max值
    ' 80% 预警

    Dim i%, j%, RowNumb%, DstColNumb%, DstRowNumb%, InsertRow%
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
        If Application.WorksheetFunction.countif(Sheets("省份统计").[G:G],Cells(i,1)) = 0 Then Cells(i,5)= "新增"   'TODO:分析新增地区卡类
    Next i
    [I1] = "总用卡数"
    [I2] = Application.WorksheetFunction.Sum([D:D])


    '------------------- 省份统计表添加新增的地区卡类型 -------------------
    Dim ProvCell As Range
    For i = 1 To RowNumb
        ' 获取是否为本周新增
        If Cells(i,5)="新增" Then 
            Set ProvCell = Sheets("省份统计").Columns(1).Find(Sheets("Data").Cells(i,2),,,,,xlPrevious)
            If ProvCell Is Nothing Then
                Sheets("Data").Cells(i,6) = "新增省份"
                DstRowNumb = Sheets("省份统计").[A1048576].End(xlUp).Row
                InsertRow = DstRowNumb               
            Else
                InsertRow = ProvCell.Row + 1
            End If

            With Sheets("省份统计")
                DstColNumb = .Cells(1, Columns.Count).End(xlToLeft).Column
                .Rows(InsertRow).Insert
                .Cells(InsertRow,1) = Sheets("Data").Cells(i,2)
                .Cells(InsertRow,8) = Sheets("Data").Cells(i,3)
                .Cells(InsertRow,7) = .Cells(InsertRow,1) & .Cells(InsertRow,8)
                For j = 10 To DstColNumb
                    .Cells(InsertRow,j) = 0
                Next j
            End With               
        End If
    Next i

    '------------------- 卡类统计表添加新增的卡类型地区 -------------------
    Dim CardTypeCell As Range
    ' Const NoCardsProv_Numb As Integer = 7  '此数量是
    For i = 1 To RowNumb
        ' 获取是否为本周新增
        If Cells(i,5)="新增" Then 
            Set CardTypeCell = Sheets("卡类统计").Columns(1).Find(Sheets("Data").Cells(i,3),,,,,xlPrevious)
            If CardTypeCell Is Nothing Then
                Sheets("Data").Cells(i,7) = "新增卡类型"
                DstRowNumb = Sheets("卡类统计").[A1048576].End(xlUp).Row
                InsertRow = DstRowNumb               
            Else
                InsertRow = CardTypeCell.Row + 1
            End If

' 如果不是新增的要把其他信息复制过来
            With Sheets("卡类统计")
                DstColNumb = .Cells(1, Columns.Count).End(xlToLeft).Column
                .Rows(InsertRow).Insert
                .Cells(InsertRow,1) = Sheets("Data").Cells(i,3)
                .Cells(InsertRow,7) = Sheets("Data").Cells(i,2)
                For j = 9 To DstColNumb
                    .Cells(InsertRow,j) = 0
                Next j
            End With               
        End If
    Next i
    'TODO:卡类有，但没发，本周开发，这种新插入的行，怎么处理

有新增的地区或卡，要停下来，填完完整的数据后再继续运行程序，inputbox 参数，参数为空，则 exit sub

    '------------------- 省份统计表数据获取 ------------------
    vlookup 获取本周的值
    [J:J].Replace "#N/A",""
    本周总计值计算
    检查本周值是否正确
    计算本周增加数
    增加数最大的进行颜色填充，根据值的大小确定填充几个

    计算已经发卡数 D列
    计算剩余量 Elie
    计算投放进度 F列
    检查 已发卡数和本周总数是否一致

'------------------- 省份累计表处理 ------------------
插入新列
更新已投放数
获取本周数据
检验总数与周增长数是否和上一个表一致


'------------------- 卡类统计表数据获取 ------------------
    vlookup 获取本周的值
    [J:J].Replace "#N/A",""
    本周总计值计算
    检查本周值是否正确
    计算本周增加数
    增加数最大的进行颜色填充，根据值的大小确定填充几个
    购卡数量更新

    计算已经发卡数 D列
    计算投放进度
    计算所有一头卡投放进度 F列
    检查 已发卡数和本周总数是否一致






    '------------------- 获取剩余数、进度 -------------------


    进度如果>1,或者限制数为0的如何处理，新爆仓的怎么办
    保留两位小数


前三名填充颜色

    '------------------- 获取将之前周的统计数据 -------------------
    Usedrange.replace "#N/A","0"

    '------------------- 工作表按省份名称排序 -------------------


    '------------------- 处理新表样式 -------------------



排序

    '------------------- 周检查是否有更新 -------------------
例如总线制数、总购卡数


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

生成xlsx
标准命名
删除 data 表格

' - 省份统计表
[G:G].Hidden = True
Merge 单元格