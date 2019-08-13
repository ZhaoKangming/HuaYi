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


    80% 预警

    Dim i%, RowNumbs%
    Dim Prov_Dict As Object, Current_Card_Dict As Dict
    
    Workbooks("HYCards-DataTools.xlsm").Activate

    '------------------- 初步处理原始 Data 与表格-------------------
    Sheets("Data").Activate
    Rows(1).Delete
    Columns(1).ClearContents
    Columns(1).ColumnWidth = 45
    Columns(3).ColumnWidth = 35
    With Columns(4)
        .NumberFormatLocal = "G/通用格式"   '把单元格设置为常规
        .Value = .Value   '取值
    End With
    RowNumbs = Sheets("Data").[B99999].End(xlUp).Row
    For i = 1 To RowNumbs
        Cells(i,1) = Cells(i,2) & Cells(i,3)  ' 合并省份及学术卡类型
    Next


    '------------------- 分析新增的地区及卡类型-------------------
    
    For i = 1 To RowNumbs
        Cells(i,1) = Cells(i,2) & Cells(i,3)  ' 合并省份及学术卡类型
    Next

    '------------------- 分析新增地区卡类-------------------


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


