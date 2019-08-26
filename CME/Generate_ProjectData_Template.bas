'使用说明：从网页的课程列表处开始复制到课件末尾，粘贴在新表格中的[F1]单元格，粘贴模式为：匹配目标格式

Sub Generate_ProjectData_Template()
    Application.ScreenUpdating = False
    '--------------------- 处理项目信息总表 ---------------------
    Dim i%, j%, DelRowStr_Arr
    Columns(6).ColumnWidth = 80
    With Sheets("总表").UsedRange
        .Replace " ",""
        .Replace "关注度：",""
    End With

    DelRowStr_Arr = Array("单位*", "授课老师*", "*课程列表*", "*类*学分*")

    For i = [F1048576].End(xlUp).Row To 1 Step -1
        For j = 0 To UBound(DelRowStr_Arr)
            If Cells(i,6) like DelRowStr_Arr(j) Then
                Cells(i,6).Delete Shift:=xlUp
                Exit For
            End if
        Next j
        If Trim(Cells(i,6)) = "" Then Cells(i,6).Delete Shift:=xlUp

        If Cells(i,6) like "*项目负责人*单位*" Then 
            Cells(i,6) = Trim(Cells(i,6))
            Cells(i,6) = Mid(Cells(i,6),7,InStr(Cells(i,6),"单位")-7)
        End If
    Next i
    [F3].Delete Shift:=xlUp
    Sheets("总表").UsedRange.Replace " ",""

    For i = 3 To [F1048576].End(xlUp).Row       '给课题名称添加序号
        Cells(i,6) = i - 2 & "-" & Cells(i,6)
    Next i

    For i = 1 To [F1048576].End(xlUp).Row -4    '在课题名称中添加新行
        Rows(5).insert
        [B5:D5].Merge
        [B5].HorizontalAlignment = xlLeft
    Next i

    For i = [F1048576].End(xlUp).Row To 1 Step -1   '删除因为增加新行导致产生的新的空格
        If Trim(Cells(i,6)) = "" Then Cells(i,6).Delete Shift:=xlUp
    Next i

    For i = 1 To [F1048576].End(xlUp).Row       '数据移位
        Cells(i+1,2) = Cells(i,6)
    Next i

    [F:F].Delete

    '--------------------- 设置项目数据的统计周期 ---------------------
    Dim Statistical_Period$, SP_Result
    Statistical_Period = inputbox("请输入统计的周期,或輸入 now ","统计周期")
    If Statistical_Period = "" Then
        SP_Result = MsgBox("使用 190207 - NOW？", vbYesNo + vbQuestion + vbDefaultButton1, "核对统计周期")
        If SP_Result = vbYes Then Statistical_Period = "now"
        If SP_Result = vbNo Then Statistical_Period = inputbox("请输入统计的周期","统计周期")
    End if
    If Statistical_Period like "*now*" Then Statistical_Period = "2019年2月7日-" & Format(Now,"yyyy年m月d日")

    Sheets("专业分析").[B2].Value = Statistical_Period
    Sheets("职称分析").[B2].Value = Statistical_Period
    Sheets("省市分布分析").[C2].Value = Statistical_Period
    Sheets("医院等级分析").[B2].Value = Statistical_Period

    '--------------------- 设置统计周期区段 ---------------------
    Dim Query_Result
    'TODO:自动生成统计的区段
    'TODO:获取某个月的最后一天

    'TODO:区分单次查询还是项目月报，单次查询的话 统计区段只有一次，就是 Statistical_Period
    Query_Result = MsgBox("是单次查询么?", vbYesNo + vbQuestion + vbDefaultButton1, "报表性质")
    If Query_Result = vbYes Then
        Sheets("学习人数汇总").[A3].Value = Statistical_Period
        Sheets("学习基本情况").[A3].Value = Statistical_Period
    End if

    '--------------------- 选定相应单元格，使技术粘贴目标数据至此 ---------------------
    Sheets("总表").[A1].Select
    Sheets("学习人数汇总").[B3].Select
    Sheets("学习基本情况").[B3].Select
    Sheets("专业分析").[D1].Select
    Sheets("职称分析").[D1].Select
    Sheets("省市分布分析").[E1].Select
    Sheets("医院等级分析").[D1].Select



    'TODO:自动生成新表并改名

    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Msgbox "已经处理完成！"
End Sub