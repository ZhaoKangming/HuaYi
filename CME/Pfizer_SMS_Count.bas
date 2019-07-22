'【功能】统计每周的辉瑞尚医短信推广效果

Sub Pfizer_SMS_Count()
    Application.ScreenUpdating = False

    Dim i%, Result, MsgboxText$
    Dim ClickRecord_Wkb As workbook, LearnRecord_Wkb As workbook, Report_Wkb As workbook

    '------------------- 提醒提前确定比较条件 -------------------
    MsgboxText = "请问是否完成了以下处理？" & vblf & vblf &  _
                "1. 复制短信表及添加省份数量" & vblf &  _
                "2. 确定课程列表" & vblf &  _
                "3. 仅保留起止时间内数据" & vblf &  _
                "4. 正确给表格命名为：点击/学习"
    '如果课程数量多的话，使用Array来判断            
    'Dim CourseList_Arr As Array 
    'CourseList_Arr = Array("冠心病相关疾病的规范化治疗","高血压患者相关疾病管理")
    Result = MsgBox(MsgboxText, vbYesNo + vbQuestion + vbDefaultButton1, "数据准备")
    If Result = vbNo Then Exit Sub

    '------------------- 切换工作表 -------------------
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*点击*" Then Set ClickRecord_Wkb = Workbooks(Workbooks(i).Name)
        If Workbooks(i).Name like "*学习*" Then Set LearnRecord_Wkb = Workbooks(Workbooks(i).Name)
        If Workbooks(i).Name like "*辉瑞尚医*" Then Set Report_Wkb = Workbooks(Workbooks(i).Name)        
    Next

    '------------------- 医生类型分析 -------------------
    ClickRecord_Wkb.Activate
    With ClickRecord_Wkb.Sheets(1)
        For i = 2 To .[A1048576].End(xlUp).Row
            If .Cells(i,4) Like "*药*" Or .Cells(i,4) Like "*护*" Or .Cells(i,4) Like "*技*" Then 
                .Cells(i,5) = "药技护"
            Else
                .Cells(i,5) = .Cells(i,2) & "医生"
            End If
        Next
    End With

    '------------------- 短信点击情况分析 -------------------
    Report_Wkb.Activate
    Dim Sum_RowNumb%
    With Report_Wkb.Sheets("辉瑞尚医-短信推广")
        .[B2].Value ="辉瑞·尚医项目短信推广情况 - " & Format(Now,"yyyymmdd")
        Sum_RowNumb = .Columns(2).Find("总计",LookIn:=xlValues).Row
        For i = 4 To Sum_RowNumb - 1
            .Cells(i,4) = Application.WorksheetFunction.Countif(ClickRecord_Wkb.Sheets(1).[B:B],.Cells(i,2))
            .Cells(i,5) = .Cells(i,4)/.Cells(i,3)
            .Cells(i,6) = Application.WorksheetFunction.Countif(ClickRecord_Wkb.Sheets(1).[E:E],.Cells(i,2)&"医生")
            .Cells(i,7) = .Cells(i,6)/.Cells(i,3)
        Next
        .Cells(Sum_RowNumb,3) = Application.WorksheetFunction.Sum(.Range(.[C4],.Cells(Sum_RowNumb-1,3)))
        .Cells(Sum_RowNumb,4) = Application.WorksheetFunction.Sum(.Range(.[D4],.Cells(Sum_RowNumb-1,4)))
        .Cells(Sum_RowNumb,5) = .Cells(Sum_RowNumb,4)/.Cells(Sum_RowNumb,3)
        .Cells(Sum_RowNumb,6) = Application.WorksheetFunction.Sum(.Range(.[F4],.Cells(Sum_RowNumb-1,6)))
        .Cells(Sum_RowNumb,7) = .Cells(Sum_RowNumb,6)/.Cells(Sum_RowNumb,3)
    End With

    '------------------- 获取有效的学习记录 -------------------
    LearnRecord_Wkb.Activate
    Application.DisplayAlerts = False
    With LearnRecord_Wkb.Sheets(2)
        For i = .[A1048576].End(xlUp).Row To 1 Step -1
            If .Cells(i, 11) & .Cells(i, 13) <> "冠心病相关疾病的规范化治疗医生" And .Cells(i, 11) & .Cells(i, 13) <> "高血压患者相关疾病管理医生" Then 
                Rows(i).EntireRow.Delete
            End If
        Next
    End With
    Application.DisplayAlerts = True

    '------------------- 学习情况分析 -------------------
    Report_Wkb.Activate
    With Report_Wkb.Sheets("辉瑞尚医-短信推广")
        For i = 4 To Sum_RowNumb - 1
            .Cells(i,8) = Application.WorksheetFunction.Countif(LearnRecord_Wkb.Sheets(2).[B:B],.Cells(i,2))
            .Cells(i,9) = .Cells(i,8)/.Cells(i,3)

            
            .Cells(i,10) = Application.WorksheetFunction.Countif(ClickRecord_Wkb.Sheets(1).[E:E],.Cells(i,2)&"医生")
            .Cells(i,11) = .Cells(i,6)/.Cells(i,3)
            .Cells(i,12) = .Cells(i,6)/.Cells(i,3)
        Next
        .Cells(Sum_RowNumb,8) = Application.WorksheetFunction.Sum(.Range(.[C4],.Cells(Sum_RowNumb-1,3)))
        .Cells(Sum_RowNumb,9) = Application.WorksheetFunction.Sum(.Range(.[D4],.Cells(Sum_RowNumb-1,4)))
        .Cells(Sum_RowNumb,10) = .Cells(Sum_RowNumb,4)/.Cells(Sum_RowNumb,3)
        .Cells(Sum_RowNumb,11) = Application.WorksheetFunction.Sum(.Range(.[F4],.Cells(Sum_RowNumb-1,6)))
        .Cells(Sum_RowNumb,12) = .Cells(Sum_RowNumb,6)/.Cells(Sum_RowNumb,3)
    End With

    '------------------- 保存表格并释放变量 -------------------
    ClickRecord_Wkb.Save
    LearnRecord_Wkb.Save
    Report_Wkb.Save
    Set ClickRecord_Wkb = Nothing
    Set LearnRecord_Wkb = Nothing
    Set Report_Wkb = Nothing

    Msgbox "已经完成短信数据的统计！"
    Application.ScreenUpdating = True

End Sub