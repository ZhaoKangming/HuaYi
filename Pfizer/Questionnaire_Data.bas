' 需要提前处理的工作：
' [1] 仅保留目标城市数据，并将此表命名为 "Data"
' [2] 将以文本存储的数字转化为数值


Sub Questionnaire_Data_Analyse()
    Application.ScreenUpdating = False
    Dim LastRow&, i&, DataError as Boolean, ICU_Arr

    DataError = False
    Sheets("Data").Activate
    LastRow = Sheets("Data").[A1048576].End(xlUp).Row
    ' With Sheets("Data").UsedRange
    '     .Replace " ",""
    
    ' End With


    '------------------- 删除没有用的数据 -------------------
    ' Columns(10).Delete '删除来源列
    ' Columns(9).Delete '删除ip列
    ' Columns(6).Delete '删除序号列

    '------------------- 处理来源端口数据 -------------------
    ' For i = 2 To LastRow
    '     If Left(Cells(i,8),2) = "pc" Then 
    '         Cells(i,8) = "PC"
    '     ElseIf Left(Cells(i,8),2) = "mo" Then
    '         Cells(i,8) = "Mobile"
    '     Else
    '         Cells(i,8) = "Others"
    '         DataError = True
    '     End If
    ' Next
    ' If DataError = True Then Msgbox "存在其他来源的问卷数据！"

    '------------------- 数据清洗 -------------------
    

    '------------------- 获取标准职称 -------------------



    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Msgbox "数据已经处理完成！"
End Sub