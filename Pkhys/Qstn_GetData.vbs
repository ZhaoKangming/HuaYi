' 此程序用于处理从问卷星到处的数据表，并按需求进行统计
' 此工作已经交接，无需继续开发
Sub GetDate()
    Call DateClean
    Call CreatSheet
    Call DateCount
    Call Beautify
    ThisWorkbook.Save
    Msgbox "Data has been processed!"
End Sub

Sub DateClean()
    Dim QuestTitleArr,LastRow%, i%
    LastRow = [A10000].End(xlUp).Row
    
    QuestTitleArr = Array("T1-A","T1-B","T1-C","T1-D","T2-A","T2-B","T2-C","T2-D","T3-A","T3-B","T3-D","T4" _
                            "T5","T6","T7-A","T7-B","T7-C","T7-D","T1","T2","T3","T4","T5","T6","T7")
    
    Rows(2,8).Delete
    Columns(2,6).Delete
    'Columns("B:T").ColumnWidth = 5

    Range([B2],Cells(LastRow,5)).Interior.Color = &H00ffcc
    Range([F2],Cells(LastRow,9)).Interior.Color = &H33ff99
    Range([J2],Cells(LastRow,13)).Interior.Color = &H33ffcc
    Range([N2],Cells(LastRow,14)).Interior.Color = &H33ffff
    Range([O2],Cells(LastRow,15)).Interior.Color = &H33ccff
    Range([P2],Cells(LastRow,16)).Interior.Color = &H3399ff
    Range([Q2],Cells(LastRow,20)).Interior.Color = &H3366ff

    Range([A2],Cells(LastRow,1)).ClearContents
    For i = 2 To LastRow
        Cells(i,1) = i - 1
        Cells(i,21) = Application.WorksheetFunction.Concat(Range(Cells(i,2),Cells(i,5)))
        Cells(i,22) = Application.WorksheetFunction.Concat(Range(Cells(i,6),Cells(i,9)))
        Cells(i,23) = Application.WorksheetFunction.Concat(Range(Cells(i,10),Cells(i,13)))
        Cells(i,24) = Cells(i,14)
        Cells(i,25) = Cells(i,15)
        Cells(i,26) = Cells(i,16)
        Cells(i,27) = Application.WorksheetFunction.Concat(Range(Cells(i,17),Cells(i,20)))
    Next

    Columns("B:T").Delete
    

    ' With Range([B2],Cells(LastRow,8))
    '     .Replace 


    ' End With


End Sub

Sub CreatSheet()

End Sub

Sub DateCount()
    Dim n

End Sub

Sub Beautify()

End Sub