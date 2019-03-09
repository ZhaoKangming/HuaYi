Sub CountNumbs()
    Application.ScreenUpdating = False
    Call DateClean


    Call Beautify

    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "Finished！"
End Sub

Sub DateClean()
    Sheets("Sheet1").UsedRange.Replace " ",""
    With Sheets("Sheet1").UsedRange
        .Replace "NULL","其他"
        .Replace "-请选择-","其他"
        .Replace "？",""
        .Replace "~?",""
        .Replace "！",""
        .Replace "!",""
        .Replace "~*",""
        .Replace "","其他"
    End With
End Sub




Sub Beautify()

    With Sheets("Sheet1").UsedRange
        .Font.Name = "微软雅黑"
        .Font.Size = 12    
    End With

End Sub



医院等级里有Null和-请选择-
    Dim RowsNumb%
    RowsNumb = [a99999].End(xlUp).Row
