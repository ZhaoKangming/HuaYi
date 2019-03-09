Sub CountNumbs()
    Application.ScreenUpdating = False
    Call DateClean
    Call GetList

    Call Beautify

    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "Finished！"
End Sub

Sub DateClean()
    Sheets("Sheet1").UsedRange.Replace " ",""
    Sheets("Sheet1").[E:E].Replace "*医院",""
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
    'MsgBox "Date have been cleaned!"
End Sub


Sub Beautify()

    With Sheets("Sheet1").UsedRange
        .Font.Name = "微软雅黑"
        .Font.Size = 12    
    End With

    Sheets("Sheet3").Delete
End Sub

Sub GetFrequency(ValueRange As Range,StartCell As Range)
    Dim rng As Range, arr, d As Object, i%, SCRow%
    Set d = CreateObject("scripting.dictionary")
    For Each rng In ValueRange
        If rng <> "" And Not d.exists(rng.Value) Then d(rng.Value)= rng.Value
    Next
    arr = d.items
    SCRow = StartCell.Row
    Cells(SCRow + 1,StartCell.Column).Resize(d.Count-1, 1).EntireRow.Insert shift:=xlDown
    For i = 0 To d.Count - 1
        Cells(i + SCRow, StartCell.Column) = arr(i)
    Next
End Sub

Sub GetList()
    Dim RowsNumb%, i%, CitiesNumb%
    RowsNumb = Sheets("Sheet1").[a99999].End(xlUp).Row
    Call GetFrequency(Sheets("Sheet1").Range([A2], Cells(RowsNumb, 1)),Sheets("Sheet2").[A2])

    For i = 2 to Sheets("Sheet2").[a99999].End(xlUp).Row


End Sub






    
