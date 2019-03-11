Sub CountNumbs()
    Application.ScreenUpdating = False
    Call DateClean
    Call GetList
    Call GetNumbs
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

    With Sheets("Sheet2").UsedRange
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
    Set d = Nothing
End Sub

Sub GetList()
    Dim RowsNumb%, i%, CitiesNumb%, ProvienceRows%, ProvienceRng As Range, ProFirstRow%
    Dim GetF_VR As Range, GetF_SC As Range, cellrng As Range, dict As Object
    Sheets("Sheet2").[A1] = "省份"
    Sheets("Sheet2").[B1] = "城市"
    Sheets("Sheet2").[C1] = "医生"
    Sheets("Sheet2").[D1] = "护士"
    Sheets("Sheet2").[E1] = "技师"
    Sheets("Sheet2").[F1] = "药师"
    RowsNumb = Sheets("Sheet1").[a99999].End(xlUp).Row
    ProvienceRng = Sheets("Sheet1").Range([A2], Cells(RowsNumb, 1))

    Set dict = CreateObject("scripting.dictionary")
    For Each cellrng In Sheets("Sheet1").Range([B2], Cells(RowsNumb, 2))
        If cellrng <> "" And Not dict.exists(cellrng.Value) Then dict(cellrng.Value)= cellrng.Value
    Next
    CitiesNumb = dict.Count
    Set dict = Nothing
    Call GetFrequency(ProvienceRng,Sheets("Sheet2").[A2])

    For i = 2 to CitiesNumb + 1
        If Sheets("Sheet2").Cells(i,1) <> "" Then
            ProvienceRows = Application.WorksheetFunction.Countif(ProvienceRng,Sheets("Sheet2").Cells(i,1)) 
            ProFirstRow = ProvienceRng.Find(Sheets("Sheet2").Cells(i,1)).Row
            GetF_VR = Sheets("Sheet1").Range(Cells(ProFirstRow,2), Cells(ProFirstRow + ProvienceRows - 1, 2))
            GetF_SC = Sheets("Sheet2").Cells(i,2)
            Call GetFrequency(GetF_VR,GetF_SC)
        End If
    Next
End Sub

Sub GetNumbs()
    Dim Arr, i%
    Arr = Array("中医主管护师","中医护师","中医护士","中医副主任护师","副主任护师","主任护师","主管护师","护师", _
                "护士","见习护士","见习护师","护师（新晋）","中医主任护师")
    Arr = Array("药具士","主管药具师","副主任药师","主任药师","药剂士","药剂师","主管药师","中医副主任药师","中医 _
                主任药师","中医主管药师","中医药士","中医药师","见习药剂士","见习药剂师","主管药师（新晋）","药具师")
    Arr = Array("技师","主管技师","主任技师","副主任技师","见习技师","中医技师","中医主任技师","中医主管技师", _
                "中医副主任技师","主管技师（新晋）")
    For i = 0 To UBound(Arr)
        Debug.Print Arr(i)

　　Next

End Sub




