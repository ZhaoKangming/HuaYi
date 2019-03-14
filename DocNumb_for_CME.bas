Sub CountNumbs()
    Application.ScreenUpdating = False
    Call DateClean
    Call GetList
    Call DocTitle
    Call GetNumbs
    Call Beautify
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "Finished！"
End Sub

'todo 检查是否有非汉字
'替换空格

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

Sub GetFrequency(ByVal ValueRange As Range,ByVal StartCell As Range)
    Dim rng As Range, arr, d As Object, i%, SCRow%
    Set d = CreateObject("scripting.dictionary")
    For Each rng In ValueRange
        If rng <> "" And Not d.exists(rng.Value) Then d(rng.Value)= rng.Value
    Next
    arr = d.items
    SCRow = StartCell.Row
    If d.Count > 1 Then Cells(SCRow + 1,StartCell.Column).Resize(d.Count-1, 1).EntireRow.Insert shift:=xlDown
    For i = 0 To d.Count - 1
        Cells(i + SCRow, StartCell.Column) = arr(i)
    Next
    Set d = Nothing
End Sub

Sub GetList()
    Dim RowsNumb%, i%, CitiesNumb%, ProvienceRows%, ProvienceRng As Range, FirstColRng As Range, ProFirstRow%
    Dim GetF_VR As Range, GetF_SC As Range, cellrng As Range, dict As Object
    Sheets("Sheet2").[A1] = "省份"
    Sheets("Sheet2").[B1] = "城市"
    Sheets("Sheet2").[C1] = "医生"
    Sheets("Sheet2").[D1] = "护士"
    Sheets("Sheet2").[E1] = "技师"
    Sheets("Sheet2").[F1] = "药师"
    RowsNumb = Sheets("Sheet1").[a99999].End(xlUp).Row
    Set ProvienceRng = Range([A2], Cells(RowsNumb, 1))
    Set FirstColRng = Range([A1], Cells(RowsNumb, 1))

    Set dict = CreateObject("scripting.dictionary")
    For Each cellrng In Range([B2], Cells(RowsNumb, 2))
        If cellrng <> "" And Not dict.exists(cellrng.Value) Then dict(cellrng.Value)= cellrng.Value
    Next
    CitiesNumb = dict.Count
    Set dict = Nothing
    Sheets("Sheet2").Activate
    Call GetFrequency(ProvienceRng,Sheets("Sheet2").[A2])

    For i = 2 to CitiesNumb + 1
        If Cells(i,1) <> "" Then
            ProvienceRows = Application.WorksheetFunction.Countif(ProvienceRng,Cells(i,1)) 
            ProFirstRow = FirstColRng.Find(Cells(i,1)).Row
            Sheets("Sheet1").Activate
            Set GetF_VR = Range(Cells(ProFirstRow,2), Cells(ProFirstRow + ProvienceRows - 1, 2))
            Set GetF_SC = Cells(i,2)
            Sheets("Sheet2").Activate
            Call GetFrequency(GetF_VR,GetF_SC)
        End If
    Next
End Sub

Sub DocTitle()
    Dim YiShengArr, HuShiArr, YaoShiArr, JiShiArr, RowsNumb%, i%, j%, ConfirmDocTitle As Boolean
    Dim ZRDocArr, FZRDocArr, ZZDocArr, YSDocArr, QTDocArr
    YiShengArr = Array("编辑","编审","副编审","副教授","副研究馆员","副研究员","副主任检验师", _
                        "副主任医师","高级工程师","高级会计师","高级经济师","高级统计师","工程师", _
                        "馆员","会计师","会计员","检验师","检验士","见习检验师","见习检验士", _
                        "见习医师","见习医士","讲师","教授","经济师","经济员","其他","实习研究员", _
                        "统计师","统计员","无职称","乡村医生","小学初级教师","小学高级教师", _
                        "小学中级教师","研究馆员","研究员","医师","医士","中西医结合副主任医师", _
                        "中西医结合医师","中西医结合主任医师","中西医结合主治医师","中学初级教师", _
                        "中学高级教师","中学特级教师","中学中级教师","中医保健按摩","中医保健按摩及中医美容", _
                        "中医副主任医师","中医美容","中医医师","中医医士","中医主任医师","中医主治医师", _
                        "主管检验师","主管医师","主任检验师","主任医师","主治医师","助教","助理编辑", _
                        "助理工程师","助理馆员","助理会计师","助理经济师","助理统计师","助理研究员")
    HuShiArr = Array("副主任护师","护师","护士","见习护师","见习护士","主管护师","主任护师")
    YaoShiArr = Array("副主任药师","见习药剂师","见习药剂士","药剂师","药剂士","中医副主任药师", _
                    "中医药师","中医药士","中医主管药师","中医主任药师","主管药师","主任药师")
    JiShiArr = Array("副主任技师","技师","技士","技术员","见习技师","见习技士","中医副主任技师", _
                    "中医技师","中医技士","中医主管技师","中医主任技师","主管技师","主任技师")
    ZRDocArr = Array()
    FZRDocArr = Array()
    ZZDocArr = Array()
    YSDocArr = Array()
    QTDocArr = Array()

    Sheets("Sheet1").Activate
    RowsNumb = Sheets("Sheet1").[a99999].End(xlUp).Row
    For i = 2 to RowsNumb
        ConfirmDocTitle = False
        For j = 0 To UBound(YiShengArr)
                If Cells(i,6) = HuShiArr(j) Then 
                    Cells(i,10) = Cells(i,2) & "医生"
                    ConfirmDocTitle = True
                    Exit For
                End If
            Next j

        If ConfirmDocTitle = False Then
            For j = 0 To UBound(HuShiArr)
                If Cells(i,6) = HuShiArr(j) Then 
                    Cells(i,10) = Cells(i,2) & "护士"
                    ConfirmDocTitle = True
                    Exit For
                End If
            Next j
        End If

        If ConfirmDocTitle = False Then
            For j = 0 To UBound(YaoShiArr)
                If Cells(i,6) = YaoShiArr(j) Then 
                    Cells(i,10) = Cells(i,2) & "药师"
                    ConfirmDocTitle = True
                    Exit For
                End If
            Next j
        End If

        If ConfirmDocTitle = False Then
            For j = 0 To UBound(JiShiArr)
                If Cells(i,6) = JiShiArr(j) Then 
                    Cells(i,10) = Cells(i,2) & "技师"
                    ConfirmDocTitle = True
                    Exit For
                End If
            Next j
        End If

        If ConfirmDocTitle = False Then Cells(i,10) = Cells(i,2) & "医生"
    Next

    For i = 2 to RowsNumb
        If Right(Cells(i,10),2) = "医生" Then
            If Cells(i,6) = "中西医结合主任医师" Or Cells(i,6) = "中医主任医师" Or Cells(i,6) = "主任医师"  Then Cells(i,11)="主任医师"
            


        End If
    Next

End Sub

Sub GetNumbs()
    Dim i%, j%, DocTitleRng As Range, HospTitleRng As Range, DocClassRng As Range
    Sheets("Sheet1").Activate
    j = Sheets("Sheet1").[a99999].End(xlUp).Row
    Set DocTitleRng = Range([J2],Cells(j,10))
    Set HospTitleRng = Range([G2],Cells(j,7))
    Set DocClassRng = Range([K2],Cells(j,11))
    Sheets("Sheet2").Activate
    For i = 2 To [b99999].End(xlUp).Row
        Cells(i,3) = Application.WorksheetFunction.Countif(DocTitleRng,Cells(i,2)&"医生")
        Cells(i,4) = Application.WorksheetFunction.Countif(DocTitleRng,Cells(i,2)&"护士")
        Cells(i,5) = Application.WorksheetFunction.Countif(DocTitleRng,Cells(i,2)&"技师")
        Cells(i,6) = Application.WorksheetFunction.Countif(DocTitleRng,Cells(i,2)&"药师")
    Next

    Sheets("Sheet3").Activate
    Cells(1,1) = "医生职称"
    Cells(2,1) = "主任医师"
    Cells(3,1) = "副主任医师"
    Cells(4,1) = "主治医师"
    Cells(5,1) = "医师"
    Cells(6,1) = "其他"
    Cells(7,1) = "总计"

    Cells(1,4) = "医院级别"
    Cells(2,4) = "三甲"
    Cells(3,4) = "三乙"
    Cells(4,4) = "二甲"
    Cells(5,4) = "二乙"
    Cells(6,4) = "一甲"
    Cells(7,4) = "一乙"
    Cells(8,4) = "其他"
    Cells(9,4) = "总计"

    For i = 2 To 6
        Cells(i,1) = Application.WorksheetFunction.Countif(DocClassRng,Cells(i,1))
    Next
    Cells(7,1) = Application.WorksheetFunction.Sum([A2:A6])
    

    For i = 2 To 8
        Cells(i,4) = Application.WorksheetFunction.Countif(HospTitleRng,Cells(i,4))
    Next
    Cells(9,4) = Application.WorksheetFunction.Sum([D2:D8])
    'todo 如果和总数不一致报错
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

    With Sheets("Sheet3").UsedRange
        .Font.Name = "微软雅黑"
        .Font.Size = 12    
    End With
End Sub


