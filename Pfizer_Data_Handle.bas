Sub Pfizer_Data_Handle()
    Application.ScreenUpdating = False
    
'激活数据源表
    Dim i%, Src_Wkb, Dst_Wkb, RowNumbs%
    Dim Temp_Dict As object
    Dim CellRng As Range, Temp_Rng As Range
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name like "*学习记录*" Then Workbooks(i).Activate       
    Next
    If Not ActiveWorkbook.Name like "*学习记录*" Then 
        Msgbox "Cannot find the workbook!"
        Exit Sub
    End If

    Src_Wkb = Workbooks(ActiveWorkbook.Name)
    RowNumbs = Sheets("Sheet1").[a99999].End(xlUp).Row
    Set Temp_Dict = CreateObject("scripting.dictionary")
    For i = 2 to RowNumbs
        If Cells(i,1)
    Next
		

End Sub



	