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

    '复制工作表
    Sheets("Sheet1").Copy Before:=Sheets("Sheet1")
    Sheets("Sheet1 (2)").Name = "TEMP"

    '只保留医生数据
    Sheets("TEMP").Select
    Columns("N:P").Delete
    RowNumbs = Sheets("TEMP").[a99999].End(xlUp).Row
    For i = RowNumbs to 2 Step-1
        If Cells(i,)
    Next

    Set Temp_Dict = CreateObject("scripting.dictionary")
    
        If Cells(i,1)
    Next
		

End Sub



