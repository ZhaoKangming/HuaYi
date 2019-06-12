Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '暂停功能
Public Declare PtrSafe Function MsgBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long 'AutoClose


Sub Download_Report()
    Dim lastRow_Temp%, i%
    Application.ScreenUpdating = False
    Sheets("Tools").Select
    lastRow_Temp = Range("a1048576").End(xlUp).Row
    Columns("F:H").ClearContents
    Columns(6).ColumnWidth = 27
    Columns(7).ColumnWidth = 40

    '-------------------- 生成文件名 --------------------
    For i = 2 To lastRow_Temp
        If Cells(i, 3) = "报告2" Then Cells(i, 7) = "_R2"
        If Cells(i, 3) = "病例2" Then Cells(i, 7) = "_C2"
        Cells(i, 3) = Left(Cells(i, 3), 2)
    Next
    Columns("D:D").NumberFormatLocal = "yymmdd"

    For i = 2 To lastRow_Temp
        Cells(i, 6) = Cells(i, 1) & "_" & Cells(i, 2) & Cells(i, 7) & "_" & Application.WorksheetFunction.Text(Cells(i, 4), "yymmdd")
    Next

    Columns(7).ClearContents

    '-------------------- 提取超链接 --------------------
    Dim ArrayHL(0 To 200, 0 To 1) As String '200这个数是随意写的，只要大于最大行数就行
    For i = 2 To lastRow_Temp
        ArrayHL(i - 2, 0) = Cells(i, 5).Hyperlinks(1).Address
    Next
    Range("G2").Resize(lastRow_Temp, 1).Value = ArrayHL

    ' -------------------- 下载报告病例文件 --------------------
    Dim Rep_Folder$, Case_Folder$, extensionName$, FileName$, FileLink$, myFolder$
    Rep_Folder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\报告\"
    Case_Folder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\病例\"
    MkDir Rep_Folder
    MkDir Case_Folder

    For i = 2 To lastRow_Temp
        If Right(Cells(i, 7), 6) Like "*.*" Then
            extensionName = Split(Right(Cells(i, 7), 6), ".")(1)
        Else
            MsgBox Cells(i, 1) & " 的文件无扩展名！"
            Exit Sub
        End If
        If Not extensionName Like "*doc*" Then
            MsgBox Cells(i, 1) & " 的文件格式不是Word文件！"
            Exit Sub
        End If
        FileName = Cells(i, 6) & "." & extensionName
        FileLink = Cells(i, 7)
        myFolder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\" & Cells(i, 3) & "\"
        If URLDownloadToFile(0&, FileLink, myFolder & FileName, 0&, 0&) = 0 Then
        Else
            MsgBox "Failure"
        End If
    Next
    Sleep 5

    '-------------------- 核对是否全部下载了 --------------------
    Dim RFNumb%, CFNumb%, MyFile
    RFNumb = 0
    MyFile = Dir(Rep_Folder, 16)
    Do While MyFile <> ""
        If MyFile <> "." And MyFile <> ".." Then RFNumb = RFNumb + 1
        MyFile = Dir
    Loop

    CFNumb = 0
    MyFile = Dir(Case_Folder, 16)
    Do While MyFile <> ""
        If MyFile <> "." And MyFile <> ".." Then CFNumb = CFNumb + 1
        MyFile = Dir
    Loop

    If RFNumb + CFNumb <> lastRow_Temp - 1 Then
        MsgBox "存在未下载文件 " & lastRow_Temp - 1 - RFNumb - CFNumb & " 个！"
        Exit Sub
    End If

    '-------------------- 将文件复制到原始文件夹中 --------------------
    Dim SrcRep_Folder$, SrcCase_Folder$
    SrcRep_Folder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\文档\华医网\赋能起航\报告病例审核\原始报告\"
    SrcCase_Folder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\文档\华医网\赋能起航\报告病例审核\原始病例\"

    MyFile = Dir(Rep_Folder, 16)
    Do While MyFile <> ""
        If MyFile <> "." And MyFile <> ".." Then
            FileCopy Rep_Folder & MyFile, SrcRep_Folder & MyFile
        End If
        MyFile = Dir
    Loop
    Sleep 5
    
    MyFile = Dir(Case_Folder, 16)
    Do While MyFile <> ""
        If MyFile <> "." And MyFile <> ".." Then
            FileCopy Case_Folder & MyFile, SrcCase_Folder & MyFile
        End If
        MyFile = Dir
    Loop
    Sleep 5
    
    ActiveWorkbook.Save
    [A2].Select
    Application.ScreenUpdating = True
    MsgBox "Downloaded!"
End Sub
Sub Clear_Sheet()
    Rows("2:200").Clear
End Sub

Sub Copy_Info()
    Dim ThisRow%, dst_sht As Worksheet, dst_lastrow%, i%
    Sheets("Tools").Activate
    ThisRow = Selection.Row
    Set dst_sht = Worksheets(Left(Cells(ThisRow, 3), 2) & "-总结库")
    dst_lastrow = dst_sht.[E999999].End(xlUp).Row
    
    For i = 1 To 3
        dst_sht.Cells(dst_lastrow, i) = Sheets("Tools").Cells(ThisRow, i).Value
    Next
    dst_sht.Cells(dst_lastrow, 4) = Application.WorksheetFunction.Text(Sheets("Tools").Cells(ThisRow, 4), "yymmdd")
    
    Sheets("Tools").Rows(Selection.Row).Delete
    dst_sht.Activate
    dst_sht.Cells(dst_lastrow + 1, 5).Select
    Set dst_sht = Nothing
    Call AutoClose
End Sub

Sub Copy_Wrong_Info()
    Dim ThisRow%, dst_sht As Worksheet, dst_lastrow%, i%
    Sheets("Tools").Activate
    ThisRow = Selection.Row
    Set dst_sht = Worksheets(Left(Cells(ThisRow, 3), 2) & "-修改记录")
    dst_lastrow = dst_sht.[C999999].End(xlUp).Row
    
    For i = 1 To 3
        dst_sht.Cells(dst_lastrow + 1, i + 2) = Sheets("Tools").Cells(ThisRow, i).Value
    Next
    dst_sht.Cells(dst_lastrow + 1, 6) = Application.WorksheetFunction.Text(Sheets("Tools").Cells(ThisRow, 4), "yymmdd")
    Sheets("Tools").Rows(Selection.Row).Delete
    dst_sht.Activate
    dst_sht.Cells(dst_lastrow + 1, 7).Select
    Set dst_sht = Nothing
    Call AutoClose
End Sub

Public Sub AutoClose()
    '过程,"弹出对话","对话框标题",图标类型,默认参数,N毫秒后自动关闭
    MsgBoxTimeOut 0, "录入完毕!!", "提示", 64, 0, 300
End Sub
Sub Delete_Row2()
    Rows(2).Delete
End Sub


Sub Get_Files()
    Dim Folder_Arr, fso As Object, Src_Dir$, Dst_Dir$, MyFile, New_Dir$, New_File
    
    Folder_Arr = Array("合格报告", "原始报告", "合格病例", "原始病例")
    Dst_Dir = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\赋能起航-报告病例审核-" & Format(Now, "yymmdd") & "\"
    MkDir Dst_Dir
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 0 To 3
        Src_Dir = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\文档\华医网\赋能起航\报告病例审核\" & Folder_Arr(i)
        fso.CopyFolder Src_Dir, Dst_Dir
        Sleep 2
        
        New_Dir = Dst_Dir & Folder_Arr(i) & "\"
        MyFile = Dir(New_Dir, 16)
        Do While MyFile <> ""
            If MyFile <> "." And MyFile <> ".." Then
                If MyFile Like "*_Y*" Then New_File = Right(MyFile, Len(MyFile) - InStr(MyFile, "Y") + 1)
                If MyFile Like "*_A*" Then New_File = Right(MyFile, Len(MyFile) - InStr(MyFile, "A") + 1)
                Name New_Dir & MyFile As New_Dir & New_File
            End If
            MyFile = Dir
        Loop
    Next

    Set fso = Nothing
    'TODO:文件打包
    MsgBox "Finished the work!"
    
End Sub


Sub Only_Chinese()
    
    Dim oRegExp As Object, oMatches As Object, sText As String
    sText = ActiveCell.Value
    Set oRegExp = CreateObject("vbscript.regexp")
    With oRegExp
        .Global = True
        .IgnoreCase = True        
        '------------保留中文------------------
        .Pattern = "[^\u4e00-\u9fa5]+"
        Set oMatches = .Execute(sText)
        ActiveCell.Value = .Replace(sText, "")
    End With

    Set oRegExp = Nothing
    Set oMatches = Nothing
    Sheets("Tools").Activate
    
End Sub



