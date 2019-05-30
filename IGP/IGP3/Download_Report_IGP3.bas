Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) '暂停功能

Sub Get_Report()
    Dim lastRow_Temp%, i%
    Application.ScreenUpdating = False
    Sheets("Tools").Select
    lastRow_Temp = Range("a1048576").End(xlUp).Row
    Columns("F:H").Delete
    Columns(6).ColumnWidth = 27
    Columns(7).ColumnWidth = 40

    '-------------------- 生成文件名 --------------------
    For i = 2 To lastRow_Temp
        If Cells(i,3) = "报告2" Then Cells(i,7) = "_R2"
        If Cells(i,3) = "病例2" Then Cells(i,7) = "_C2"
        Cells(i,3) = Left(Cells(i,3),2)
    Next
    Columns("D:D").NumberFormatLocal = "yymmdd"

    For i = 2 To lastRow_Temp
        Cells(i,6) = Cells(i,1) & "_" & Cells(i,2) & Cells(i,7) & "_" & Application.WorksheetFunction.Text(Cells(i,4),"yymmdd")
    Next

    Columns(7).ClearContents

    '-------------------- 提取超链接 --------------------
    Dim ArrayHL(0 To 200, 0 To 1) As String '200这个数是随意写的，只要大于最大行数就行
    For  i = 2 To lastRow_Temp
        ArrayHL( i - 2, 0) = Cells(i,5).Hyperlinks(1).Address
    Next
    Range("G2").Resize(lastRow_Temp, 1).Value = ArrayHL

    ' -------------------- 下载报告病例文件 --------------------
    Dim myFolder$, Rep_Folder$, Case_Folder$, extensionName$, FileName$, FileLink$
    Rep_Folder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\报告\"
    Case_Folder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\病例\"
    MkDir Rep_Folder
    MkDir Case_Folder

    For  i = 2 To lastRow_Temp
        If Right(Cells(i, 7), 6) Like "*.*" Then
            extensionName = Split(Right(Cells(i, 7), 6), ".")(1)
        Else
            MsgBox Cells(i, 1) & " 的文件无扩展名！"
            Exit Sub
        End If
        If Not extensionName like "*doc*" Then
            Msgbox Cells(i,1) & " 的文件格式不是Word文件！"
            Exit Sub
        End If
        FileName = Cells(i,6) & "." & extensionName
        FileLink = Cells(i,7)
        myFolder = "C:\Users\ZhaoKangming\OneDrive - cnu.edu.cn\桌面\" & Cells(i,3) & "\"
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
        Msgbox "存在未下载文件 " & lastRow_Temp - 1 - RFNumb - CFNumb & " 个！"
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

