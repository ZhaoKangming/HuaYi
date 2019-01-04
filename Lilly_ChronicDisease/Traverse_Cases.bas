Sub Travere_Folder()
Dim folderPath$, MyName$, FileNames$, File_Count%, Mypath$, i%, folderPath_Now$
Application.ScreenUpdating = False
folderPath = InputBox("输入病例路径") & "\"
MkDir folderPath & "Checked\"
MyName = Dir(folderPath, vbNormal)
FileNames = Dir(folderPath, vbNormal)
Dim File_Arr() As String
File_Count = 0
Do While FileNames <> ""
    DoEvents
    If FileNames <> "." And FileNames <> ".." Then
        If (GetAttr(folderPath & FileNames) And vbNormal) = vbNormal Then
            File_Count = File_Count + 1
            ReDim Preserve File_Arr(File_Count)
            File_Arr(File_Count) = FileNames
        End If
    End If
    FileNames = Dir
Loop
'获取当前目录下的所有文件名存入File_Arr数组中。路径在folderPath变量中。
For i = 1 To UBound(File_Arr)
    folderPath_Now = folderPath & File_Arr(i)
    Workbooks.Open (folderPath_Now), UpdateLinks:=0  '打开工作簿后不更新链接
    '执行完成后保存、关闭直至全部文件执行完成。
    Call Case_Check
Next i
Application.ScreenUpdating = True
MsgBox "全部执行完成！！"

End Sub
