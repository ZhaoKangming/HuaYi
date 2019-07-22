' 此过程用于从医生资料库中取出指定医生的资料
'TODO:提高筛选取出效率
Sub CopyDocFolders()
    Dim i%, OKNumb%, FolderPath$, MyFile, myFolder$, myNewFilePath$
    OKNumb = 0
    Set fso = CreateObject("Scripting.FileSystemObject")  
    ' FolderPath = InputBox("请输入新建文件夹的路径，不以'\'结尾", "输入地址") & "\"
    FolderPath = "H:\【汇总】2019赋能起航-医护支持文件扫描件\"
    myNewFilePath = "H:\DocInfo_" & Format(Now,"yymmdd") & "\"    '要移动的位置
    Mkdir myNewFilePath
    UsedRange.Replace " ",""
    For i = 1 To [A1048576].End(xlUp).Row
        MyFile = Dir(FolderPath, 16)
        Do While MyFile <> ""
            If MyFile like "*" & Cells(i,1) & "*" Then 
                Cells(i,3) = "OK"
                OKNumb = OKNumb + 1
                myFolder = FolderPath & MyFile  '要移动的文件夹
                If Not fso.FolderExists(myNewFilePath & MyFile) Then fso.CopyFolder myFolder, myNewFilePath
                Exit Do
            End If
            MyFile = Dir
        Loop
    Next
    Set fso = nothing
    ThisWorkbook.Save
    If OKNumb = [A1048576].End(xlUp).Row Then Msgbox "已经全部处理完成！"
    If OKNumb < [A1048576].End(xlUp).Row Then Msgbox "处理完成！" & vblf & vblf & "缺少 " & [A1048576].End(xlUp).Row - OKNumb & " 个文件夹！"
End Sub